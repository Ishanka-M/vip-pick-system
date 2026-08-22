"""
invoice_register.py — every uploaded document, kept
====================================================
Two reports, downloaded separately:

  Summary   one row per Tax Invoice / Delivery Challan
  Details   one row per document line

`Korber Pick` starts as **No** with the reason as the remark. When the same
invoice is picked later the row is updated in place — Yes, remark cleared. It is
never a second row, so the register stays one row per invoice for good.
"""
from __future__ import annotations

import io
import json
import re
from datetime import datetime, timedelta
from typing import Any

import pandas as pd

from doc_parser import ParsedDoc, base_item, clean_item

# Bumped whenever this module's public surface changes; app.py refuses to run
# against a stale copy instead of dying with a redacted TypeError.
API = 17

# Warehouse execution status — independent of KORBER_PICK (this app's own pallet
# allocation). These three track the physical pick / pack / dispatch, driven by
# the Pick_Live_status report, the packing QR scan, and the sales report match.
STATUS_COLS = ["PICKING", "PACKING", "DISPATCH"]
STATUS_PENDING = "Pending"
STATUS_DONE = "Completed"

# KORBER_PICK — this app's own pallet allocation for the document.
#   No      nothing was picked
#   Partial some of it went out; the rest is still owed and the document will
#           come back through the pick for its balance
#   Yes     the whole document quantity has been picked
PICK_NO, PICK_PART, PICK_YES = "No", "Partial", "Yes"
# higher wins on merge — a re-upload must never walk a document backwards
PICK_RANK = {PICK_NO: 0, PICK_PART: 1, PICK_YES: 2}


def pick_rank(v: Any) -> int:
    return PICK_RANK.get(str(v or "").strip().title(), 0)

SUMMARY_COLS = [
    "TAX_INVOICE_DATE", "TAX_INVOICE_NO", "AR_INVOICE_NO", "CUSTOMER_NAME", "QTY",
    "KORBER_PICK", "REMARK", "MRP",
    "DOC_TYPE", "CUSTOMER_CODE", "CONTACT_PERSON", "CONTACT_EMAIL", "LINES",
    "PICKED_QTY", "PLANT", "RUN_ID", "PICKED_AT", "FIRST_SEEN", "UPDATED_AT",
    "UPDATED_BY", "SOURCE_FILE", "TOTAL_INCL_TAX",
] + STATUS_COLS

DETAIL_COLS = [
    "TAX_INVOICE_DATE", "TAX_INVOICE_NO", "AR_INVOICE_NO", "CUSTOMER_NAME", "DOC_TYPE",
    "LINE", "ITEM_CODE", "BASE_ID", "DESCRIPTION", "DOC_QTY", "UOM",
    "KORBER_PICK", "PICKED_QTY", "ITEM_NUMBER", "LOT_NUMBER", "PALLETS", "LOCATIONS",
    "PLANT", "REMARK", "RUN_ID", "UPDATED_AT", "UNIT_PRICE", "LINE_TOTAL",
] + STATUS_COLS

DEFAULT_MRP = [{"NAME": "Sharma, Rahul", "EMAIL": "rahul.sharma1@donaldson.com"}]

# A duplicate is not a pending job — the invoice was already picked in an earlier
# run and is already sitting in the register as Yes. Registering it again would
# invent a second, permanently-pending copy of work that is finished.
DUPLICATE_PAT = r"duplicate"
# ...but "DUPLICATE (batch)" only means the same file was dropped in twice in one
# upload. The first copy is the real document and still needs its own row, so
# only the two "it already exists elsewhere" cases skip registration.
DUP_SKIP_PAT = r"duplicate\s*\(\s*(already processed|other user)"


# --------------------------------------------------------------------------- #
# MRP contacts
# --------------------------------------------------------------------------- #
def _qty_str(q: float) -> str:
    return str(int(round(q))) if abs(q - round(q)) < 1e-9 else f"{q:g}"


def _norm_name(v: Any) -> str:
    """'Sharma, Rahul' == 'rahul sharma' — order and punctuation don't matter."""
    parts = re.split(r"[,\s]+", str(v or "").strip().lower())
    return " ".join(sorted(p for p in parts if p))


def _norm_mail(v: Any) -> str:
    return str(v or "").strip().lower()


def load_contacts(raw: Any) -> list[dict]:
    """APP_SETTINGS holds this as JSON; fall back to the built-in contact."""
    if isinstance(raw, list):
        rows = raw
    else:
        try:
            rows = json.loads(raw) if str(raw or "").strip() else []
        except Exception:
            rows = []
    out = [{"NAME": str(r.get("NAME", "")).strip(),
            "EMAIL": str(r.get("EMAIL", "")).strip()}
           for r in rows if isinstance(r, dict)
           and (str(r.get("NAME", "")).strip() or str(r.get("EMAIL", "")).strip())]
    return out or [dict(c) for c in DEFAULT_MRP]


def dump_contacts(df: Any) -> str:
    rows = df.to_dict("records") if isinstance(df, pd.DataFrame) else list(df or [])
    keep = [{"NAME": str(r.get("NAME", "")).strip(), "EMAIL": str(r.get("EMAIL", "")).strip()}
            for r in rows
            if str(r.get("NAME", "")).strip() or str(r.get("EMAIL", "")).strip()]
    return json.dumps(keep, ensure_ascii=False)


def mrp_flag(doc: ParsedDoc, contacts: list[dict]) -> str:
    """
    Yes when the document's contact is one of the MRP contacts — matched on the
    email when the document carries one, otherwise on the name.
    A challan has no contact block at all, so it stays No.
    """
    name, mail = _norm_name(doc.contact_person), _norm_mail(doc.contact_email)
    if not name and not mail:
        return "No"
    for c in contacts:
        cn, cm = _norm_name(c.get("NAME")), _norm_mail(c.get("EMAIL"))
        if cm and mail and cm == mail:
            return "Yes"
        if cn and name and cn == name and not (cm and mail):
            return "Yes"
        if cn and name and cn == name and cm and mail and cm == mail:
            return "Yes"
    return "No"


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #
def build(docs: list[ParsedDoc], res: dict[str, Any], contacts: list[dict],
          user: str = "", plant: str = "") -> tuple[pd.DataFrame, pd.DataFrame]:
    """Summary + details for every document in this run, picked or not."""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_id = str(res.get("run_id", ""))
    picked = {str(x) for x in (res.get("accepted", pd.DataFrame())
                               .get("DOC_NUMBER", pd.Series(dtype=str)))}
    reasons: dict[str, str] = {}
    rej = res.get("rejected")
    if rej is not None and len(rej):
        for _, r in rej.iterrows():
            num = str(r.get("DOC_NUMBER", ""))
            why = str(r.get("REASON", "")).strip()
            det = str(r.get("DETAIL", "")).strip()
            # a real failure must not be overwritten by the batch-duplicate note
            if num in reasons and re.search(DUPLICATE_PAT, why, re.I):
                continue
            reasons[num] = (f"{why} — {det}" if det else why)[:400]

    alloc = res.get("allocations")
    alloc = alloc if alloc is not None and len(alloc) else pd.DataFrame()

    # partial pick: what each document / line had already received before this
    # run, and which documents went out short
    prev_map = {str(k): {int(l): float(q) for l, q in (v or {}).items()}
                for k, v in (res.get("picked_before") or {}).items()}
    partial_now = {str(x) for x in (res.get("partial") or [])}
    acc = res.get("accepted")
    short_by_doc: dict[str, float] = {}
    if acc is not None and len(acc) and "SHORT_QTY" in acc.columns:
        short_by_doc = {str(n): float(q or 0.0)
                        for n, q in zip(acc["DOC_NUMBER"], acc["SHORT_QTY"])}

    srows: list[dict] = []
    drows: list[dict] = []
    seen: set[str] = set()
    skipped: list[str] = []

    for doc in docs:
        num = str(doc.doc_number).strip()
        if not num or num in seen:
            continue
        seen.add(num)
        if num not in picked and re.search(DUP_SKIP_PAT, reasons.get(num, ""), re.I):
            skipped.append(num)          # already in the register from its real run
            continue
        is_inv = doc.doc_type.upper().startswith("INVOICE")
        ok = num in picked
        doc_qty = float(sum(l.qty for l in doc.lines))
        prev = prev_map.get(num, {})

        a = alloc[alloc["DOC_NUMBER"].astype(str) == num] if len(alloc) else pd.DataFrame()
        # what the document has had altogether — this run plus any earlier
        # partial pick, or the earlier partial alone if this run took nothing
        picked_qty = float(pd.to_numeric(a["QTY_PICKED"], errors="coerce").sum()) \
            if len(a) else 0.0
        picked_qty += float(sum(prev.values()))
        plants = ", ".join(sorted({str(x) for x in a["PLANT"]})) if len(a) else plant

        if ok and num in partial_now:
            state = PICK_PART
            short = short_by_doc.get(num, max(0.0, doc_qty - picked_qty))
            remark = f"PARTIAL PICK — {_qty_str(short)} of {_qty_str(doc_qty)} still owed"
        elif ok:
            state = PICK_YES
            remark = ""
        elif picked_qty > 0:
            # not picked in this run, but an earlier run took part of it
            state = PICK_PART
            remark = (f"PARTIAL PICK — {_qty_str(max(0.0, doc_qty - picked_qty))} of "
                      f"{_qty_str(doc_qty)} still owed · "
                      f"{reasons.get(num, 'not picked this run')}")[:400]
        else:
            state = PICK_NO
            remark = reasons.get(num, "Not picked")

        srows.append({
            "TAX_INVOICE_DATE": doc.doc_date, "TAX_INVOICE_NO": num,
            "AR_INVOICE_NO": doc.ref_number if is_inv else "",
            "CUSTOMER_NAME": doc.customer, "QTY": doc_qty,
            "KORBER_PICK": state, "REMARK": remark,
            "MRP": mrp_flag(doc, contacts),
            "DOC_TYPE": doc.doc_type, "CUSTOMER_CODE": doc.customer_code,
            "CONTACT_PERSON": doc.contact_person, "CONTACT_EMAIL": doc.contact_email,
            "LINES": len(doc.lines), "PICKED_QTY": picked_qty, "PLANT": plants,
            "RUN_ID": run_id if ok else "", "PICKED_AT": now if ok else "",
            "FIRST_SEEN": now, "UPDATED_AT": now, "UPDATED_BY": user,
            "SOURCE_FILE": doc.source_file,
            "TOTAL_INCL_TAX": doc.total_incl_tax if doc.total_incl_tax is not None else "",
            "PICKING": STATUS_PENDING, "PACKING": STATUS_PENDING,
            "DISPATCH": STATUS_PENDING,
        })

        for ln in doc.lines:
            la = a[pd.to_numeric(a["DOC_LINE"], errors="coerce") == ln.line_no] \
                if len(a) else pd.DataFrame()
            ln_picked = (float(pd.to_numeric(la["QTY_PICKED"], errors="coerce").sum())
                         if len(la) else 0.0) + float(prev.get(ln.line_no, 0.0))
            # a line of a partial document is answered on its own numbers:
            # a short line is Partial (or No), a line that went in full is Yes
            if ln_picked <= 0:
                ln_state = PICK_NO
            elif ln_picked + 1e-9 >= float(ln.qty):
                ln_state = PICK_YES
            else:
                ln_state = PICK_PART
            drows.append({
                "TAX_INVOICE_DATE": doc.doc_date, "TAX_INVOICE_NO": num,
                "AR_INVOICE_NO": doc.ref_number if is_inv else "",
                "CUSTOMER_NAME": doc.customer, "DOC_TYPE": doc.doc_type,
                "LINE": ln.line_no, "ITEM_CODE": ln.item_code, "BASE_ID": ln.base,
                "DESCRIPTION": ln.description, "DOC_QTY": float(ln.qty),
                "UOM": ln.uom, "KORBER_PICK": ln_state,
                "PICKED_QTY": ln_picked,
                "ITEM_NUMBER": ", ".join(dict.fromkeys(str(x) for x in la["ITEM_NUMBER"]))
                if len(la) else "",
                "LOT_NUMBER": ", ".join(dict.fromkeys(str(x) for x in la["LOT_NUMBER"]))
                if len(la) else "",
                "PALLETS": ", ".join(dict.fromkeys(str(x) for x in la["PALLET"]))
                if len(la) else "",
                "LOCATIONS": ", ".join(dict.fromkeys(str(x) for x in la["LOCATION_ID"]))
                if len(la) else "",
                "PLANT": plants, "REMARK": remark, "RUN_ID": run_id if ok else "",
                "UPDATED_AT": now,
                "UNIT_PRICE": ln.unit_price if ln.unit_price is not None else "",
                "LINE_TOTAL": ln.line_total if ln.line_total is not None else "",
                "PICKING": STATUS_PENDING, "PACKING": STATUS_PENDING,
                "DISPATCH": STATUS_PENDING,
            })

    out_s = pd.DataFrame(srows, columns=SUMMARY_COLS)
    out_d = pd.DataFrame(drows, columns=DETAIL_COLS)
    out_s.attrs["skipped_duplicates"] = skipped
    out_d.attrs["skipped_duplicates"] = skipped
    return out_s, out_d


# --------------------------------------------------------------------------- #
# merge into what is already stored
# --------------------------------------------------------------------------- #
def is_duplicate_row(df: pd.DataFrame) -> pd.Series:
    """Rows that only exist because an invoice was uploaded a second time."""
    if df is None or not len(df):
        return pd.Series(dtype=bool)
    rem = df.get("REMARK", pd.Series("", index=df.index)).astype(str)
    pick = df.get("KORBER_PICK", pd.Series("", index=df.index)).astype(str).str.strip()
    return rem.str.contains(DUPLICATE_PAT, case=False, na=False) & (pick != "Yes")


def strip_duplicates(summary: pd.DataFrame, details: pd.DataFrame | None = None
                     ) -> dict[str, Any]:
    """Drop duplicate rows from the register — and their detail lines with them."""
    if summary is None or not len(summary):
        return {"summary": summary, "details": details, "dropped": 0, "invoices": []}
    bad = is_duplicate_row(summary)
    nums = [_id_str(x) for x in summary.loc[bad, "TAX_INVOICE_NO"]]
    out_s = summary[~bad].reset_index(drop=True)
    out_d = details
    if details is not None and len(details) and nums:
        out_d = details[~details["TAX_INVOICE_NO"].map(_id_str).isin(set(nums))] \
            .reset_index(drop=True)
    return {"summary": out_s, "details": out_d, "dropped": int(bad.sum()),
            "invoices": nums}


def merge_summary(old: pd.DataFrame | None, new: pd.DataFrame) -> dict[str, Any]:
    """
    One row per invoice, updated in place.

    A `No` that is picked later becomes `Yes` and loses its remark. A `Yes` is
    not pushed back to `No` by a later run that happened to skip it as a
    duplicate — only deleting the load does that (`mark_unpicked`).

    `Partial` sits between the two: a partly picked document may still be
    completed later (`Partial` -> `Yes`), but a later run that takes nothing
    must not erase what already went out (`Partial` never falls back to `No`).
    """
    cols = SUMMARY_COLS
    if new is not None and len(new):          # belt and braces — never store one
        new = new[~is_duplicate_row(new)].reset_index(drop=True)
    if old is None or not len(old):
        return {"data": new.reindex(columns=cols), "new": len(new), "updated": 0,
                "picked_now": [_id_str(n) for n, k in
                               zip(new["TAX_INVOICE_NO"], new["KORBER_PICK"])
                               if k == "Yes"]}

    o = old.reindex(columns=cols).copy()
    o["TAX_INVOICE_NO"] = o["TAX_INVOICE_NO"].map(_id_str)
    o = o.drop_duplicates(subset=["TAX_INVOICE_NO"], keep="last")
    idx = {_id_str(v): i for i, v in enumerate(o["TAX_INVOICE_NO"])}
    rows = o.to_dict("records")

    n_new = n_upd = 0
    picked_now: list[str] = []
    for _, r in new.iterrows():
        num = _id_str(r["TAX_INVOICE_NO"])
        rec = dict(r)
        rec["TAX_INVOICE_NO"] = num
        if num not in idx:
            rows.append(rec)
            n_new += 1
            if rec["KORBER_PICK"] == PICK_YES:
                picked_now.append(num)
            continue

        cur = rows[idx[num]]
        was = str(cur.get("KORBER_PICK", "No"))
        rec["FIRST_SEEN"] = cur.get("FIRST_SEEN") or rec["FIRST_SEEN"]
        # Picking/Packing/Dispatch track the physical warehouse execution, not
        # this run's own pallet allocation — a re-upload of the same invoice
        # must never roll a Completed status back to Pending.
        for k in STATUS_COLS:
            if str(cur.get(k, "")) == STATUS_DONE:
                rec[k] = STATUS_DONE
        now = str(rec.get("KORBER_PICK", PICK_NO))
        if pick_rank(now) < pick_rank(was):
            # already further along — keep the pick, just refresh what the
            # document itself says
            for k in ("KORBER_PICK", "REMARK", "RUN_ID", "PICKED_AT", "PICKED_QTY",
                      "PLANT"):
                rec[k] = cur.get(k, "")
        elif pick_rank(now) > pick_rank(was):
            if now == PICK_YES:
                rec["REMARK"] = ""
                picked_now.append(num)
        rows[idx[num]] = rec
        n_upd += 1

    out = pd.DataFrame(rows, columns=cols)
    out = out.drop_duplicates(subset=["TAX_INVOICE_NO"], keep="last").reset_index(drop=True)
    return {"data": out, "new": n_new, "updated": n_upd, "picked_now": picked_now}


def merge_details(old: pd.DataFrame | None, new: pd.DataFrame) -> pd.DataFrame:
    """
    Detail rows are replaced wholesale for the invoices in this run — except the
    Picking/Packing/Dispatch status, which belongs to the physical warehouse
    execution and must survive a re-upload of the same invoice/line.
    """
    if new is None or not len(new):
        return (old if old is not None else pd.DataFrame(columns=DETAIL_COLS))
    # same belt and braces as merge_summary: a line that only exists because the
    # invoice was uploaded twice is not a line, and its summary row is dropped
    new = new[~is_duplicate_row(new)]
    if not len(new):
        return (old if old is not None else pd.DataFrame(columns=DETAIL_COLS))
    new = new.reindex(columns=DETAIL_COLS).copy()
    if old is None or not len(old):
        return new
    nums = {_id_str(x) for x in new["TAX_INVOICE_NO"]}
    old = old.reindex(columns=DETAIL_COLS)
    touched = old["TAX_INVOICE_NO"].map(_id_str).isin(nums)
    prior = old[touched]
    if len(prior):
        # Carry every Completed status forward by an indexed lookup rather than
        # a per-row scan — .iterrows() over the prior rows cost 136 ms on a
        # 30 000-line register, and this runs inside the save lock.
        pkey = (prior["TAX_INVOICE_NO"].map(_id_str) + "|"
                + prior["LINE"].astype(str))
        nkey = (new["TAX_INVOICE_NO"].map(_id_str) + "|"
                + new["LINE"].astype(str))
        for k in STATUS_COLS:
            done = pkey[prior[k].astype(str) == STATUS_DONE]
            if len(done):
                new.loc[nkey.isin(set(done)), k] = STATUS_DONE

        # A line never walks backwards — Yes stays Yes, and a Partial line is
        # not reset to No by a later run that took nothing. merge_summary()
        # protects the summary the same way; without this the Details tab
        # answered "No, pallets unknown" for an invoice the Summary tab still
        # calls picked, and the allocation behind it was lost.
        prior_rank = prior["KORBER_PICK"].map(pick_rank)
        new_rank = new["KORBER_PICK"].map(pick_rank)
        best = dict(zip(pkey, prior_rank))
        was_yes = pkey[prior_rank > 0]
        if len(was_yes):
            keep = nkey.isin(set(was_yes)) & \
                (new_rank.values < nkey.map(best).fillna(0).values)
            if keep.any():
                cols = ["KORBER_PICK", "PICKED_QTY", "ITEM_NUMBER", "LOT_NUMBER",
                        "PALLETS", "LOCATIONS", "PLANT", "REMARK", "RUN_ID"]
                src = prior.set_index(pd.Index(pkey.to_numpy(dtype=object)))
                src = src[~src.index.duplicated(keep="last")]
                take = nkey[keep]
                new = _writable(new, *cols)
                for c in cols:
                    new.loc[keep, c] = take.map(src[c]).values
    return pd.concat([old[~touched], new], ignore_index=True)


def _writable(df: pd.DataFrame, *cols: str) -> pd.DataFrame:
    """
    Widen the columns we are about to write into.

    pandas 3 refuses to put a number in a `str`-dtype column, and a frame read
    from the sheet is all strings. Only the columns being written are cast, so
    a 30 000-row register does not get copied wholesale for the sake of one
    cell.
    """
    for c in cols:
        if c in df.columns and df[c].dtype != object:
            df[c] = df[c].astype(object)
    return df


def mark_unpicked(summary: pd.DataFrame, invoice_no: str,
                  remark: str = "Load deleted") -> pd.DataFrame:
    """Deleting a load frees the stock, so the register has to say No again."""
    if summary is None or not len(summary):
        return summary
    d = summary.copy()
    m = d["TAX_INVOICE_NO"].map(_id_str) == _id_str(invoice_no).strip()
    if not m.any():
        return d
    _writable(d, "PICKED_QTY")
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    d.loc[m, "KORBER_PICK"] = "No"
    d.loc[m, "REMARK"] = f"{remark} · {stamp}"
    for c in ("RUN_ID", "PICKED_AT"):
        d.loc[m, c] = ""
    d.loc[m, "PICKED_QTY"] = 0.0
    d.loc[m, "UPDATED_AT"] = stamp
    return d


def mark_unpicked_details(details: pd.DataFrame, invoice_no: str,
                          remark: str = "Load deleted") -> pd.DataFrame:
    """
    The detail rows have to be released with the summary.

    Deleting a load hands the pallets back, so a detail line that still named a
    pallet, lot and location would be pointing at stock that is free again —
    and the register would answer "Yes, picked" on the Details tab while the
    Summary tab says "No". Picking/Packing/Dispatch are *not* touched: those
    record what the floor did, not this app's allocation.
    """
    if details is None or not len(details):
        return details
    d = details.copy()
    m = d["TAX_INVOICE_NO"].map(_id_str) == _id_str(invoice_no).strip()
    if not m.any():
        return d
    _writable(d, "PICKED_QTY")
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    d.loc[m, "KORBER_PICK"] = "No"
    d.loc[m, "REMARK"] = f"{remark} · {stamp}"
    d.loc[m, "PICKED_QTY"] = 0.0
    for c in ("ITEM_NUMBER", "LOT_NUMBER", "PALLETS", "LOCATIONS", "RUN_ID"):
        if c in d.columns:
            d.loc[m, c] = ""
    d.loc[m, "UPDATED_AT"] = stamp
    return d


# --------------------------------------------------------------------------- #
# Picking / Packing / Dispatch — warehouse execution status
# --------------------------------------------------------------------------- #
# These three are independent of KORBER_PICK (this app's own pallet allocation)
# and only ever move Pending -> Completed. A stale or re-uploaded report, or a
# re-run of the pick, must never undo progress already confirmed on the floor.
def _norm_col(s: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s or "").lower())


def _find_col(df: pd.DataFrame, *names: str) -> str | None:
    """Best-effort header match for an uploaded report — exact normalized name
    first, then substring, so small formatting differences don't break it."""
    if df is None or not len(df.columns):
        return None
    norm = {c: _norm_col(c) for c in df.columns}
    for name in names:
        want = _norm_col(name)
        for c, n in norm.items():
            if n == want:
                return c
    for name in names:
        want = _norm_col(name)
        if not want:
            continue
        for c, n in norm.items():
            if want in n:
                return c
    return None


def _id_str(x: Any) -> str:
    """
    Excel stores an all-digit invoice / Load Id as a number, so it reads back
    '333262712441.0' — one character different from the string version this
    app uses everywhere else. Strip a trailing whole-number '.0' before any
    matching is done.
    """
    if isinstance(x, float):
        if pd.isna(x):
            return ""
        return str(int(x)) if x.is_integer() else str(x)
    s = str(x or "").strip()
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def _to_num(x: Any) -> float | None:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = str(x).replace(",", "").strip()
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


# --------------------------------------------------------------------------- #
# Pull the live status straight off the Körber dashboard
# --------------------------------------------------------------------------- #
# The dashboard serves the same table the Excel export holds, so the fetch only
# has to end with a DataFrame — after that it is the ordinary upload path. The
# response could be an HTML page, a CSV, a JSON array or a real .xlsx, so sniff
# rather than assume: a warehouse endpoint's content type is often wrong.
FETCH_TIMEOUT = 30


def _table_from_bytes(body: bytes, content_type: str = "") -> tuple[pd.DataFrame, str]:
    ct = (content_type or "").lower()
    head = body[:512].lstrip()

    if body[:4] == b"PK\x03\x04":                       # a real xlsx
        return pd.read_excel(io.BytesIO(body)), "excel"

    if head[:1] in (b"{", b"[") or "json" in ct:
        import json as _json
        data = _json.loads(body.decode("utf-8", "replace"))
        # the rows may sit under a key — take the first list of objects found
        if isinstance(data, dict):
            for v in data.values():
                if isinstance(v, list) and v and isinstance(v[0], dict):
                    data = v
                    break
        return pd.DataFrame(data), "json"

    text = body.decode("utf-8", "replace")
    if "<table" in text.lower() or "html" in ct:
        tables = pd.read_html(io.StringIO(text))
        if not tables:
            raise ValueError("the page has no table in it")
        # the widest table with a Load Id column, else simply the widest
        best = max((t for t in tables if _find_col(t, "Load Id")),
                   key=lambda t: t.shape[1], default=None)
        return (best if best is not None
                else max(tables, key=lambda t: t.shape[1])), "html"

    return pd.read_csv(io.StringIO(text), sep=None, engine="python"), "csv"


def fetch_live_status(url: str, user: str = "", password: str = "",
                      token: str = "", timeout: int = FETCH_TIMEOUT) -> dict[str, Any]:
    """
    GET the dashboard and hand back the table it is showing.

    Returns the same shape whatever went wrong, so the caller can show the
    reason instead of a traceback — a warehouse box being unreachable is an
    ordinary Tuesday, not an exception.
    """
    out = {"data": pd.DataFrame(), "kind": "", "rows": 0, "error": "", "url": url}
    if not str(url or "").strip():
        out["error"] = "No URL."
        return out
    try:
        import requests
    except Exception:
        out["error"] = "The `requests` package is not installed."
        return out

    try:
        headers = {"Accept": "text/html,application/json,text/csv,*/*"}
        if token:
            headers["Authorization"] = f"Bearer {token}"
        auth = (user, password) if user else None
        r = requests.get(str(url).strip(), timeout=timeout, headers=headers, auth=auth)
        r.raise_for_status()
        df, kind = _table_from_bytes(r.content, r.headers.get("Content-Type", ""))
    except Exception as ex:                                   # noqa: BLE001
        out["error"] = f"{type(ex).__name__}: {str(ex)[:300]}"
        return out

    if df is None or not len(df):
        out["error"] = "The response held no rows."
        return out
    df = df.dropna(axis=1, how="all")
    if not _find_col(df, "Load Id"):
        out["error"] = ("No 'Load Id' column in the table that came back — "
                        f"found: {', '.join(str(c) for c in df.columns[:12])}")
        out["data"], out["kind"], out["rows"] = df, kind, len(df)
        return out
    out.update({"data": df, "kind": kind, "rows": len(df)})
    return out


LIVE_STATUS_REPORT_COLS = ["LOAD_ID", "MATCHED", "PICKING", "DISPATCH"]


def apply_pick_live_status(summary: pd.DataFrame, details: pd.DataFrame,
                           live_df: pd.DataFrame) -> dict[str, Any]:
    """
    Pick_Live_status report -> Picking / Dispatch, matched on Load Id == Tax
    Invoice No.

        Open Pick == 0                                   -> Picking  = Completed
        Shipped Pick != 0  OR  Total Pick == Shipped Pick -> Dispatch = Completed
        (the second rule needs Total Pick > 0 — two blank/zero columns are not
        a shipment)
    """
    empty = pd.DataFrame(columns=LIVE_STATUS_REPORT_COLS)
    out = {"summary": summary, "details": details, "report": empty,
          "picking_done": 0, "dispatch_done": 0, "unmatched": [], "error": ""}
    if live_df is None or not len(live_df) or summary is None or not len(summary):
        return out

    c_load = _find_col(live_df, "Load Id")
    c_open = _find_col(live_df, "Open Pick")
    c_total = _find_col(live_df, "Total Pick")
    c_ship = _find_col(live_df, "Shipped Pick")
    if not c_load:
        out["error"] = "Could not find a 'Load Id' column in this file."
        return out

    rows, pick_set, disp_set = [], {}, {}
    for _, r in live_df.iterrows():
        lid = _id_str(r.get(c_load, "")).strip()
        if not lid:
            continue
        op = _to_num(r.get(c_open)) if c_open else None
        tp = _to_num(r.get(c_total)) if c_total else None
        sp = _to_num(r.get(c_ship)) if c_ship else None
        p_done = op is not None and op == 0
        d_done = (sp is not None and sp != 0) or \
                 (tp is not None and sp is not None and tp > 0 and tp == sp)
        pick_set[lid] = pick_set.get(lid, False) or p_done
        disp_set[lid] = disp_set.get(lid, False) or d_done

    known = {_id_str(x).strip() for x in summary["TAX_INVOICE_NO"]}
    for lid in sorted(set(pick_set) | set(disp_set)):
        rows.append({"LOAD_ID": lid, "MATCHED": lid in known,
                    "PICKING": STATUS_DONE if pick_set.get(lid) else "",
                    "DISPATCH": STATUS_DONE if disp_set.get(lid) else ""})

    s = summary.copy()
    # _id_str, not astype(str): a register read from Excel rather than the sheet
    # gives '30426013174.0' and would match nothing.
    inv = s["TAX_INVOICE_NO"].map(_id_str).str.strip()
    n_pick = n_disp = 0
    for lid, done in pick_set.items():
        if not done:
            continue
        m = (inv == lid) & (s["PICKING"] != STATUS_DONE)
        n_pick += int(m.sum())
        s.loc[m, "PICKING"] = STATUS_DONE
    for lid, done in disp_set.items():
        if not done:
            continue
        m = (inv == lid) & (s["DISPATCH"] != STATUS_DONE)
        n_disp += int(m.sum())
        s.loc[m, "DISPATCH"] = STATUS_DONE

    d = details
    if details is not None and len(details):
        d = details.copy()
        dinv = d["TAX_INVOICE_NO"].map(_id_str).str.strip()
        for lid, done in pick_set.items():
            if done:
                d.loc[(dinv == lid) & (d["PICKING"] != STATUS_DONE), "PICKING"] = STATUS_DONE
        for lid, done in disp_set.items():
            if done:
                d.loc[(dinv == lid) & (d["DISPATCH"] != STATUS_DONE), "DISPATCH"] = STATUS_DONE

    return {"summary": s, "details": d, "report": pd.DataFrame(rows, columns=LIVE_STATUS_REPORT_COLS),
           "picking_done": n_pick, "dispatch_done": n_disp,
           "unmatched": [r["LOAD_ID"] for r in rows if not r["MATCHED"]], "error": ""}


def apply_status_scan(summary: pd.DataFrame, details: pd.DataFrame, load_id: str,
                      column: str = "PACKING", user: str = "") -> dict[str, Any]:
    """
    One LOAD_ID -> that column = Completed, on the summary and its detail lines.

    Drives both floor screens: the Packing station sets `PACKING`, the Dispatch
    station sets `DISPATCH`. The LOAD_ID may come from a QR scan or be typed in
    by hand — identical either way, so both routes share this.
    """
    if column not in STATUS_COLS:
        raise ValueError(f"{column!r} is not one of {STATUS_COLS}")
    lid = _id_str(load_id).strip()
    out = {"found": False, "already": False, "changed": False, "invoice": None,
          "column": column, "summary": summary, "details": details}
    if not lid or summary is None or not len(summary):
        return out
    inv = summary["TAX_INVOICE_NO"].map(_id_str).str.strip()
    m = inv == lid
    if not m.any():
        return out

    out["found"] = True
    out["invoice"] = summary.loc[m].iloc[0].to_dict()
    out["already"] = str(out["invoice"].get(column, "")) == STATUS_DONE

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    s = summary.copy()
    s.loc[m, column] = STATUS_DONE
    s.loc[m, "UPDATED_AT"] = stamp
    if user:
        s.loc[m, "UPDATED_BY"] = user
    out["summary"] = s
    changed = not out["already"]

    if details is not None and len(details):
        d = details.copy()
        dm = d["TAX_INVOICE_NO"].map(_id_str).str.strip() == lid
        # `already` speaks for the summary row only. The detail lines can still
        # be Pending — a line added by a later re-upload, or a backfill that ran
        # after the scan — and skipping the write on `already` left them Pending
        # for good.
        changed = changed or bool((dm & (d[column].astype(str) != STATUS_DONE)).any())
        d.loc[dm, column] = STATUS_DONE
        d.loc[dm, "UPDATED_AT"] = stamp
        out["details"] = d
    out["changed"] = changed
    return out


def apply_packing_scan(summary: pd.DataFrame, details: pd.DataFrame, load_id: str,
                       user: str = "") -> dict[str, Any]:
    """OUTBOUND PICK SHEET LOAD_ID -> Packing = Completed."""
    return apply_status_scan(summary, details, load_id, "PACKING", user)


def apply_dispatch_scan(summary: pd.DataFrame, details: pd.DataFrame, load_id: str,
                        user: str = "") -> dict[str, Any]:
    """OUTBOUND PICK SHEET LOAD_ID -> Dispatch = Completed."""
    return apply_status_scan(summary, details, load_id, "DISPATCH", user)


# --------------------------------------------------------------------------- #
# Backfill — invoices picked before the register existed
# --------------------------------------------------------------------------- #
BACKFILL_COLS = ["TAX_INVOICE_NO", "DOC_TYPE", "TAX_INVOICE_DATE", "LINES", "QTY",
                 "PLANT", "PICKED_AT", "SOURCE"]

BACKFILL_REMARK = "Backfilled from the pick history"


def _join_unique(s: pd.Series) -> str:
    return ", ".join(dict.fromkeys(str(x) for x in s if str(x).strip()))


def _first_text(s: pd.Series) -> str:
    for x in s:
        t = str(x).strip()
        if t:
            return t
    return ""


def backfill_from_history(ledger: pd.DataFrame, registry: pd.DataFrame | None,
                          summary: pd.DataFrame | None, details: pd.DataFrame | None,
                          user: str = "") -> dict[str, Any]:
    """
    Rebuild register rows for invoices that were picked before the register
    existed — straight from `PALLET_LEDGER`, no files needed.

    The ledger is the honest record of what was actually taken off a pallet:
    one row per document line per pallet, with the item, lot, location and
    quantity. A pick is all-or-nothing, so the picked quantity of a line *is*
    the document quantity, and the register row can be reconstructed exactly.

    `DOC_REGISTRY` fills in the document date and AR number. Only its first
    few columns are read, deliberately: a stale header on that sheet shifted
    everything after `PICKED_QTY` by two columns on older rows, so the tail is
    not trustworthy, while `DOC_NUMBER … PICKED_QTY` are.

    The customer name cannot be recovered this way — it lives only on the PDF —
    so those rows are left blank and can be filled by re-uploading the invoice.
    """
    out = {"summary": summary, "details": details, "report":
           pd.DataFrame(columns=BACKFILL_COLS), "added": 0, "invoices": []}
    if ledger is None or not len(ledger) or "DOC_NUMBER" not in ledger.columns:
        return out

    led = ledger.copy()
    led["_DOC"] = led["DOC_NUMBER"].map(_id_str).str.strip()
    led = led[led["_DOC"] != ""]
    if not len(led):
        return out
    led["_QTY"] = pd.to_numeric(led.get("QTY_PICKED"), errors="coerce").fillna(0.0)
    led["_LINE"] = pd.to_numeric(led.get("DOC_LINE"), errors="coerce").fillna(0).astype(int)

    known = set()
    if summary is not None and len(summary):
        known = set(summary["TAX_INVOICE_NO"].map(_id_str).str.strip())
    todo = [d for d in dict.fromkeys(led["_DOC"]) if d not in known]
    if not todo:
        return out
    led = led[led["_DOC"].isin(set(todo))]

    # document date / AR number from the registry's trustworthy prefix, plus
    # the document quantity where the registry has it — a partial pick took
    # less than the document asked for, so the ledger alone would understate it
    meta: dict[str, dict] = {}
    if registry is not None and len(registry) and "DOC_NUMBER" in registry.columns:
        for col in ("DOC_DATE", "REF_NUMBER", "DOC_QTY", "PICK_STATUS"):
            if col not in registry.columns:
                registry = registry.assign(**{col: ""})
        for _, r in registry.iterrows():
            meta[_id_str(r["DOC_NUMBER"]).strip()] = {
                "date": str(r.get("DOC_DATE", "") or "").strip(),
                "ref": str(r.get("REF_NUMBER", "") or "").strip(),
                "doc_qty": _to_num(r.get("DOC_QTY")),
                "partial": str(r.get("PICK_STATUS", "")).strip().upper() == "PARTIAL"}

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ---- one detail row per document line ----
    g = led.groupby(["_DOC", "_LINE"], sort=True)
    det = pd.DataFrame({
        "PICKED_QTY": g["_QTY"].sum(),
        "ITEM_CODE": g["DOC_ITEM_CODE"].agg(_first_text) if "DOC_ITEM_CODE" in led
        else g["_DOC"].agg(lambda _s: ""),
        "DESCRIPTION": g["DESCRIPTION"].agg(_first_text) if "DESCRIPTION" in led
        else g["_DOC"].agg(lambda _s: ""),
        "ITEM_NUMBER": g["ITEM_NUMBER"].agg(_join_unique) if "ITEM_NUMBER" in led
        else g["_DOC"].agg(lambda _s: ""),
        "LOT_NUMBER": g["LOT_NUMBER"].agg(_join_unique) if "LOT_NUMBER" in led
        else g["_DOC"].agg(lambda _s: ""),
        "PALLETS": g["PALLET"].agg(_join_unique) if "PALLET" in led
        else g["_DOC"].agg(lambda _s: ""),
        "LOCATIONS": g["LOCATION_ID"].agg(_join_unique) if "LOCATION_ID" in led
        else g["_DOC"].agg(lambda _s: ""),
        "PLANT": g["PLANT"].agg(_join_unique) if "PLANT" in led
        else g["_DOC"].agg(lambda _s: ""),
        "UOM": g["UOM"].agg(_first_text) if "UOM" in led
        else g["_DOC"].agg(lambda _s: ""),
        "RUN_ID": g["RUN_ID"].agg(_first_text) if "RUN_ID" in led
        else g["_DOC"].agg(lambda _s: ""),
        "DOC_TYPE": g["DOC_TYPE"].agg(_first_text) if "DOC_TYPE" in led
        else g["_DOC"].agg(lambda _s: ""),
    }).reset_index().rename(columns={"_DOC": "TAX_INVOICE_NO", "_LINE": "LINE"})

    det["BASE_ID"] = det["ITEM_CODE"].map(base_item)
    # A full pick is all-or-nothing, so the picked quantity of a line *is* the
    # document quantity. A partial pick is the exception, and the ledger cannot
    # tell them apart on its own — DOC_REGISTRY's PICK_STATUS can.
    det["DOC_QTY"] = det["PICKED_QTY"]
    det["KORBER_PICK"] = PICK_YES
    det["REMARK"] = BACKFILL_REMARK
    det["UPDATED_AT"] = now
    det["TAX_INVOICE_DATE"] = det["TAX_INVOICE_NO"].map(
        lambda d: meta.get(d, {}).get("date", ""))
    det["AR_INVOICE_NO"] = det["TAX_INVOICE_NO"].map(
        lambda d: meta.get(d, {}).get("ref", ""))
    for c in STATUS_COLS:
        det[c] = STATUS_PENDING
    det = det.reindex(columns=DETAIL_COLS, fill_value="")

    # ---- one summary row per document ----
    dg = det.groupby("TAX_INVOICE_NO", sort=True)
    sm = pd.DataFrame({
        "LINES": dg["LINE"].nunique(),
        "QTY": dg["DOC_QTY"].sum(),
        "PICKED_QTY": dg["PICKED_QTY"].sum(),
        "PLANT": dg["PLANT"].agg(_join_unique),
        "RUN_ID": dg["RUN_ID"].agg(_first_text),
        "DOC_TYPE": dg["DOC_TYPE"].agg(_first_text),
        "TAX_INVOICE_DATE": dg["TAX_INVOICE_DATE"].agg(_first_text),
        "AR_INVOICE_NO": dg["AR_INVOICE_NO"].agg(_first_text),
    }).reset_index()

    lg = led.groupby("_DOC", sort=True)
    picked_at = lg["PICK_DATE"].agg(_first_text) if "PICK_DATE" in led.columns else None
    source = lg["SOURCE_FILE"].agg(_first_text) if "SOURCE_FILE" in led.columns else None
    sm["PICKED_AT"] = (sm["TAX_INVOICE_NO"].map(picked_at).fillna("")
                       if picked_at is not None else "")
    sm["SOURCE_FILE"] = (sm["TAX_INVOICE_NO"].map(source).fillna("")
                         if source is not None else "")
    # a document the registry marked PARTIAL still owes its balance
    part = sm["TAX_INVOICE_NO"].map(lambda d: bool(meta.get(d, {}).get("partial")))
    want = sm["TAX_INVOICE_NO"].map(lambda d: meta.get(d, {}).get("doc_qty"))
    sm["QTY"] = [float(w) if (pt and w) else float(q)
                 for pt, w, q in zip(part, want, sm["QTY"])]
    sm["KORBER_PICK"] = [PICK_PART if pt else PICK_YES for pt in part]
    sm["REMARK"] = [
        (f"{BACKFILL_REMARK} · PARTIAL PICK — "
         f"{_qty_str(max(0.0, float(q) - float(g)))} of {_qty_str(float(q))} still owed")
        if pt else BACKFILL_REMARK
        for pt, q, g in zip(part, sm["QTY"], sm["PICKED_QTY"])]
    sm["MRP"] = "No"
    sm["CUSTOMER_NAME"] = ""            # only the PDF has it
    sm["FIRST_SEEN"] = sm["PICKED_AT"].where(sm["PICKED_AT"].astype(bool), now)
    sm["UPDATED_AT"] = now
    sm["UPDATED_BY"] = user
    for c in STATUS_COLS:
        sm[c] = STATUS_PENDING
    sm = sm.reindex(columns=SUMMARY_COLS, fill_value="")

    report = pd.DataFrame({
        "TAX_INVOICE_NO": sm["TAX_INVOICE_NO"], "DOC_TYPE": sm["DOC_TYPE"],
        "TAX_INVOICE_DATE": sm["TAX_INVOICE_DATE"], "LINES": sm["LINES"],
        "QTY": sm["QTY"], "PLANT": sm["PLANT"], "PICKED_AT": sm["PICKED_AT"],
        "SOURCE": "PALLET_LEDGER",
    }).reindex(columns=BACKFILL_COLS)

    merged_s = merge_summary(summary if summary is not None and len(summary) else None, sm)
    merged_d = merge_details(details if details is not None and len(details) else None, det)
    return {"summary": merged_s["data"], "details": merged_d, "report": report,
            "added": len(sm), "invoices": [str(x) for x in sm["TAX_INVOICE_NO"]]}


def enrich_from_history(summary: pd.DataFrame, details: pd.DataFrame,
                        ledger: pd.DataFrame | None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    An uploaded invoice that was in fact picked earlier must not be filed as
    `No`. Where the ledger has a pick for it, mark it Yes and carry the pallet
    detail across — so re-uploading a PDF fills in the customer name the ledger
    never had, without losing what the ledger knows.
    """
    if summary is None or not len(summary) or ledger is None or not len(ledger) \
            or "DOC_NUMBER" not in ledger.columns:
        return summary, details

    led = ledger.copy()
    led["_DOC"] = led["DOC_NUMBER"].map(_id_str).str.strip()
    led["_QTY"] = pd.to_numeric(led.get("QTY_PICKED"), errors="coerce").fillna(0.0)
    led["_LINE"] = pd.to_numeric(led.get("DOC_LINE"), errors="coerce").fillna(0).astype(int)

    picked = led.groupby("_DOC")["_QTY"].sum()
    runs = led.groupby("_DOC")["RUN_ID"].agg(_first_text) if "RUN_ID" in led else None
    when = led.groupby("_DOC")["PICK_DATE"].agg(_first_text) if "PICK_DATE" in led else None
    plants = led.groupby("_DOC")["PLANT"].agg(_join_unique) if "PLANT" in led else None

    s = _writable(summary.copy(), "PICKED_QTY")
    key = s["TAX_INVOICE_NO"].map(_id_str).str.strip()
    hit = key.isin(set(picked.index))
    if hit.any():
        got = key[hit].map(picked).astype(float)
        want = pd.to_numeric(s.loc[hit, "QTY"], errors="coerce").fillna(0.0)
        # the ledger says what came off the pallets; if that is less than the
        # document asks for, this is a partial pick, not a completed one
        full = (got.values + 1e-9) >= want.values
        s.loc[hit, "KORBER_PICK"] = [PICK_YES if f else PICK_PART for f in full]
        s.loc[hit, "REMARK"] = [
            "" if f else f"PARTIAL PICK — {_qty_str(max(0.0, w - g))} of "
                         f"{_qty_str(w)} still owed"
            for f, g, w in zip(full, got.values, want.values)]
        s.loc[hit, "PICKED_QTY"] = got.values
        if runs is not None:
            s.loc[hit, "RUN_ID"] = key[hit].map(runs).values
        if when is not None:
            s.loc[hit, "PICKED_AT"] = key[hit].map(when).values
        if plants is not None:
            s.loc[hit, "PLANT"] = key[hit].map(plants).values

    d = details
    if details is not None and len(details):
        d = _writable(details.copy(), "PICKED_QTY")
        dkey = (d["TAX_INVOICE_NO"].map(_id_str).str.strip() + "|"
                + pd.to_numeric(d["LINE"], errors="coerce").fillna(0)
                .astype(int).astype(str))
        lg = led.groupby([led["_DOC"], led["_LINE"]])
        line_qty = lg["_QTY"].sum()
        line_qty.index = [f"{a}|{b}" for a, b in line_qty.index]
        dhit = dkey.isin(set(line_qty.index))
        if dhit.any():
            got = dkey[dhit].map(line_qty).astype(float)
            want = pd.to_numeric(d.loc[dhit, "DOC_QTY"], errors="coerce").fillna(0.0)
            full = (got.values + 1e-9) >= want.values
            d.loc[dhit, "KORBER_PICK"] = [PICK_YES if f else PICK_PART for f in full]
            d.loc[dhit, "REMARK"] = ["" if f else "PARTIAL PICK" for f in full]
            d.loc[dhit, "PICKED_QTY"] = got.values
            for col, src in (("ITEM_NUMBER", "ITEM_NUMBER"), ("LOT_NUMBER", "LOT_NUMBER"),
                             ("PALLETS", "PALLET"), ("LOCATIONS", "LOCATION_ID"),
                             ("PLANT", "PLANT")):
                if src in led.columns:
                    vals = lg[src].agg(_join_unique)
                    vals.index = [f"{a}|{b}" for a, b in vals.index]
                    d.loc[dhit, col] = dkey[dhit].map(vals).values
    return s, d


SALES_REPORT_COLS = ["TAX_INVOICE_NO", "CUSTOMER_ITEM", "ITEM_CODE", "BASE_ID",
                     "SALES_QTY", "IN_WMS_MASTER", "WMS_QTY", "STATUS"]


def apply_sales_report(summary: pd.DataFrame, details: pd.DataFrame,
                       sales_df: pd.DataFrame, wms_master: pd.DataFrame | None = None,
                       wms_detail: pd.DataFrame | None = None) -> dict[str, Any]:
    """
    Invoice sales report -> reconciled against the **actual WMS output**.

    `OUTBOUND_MASTER.LOAD_ID` answers "was this invoice really picked and
    pushed to the WMS at all?", and `OUTBOUND_DETAIL.DISPLAY_ITEM_NUMBER` /
    `QTY` answers "what was actually sent". `INVOICE_DETAIL` exists for every
    uploaded document whether or not it was ever picked, so it is not proof of
    anything by itself.

    **The report has one row per uploaded sales-report row**, not per register
    line. The register only holds the invoices this app has parsed — typically
    a fraction of what the ERP exports — so driving the report off the register
    hid every invoice the ERP knows about but the register has not seen, which
    is exactly the reconciliation the report exists to show.

    Matching: Tax Invoice No. == `OUTBOUND_MASTER.LOAD_ID`, then item
    (Item Code / Customer Item, base-ID matched like everywhere else in this
    app) + Qty against `OUTBOUND_DETAIL` for that same order. Where the
    quantities agree, the matching register line (if any) is marked
    Picking = Completed; once every line of an invoice is confirmed, the
    invoice itself is marked Picking = Completed.

    Falls back to comparing against the register's own `DOC_QTY` when no WMS
    output has been saved yet — degraded, but still useful before a pick run.
    """
    empty = pd.DataFrame(columns=SALES_REPORT_COLS)
    out = {"summary": summary, "details": details, "report": empty,
          "matched": 0, "invoices_completed": 0, "error": "", "used_wms": False}
    if sales_df is None or not len(sales_df):
        return out

    c_inv = _find_col(sales_df, "Tax Invoice No.", "Tax Invoice No")
    c_item = _find_col(sales_df, "Item Code")
    c_citem = _find_col(sales_df, "Customer Item")
    c_qty = _find_col(sales_df, "QTY", "Quantity")
    if not c_inv or not c_qty or (not c_item and not c_citem):
        out["error"] = "Could not find Tax Invoice No. / Item Code / QTY columns."
        return out

    # Register under *every* candidate base, not just one preferred column —
    # this ERP export sometimes puts "P601560 710" (a space, not a hyphen) in
    # Item Code, which base_item() can't split, while Customer Item ("P601560")
    # is already clean. Whichever one actually matches the WMS side wins; the
    # other key is simply never looked up, so this never double-counts.
    sales_qty: dict[tuple[str, str], float] = {}
    for _, r in sales_df.iterrows():
        inv = _id_str(r.get(c_inv, "")).strip()
        if not inv:
            continue
        bases = {b for b in (base_item(r.get(c_item)) if c_item else "",
                             base_item(r.get(c_citem)) if c_citem else "") if b}
        qty = _to_num(r.get(c_qty)) or 0.0
        for base in bases:
            sales_qty[(inv, base)] = sales_qty.get((inv, base), 0.0) + qty

    # ---- what the WMS actually received, if it has been saved ----
    # OUTBOUND_MASTER carries both LOAD_ID and DISPLAY_ORDER_NUMBER — they are
    # the same Invoice / DC No, but LOAD_ID is the one asked for here.
    c_master_inv = _find_col(wms_master, "LOAD_ID", "DISPLAY_ORDER_NUMBER")
    c_wd_inv = _find_col(wms_detail, "DISPLAY_ORDER_NUMBER", "LOAD_ID")
    c_wd_item = _find_col(wms_detail, "DISPLAY_ITEM_NUMBER")
    c_wd_qty = _find_col(wms_detail, "QTY")
    use_wms = bool(wms_master is not None and len(wms_master) and c_master_inv
                   and wms_detail is not None and len(wms_detail)
                   and c_wd_inv and c_wd_item)
    out["used_wms"] = use_wms
    sent_invoices: set[str] = set()
    wms_qty: dict[tuple[str, str], float] = {}
    if use_wms:
        sent_invoices = {_id_str(x).strip() for x in wms_master[c_master_inv]}
        wd = wms_detail.copy()
        wd["_INV"] = wd[c_wd_inv].map(_id_str).str.strip()
        wd["_BASE"] = wd[c_wd_item].map(base_item)
        wd["_QTY"] = pd.to_numeric(wd[c_wd_qty], errors="coerce").fillna(0.0) if c_wd_qty \
            else 0.0
        for (inv, base), grp in wd.groupby(["_INV", "_BASE"]):
            wms_qty[(inv, base)] = wms_qty.get((inv, base), 0.0) + float(grp["_QTY"].sum())

    # the register's own quantities — only needed for the no-WMS fallback
    doc_qty_by_key: dict[tuple[str, str], float] = {}
    if not use_wms and details is not None and len(details):
        for _, r in details.iterrows():
            key = (str(r["TAX_INVOICE_NO"]).strip(), base_item(r["ITEM_CODE"]))
            doc_qty_by_key[key] = doc_qty_by_key.get(key, 0.0) + \
                (_to_num(r.get("DOC_QTY")) or 0.0)

    # ---- one report row per uploaded sales-report row ----
    rep_rows: list[dict] = []
    matched_keys: set[tuple[str, str]] = set()
    n_matched = 0
    for _, r in sales_df.iterrows():
        inv = _id_str(r.get(c_inv, "")).strip()
        b1 = base_item(r.get(c_item)) if c_item else ""
        b2 = base_item(r.get(c_citem)) if c_citem else ""
        row = {"TAX_INVOICE_NO": inv,
               "CUSTOMER_ITEM": str(r.get(c_citem, "") or "") if c_citem else "",
               "ITEM_CODE": str(r.get(c_item, "") or "") if c_item else "",
               "BASE_ID": "", "SALES_QTY": _to_num(r.get(c_qty)) or 0.0,
               "IN_WMS_MASTER": "", "WMS_QTY": "", "STATUS": ""}

        if not inv:
            row["STATUS"] = "No Tax Invoice No."
            rep_rows.append(row)
            continue
        if not (b1 or b2):
            row["STATUS"] = "No item code"
            rep_rows.append(row)
            continue

        if use_wms:
            sent = inv in sent_invoices
            row["IN_WMS_MASTER"] = "Yes" if sent else "No"
            if not sent:
                row["STATUS"] = "Not in OUTBOUND_MASTER — not picked yet"
                rep_rows.append(row)
                continue
            # whichever base the WMS side actually carries
            hit = next((b for b in (b1, b2) if b and (inv, b) in wms_qty), "")
            row["BASE_ID"] = hit or (b1 or b2)
            if not hit:
                row["STATUS"] = "Not in OUTBOUND_DETAIL"
                rep_rows.append(row)
                continue
            wq = wms_qty[(inv, hit)]
            sq = sales_qty.get((inv, hit), 0.0)
            row["WMS_QTY"] = wq
            if abs(sq - wq) <= 0.01:
                row["STATUS"] = "Matched"
                matched_keys.add((inv, hit))
                n_matched += 1
            else:
                row["STATUS"] = f"Qty mismatch (sales {sq:g} vs WMS {wq:g})"
        else:
            hit = next((b for b in (b1, b2) if b and (inv, b) in doc_qty_by_key), "")
            row["BASE_ID"] = hit or (b1 or b2)
            if not hit:
                row["STATUS"] = "Not in the invoice register"
                rep_rows.append(row)
                continue
            dq = doc_qty_by_key[(inv, hit)]
            sq = sales_qty.get((inv, hit), 0.0)
            row["WMS_QTY"] = dq
            if abs(sq - dq) <= 0.01:
                row["STATUS"] = "Matched (register qty — no WMS output yet)"
                matched_keys.add((inv, hit))
                n_matched += 1
            else:
                row["STATUS"] = f"Qty mismatch (sales {sq:g} vs register {dq:g})"
        rep_rows.append(row)

    # ---- carry the confirmation into the register ----
    d = details
    s = summary
    n_inv_done = 0
    if details is not None and len(details) and matched_keys:
        d = details.copy()
        # _id_str on both sides: matched_keys was built with it, so plain
        # astype(str) here would miss every invoice number Excel handed back
        # as a float.
        d_inv = d["TAX_INVOICE_NO"].map(_id_str).str.strip()
        keys = list(zip(d_inv, d["ITEM_CODE"].map(base_item)))
        hit_mask = pd.Series([k in matched_keys for k in keys], index=d.index)
        d.loc[hit_mask & (d["PICKING"] != STATUS_DONE), "PICKING"] = STATUS_DONE

        if summary is not None and len(summary):
            s = summary.copy()
            all_done = d.groupby(d_inv)["PICKING"].apply(
                lambda x: bool(len(x)) and (x == STATUS_DONE).all())
            inv_col = s["TAX_INVOICE_NO"].map(_id_str).str.strip()
            for inv, done in all_done.items():
                if not done:
                    continue
                m = (inv_col == str(inv).strip()) & (s["PICKING"] != STATUS_DONE)
                n_inv_done += int(m.sum())
                s.loc[m, "PICKING"] = STATUS_DONE

    return {"summary": s, "details": d,
            "report": pd.DataFrame(rep_rows, columns=SALES_REPORT_COLS),
            "matched": n_matched, "invoices_completed": n_inv_done, "error": "",
            "used_wms": use_wms}


# --------------------------------------------------------------------------- #
# dashboard — how much is still left
# --------------------------------------------------------------------------- #
PENDING_COLS = ["DAYS", "TAX_INVOICE_DATE", "TAX_INVOICE_NO", "CUSTOMER_NAME", "QTY",
                "SHORT_QTY", "REASON", "DOC_TYPE", "MRP", "REMARK"]

# order matters — the first pattern that matches a remark wins
_REASONS: list[tuple[str, str]] = [
    ("Partially picked", r"partial pick"),
    ("On another pick task", r"another pick task|on pick task"),
    ("Stock short", r"stock short"),
    ("Pallet over-pick", r"over-pick"),
    ("Qty verify failed", r"qty verify"),
    ("Document incomplete", r"incomplete document"),
    ("Load deleted", r"load deleted"),
]


# Tried in order, each one only on what is still unparsed. Order matters:
#   * "12-AUG-2026" is what the Donaldson PDFs carry;
#   * ISO must beat dayfirst, or "2026-08-01" is read as 8 January;
#   * day-first slashes come last, because this is an Indian ERP — "01/08/2026"
#     is 1 August, never 8 January.
_DATE_FORMATS = ("%d-%b-%Y", "%d-%B-%Y", "%d/%b/%Y", "ISO8601",
                 "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y/%m/%d")


def parse_dates(s: pd.Series) -> pd.Series:
    """
    A whole date column in one pass — every row read by the same rules.

    Not `pd.to_datetime(col, dayfirst=True)`: pandas infers a single format from
    the first value and NaTs every row that does not share it, so one ISO date
    at the top of the register silently blanked every `01/08/2026` below it.
    Not `.map()` either — that is 306 ms on a 5 000-row register on every rerun
    of the Dashboard. So: vectorised, but one explicit format at a time.
    """
    txt = pd.Series(s, dtype="object").astype("string")
    todo = txt.notna() & (txt.str.strip() != "")
    out = pd.Series(pd.NaT, index=txt.index, dtype="datetime64[ns]")
    for fmt in _DATE_FORMATS:
        if not todo.any():
            return out
        got = pd.to_datetime(txt[todo], format=fmt, errors="coerce")
        hit = got.notna()
        if hit.any():
            out.loc[got.index[hit]] = got[hit]
            todo = todo & ~todo.index.isin(got.index[hit])
    if todo.any():                       # anything left: let pandas guess, row by row
        got = pd.to_datetime(txt[todo], dayfirst=True, errors="coerce", format="mixed")
        out.loc[got.index] = got
    return out


def parse_date(v: Any) -> Any:
    """One value, by exactly the same rules as the column — never diverge."""
    return parse_dates(pd.Series([v])).iloc[0]


# --------------------------------------------------------------------------- #
# Date filtering — "today", "this week", a range, or everything
# --------------------------------------------------------------------------- #
# Which date the question is about. They answer different questions and the
# right one depends on what is being asked:
#   Invoice date  — "what has the customer been invoiced for in this period"
#   Picked date   — "what did the warehouse actually pick today"
#   Last updated  — "what moved at all today" (a scan, a status report, a pick)
DATE_BASES = {"Invoice date": "TAX_INVOICE_DATE",
              "Picked date": "PICKED_AT",
              "Last updated": "UPDATED_AT"}

# A blank invoice date must not hide a row — the date simply could not be read
# off the document, and the invoice is still outstanding work. A blank picked /
# updated date is different: it means the thing being asked about never
# happened, so the row genuinely falls outside the window.
KEEP_UNDATED = {"TAX_INVOICE_DATE"}

DATE_PRESETS = ["Today", "Yesterday", "Last 7 days", "Last 30 days",
                "This month", "All time", "Custom range"]


def date_preset(name: str, today: Any = None) -> tuple[Any, Any]:
    """A preset name -> (from, to) as dates. `All time` is (None, None)."""
    t = pd.Timestamp(today).date() if today is not None else datetime.now().date()
    if name == "Today":
        return t, t
    if name == "Yesterday":
        y = t - timedelta(days=1)
        return y, y
    if name == "Last 7 days":
        return t - timedelta(days=6), t
    if name == "Last 30 days":
        return t - timedelta(days=29), t
    if name == "This month":
        return t.replace(day=1), t
    return None, None                       # All time / Custom range


def filter_by_date(df: pd.DataFrame, date_from: Any = None, date_to: Any = None,
                   column: str = "TAX_INVOICE_DATE") -> pd.DataFrame:
    """
    Rows inside a date window, read off `column`.

    The same rule drives the Dashboard and the Register, so a range that says
    "42 invoices" on one tab says 42 on the other.
    """
    if df is None or not len(df) or (date_from is None and date_to is None):
        return df
    col = column if column in df.columns else "TAX_INVOICE_DATE"
    if col not in df.columns:
        return df
    d = parse_dates(df[col])
    keep = pd.Series(True, index=df.index)
    if date_from is not None:
        keep &= d >= pd.Timestamp(date_from)
    if date_to is not None:
        # to the end of that day, or a same-day range would match nothing
        keep &= d <= pd.Timestamp(date_to) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
    if col in KEEP_UNDATED:
        keep |= d.isna()
    return df[keep]


def classify(remark: Any) -> str:
    t = str(remark or "").lower()
    if not t.strip():
        return "Not picked yet"
    for name, pat in _REASONS:
        if re.search(pat, t):
            return name
    return "Other"


def dashboard(summary: pd.DataFrame, date_from: Any = None, date_to: Any = None,
              doc_types: list[str] | None = None,
              date_col: str = "TAX_INVOICE_DATE") -> dict[str, Any]:
    """
    Everything the Pending vs picked view needs, in one pass.

    The question is always "how much is still left" — so the numbers are counted
    both ways: invoices and quantity, because one pending invoice for 1 000 units
    is not the same problem as ten pending invoices for 2 units each.
    """
    empty = {"kpi": {k: 0 for k in ("total", "picked", "pending", "partial", "qty",
                                    "qty_picked", "qty_pending", "qty_owed", "pct",
                                    "oldest", "picking_done", "packing_done",
                                    "dispatch_done")},
             "by_reason": pd.DataFrame(columns=["REASON", "INVOICES", "QTY"]),
             "by_customer": pd.DataFrame(columns=["CUSTOMER_NAME", "INVOICES", "QTY"]),
             "pending": pd.DataFrame(columns=PENDING_COLS),
             "picked": pd.DataFrame(columns=PENDING_COLS),
             "summary": pd.DataFrame(columns=SUMMARY_COLS), "invoices": [],
             "duplicates": 0}
    if summary is None or not len(summary):
        return empty

    d = summary.reindex(columns=SUMMARY_COLS).copy()
    d["QTY"] = pd.to_numeric(d["QTY"], errors="coerce").fillna(0.0)
    d["PICKED_QTY"] = pd.to_numeric(d["PICKED_QTY"], errors="coerce").fillna(0.0)
    d["_DATE"] = parse_dates(d["TAX_INVOICE_DATE"])
    d["KORBER_PICK"] = d["KORBER_PICK"].astype(str).str.strip().str.title()

    if doc_types:
        d = d[d["DOC_TYPE"].astype(str).isin(doc_types)]
    d = filter_by_date(d, date_from, date_to, date_col)
    # Legacy rows written before duplicates were excluded — out of every number.
    dup = is_duplicate_row(d)
    n_dup = int(dup.sum())
    d = d[~dup]
    if not len(d):
        return {**empty, "duplicates": n_dup}

    today = pd.Timestamp(datetime.now().date())
    d["DAYS"] = (today - d["_DATE"]).dt.days
    d["REASON"] = d["REMARK"].map(classify)
    d.loc[d["KORBER_PICK"] == PICK_YES, "REASON"] = "Picked"
    d.loc[d["KORBER_PICK"] == PICK_PART, "REASON"] = "Partially picked"
    d["SHORT_QTY"] = (d["QTY"] - d["PICKED_QTY"]).clip(lower=0)

    # A partly picked invoice is not finished, so it stays on the pending side —
    # but only the balance is still owed, and that is what the floor has to
    # chase. qty_owed answers "how much is actually left to pick".
    part = d[d["KORBER_PICK"] == PICK_PART].copy()
    pend = d[d["KORBER_PICK"] != PICK_YES].copy()
    done = d[d["KORBER_PICK"] == PICK_YES].copy()

    kpi = {
        "total": int(len(d)), "picked": int(len(done)), "pending": int(len(pend)),
        "partial": int(len(part)),
        "qty": float(d["QTY"].sum()), "qty_picked": float(done["QTY"].sum()),
        "qty_pending": float(pend["QTY"].sum()),
        "qty_owed": float(pend["SHORT_QTY"].sum()),
        "pct": (float(len(done)) / len(d) * 100.0) if len(d) else 0.0,
        "oldest": int(pend["DAYS"].max()) if len(pend) and pend["DAYS"].notna().any() else 0,
        "picking_done": int((d.get("PICKING", "") == STATUS_DONE).sum()),
        "packing_done": int((d.get("PACKING", "") == STATUS_DONE).sum()),
        "dispatch_done": int((d.get("DISPATCH", "") == STATUS_DONE).sum()),
    }

    by_reason = (pend.groupby("REASON")
                 .agg(INVOICES=("TAX_INVOICE_NO", "count"), QTY=("QTY", "sum"))
                 .reset_index().sort_values("QTY", ascending=False)
                 if len(pend) else empty["by_reason"])
    by_customer = (pend.groupby("CUSTOMER_NAME")
                   .agg(INVOICES=("TAX_INVOICE_NO", "count"), QTY=("QTY", "sum"))
                   .reset_index().sort_values("QTY", ascending=False)
                   if len(pend) else empty["by_customer"])

    def _tidy(x: pd.DataFrame) -> pd.DataFrame:
        if not len(x):
            return pd.DataFrame(columns=PENDING_COLS)
        return (x.sort_values(["DAYS", "QTY"], ascending=[False, False])
                 .reindex(columns=PENDING_COLS).reset_index(drop=True))

    return {"kpi": kpi, "by_reason": by_reason, "by_customer": by_customer,
            "pending": _tidy(pend), "picked": _tidy(done),
            "summary": d.reindex(columns=SUMMARY_COLS).reset_index(drop=True),
            "invoices": [_id_str(x) for x in d["TAX_INVOICE_NO"]],
            "duplicates": n_dup}


def details_for(details: pd.DataFrame, invoices: list[str]) -> pd.DataFrame:
    """Detail lines belonging to a set of invoices — keeps the two reports in step."""
    if details is None or not len(details):
        return pd.DataFrame(columns=DETAIL_COLS)
    keep = {_id_str(x) for x in (invoices or [])}
    d = details.reindex(columns=DETAIL_COLS)
    return d[d["TAX_INVOICE_NO"].map(_id_str).isin(keep)].reset_index(drop=True)


def pending_excel(dash: dict[str, Any]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine=_XL_ENGINE) as xw:
        k = dash["kpi"]
        pd.DataFrame([
            {"Measure": "Invoices total", "Value": k["total"]},
            {"Measure": "Picked", "Value": k["picked"]},
            {"Measure": "Pending", "Value": k["pending"]},
            {"Measure": "Qty total", "Value": k["qty"]},
            {"Measure": "Qty picked", "Value": k["qty_picked"]},
            {"Measure": "Qty pending", "Value": k["qty_pending"]},
            {"Measure": "Complete %", "Value": round(k["pct"], 1)},
            {"Measure": "Oldest pending (days)", "Value": k["oldest"]},
        ]).to_excel(xw, sheet_name="Status", index=False)
        for name, key in [("Pending", "pending"), ("By reason", "by_reason"),
                          ("By customer", "by_customer"), ("Picked", "picked")]:
            df = dash.get(key)
            (df if df is not None and len(df) else pd.DataFrame({"info": ["- none -"]})) \
                .to_excel(xw, sheet_name=name, index=False)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# excel
# --------------------------------------------------------------------------- #
# xlsxwriter writes the same sheet in less than half the time and a third of the
# bytes (30 000 rows: 9.8 s / 1.9 MB -> 4.3 s / 0.6 MB). It is optional, so fall
# back to openpyxl — which is always present for reading — when it is missing.
try:                                     # noqa: SIM105
    import xlsxwriter as _xlsxwriter     # noqa: F401
    _XL_ENGINE = "xlsxwriter"
except Exception:                        # pragma: no cover
    _XL_ENGINE = "openpyxl"

# Column widths are cosmetic; measuring every row of a 30 000-row export costs
# more than it is worth, and the widest of the first few hundred is just as good.
_WIDTH_SAMPLE = 400


def _col_width(col: pd.Series, name: str) -> int:
    if not len(col):
        return len(name) + 2
    longest = int(col.head(_WIDTH_SAMPLE).astype(str).str.len().max() or 0)
    return max(len(name) + 2, min(38, longest + 2))


def to_excel(df: pd.DataFrame, sheet: str, cols: list[str] | None = None) -> bytes:
    buf = io.BytesIO()
    d = (df if df is not None and len(df) else pd.DataFrame(columns=cols or []))
    if cols:
        d = d.reindex(columns=cols)
    name = sheet[:31]
    with pd.ExcelWriter(buf, engine=_XL_ENGINE) as xw:
        d.to_excel(xw, sheet_name=name, index=False)
        ws = xw.sheets[name]
        widths = [_col_width(d[c], str(c)) for c in d.columns]
        if _XL_ENGINE == "xlsxwriter":
            for i, w in enumerate(widths):
                ws.set_column(i, i, w)
            ws.freeze_panes(1, 0)
            if len(d.columns):
                ws.autofilter(0, 0, max(len(d), 1), len(d.columns) - 1)
        else:
            for i, w in enumerate(widths, start=1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


def summary_excel(df: pd.DataFrame) -> bytes:
    return to_excel(df, "Invoice Summary", SUMMARY_COLS)


def details_excel(df: pd.DataFrame) -> bytes:
    return to_excel(df, "Invoice Details", DETAIL_COLS)


def sales_reconciliation_excel(df: pd.DataFrame) -> bytes:
    return to_excel(df, "Reconciliation", SALES_REPORT_COLS)


def live_status_excel(df: pd.DataFrame) -> bytes:
    return to_excel(df, "Pick Live Status", LIVE_STATUS_REPORT_COLS)


def search(df: pd.DataFrame, query: str) -> pd.DataFrame:
    """Item codes match on the base id too, like everywhere else in the app."""
    if df is None or not len(df) or not str(query).strip():
        return df if df is not None else pd.DataFrame()
    joined = None
    for c in df.columns:
        col = df[c].astype("string").fillna("").str.lower()
        joined = col if joined is None else joined.str.cat(col, sep=" | ")
    mask = pd.Series(True, index=df.index)
    for term in str(query).split():
        t = term.strip().lower()
        if not t:
            continue
        m = joined.str.contains(re.escape(t), na=False)
        b = base_item(t).lower()
        if b and b != t:
            m |= joined.str.contains(re.escape(b), na=False)
        c = clean_item(t).lower()
        if c and c != t:
            m |= joined.str.contains(re.escape(c), na=False)
        mask &= m
    return df[mask].reset_index(drop=True)
