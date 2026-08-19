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
from datetime import datetime
from typing import Any

import pandas as pd

from doc_parser import ParsedDoc, base_item, clean_item

# Bumped whenever this module's public surface changes; app.py refuses to run
# against a stale copy instead of dying with a redacted TypeError.
API = 11

# Warehouse execution status — independent of KORBER_PICK (this app's own pallet
# allocation). These three track the physical pick / pack / dispatch, driven by
# the Pick_Live_status report, the packing QR scan, and the sales report match.
STATUS_COLS = ["PICKING", "PACKING", "DISPATCH"]
STATUS_PENDING = "Pending"
STATUS_DONE = "Completed"

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
        remark = "" if ok else reasons.get(num, "Not picked")
        doc_qty = float(sum(l.qty for l in doc.lines))

        a = alloc[alloc["DOC_NUMBER"].astype(str) == num] if len(alloc) else pd.DataFrame()
        picked_qty = float(pd.to_numeric(a["QTY_PICKED"], errors="coerce").sum()) \
            if len(a) else 0.0
        plants = ", ".join(sorted({str(x) for x in a["PLANT"]})) if len(a) else plant

        srows.append({
            "TAX_INVOICE_DATE": doc.doc_date, "TAX_INVOICE_NO": num,
            "AR_INVOICE_NO": doc.ref_number if is_inv else "",
            "CUSTOMER_NAME": doc.customer, "QTY": doc_qty,
            "KORBER_PICK": "Yes" if ok else "No", "REMARK": remark,
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
            drows.append({
                "TAX_INVOICE_DATE": doc.doc_date, "TAX_INVOICE_NO": num,
                "AR_INVOICE_NO": doc.ref_number if is_inv else "",
                "CUSTOMER_NAME": doc.customer, "DOC_TYPE": doc.doc_type,
                "LINE": ln.line_no, "ITEM_CODE": ln.item_code, "BASE_ID": ln.base,
                "DESCRIPTION": ln.description, "DOC_QTY": float(ln.qty),
                "UOM": ln.uom, "KORBER_PICK": "Yes" if ok else "No",
                "PICKED_QTY": float(pd.to_numeric(la["QTY_PICKED"],
                                                  errors="coerce").sum())
                if len(la) else 0.0,
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
    nums = [str(x) for x in summary.loc[bad, "TAX_INVOICE_NO"]]
    out_s = summary[~bad].reset_index(drop=True)
    out_d = details
    if details is not None and len(details) and nums:
        out_d = details[~details["TAX_INVOICE_NO"].astype(str).isin(set(nums))] \
            .reset_index(drop=True)
    return {"summary": out_s, "details": out_d, "dropped": int(bad.sum()),
            "invoices": nums}


def merge_summary(old: pd.DataFrame | None, new: pd.DataFrame) -> dict[str, Any]:
    """
    One row per invoice, updated in place.

    A `No` that is picked later becomes `Yes` and loses its remark. A `Yes` is
    not pushed back to `No` by a later run that happened to skip it as a
    duplicate — only deleting the load does that (`mark_unpicked`).
    """
    cols = SUMMARY_COLS
    if new is not None and len(new):          # belt and braces — never store one
        new = new[~is_duplicate_row(new)].reset_index(drop=True)
    if old is None or not len(old):
        return {"data": new.reindex(columns=cols), "new": len(new), "updated": 0,
                "picked_now": [n for n, k in zip(new["TAX_INVOICE_NO"], new["KORBER_PICK"])
                               if k == "Yes"]}

    o = old.reindex(columns=cols).copy()
    o["TAX_INVOICE_NO"] = o["TAX_INVOICE_NO"].astype(str)
    o = o.drop_duplicates(subset=["TAX_INVOICE_NO"], keep="last")
    idx = {str(v): i for i, v in enumerate(o["TAX_INVOICE_NO"])}
    rows = o.to_dict("records")

    n_new = n_upd = 0
    picked_now: list[str] = []
    for _, r in new.iterrows():
        num = str(r["TAX_INVOICE_NO"])
        rec = dict(r)
        if num not in idx:
            rows.append(rec)
            n_new += 1
            if rec["KORBER_PICK"] == "Yes":
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
        if was == "Yes" and rec["KORBER_PICK"] != "Yes":
            # already picked — keep the pick, just refresh what the document says
            for k in ("KORBER_PICK", "REMARK", "RUN_ID", "PICKED_AT", "PICKED_QTY",
                      "PLANT"):
                rec[k] = cur.get(k, "")
        elif was != "Yes" and rec["KORBER_PICK"] == "Yes":
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
    new = new.reindex(columns=DETAIL_COLS).copy()
    if old is None or not len(old):
        return new
    nums = {str(x) for x in new["TAX_INVOICE_NO"]}
    old = old.reindex(columns=DETAIL_COLS)
    touched = old["TAX_INVOICE_NO"].astype(str).isin(nums)
    prior = old[touched]
    if len(prior):
        # Carry every Completed status forward by an indexed lookup rather than
        # a per-row scan — .iterrows() over the prior rows cost 136 ms on a
        # 30 000-line register, and this runs inside the save lock.
        pkey = (prior["TAX_INVOICE_NO"].astype(str) + "|"
                + prior["LINE"].astype(str))
        nkey = (new["TAX_INVOICE_NO"].astype(str) + "|"
                + new["LINE"].astype(str))
        for k in STATUS_COLS:
            done = pkey[prior[k].astype(str) == STATUS_DONE]
            if len(done):
                new.loc[nkey.isin(set(done)), k] = STATUS_DONE
    return pd.concat([old[~touched], new], ignore_index=True)


def mark_unpicked(summary: pd.DataFrame, invoice_no: str,
                  remark: str = "Load deleted") -> pd.DataFrame:
    """Deleting a load frees the stock, so the register has to say No again."""
    if summary is None or not len(summary):
        return summary
    d = summary.copy()
    m = d["TAX_INVOICE_NO"].astype(str) == str(invoice_no).strip()
    if not m.any():
        return d
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    d.loc[m, "KORBER_PICK"] = "No"
    d.loc[m, "REMARK"] = f"{remark} · {stamp}"
    for c in ("RUN_ID", "PICKED_AT"):
        d.loc[m, c] = ""
    d.loc[m, "PICKED_QTY"] = 0.0
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

    known = {str(x).strip() for x in summary["TAX_INVOICE_NO"]}
    for lid in sorted(set(pick_set) | set(disp_set)):
        rows.append({"LOAD_ID": lid, "MATCHED": lid in known,
                    "PICKING": STATUS_DONE if pick_set.get(lid) else "",
                    "DISPATCH": STATUS_DONE if disp_set.get(lid) else ""})

    s = summary.copy()
    inv = s["TAX_INVOICE_NO"].astype(str).str.strip()
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
        dinv = d["TAX_INVOICE_NO"].astype(str).str.strip()
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
    out = {"found": False, "already": False, "invoice": None, "column": column,
          "summary": summary, "details": details}
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

    if details is not None and len(details):
        d = details.copy()
        dm = d["TAX_INVOICE_NO"].map(_id_str).str.strip() == lid
        d.loc[dm, column] = STATUS_DONE
        d.loc[dm, "UPDATED_AT"] = stamp
        out["details"] = d
    return out


def apply_packing_scan(summary: pd.DataFrame, details: pd.DataFrame, load_id: str,
                       user: str = "") -> dict[str, Any]:
    """OUTBOUND PICK SHEET LOAD_ID -> Packing = Completed."""
    return apply_status_scan(summary, details, load_id, "PACKING", user)


def apply_dispatch_scan(summary: pd.DataFrame, details: pd.DataFrame, load_id: str,
                        user: str = "") -> dict[str, Any]:
    """OUTBOUND PICK SHEET LOAD_ID -> Dispatch = Completed."""
    return apply_status_scan(summary, details, load_id, "DISPATCH", user)


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
        keys = list(zip(d["TAX_INVOICE_NO"].astype(str).str.strip(),
                        d["ITEM_CODE"].map(base_item)))
        hit_mask = pd.Series([k in matched_keys for k in keys], index=d.index)
        d.loc[hit_mask & (d["PICKING"] != STATUS_DONE), "PICKING"] = STATUS_DONE

        if summary is not None and len(summary):
            s = summary.copy()
            all_done = d.groupby("TAX_INVOICE_NO")["PICKING"].apply(
                lambda x: bool(len(x)) and (x == STATUS_DONE).all())
            inv_col = s["TAX_INVOICE_NO"].astype(str).str.strip()
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
    ("On another pick task", r"another pick task|on pick task"),
    ("Stock short", r"stock short"),
    ("Pallet over-pick", r"over-pick"),
    ("Qty verify failed", r"qty verify"),
    ("Document incomplete", r"incomplete document"),
    ("Load deleted", r"load deleted"),
]


def parse_date(v: Any) -> Any:
    """'12-AUG-2026' first, anything else after."""
    d = pd.to_datetime(v, format="%d-%b-%Y", errors="coerce")
    if pd.isna(d):
        d = pd.to_datetime(v, dayfirst=True, errors="coerce")
    return d


def parse_dates(s: pd.Series) -> pd.Series:
    """
    Same rules as `parse_date`, one pass over the column.

    `.map(parse_date)` calls pd.to_datetime once per row — 306 ms on a 5 000-row
    register, on every single rerun of the Dashboard. Vectorised it is ~2 ms.
    """
    txt = s.astype("string")
    out = pd.to_datetime(txt, format="%d-%b-%Y", errors="coerce")
    rest = out.isna() & txt.notna() & (txt.str.strip() != "")
    if rest.any():
        out.loc[rest] = pd.to_datetime(txt[rest], dayfirst=True, errors="coerce")
    return out


def classify(remark: Any) -> str:
    t = str(remark or "").lower()
    if not t.strip():
        return "Not picked yet"
    for name, pat in _REASONS:
        if re.search(pat, t):
            return name
    return "Other"


def dashboard(summary: pd.DataFrame, date_from: Any = None, date_to: Any = None,
              doc_types: list[str] | None = None) -> dict[str, Any]:
    """
    Everything the Pending vs picked view needs, in one pass.

    The question is always "how much is still left" — so the numbers are counted
    both ways: invoices and quantity, because one pending invoice for 1 000 units
    is not the same problem as ten pending invoices for 2 units each.
    """
    empty = {"kpi": {k: 0 for k in ("total", "picked", "pending", "qty", "qty_picked",
                                    "qty_pending", "pct", "oldest", "picking_done",
                                    "packing_done", "dispatch_done")},
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
    if date_from is not None:
        d = d[d["_DATE"].isna() | (d["_DATE"] >= pd.Timestamp(date_from))]
    if date_to is not None:
        d = d[d["_DATE"].isna() | (d["_DATE"] <= pd.Timestamp(date_to))]
    # Legacy rows written before duplicates were excluded — out of every number.
    dup = is_duplicate_row(d)
    n_dup = int(dup.sum())
    d = d[~dup]
    if not len(d):
        return {**empty, "duplicates": n_dup}

    today = pd.Timestamp(datetime.now().date())
    d["DAYS"] = (today - d["_DATE"]).dt.days
    d["REASON"] = d["REMARK"].map(classify)
    d.loc[d["KORBER_PICK"] == "Yes", "REASON"] = "Picked"
    d["SHORT_QTY"] = (d["QTY"] - d["PICKED_QTY"]).clip(lower=0)

    pend = d[d["KORBER_PICK"] != "Yes"].copy()
    done = d[d["KORBER_PICK"] == "Yes"].copy()

    kpi = {
        "total": int(len(d)), "picked": int(len(done)), "pending": int(len(pend)),
        "qty": float(d["QTY"].sum()), "qty_picked": float(done["QTY"].sum()),
        "qty_pending": float(pend["QTY"].sum()),
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
            "invoices": [str(x) for x in d["TAX_INVOICE_NO"]],
            "duplicates": n_dup}


def details_for(details: pd.DataFrame, invoices: list[str]) -> pd.DataFrame:
    """Detail lines belonging to a set of invoices — keeps the two reports in step."""
    if details is None or not len(details):
        return pd.DataFrame(columns=DETAIL_COLS)
    keep = {str(x) for x in (invoices or [])}
    d = details.reindex(columns=DETAIL_COLS)
    return d[d["TAX_INVOICE_NO"].astype(str).isin(keep)].reset_index(drop=True)


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
