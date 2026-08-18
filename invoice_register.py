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
API = 5

SUMMARY_COLS = [
    "TAX_INVOICE_DATE", "TAX_INVOICE_NO", "AR_INVOICE_NO", "CUSTOMER_NAME", "QTY",
    "KORBER_PICK", "REMARK", "MRP",
    "DOC_TYPE", "CUSTOMER_CODE", "CONTACT_PERSON", "CONTACT_EMAIL", "LINES",
    "PICKED_QTY", "PLANT", "RUN_ID", "PICKED_AT", "FIRST_SEEN", "UPDATED_AT",
    "UPDATED_BY", "SOURCE_FILE",
]

DETAIL_COLS = [
    "TAX_INVOICE_DATE", "TAX_INVOICE_NO", "AR_INVOICE_NO", "CUSTOMER_NAME", "DOC_TYPE",
    "LINE", "ITEM_CODE", "BASE_ID", "DESCRIPTION", "DOC_QTY", "UOM",
    "KORBER_PICK", "PICKED_QTY", "ITEM_NUMBER", "LOT_NUMBER", "PALLETS", "LOCATIONS",
    "PLANT", "REMARK", "RUN_ID", "UPDATED_AT",
]

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
    """Detail rows are replaced wholesale for the invoices in this run."""
    if new is None or not len(new):
        return (old if old is not None else pd.DataFrame(columns=DETAIL_COLS))
    if old is None or not len(old):
        return new.reindex(columns=DETAIL_COLS)
    nums = {str(x) for x in new["TAX_INVOICE_NO"]}
    keep = old.reindex(columns=DETAIL_COLS)
    keep = keep[~keep["TAX_INVOICE_NO"].astype(str).isin(nums)]
    return pd.concat([keep, new.reindex(columns=DETAIL_COLS)], ignore_index=True)


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
                                    "qty_pending", "pct", "oldest")},
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
    d["_DATE"] = d["TAX_INVOICE_DATE"].map(parse_date)
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
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
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
def to_excel(df: pd.DataFrame, sheet: str, cols: list[str] | None = None) -> bytes:
    buf = io.BytesIO()
    d = (df if df is not None and len(df) else pd.DataFrame(columns=cols or []))
    if cols:
        d = d.reindex(columns=cols)
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        d.to_excel(xw, sheet_name=sheet[:31], index=False)
        ws = xw.sheets[sheet[:31]]
        for i, c in enumerate(d.columns, start=1):
            width = max(len(str(c)) + 2,
                        min(38, int(d[c].astype(str).str.len().max() or 0) + 2)
                        if len(d) else len(str(c)) + 2)
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


def summary_excel(df: pd.DataFrame) -> bytes:
    return to_excel(df, "Invoice Summary", SUMMARY_COLS)


def details_excel(df: pd.DataFrame) -> bytes:
    return to_excel(df, "Invoice Details", DETAIL_COLS)


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
