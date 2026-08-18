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
API = 1

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
            reasons[num] = (f"{why} — {det}" if det else why)[:400]

    alloc = res.get("allocations")
    alloc = alloc if alloc is not None and len(alloc) else pd.DataFrame()

    srows: list[dict] = []
    drows: list[dict] = []
    seen: set[str] = set()

    for doc in docs:
        num = str(doc.doc_number).strip()
        if not num or num in seen:
            continue
        seen.add(num)
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

    return (pd.DataFrame(srows, columns=SUMMARY_COLS),
            pd.DataFrame(drows, columns=DETAIL_COLS))


# --------------------------------------------------------------------------- #
# merge into what is already stored
# --------------------------------------------------------------------------- #
def merge_summary(old: pd.DataFrame | None, new: pd.DataFrame) -> dict[str, Any]:
    """
    One row per invoice, updated in place.

    A `No` that is picked later becomes `Yes` and loses its remark. A `Yes` is
    not pushed back to `No` by a later run that happened to skip it as a
    duplicate — only deleting the load does that (`mark_unpicked`).
    """
    cols = SUMMARY_COLS
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
