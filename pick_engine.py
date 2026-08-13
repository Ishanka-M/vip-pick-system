"""
pick_engine.py — Inventory matching, pallet-level allocation, WMS output build
=============================================================================
Rules
  * Item match  : base ID only (P162400-000-140 -> P162400)
  * Plant       : user confirm කරපු plant(s) එකෙන් විතරයි
  * Whole doc   : Invoice / DC එකේ හැම line එකකටම stock තිබ්බොත් විතරයි pick
  * Duplicate   : එකම Invoice / DC number එකකට එක output එකයි
  * Pallet level: කොච්චර pick කරාද, balance කීයද — හැම pallet එකකටම save
"""
from __future__ import annotations

import io
import re
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from doc_parser import ParsedDoc, base_item, clean_item

# --------------------------------------------------------------------------- #
# WMS templates
# --------------------------------------------------------------------------- #
MASTER_COLS = [
    "HOST_ORDER_MASTER_ID", "HOST_GROUP_ID", "RECORD_CREATE_DATE", "PROCESSING_CODE",
    "WH_ID", "CLIENT_CODE", "ORDER_NUMBER", "DISPLAY_ORDER_NUMBER", "STORE_ORDER_NUMBER",
    "ORDER_TYPE", "CUSTOMER_CODE", "CUSTOMER_PO_NUMBER", "DEPARTMENT", "LOAD_ID",
    "LOAD_SEQ", "BOL_NUMBER", "MASTER_BOL_NUMBER", "PRO_NUMBER", "CARRIER", "CARRIER_SCAC",
    "FREIGHT_TERMS", "RUSH", "ORDER_DATE", "ARRIVE_DATE", "DATE_EXPECTED", "PROMISED_DATE",
    "WEIGHT", "CUBIC_VOLUME", "CONTAINERS", "BACKORDER", "PRE_PAID", "COD_AMOUNT",
    "INSURANCE_AMOUNT", "PIP_AMOUNT", "FREIGHT_COST", "SHIP_TO_CODE", "SHIP_TO_NAME",
    "SHIP_TO_ADDR1", "SHIP_TO_ADDR2", "SHIP_TO_ADDR3", "SHIP_TO_CITY", "SHIP_TO_STATE",
    "SHIP_TO_ZIP", "SHIP_TO_COUNTRY_CODE", "SHIP_TO_COUNTRY_NAME", "SHIP_TO_PHONE",
    "BILL_TO_CODE", "BILL_TO_NAME", "BILL_TO_ADDR1", "BILL_TO_ADDR2", "BILL_TO_ADDR3",
    "BILL_TO_CITY", "BILL_TO_STATE", "BILL_TO_ZIP", "BILL_TO_COUNTRY_CODE",
    "BILL_TO_COUNTRY_NAME", "BILL_TO_PHONE", "DELIVERY_NAME", "DELIVERY_ADDR1",
    "DELIVERY_ADDR2", "DELIVERY_ADDR3", "DELIVERY_CITY", "DELIVERY_STATE", "DELIVERY_ZIP",
    "DELIVERY_COUNTRY_CODE", "DELIVERY_COUNTRY_NAME", "DELIVERY_PHONE",
    "BILL_FRGHT_TO_CODE", "BILL_FRGHT_TO_NAME", "BILL_FRGHT_TO_ADDR1",
    "BILL_FRGHT_TO_ADDR2", "BILL_FRGHT_TO_ADDR3", "BILL_FRGHT_TO_CITY",
    "BILL_FRGHT_TO_STATE", "BILL_FRGHT_TO_ZIP", "BILL_FRGHT_TO_COUNTRY_CODE",
    "BILL_FRGHT_TO_COUNTRY_NAME", "BILL_FRGHT_TO_PHONE", "CARTON_LABEL", "VER_FLAG",
    "PARTIAL_ORDER_FLAG", "EARLIEST_SHIP_DATE", "LATEST_SHIP_DATE",
    "EARLIEST_DELIVERY_DATE", "LATEST_DELIVERY_DATE", "TEMP_LINK_ID", "SERVICE_LEVEL",
    "SHIP_VIA", "SHIP_TO_ATTENTION", "SAT_DELIVERY_FLAG", "REGISTERED_MAIL_FLAG",
    "RESTRICTED_MAIL_FLAG", "COD_FLAG", "COD_PAY_TYPE", "COD_OPTION", "INSURANCE_FLAG",
    "BILL_FRGHT_TO_ATTENTION", "SHIP_TO_RESIDENTIAL_FLAG", "CARRIER_MODE",
    "EARLIEST_APPT_TIME", "LATEST_APPT_TIME",
]

DETAIL_COLS = [
    "HOST_ORDER_DETAIL_ID", "HOST_ORDER_MASTER_ID", "HOST_GROUP_ID", "RECORD_CREATE_DATE",
    "PROCESSING_CODE", "WH_ID", "CLIENT_CODE", "ORDER_NUMBER", "DISPLAY_ORDER_NUMBER",
    "LINE_NUMBER", "ITEM_NUMBER", "DISPLAY_ITEM_NUMBER", "ITEM_DESCRIPTION", "CUST_PART",
    "LOT_NUMBER", "QTY", "UNIT_WEIGHT", "UNIT_VOLUME", "EXTENDED_WEIGHT",
    "EXTENDED_VOLUME", "HAZ_MATERIAL", "BOL_CLASS", "BOL_CODE", "ORDER_UOM",
    "HOST_WAVE_ID", "TEMP_LINK_ID", "UNIT_INSURANCE_AMOUNT",
    "GEN_ATTRIBUTE_VALUE1", "GEN_ATTRIBUTE_VALUE2", "GEN_ATTRIBUTE_VALUE3",
    "GEN_ATTRIBUTE_VALUE4", "GEN_ATTRIBUTE_VALUE5", "GEN_ATTRIBUTE_VALUE6",
    "GEN_ATTRIBUTE_VALUE7", "GEN_ATTRIBUTE_VALUE8", "GEN_ATTRIBUTE_VALUE9",
    "GEN_ATTRIBUTE_VALUE10", "GEN_ATTRIBUTE_VALUE11", "HOLD_REASON_ID", "PACKING_INST",
    "REIMA_LINE", "VAS_INST",
]

MASTER_FIXED = {
    "PROCESSING_CODE": "NEW", "ORDER_TYPE": "Sales Orders", "BACKORDER": "N",
    "PARTIAL_ORDER_FLAG": "N", "SAT_DELIVERY_FLAG": "N", "REGISTERED_MAIL_FLAG": "N",
    "RESTRICTED_MAIL_FLAG": "N", "COD_FLAG": "N", "COD_PAY_TYPE": "N", "COD_OPTION": "N",
    "INSURANCE_FLAG": "N", "SHIP_TO_RESIDENTIAL_FLAG": "N",
}
DETAIL_FIXED = {"PROCESSING_CODE": "NEW"}

# GEN_ATTRIBUTE_VALUEn  ->  Inventory column
GEN_MAP = {
    "GEN_ATTRIBUTE_VALUE1": "color",
    "GEN_ATTRIBUTE_VALUE2": "size",
    "GEN_ATTRIBUTE_VALUE3": "style",
    "GEN_ATTRIBUTE_VALUE4": "supplier",
    "GEN_ATTRIBUTE_VALUE5": "plant",
    "GEN_ATTRIBUTE_VALUE6": "client_so",
    "GEN_ATTRIBUTE_VALUE7": "client_so_line",
    "GEN_ATTRIBUTE_VALUE8": "po_cust_dec",
    "GEN_ATTRIBUTE_VALUE9": "customer_ref_number",
    "GEN_ATTRIBUTE_VALUE10": "item_id",
    "GEN_ATTRIBUTE_VALUE11": "invoice_number_1",
}

ALLOC_COLS = [
    "RUN_ID", "PICK_DATE", "DOC_TYPE", "DOC_NUMBER", "DOC_LINE", "DOC_ITEM_CODE",
    "BASE_ID", "ITEM_NUMBER", "DESCRIPTION", "PALLET", "LOCATION_ID", "LOT_NUMBER",
    "PLANT", "UOM", "QTY_BEFORE", "QTY_PICKED", "QTY_BALANCE", "FIFO_DATE",
    "GRN_NUMBER", "STORED_ATTRIBUTE_ID", "ROW_KEY", "SOURCE_FILE",
]

# --------------------------------------------------------------------------- #
# Inventory normalisation
# --------------------------------------------------------------------------- #
INV_ALIASES: dict[str, list[str]] = {
    "wh_id": ["wh id", "warehouse id"],
    "client_code": ["client code"],
    "pallet": ["pallet", "pallet id", "lpn"],
    "location_id": ["location id", "location"],
    "item_number": ["item number", "item no"],
    "display_item_number": ["display item number"],
    "description": ["description", "item description"],
    "lot_number": ["lot number", "lot"],
    "actual_qty": ["actual qty", "available qty", "qty", "quantity"],
    "unavailable_qty": ["unavailable qty", "unavailable"],
    "uom": ["uom", "unit of measure"],
    "status": ["status"],
    "stored_attribute_id": ["stored attribute id"],
    "fifo_date": ["fifo date"],
    "grn_number": ["grn number", "grn"],
    "cbm": ["cbm"],
    "color": ["color", "colour"],
    "size": ["size"],
    "style": ["style"],
    "supplier": ["supplier"],
    "plant": ["plant"],
    "client_so": ["client so"],
    "client_so_line": ["client so line"],
    "po_cust_dec": ["po cust dec"],
    "customer_ref_number": ["customer ref number"],
    "item_id": ["item id"],
    "invoice_number_1": ["invoice number 1"],
}


def _norm(s: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def normalize_inventory(df: pd.DataFrame) -> pd.DataFrame:
    """Inventory report -> canonical columns + numeric qty + base id."""
    src = {c: _norm(c) for c in df.columns}
    used: set[str] = set()
    out = pd.DataFrame(index=df.index)

    for canon, aliases in INV_ALIASES.items():
        found = None
        for a in aliases:                                   # exact match first
            na = _norm(a)
            for col, nc in src.items():
                if nc == na and col not in used:
                    found = col
                    break
            if found:
                break
        if found is None:                                   # then contains
            for a in aliases:
                na = _norm(a)
                for col, nc in src.items():
                    if na and na in nc and col not in used:
                        found = col
                        break
                if found:
                    break
        if found is not None:
            used.add(found)
            out[canon] = df[found]
        else:
            out[canon] = ""

    out["actual_qty"] = pd.to_numeric(out["actual_qty"], errors="coerce").fillna(0.0)
    out["unavailable_qty"] = pd.to_numeric(out["unavailable_qty"], errors="coerce").fillna(0.0)
    # raw = WMS එකේ තියෙන හරියටම item number (trailing '.' වගේ ඒවත් එහෙමම තියෙන්න ඕන)
    out["item_number_raw"] = out["item_number"].astype(str).str.strip().replace(
        {"nan": "", "None": ""})
    out["item_number"] = out["item_number"].map(clean_item)      # matching key
    out["base_id"] = out["item_number"].map(base_item)
    out["_fifo"] = pd.to_datetime(out["fifo_date"], errors="coerce", dayfirst=True)
    for c in ("pallet", "location_id", "lot_number", "plant", "uom", "status",
              "stored_attribute_id"):
        out[c] = out[c].astype(str).replace({"nan": "", "None": "", "NaT": ""}).str.strip()
    out["row_key"] = (
        out["pallet"] + "|" + out["location_id"] + "|" + out["item_number"] + "|"
        + out["lot_number"] + "|" + out["stored_attribute_id"].astype(str)
    )
    out["free_qty"] = (out["actual_qty"] - out["unavailable_qty"]).clip(lower=0)
    out = out[out["item_number"] != ""].reset_index(drop=True)
    return out


def plant_summary(inv: pd.DataFrame) -> pd.DataFrame:
    g = (inv.groupby("plant", dropna=False)
           .agg(Pallets=("pallet", "nunique"), Items=("item_number", "nunique"),
                Qty=("free_qty", "sum"))
           .reset_index().rename(columns={"plant": "Plant"}))
    g["Qty"] = g["Qty"].astype(int)
    return g.sort_values("Qty", ascending=False).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
@dataclass
class EngineConfig:
    wh_id: str = "INMM01"
    client_code: str = "INM0DONA"
    order_type: str = "Sales Orders"
    plants: list[str] = field(default_factory=list)
    statuses: list[str] = field(default_factory=lambda: ["Available"])
    strategy: str = "FIFO"              # FIFO | LEAST_PALLETS | SINGLE_PALLET_FIRST
    exact_item_first: bool = True
    blank_fill: str = "TBC"
    fill_item_number_col: bool = False
    merge_same_item_lines: bool = False
    override_doc_check: bool = False    # ⚠️ manual verify කරලා විතරක් — stock check bypass වෙන්නේ නෑ
    pick_date: datetime = field(default_factory=datetime.now)


# --------------------------------------------------------------------------- #
# Allocation
# --------------------------------------------------------------------------- #
def _order_pool(pool: pd.DataFrame, need: float, doc_item: str, cfg: EngineConfig) -> pd.DataFrame:
    p = pool.copy()
    p["_exact"] = (p["item_number"] == clean_item(doc_item)).astype(int)
    asc_exact = not cfg.exact_item_first

    if cfg.strategy == "LEAST_PALLETS":
        p = p.sort_values(["_exact", "avail", "_fifo"], ascending=[asc_exact, False, True])
    elif cfg.strategy == "SINGLE_PALLET_FIRST":
        p["_covers"] = (p["avail"] >= need).astype(int)
        p = p.sort_values(["_exact", "_covers", "avail", "_fifo"],
                          ascending=[asc_exact, False, True, True])
    else:                                                    # FIFO
        p = p.sort_values(["_exact", "_fifo", "avail"], ascending=[asc_exact, True, True])
    return p


def _detail_key(a: dict) -> tuple:
    return (a["ITEM_NUMBER"], a["LOT_NUMBER"], a["UOM"], a["PLANT"],
            a["_attrs"])


def run_pick(
    docs: list[ParsedDoc],
    inv_raw: pd.DataFrame,
    cfg: EngineConfig,
    ledger: pd.DataFrame | None = None,
    processed_docs: set[str] | None = None,
) -> dict[str, Any]:
    """මුළු pipeline එක — validate -> duplicate -> allocate -> WMS output."""
    run_id = datetime.now().strftime("%Y%m%d%H%M%S") + "-" + uuid.uuid4().hex[:4]
    stamp = cfg.pick_date.strftime("%Y-%m-%d %H:%M:%S")
    processed_docs = {str(x).strip() for x in (processed_docs or set())}

    inv = normalize_inventory(inv_raw)
    if cfg.statuses:
        inv = inv[inv["status"].isin(cfg.statuses)]
    if cfg.plants:
        inv = inv[inv["plant"].isin(cfg.plants)]

    # ---- remaining qty per inventory row (previous commitments subtracted) ---
    remaining: dict[str, float] = {}
    for _, r in inv.iterrows():
        remaining[r["row_key"]] = remaining.get(r["row_key"], 0.0) + float(r["free_qty"])

    if ledger is not None and len(ledger):
        led = ledger.copy()
        if "ROW_KEY" in led.columns and "QTY_PICKED" in led.columns:
            led["QTY_PICKED"] = pd.to_numeric(led["QTY_PICKED"], errors="coerce").fillna(0)
            for k, v in led.groupby("ROW_KEY")["QTY_PICKED"].sum().items():
                if k in remaining:
                    remaining[k] = max(0.0, remaining[k] - float(v))

    inv = inv.drop_duplicates(subset=["row_key"], keep="first").reset_index(drop=True)
    inv["avail"] = inv["row_key"].map(remaining).fillna(0.0)
    inv = inv[inv["avail"] > 0].reset_index(drop=True)

    allocations: list[dict] = []
    rejected: list[dict] = []
    shortages: list[dict] = []
    accepted: list[dict] = []
    detail_rows: list[dict] = []
    master_rows: list[dict] = []

    seen_batch: set[str] = set()

    for doc in docs:
        num = str(doc.doc_number).strip()

        # ---------- duplicate ----------
        if num in seen_batch:
            rejected.append({"DOC_NUMBER": num, "DOC_TYPE": doc.doc_type,
                             "REASON": "DUPLICATE (batch)",
                             "DETAIL": "මේ upload එකේම එකම number එක ආපහු — එකක් විතරක් ගත්තා",
                             "SOURCE_FILE": doc.source_file})
            continue
        if num in processed_docs:
            rejected.append({"DOC_NUMBER": num, "DOC_TYPE": doc.doc_type,
                             "REASON": "DUPLICATE (already processed)",
                             "DETAIL": "කලින් run එකක pick කරලා තියෙනවා",
                             "SOURCE_FILE": doc.source_file})
            continue
        seen_batch.add(num)

        # ---------- document completeness ----------
        ok, problems = doc.completeness()
        doc_check = "OK"
        if not ok:
            if not cfg.override_doc_check:
                rejected.append({"DOC_NUMBER": num, "DOC_TYPE": doc.doc_type,
                                 "REASON": "INCOMPLETE DOCUMENT",
                                 "DETAIL": " · ".join(problems),
                                 "SOURCE_FILE": doc.source_file})
                continue
            doc_check = "MANUAL OVERRIDE — " + " · ".join(problems)

        # ---------- dry-run allocation (all-or-nothing) ----------
        trial = {k: v for k, v in remaining.items()}
        doc_alloc: list[dict] = []
        doc_short: list[dict] = []

        for ln in doc.lines:
            need = float(ln.qty)
            pool = inv[inv["base_id"] == ln.base].copy()
            pool["avail"] = pool["row_key"].map(trial).fillna(0.0)
            pool = pool[pool["avail"] > 0]
            have = float(pool["avail"].sum())

            if pool.empty or have + 1e-9 < need:
                doc_short.append({
                    "DOC_NUMBER": num, "DOC_TYPE": doc.doc_type, "DOC_LINE": ln.line_no,
                    "DOC_ITEM_CODE": ln.item_code, "BASE_ID": ln.base,
                    "DESCRIPTION": ln.description, "REQUIRED": need,
                    "AVAILABLE": have, "SHORT": max(0.0, need - have),
                    "REASON": "Item not in inventory/plant" if pool.empty else "Stock short",
                })
                continue

            for _, r in _order_pool(pool, need, ln.item_code, cfg).iterrows():
                if need <= 1e-9:
                    break
                take = min(need, float(trial.get(r["row_key"], 0.0)))
                if take <= 0:
                    continue
                before = float(trial.get(r["row_key"], 0.0))
                trial[r["row_key"]] = before - take
                need -= take
                doc_alloc.append({
                    "RUN_ID": run_id, "PICK_DATE": stamp, "DOC_TYPE": doc.doc_type,
                    "DOC_NUMBER": num, "DOC_LINE": ln.line_no,
                    "DOC_ITEM_CODE": ln.item_code, "BASE_ID": ln.base,
                    "ITEM_NUMBER": r["item_number_raw"] or r["item_number"],
                    "DESCRIPTION": r["description"] or ln.description,
                    "PALLET": r["pallet"], "LOCATION_ID": r["location_id"],
                    "LOT_NUMBER": r["lot_number"], "PLANT": r["plant"],
                    "UOM": r["uom"] or ln.uom, "QTY_BEFORE": before,
                    "QTY_PICKED": take, "QTY_BALANCE": before - take,
                    "FIFO_DATE": r["fifo_date"], "GRN_NUMBER": r["grn_number"],
                    "STORED_ATTRIBUTE_ID": r["stored_attribute_id"],
                    "ROW_KEY": r["row_key"], "SOURCE_FILE": doc.source_file,
                    "_attrs": tuple(str(r.get(v, "") or "") for v in GEN_MAP.values()),
                })

        # ---------- whole document must be complete ----------
        if doc_short:
            shortages.extend(doc_short)
            miss = ", ".join(
                f"L{s['DOC_LINE']} {s['DOC_ITEM_CODE']} (need {s['REQUIRED']:g}, "
                f"have {s['AVAILABLE']:g})" for s in doc_short
            )
            rejected.append({"DOC_NUMBER": num, "DOC_TYPE": doc.doc_type,
                             "REASON": "STOCK SHORT — pick කළේ නෑ",
                             "DETAIL": miss, "SOURCE_FILE": doc.source_file})
            continue

        # ---------- commit ----------
        remaining = trial
        allocations.extend(doc_alloc)

        # ---------- WMS Detail (pallet allocations -> order lines) ----------
        groups: dict[tuple, dict] = {}
        order: list[tuple] = []
        for a in doc_alloc:
            key = _detail_key(a) if cfg.merge_same_item_lines else (a["DOC_LINE"],) + _detail_key(a)
            if key not in groups:
                groups[key] = {"qty": 0.0, "a": a}
                order.append(key)
            groups[key]["qty"] += a["QTY_PICKED"]

        line_no = 0
        for key in order:
            g = groups[key]
            a = g["a"]
            line_no += 1
            row = {c: "" for c in DETAIL_COLS}
            row.update(DETAIL_FIXED)
            row["WH_ID"] = cfg.wh_id
            row["CLIENT_CODE"] = cfg.client_code
            row["DISPLAY_ORDER_NUMBER"] = num
            row["LINE_NUMBER"] = str(line_no)
            row["DISPLAY_ITEM_NUMBER"] = a["ITEM_NUMBER"]
            if cfg.fill_item_number_col:
                row["ITEM_NUMBER"] = a["ITEM_NUMBER"]
            row["LOT_NUMBER"] = a["LOT_NUMBER"] or cfg.blank_fill
            row["QTY"] = _qty_str(g["qty"])
            row["ORDER_UOM"] = a["UOM"] or cfg.blank_fill
            for i, (gcol, invcol) in enumerate(GEN_MAP.items()):
                val = a["_attrs"][i]
                row[gcol] = val if str(val).strip() else cfg.blank_fill
            detail_rows.append(row)

        # ---------- WMS Master ----------
        m = {c: "" for c in MASTER_COLS}
        m.update(MASTER_FIXED)
        m["ORDER_TYPE"] = cfg.order_type
        m["WH_ID"] = cfg.wh_id
        m["CLIENT_CODE"] = cfg.client_code
        m["DISPLAY_ORDER_NUMBER"] = num
        m["STORE_ORDER_NUMBER"] = num
        m["CUSTOMER_PO_NUMBER"] = num
        m["LOAD_ID"] = num
        master_rows.append(m)

        accepted.append({
            "DOC_NUMBER": num, "DOC_TYPE": doc.doc_type, "DOC_DATE": doc.doc_date,
            "REF_NUMBER": doc.ref_number, "DOC_CHECK": doc_check, "LINES": len(doc.lines),
            "WMS_LINES": line_no, "DOC_QTY": sum(l.qty for l in doc.lines),
            "PICKED_QTY": sum(a["QTY_PICKED"] for a in doc_alloc),
            "PALLETS": len({a["PALLET"] for a in doc_alloc}),
            "PLANTS": ", ".join(sorted({a["PLANT"] for a in doc_alloc})),
            "SOURCE_FILE": doc.source_file,
        })

    alloc_df = pd.DataFrame(allocations)
    if len(alloc_df):
        alloc_df = alloc_df[ALLOC_COLS]

    master_df = pd.DataFrame(master_rows, columns=MASTER_COLS)
    detail_df = pd.DataFrame(detail_rows, columns=DETAIL_COLS)

    return {
        "run_id": run_id,
        "pick_date": stamp,
        "master": master_df,
        "detail": detail_df,
        "allocations": alloc_df if len(alloc_df) else pd.DataFrame(columns=ALLOC_COLS),
        "rejected": pd.DataFrame(rejected, columns=["DOC_NUMBER", "DOC_TYPE", "REASON",
                                                    "DETAIL", "SOURCE_FILE"]),
        "shortage": pd.DataFrame(shortages),
        "accepted": pd.DataFrame(accepted),
        "balance": pallet_balance(inv, alloc_df),
        "cfg": cfg,
    }


def _qty_str(q: float) -> str:
    return str(int(round(q))) if abs(q - round(q)) < 1e-9 else f"{q:g}"


def pallet_balance(inv: pd.DataFrame, alloc_df: pd.DataFrame) -> pd.DataFrame:
    """Pick කරපු හැම pallet එකකම before / picked / balance."""
    if alloc_df is None or not len(alloc_df):
        return pd.DataFrame(columns=["PALLET", "LOCATION_ID", "ITEM_NUMBER", "LOT_NUMBER",
                                     "PLANT", "UOM", "QTY_BEFORE", "QTY_PICKED",
                                     "QTY_BALANCE", "ROW_KEY"])
    g = (alloc_df.groupby(["ROW_KEY", "PALLET", "LOCATION_ID", "ITEM_NUMBER",
                           "LOT_NUMBER", "PLANT", "UOM"], dropna=False)
         .agg(QTY_PICKED=("QTY_PICKED", "sum"), QTY_BEFORE=("QTY_BEFORE", "max"))
         .reset_index())
    g["QTY_BALANCE"] = g["QTY_BEFORE"] - g["QTY_PICKED"]
    return g[["PALLET", "LOCATION_ID", "ITEM_NUMBER", "LOT_NUMBER", "PLANT", "UOM",
              "QTY_BEFORE", "QTY_PICKED", "QTY_BALANCE", "ROW_KEY"]]


def stock_view(inv_raw: pd.DataFrame, ledger: pd.DataFrame | None = None) -> pd.DataFrame:
    """Pallet-level live view — Actual, කලින් pick කරපු ප්‍රමාණය, ඉතුරු balance."""
    inv = normalize_inventory(inv_raw)
    picked: dict[str, float] = {}
    if ledger is not None and len(ledger) and "ROW_KEY" in ledger.columns:
        led = ledger.copy()
        led["QTY_PICKED"] = pd.to_numeric(led["QTY_PICKED"], errors="coerce").fillna(0)
        picked = led.groupby("ROW_KEY")["QTY_PICKED"].sum().to_dict()
    inv["PICKED_BEFORE"] = inv["row_key"].map(picked).fillna(0.0)
    inv["BALANCE"] = (inv["free_qty"] - inv["PICKED_BEFORE"]).clip(lower=0)
    inv["item_number"] = inv["item_number_raw"].where(
        inv["item_number_raw"].astype(bool), inv["item_number"])
    out = inv.rename(columns={
        "pallet": "PALLET", "location_id": "LOCATION_ID", "item_number": "ITEM_NUMBER",
        "base_id": "BASE_ID", "description": "DESCRIPTION", "lot_number": "LOT_NUMBER",
        "plant": "PLANT", "uom": "UOM", "free_qty": "ACTUAL_QTY", "status": "STATUS",
        "fifo_date": "FIFO_DATE", "grn_number": "GRN_NUMBER", "row_key": "ROW_KEY",
    })
    return out[["PALLET", "LOCATION_ID", "ITEM_NUMBER", "BASE_ID", "DESCRIPTION",
                "LOT_NUMBER", "PLANT", "UOM", "STATUS", "ACTUAL_QTY", "PICKED_BEFORE",
                "BALANCE", "FIFO_DATE", "GRN_NUMBER", "ROW_KEY"]]


# --------------------------------------------------------------------------- #
# Excel writers (everything as TEXT)
# --------------------------------------------------------------------------- #
def _write_text_sheet(ws, df: pd.DataFrame) -> None:
    ws.append(list(df.columns))
    for _, r in df.iterrows():
        ws.append([None if (pd.isna(v) or str(v) == "") else str(v) for v in r.tolist()])
    for col in range(1, len(df.columns) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(12, min(32, len(str(df.columns[col - 1])) + 3))
        for cell in ws[letter]:
            cell.number_format = "@"


def build_wms_excel(master: pd.DataFrame, detail: pd.DataFrame) -> bytes:
    """WMS upload file — 'OutBound MASTER' + 'OutBound Detail', සියල්ල text."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "OutBound MASTER"
    _write_text_sheet(ws1, master)
    ws2 = wb.create_sheet("OutBound Detail")
    _write_text_sheet(ws2, detail)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_report_excel(res: dict[str, Any]) -> bytes:
    """Pick report — allocations, pallet balance, shortage, rejected."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, key in [("Doc Summary", "accepted"), ("Pallet Allocation", "allocations"),
                          ("Pallet Balance", "balance"), ("Shortage", "shortage"),
                          ("Rejected Docs", "rejected")]:
            df = res.get(key)
            if df is None or not len(df):
                df = pd.DataFrame({"info": ["— කිසිවක් නෑ —"]})
            df.to_excel(xw, sheet_name=name[:31], index=False)
    return buf.getvalue()
