"""
pick_engine.py
==============
VIP / EFL Pick Generation Engine.

Requament (requirement) එක අරගෙන, SKU_MASTER + Inventory_Report check කරලා
දෙ ආකාරයක output ගොනු දෙකක් හදනවා:

  1. VIP PICK         -> physical pick list (HJ Box / HJ Pcs / Pcs per box + CBM summary)
  2. INDIA SO Pick    -> WMS Sales-Order upload (OutBound MASTER + OutBound Detail)

Core logic (5 sample files මත 166/169 = 98.2% exact-match කරලා validate කළා):

    boxsize   = inventory එකේ item එකකට වැඩිපුරම තියෙන Actual Qty (carton pack size, mode)
    cbm/box   = inventory එකේ item එකේ Cbm අගය
    divisor   = boxsize     (Category == LOOSE)
              = SAP          (Category == SET)     [SKU_MASTER එකෙන්]

    HJ Box Qty = floor(Req Qty / divisor)
    HJ Pcs Qty = HJ Box Qty * boxsize
    Pcs/Box    = boxsize
    REMARKS    = "Shortage"  if HJ Pcs Qty < Req Qty (carton rounding එකේදී qty සම්පූර්ණ නොවුණොත්)

INDIA SO OutBound Detail QTY == HJ Pcs Qty (pieces) -- 169/169 exact match.
"""

from __future__ import annotations
import io
import math
from dataclasses import dataclass, field
from datetime import datetime

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------- #
# Constants captured EXACTLY from the reference INDIA_SO_Pick.xlsx
# --------------------------------------------------------------------------- #
MASTER_COLS = ["HOST_ORDER_MASTER_ID", "HOST_GROUP_ID", "RECORD_CREATE_DATE", "PROCESSING_CODE", "WH_ID", "CLIENT_CODE", "ORDER_NUMBER", "DISPLAY_ORDER_NUMBER", "STORE_ORDER_NUMBER", "ORDER_TYPE", "CUSTOMER_CODE", "CUSTOMER_PO_NUMBER", "DEPARTMENT", "LOAD_ID", "LOAD_SEQ", "BOL_NUMBER", "MASTER_BOL_NUMBER", "PRO_NUMBER", "CARRIER", "CARRIER_SCAC", "FREIGHT_TERMS", "RUSH", "ORDER_DATE", "ARRIVE_DATE", "DATE_EXPECTED", "PROMISED_DATE", "WEIGHT", "CUBIC_VOLUME", "CONTAINERS", "BACKORDER", "PRE_PAID", "COD_AMOUNT", "INSURANCE_AMOUNT", "PIP_AMOUNT", "FREIGHT_COST", "SHIP_TO_CODE", "SHIP_TO_NAME", "SHIP_TO_ADDR1", "SHIP_TO_ADDR2", "SHIP_TO_ADDR3", "SHIP_TO_CITY", "SHIP_TO_STATE", "SHIP_TO_ZIP", "SHIP_TO_COUNTRY_CODE", "SHIP_TO_COUNTRY_NAME", "SHIP_TO_PHONE", "BILL_TO_CODE", "BILL_TO_NAME", "BILL_TO_ADDR1", "BILL_TO_ADDR2", "BILL_TO_ADDR3", "BILL_TO_CITY", "BILL_TO_STATE", "BILL_TO_ZIP", "BILL_TO_COUNTRY_CODE", "BILL_TO_COUNTRY_NAME", "BILL_TO_PHONE", "DELIVERY_NAME", "DELIVERY_ADDR1", "DELIVERY_ADDR2", "DELIVERY_ADDR3", "DELIVERY_CITY", "DELIVERY_STATE", "DELIVERY_ZIP", "DELIVERY_COUNTRY_CODE", "DELIVERY_COUNTRY_NAME", "DELIVERY_PHONE", "BILL_FRGHT_TO_CODE", "BILL_FRGHT_TO_NAME", "BILL_FRGHT_TO_ADDR1", "BILL_FRGHT_TO_ADDR2", "BILL_FRGHT_TO_ADDR3", "BILL_FRGHT_TO_CITY", "BILL_FRGHT_TO_STATE", "BILL_FRGHT_TO_ZIP", "BILL_FRGHT_TO_COUNTRY_CODE", "BILL_FRGHT_TO_COUNTRY_NAME", "BILL_FRGHT_TO_PHONE", "CARTON_LABEL", "VER_FLAG", "PARTIAL_ORDER_FLAG", "EARLIEST_SHIP_DATE", "LATEST_SHIP_DATE", "EARLIEST_DELIVERY_DATE", "LATEST_DELIVERY_DATE", "TEMP_LINK_ID", "SERVICE_LEVEL", "SHIP_VIA", "SHIP_TO_ATTENTION", "SAT_DELIVERY_FLAG", "REGISTERED_MAIL_FLAG", "RESTRICTED_MAIL_FLAG", "COD_FLAG", "COD_PAY_TYPE", "COD_OPTION", "INSURANCE_FLAG", "BILL_FRGHT_TO_ATTENTION", "SHIP_TO_RESIDENTIAL_FLAG", "CARRIER_MODE", "EARLIEST_APPT_TIME", "LATEST_APPT_TIME"]

DETAIL_COLS = ["HOST_ORDER_DETAIL_ID", "HOST_ORDER_MASTER_ID", "HOST_GROUP_ID", "RECORD_CREATE_DATE", "PROCESSING_CODE", "WH_ID", "CLIENT_CODE", "ORDER_NUMBER", "DISPLAY_ORDER_NUMBER", "LINE_NUMBER", "ITEM_NUMBER", "DISPLAY_ITEM_NUMBER", "ITEM_DESCRIPTION", "CUST_PART", "LOT_NUMBER", "QTY", "UNIT_WEIGHT", "UNIT_VOLUME", "EXTENDED_WEIGHT", "EXTENDED_VOLUME", "HAZ_MATERIAL", "BOL_CLASS", "BOL_CODE", "ORDER_UOM", "HOST_WAVE_ID", "TEMP_LINK_ID", "UNIT_INSURANCE_AMOUNT", "GEN_ATTRIBUTE_VALUE1", "GEN_ATTRIBUTE_VALUE2", "GEN_ATTRIBUTE_VALUE3", "GEN_ATTRIBUTE_VALUE4", "GEN_ATTRIBUTE_VALUE5", "GEN_ATTRIBUTE_VALUE6", "GEN_ATTRIBUTE_VALUE7", "GEN_ATTRIBUTE_VALUE8", "GEN_ATTRIBUTE_VALUE9", "GEN_ATTRIBUTE_VALUE10", "GEN_ATTRIBUTE_VALUE11", "HOLD_REASON_ID", "PACKING_INST", "REIMA_LINE", "VAS_INST"]

MASTER_CONST = {
    "PROCESSING_CODE": "NEW", "WH_ID": "LPGL", "CLIENT_CODE": "INM0VIP",
    "ORDER_TYPE": "Sales Orders", "BACKORDER": "N", "PARTIAL_ORDER_FLAG": "N",
    "SAT_DELIVERY_FLAG": "N", "REGISTERED_MAIL_FLAG": "N", "RESTRICTED_MAIL_FLAG": "N",
    "COD_FLAG": "N", "COD_PAY_TYPE": "N", "COD_OPTION": "N",
    "INSURANCE_FLAG": "N", "SHIP_TO_RESIDENTIAL_FLAG": "N",
}
DETAIL_CONST = {
    "PROCESSING_CODE": "NEW", "WH_ID": "LPGL", "CLIENT_CODE": "INM0VIP",
    "ORDER_UOM": "PCS", "GEN_ATTRIBUTE_VALUE1": "TBC",
    "GEN_ATTRIBUTE_VALUE5": "TBC", "GEN_ATTRIBUTE_VALUE9": "TBC",
}


# --------------------------------------------------------------------------- #
# Configuration (UI එකෙන් override කරන්න පුළුවන්)
# --------------------------------------------------------------------------- #
@dataclass
class EngineConfig:
    rounding: str = "floor"          # floor | round | ceil  (carton count)
    per_minute_cbm: float = 1 / 3    # productivity assumption -> 0.333333
    wh_id: str = "LPGL"
    client_code: str = "INM0VIP"
    order_type: str = "Sales Orders"
    pick_date: datetime = field(default_factory=datetime.now)
    # Fixed QR codes shown across the top of the LOAD ID QR sheet
    header_qr_codes: list = field(
        default_factory=lambda: ["INM0VIP", "PKINM0", "IMSA05"]
    )


# --------------------------------------------------------------------------- #
# Flexible column resolver -- header name වෙනස් වුණත් වැඩ කරන්න
# --------------------------------------------------------------------------- #
def _find_col(df: pd.DataFrame, *candidates: str) -> str | None:
    norm = {str(c).strip().lower().replace("\n", " ").replace("_", " "): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower().replace("_", " ")
        # exact
        if key in norm:
            return norm[key]
        # contains
        for k, original in norm.items():
            if key in k:
                return original
    return None


# --------------------------------------------------------------------------- #
# Input parsing
# --------------------------------------------------------------------------- #
def parse_requirement(df: pd.DataFrame) -> pd.DataFrame:
    c_mat = _find_col(df, "Material code", "Material")
    c_qty = _find_col(df, "Req Qty", "Qty", "Quantity")
    c_del = _find_col(df, "Delivery No", "Delivery", "OBD", "Load")
    c_des = _find_col(df, "Material Description", "Description", "Desc")
    if not (c_mat and c_qty and c_del):
        raise ValueError(
            "Requirement එකේ Material / Req Qty / Delivery No columns හොයාගන්න බැරි වුණා. "
            f"හමු වුණ columns: {list(df.columns)}"
        )
    out = pd.DataFrame({
        "Material": df[c_mat].astype(str).str.strip(),
        "Description": (df[c_des].astype(str).str.strip() if c_des else ""),
        "ReqQty": pd.to_numeric(df[c_qty], errors="coerce"),
        "DeliveryNo": df[c_del],
    })
    out = out[out["Material"].notna() & (out["Material"] != "") & (out["Material"].str.lower() != "nan")]
    out = out[out["ReqQty"].notna()]
    out["ReqQty"] = out["ReqQty"].astype(int)
    return out.reset_index(drop=True)


def build_sku_lookup(df: pd.DataFrame) -> dict:
    c_mat = _find_col(df, "Material code", "Material")
    c_cat = _find_col(df, "Catergory", "Category")
    c_sap = _find_col(df, "SAP")
    c_hj = _find_col(df, "HJ")
    if not c_mat:
        raise ValueError("SKU_MASTER එකේ Material code column හොයාගන්න බැරි වුණා.")
    tmp = pd.DataFrame({
        "Material": df[c_mat].astype(str).str.strip(),
        "Category": (df[c_cat].astype(str).str.strip().str.upper() if c_cat else "LOOSE"),
        "SAP": (pd.to_numeric(df[c_sap], errors="coerce") if c_sap else 1),
        "HJ": (pd.to_numeric(df[c_hj], errors="coerce") if c_hj else 1),
    })
    # Duplicate material codes -> Category first, SAP/HJ max (validated approach)
    g = tmp.groupby("Material").agg(Category=("Category", "first"),
                                    SAP=("SAP", "max"), HJ=("HJ", "max"))
    return g.to_dict("index")


def build_inventory_lookup(df: pd.DataFrame) -> dict:
    c_item = _find_col(df, "Item Number", "Item", "Material")
    c_qty = _find_col(df, "Actual Qty", "Qty")
    c_cbm = _find_col(df, "Cbm", "CBM")
    if not (c_item and c_qty):
        raise ValueError("Inventory_Report එකේ Item Number / Actual Qty columns හොයාගන්න බැරි වුණා.")

    work = pd.DataFrame({
        "Item": df[c_item].astype(str).str.strip(),
        "Qty": pd.to_numeric(df[c_qty], errors="coerce"),
        "Cbm": (pd.to_numeric(df[c_cbm], errors="coerce") if c_cbm else 0.0),
    })

    def carton_mode(s: pd.Series):
        s = s[s > 0]
        if len(s) == 0:
            return None
        return int(s.mode().iloc[0])

    out = {}
    for item, sub in work.groupby("Item"):
        out[item] = {
            "boxsize": carton_mode(sub["Qty"]),
            "avail": int(sub["Qty"].sum(skipna=True)),
            "cbm": float(sub["Cbm"].dropna().iloc[0]) if sub["Cbm"].notna().any() else 0.0,
        }
    return out


# --------------------------------------------------------------------------- #
# Core calculation
# --------------------------------------------------------------------------- #
def _carton_count(qty: int, divisor: int, mode: str) -> int:
    if not divisor or divisor <= 0:
        return 0
    x = qty / divisor
    if mode == "round":
        return int(round(x))
    if mode == "ceil":
        return int(math.ceil(x))
    return int(math.floor(x))   # default floor


def compute_pick(requirement: pd.DataFrame, sku: dict, inv: dict,
                 cfg: EngineConfig) -> pd.DataFrame:
    """Requirement -> enriched pick rows (one per requirement line).

    Pick formula (authoritative):
        target_pcs = Req Qty * HJ / SAP
          - LOOSE :  SAP == HJ  ->  target = Req Qty
          - SET   :  HJ/SAP is the set multiplier (e.g. 3)

    A carton (inventory Actual Qty = `boxsize`) can never be split, so we pick
    only WHOLE cartons, limited by available stock:
        pick_cartons = min(target // carton, available // carton)
        picked_pcs   = pick_cartons * carton
        variance     = target - picked_pcs   (the part we cannot pick)

    The picked whole-carton portion goes to the main outputs (VIP PICK / INDIA
    SO). Any line with variance > 0 (carton remainder or short stock) is ALSO
    listed in the Cannot Pick report with picked qty + variance.

    `_status`:
        OK            -> picked == target (no variance)
        CARTON_SPLIT  -> variance from a non-divisible carton remainder
        SHORT_STOCK   -> variance because target exceeds available stock
        NO_DATA       -> material missing from SKU master and/or inventory
    """
    rows = []
    for _, r in requirement.iterrows():
        mat = r["Material"]
        qty = int(r["ReqQty"])
        s = sku.get(mat, {})
        i = inv.get(mat, {})

        category = (s.get("Category") or "LOOSE")
        sap = int(s.get("SAP") or 0)
        hj = int(s.get("HJ") or 0)
        boxsize = i.get("boxsize")
        cbm = float(i.get("cbm") or 0.0)
        avail = int(i.get("avail") or 0)

        target_exact = (qty * hj / sap) if sap else 0.0
        target = int(round(target_exact))
        unit_split = bool(sap) and abs(target_exact - target) > 1e-9

        box = int(boxsize) if boxsize else 0
        no_data = (mat not in sku) or (not box)

        if no_data:
            status = "NO_DATA"
            pick_cartons = 0
        else:
            desired_cartons = target // box
            stock_cartons = avail // box
            pick_cartons = min(desired_cartons, stock_cartons)

        picked_pcs = pick_cartons * box if box else 0
        variance = target - picked_pcs

        issues = []
        if mat not in sku:
            issues.append("Not in SKU master")
        if not box:
            issues.append("Not in inventory")
        if unit_split:
            issues.append(f"Req*HJ not divisible by SAP({sap})")

        if no_data:
            status = "NO_DATA"
        elif variance <= 0:
            status = "OK"
        elif target > avail:
            status = "SHORT_STOCK"
            issues.append(f"Req {target} > stock {avail}; picked {picked_pcs}, variance {variance}")
        else:
            status = "CARTON_SPLIT"
            issues.append(f"{target} pcs not whole cartons of {box}; picked {picked_pcs}, variance {variance}")

        remark = "" if status == "OK" else (
            f"Variance {variance}" if status in ("CARTON_SPLIT", "SHORT_STOCK") else "; ".join(issues))

        rows.append({
            "Material": mat,
            "Material des": r.get("Description", ""),
            "Qty": qty,
            "OBD": r["DeliveryNo"],
            "HJ Box Qty": int(pick_cartons),
            "HJ Pcs Qty": int(picked_pcs),
            "Pcs/Box": box,
            "REMARKS": remark,
            "Date": cfg.pick_date,
            "_status": status,
            "_category": category,
            "_sap": sap,
            "_hj": hj,
            "_target": int(target),
            "_picked": int(picked_pcs),
            "_variance": int(variance),
            "_avail": avail,
            "_issue": "; ".join(issues),
            "_cbm_per_box": cbm,
            "_line_cbm": round(pick_cartons * cbm, 4),
        })
    return pd.DataFrame(rows)


def split_reports(pick: pd.DataFrame):
    """Split into (pickable, exceptions).

    pickable   -> lines where we pick at least one whole carton (HJ Pcs Qty > 0),
                  including partial-carton lines (picked portion only).
    exceptions -> lines with variance > 0 (carton remainder / short stock) or
                  no data, showing picked qty + variance.
    """
    ok = pick[pick["HJ Pcs Qty"] > 0].reset_index(drop=True)
    bad = pick[(pick["_variance"] > 0) | (pick["_status"] == "NO_DATA")].reset_index(drop=True)

    issue_label = {
        "CARTON_SPLIT": "Carton split — partial pick",
        "SHORT_STOCK": "Insufficient stock (Req > Inventory)",
        "NO_DATA": "Missing SKU / inventory data",
    }
    if len(bad):
        exc = pd.DataFrame({
            "Material": bad["Material"],
            "Material des": bad["Material des"],
            "OBD": bad["OBD"],
            "Category": bad["_category"],
            "Req Qty": bad["Qty"],
            "SAP": bad["_sap"],
            "HJ": bad["_hj"],
            "Target Pcs": bad["_target"],
            "Pcs/Carton": bad["Pcs/Box"],
            "Picked Pcs": bad["_picked"],
            "Variance": bad["_variance"],
            "Available": bad["_avail"],
            "Issue Type": bad["_status"].map(issue_label).fillna(bad["_status"]),
            "Detail": bad["_issue"],
        })
    else:
        exc = pd.DataFrame(columns=[
            "Material", "Material des", "OBD", "Category", "Req Qty", "SAP", "HJ",
            "Target Pcs", "Pcs/Carton", "Picked Pcs", "Variance", "Available",
            "Issue Type", "Detail"])
    return ok, exc


def missing_materials(req_df, sku_df) -> list:
    """Requirement එකේ තියෙන, SKU_MASTER එකේ නැති Material codes ලැයිස්තුව
    (order of first appearance). Pick create කරන්න කලින් notify කරන්න."""
    req = parse_requirement(req_df)
    sku = build_sku_lookup(sku_df)
    req_mats = list(dict.fromkeys(req["Material"].astype(str).str.strip().tolist()))
    return [m for m in req_mats if m and m not in sku]


def dedup_load_ids(load_ids: list, existing=None):
    """Make LOAD IDs globally unique against an existing registry.

    Returns (mapping original->unique, list of newly-used unique ids).
    Duplicates get -A, -B, ... -Z, -AA ... appended.
    """
    existing = set(str(x) for x in (existing or set()))
    mapping, new_used = {}, []

    def suffix(n):
        s = ""
        n += 1
        while n:
            n, rem = divmod(n - 1, 26)
            s = chr(65 + rem) + s
        return s

    used = set(existing)
    for lid in load_ids:
        base = str(lid)
        cand = base
        k = 0
        while cand in used:
            cand = f"{base}-{suffix(k)}"
            k += 1
        used.add(cand)
        new_used.append(cand)
        mapping[lid] = cand
    return mapping, new_used


def cbm_summary(pick: pd.DataFrame, cfg: EngineConfig) -> dict:
    total_cbm = float(pick["_line_cbm"].sum())
    per_min = cfg.per_minute_cbm or (1 / 3)
    total_min = total_cbm / per_min if per_min else 0.0
    total_sec = int(round(total_min * 60))
    h, rem = divmod(total_sec, 3600)
    m, sec = divmod(rem, 60)
    return {
        "Total CBM Of Pick": round(total_cbm, 2),
        "Per Minute CBM": round(per_min, 6),
        "Total Minute for Pick": round(total_min, 2),
        "Time": f"{h}:{m:02d}:{sec:02d}",
        "Hours/Minutes/Seconds": f"{h} H {m} Min {sec} Sec",
    }


# --------------------------------------------------------------------------- #
# Output builders
# --------------------------------------------------------------------------- #
def build_india_so(pick: pd.DataFrame, cfg: EngineConfig,
                   load_id_map: dict | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """INDIA SO OutBound MASTER + Detail.

    `load_id_map` (original delivery -> globally-unique LOAD ID) is applied to the
    LOAD_ID column only; DISPLAY/STORE/CUSTOMER order numbers keep the real
    delivery number.
    """
    load_id_map = load_id_map or {}
    deliveries = list(dict.fromkeys(pick["OBD"].tolist()))  # preserve order, unique
    load_ids = [load_id_map.get(d, d) for d in deliveries]

    # ---- MASTER : one row per delivery ----
    m = pd.DataFrame(index=range(len(deliveries)), columns=MASTER_COLS)
    for k, v in MASTER_CONST.items():
        m[k] = v
    m["WH_ID"] = cfg.wh_id
    m["CLIENT_CODE"] = cfg.client_code
    m["ORDER_TYPE"] = cfg.order_type
    m["DISPLAY_ORDER_NUMBER"] = deliveries
    m["STORE_ORDER_NUMBER"] = deliveries
    m["CUSTOMER_PO_NUMBER"] = deliveries
    m["LOAD_ID"] = load_ids

    # ---- DETAIL : one row per requirement line ----
    d = pd.DataFrame(index=range(len(pick)), columns=DETAIL_COLS)
    for k, v in DETAIL_CONST.items():
        d[k] = v
    d["WH_ID"] = cfg.wh_id
    d["CLIENT_CODE"] = cfg.client_code
    d["DISPLAY_ORDER_NUMBER"] = pick["OBD"].values
    d["DISPLAY_ITEM_NUMBER"] = pick["Material"].values
    d["QTY"] = pick["HJ Pcs Qty"].values            # pieces (validated 169/169)
    # LINE_NUMBER : sequential within each delivery
    line_no = []
    counter: dict = {}
    for obd in pick["OBD"].values:
        counter[obd] = counter.get(obd, 0) + 1
        line_no.append(counter[obd])
    d["LINE_NUMBER"] = line_no

    return m, d


def build_vip_pick_sheet(pick: pd.DataFrame, cfg: EngineConfig):
    """Return (summary_dict, pick_table_df) for the VIP PICK Summary sheet."""
    table = pick[["Material", "Material des", "Qty", "OBD",
                  "HJ Box Qty", "HJ Pcs Qty", "Pcs/Box", "REMARKS", "Date"]].copy()
    table = table.rename(columns={"Material": "Material "})  # match original header
    return cbm_summary(pick, cfg), table


def build_load_id_qr(pick: pd.DataFrame) -> list:
    """Unique LOAD IDs (= Delivery No / OBD) in order of first appearance.

    Matches the reference VIP_PICK 'LOAD ID QR' sheet (one row per delivery,
    25 in the sample).
    """
    return list(dict.fromkeys(pick["OBD"].tolist()))


# --------------------------------------------------------------------------- #
# Excel writers
# --------------------------------------------------------------------------- #
def _autosize(ws, df, start_row=1):
    from openpyxl.utils import get_column_letter
    for j, col in enumerate(df.columns, start=1):
        width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200).tolist()] + [8])
        ws.column_dimensions[get_column_letter(j)].width = min(width + 2, 45)


def write_india_so_xlsx(master: pd.DataFrame, detail: pd.DataFrame) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    hdr_fill = PatternFill("solid", fgColor="1F4E78")
    align = Alignment(horizontal="center", vertical="center")

    for name, df in [("OutBound MASTER", master), ("OutBound Detail", detail)]:
        ws = wb.active if name == "OutBound MASTER" else wb.create_sheet(name)
        if name == "OutBound MASTER":
            ws.title = name
        ws.append(list(df.columns))
        for _, row in df.iterrows():
            ws.append(["" if (pd.isna(v)) else v for v in row.tolist()])
        for c in ws[1]:
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = align
        ws.freeze_panes = "A2"
        _autosize(ws, df)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _qr_png_bytes(text: str, box_size: int = 3, border: int = 2):
    """Return PNG bytes of a QR code for `text`, or None if qrcode unavailable."""
    try:
        import qrcode
    except ImportError:
        return None
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=box_size,
        border=border,
    )
    qr.add_data(str(text))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    out = io.BytesIO()
    img.save(out, format="PNG")
    out.seek(0)
    return out


def _write_load_id_qr_sheet(wb, load_ids: list, gap: int = 5, embed_qr: bool = True,
                            header_qr_codes=None):
    """Add a 'LOAD ID QR' sheet: each LOAD ID on its own row block (gap rows
    apart) with a scannable QR image placed in the 'QR Code' column.

    Mirrors the reference VIP_PICK layout (header + one ID every `gap` rows).
    If the qrcode library is missing, the QR image is skipped gracefully and
    the LOAD ID value is still written.

    `header_qr_codes` is a list of fixed labels (e.g. INM0VIP / PKINM0 / IMSA05)
    rendered as labelled QR blocks across the top-right of the sheet.
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet("LOAD ID QR")
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    hdr_fill = PatternFill("solid", fgColor="C00000")
    lbl_font = Font(bold=True, name="Arial", size=11)
    id_font = Font(bold=True, name="Arial", size=11)
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.append(["LOAD ID", "QR Code"])
    for c in ws[1]:
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = center
        c.border = border

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 28

    try:
        from openpyxl.drawing.image import Image as XLImage
        have_img = True
    except ImportError:
        have_img = False

    # ---- fixed header QR codes (INM0VIP / PKINM0 / IMSA05 ...) ----
    if header_qr_codes:
        start_col = 4          # column D
        block_w = 2            # each label/QR block spans 2 columns
        gap_cols = 1           # one blank column between blocks
        for i, code in enumerate(header_qr_codes):
            c0 = start_col + i * (block_w + gap_cols)
            c1 = c0 + block_w - 1
            l0 = get_column_letter(c0)
            l1 = get_column_letter(c1)
            ws.merge_cells(f"{l0}1:{l1}1")
            hc = ws.cell(row=1, column=c0, value=code)
            hc.font = lbl_font
            hc.alignment = center
            for cc in range(c0, c1 + 1):
                ws.cell(row=1, column=cc).border = border
            if embed_qr and have_img:
                png = _qr_png_bytes(code, box_size=4)
                if png is not None:
                    try:
                        img = XLImage(png)
                        img.width = 110
                        img.height = 110
                        ws.add_image(img, f"{l0}2")
                    except Exception:
                        pass

    row = 2
    for lid in load_ids:
        cell = ws.cell(row=row, column=1, value=lid)
        cell.font = id_font
        cell.alignment = center
        ws.cell(row=row, column=2)  # keep QR Code column present

        if embed_qr and have_img:
            png = _qr_png_bytes(lid)
            if png is not None:
                try:
                    img = XLImage(png)
                    img.width = 90
                    img.height = 90
                    ws.add_image(img, f"B{row}")
                    for r in range(row, row + gap - 1):
                        ws.row_dimensions[r].height = 20
                except Exception:
                    pass
        row += gap

    return ws


def write_vip_pick_xlsx(summary: dict, table: pd.DataFrame, load_ids=None,
                        header_qr_codes=None, exceptions=None) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    title_font = Font(bold=True, name="Arial", size=10)
    hdr_font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    hdr_fill = PatternFill("solid", fgColor="C00000")
    sum_fill = PatternFill("solid", fgColor="FFF2CC")
    short_fill = PatternFill("solid", fgColor="FFC7CE")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # --- CBM summary block (rows 1-2) ---
    sum_cols = ["Total CBM Of Pick", "Per Minute CBM", "Total Minute for Pick",
                "Time", "Hours/Minutes/Seconds"]
    ws.append(sum_cols)
    ws.append([summary[c] for c in sum_cols])
    for r in (1, 2):
        for j in range(1, len(sum_cols) + 1):
            cell = ws.cell(row=r, column=j)
            cell.fill = sum_fill
            cell.border = border
            cell.alignment = center
            if r == 1:
                cell.font = title_font
    ws.append([])  # blank row 3

    # --- pick table (header row 4) ---
    hdr_row = 4
    ws.append(list(table.columns))
    for c in ws[hdr_row]:
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = center
        c.border = border

    short_idx = list(table.columns).index("REMARKS")
    date_idx = list(table.columns).index("Date")
    for _, row in table.iterrows():
        vals = []
        for j, v in enumerate(row.tolist()):
            if j == date_idx and isinstance(v, datetime):
                v = v.strftime("%Y-%m-%d")
            vals.append("" if pd.isna(v) else v)
        ws.append(vals)
        rr = ws.max_row
        is_short = str(row["REMARKS"]).strip() != ""
        for j in range(1, len(table.columns) + 1):
            cell = ws.cell(row=rr, column=j)
            cell.border = border
            if j - 1 in (2, 4, 5, 6):   # numeric cols centered
                cell.alignment = center
            if is_short:
                cell.fill = short_fill

    ws.freeze_panes = "A5"
    _autosize(ws, table, start_row=hdr_row)

    # --- Exceptions / Cannot-Pick sheet ---
    if exceptions is not None and len(exceptions):
        ex = wb.create_sheet("Cannot Pick")
        ehdr_fill = PatternFill("solid", fgColor="C00000")
        ews_cols = list(exceptions.columns)
        ex.append(ews_cols)
        for c in ex[1]:
            c.font = hdr_font
            c.fill = ehdr_fill
            c.alignment = center
            c.border = border
        for _, row in exceptions.iterrows():
            ex.append(["" if pd.isna(v) else v for v in row.tolist()])
            for j in range(1, len(ews_cols) + 1):
                ex.cell(row=ex.max_row, column=j).border = border
                ex.cell(row=ex.max_row, column=j).fill = short_fill
        ex.freeze_panes = "A2"
        _autosize(ex, exceptions, start_row=1)

    # --- LOAD ID QR sheet ---
    if load_ids:
        _write_load_id_qr_sheet(wb, load_ids, header_qr_codes=header_qr_codes)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# One-shot pipeline
# --------------------------------------------------------------------------- #
def run_pipeline(req_df, sku_df, inv_df, cfg: EngineConfig | None = None,
                 existing_load_ids=None) -> dict:
    cfg = cfg or EngineConfig()
    requirement = parse_requirement(req_df)
    sku = build_sku_lookup(sku_df)
    inv = build_inventory_lookup(inv_df)

    full_pick = compute_pick(requirement, sku, inv, cfg)
    pick, exceptions = split_reports(full_pick)        # OK lines vs cannot-pick

    # Globally-unique LOAD IDs (against existing registry, suffix -A/-B/...)
    base_ids = build_load_id_qr(pick)
    load_id_map, load_ids = dedup_load_ids(base_ids, existing_load_ids)

    master, detail = build_india_so(pick, cfg, load_id_map)
    summary, vip_table = build_vip_pick_sheet(pick, cfg)

    return {
        "requirement": requirement,
        "full_pick": full_pick,
        "pick": pick,
        "exceptions": exceptions,
        "summary": summary,
        "vip_table": vip_table,
        "india_master": master,
        "india_detail": detail,
        "load_ids": load_ids,
        "load_id_map": load_id_map,
        "vip_pick_bytes": write_vip_pick_xlsx(
            summary, vip_table, load_ids,
            header_qr_codes=cfg.header_qr_codes, exceptions=exceptions),
        "india_so_bytes": write_india_so_xlsx(master, detail),
    }
