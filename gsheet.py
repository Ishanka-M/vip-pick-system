"""
gsheet.py
=========
Google Sheets කියවීම/ලිවීම සඳහා optional helper.

Service Account credentials, Streamlit secrets (`st.secrets["gcp_service_account"]`)
එකෙන් එනවා. Google Sheets භාවිතා කරන්නේ නැත්නම් මේ module එක import වුණත්
app එක Excel-upload mode එකෙන් සම්පූර්ණයෙන් වැඩ කරයි.

Setup (README බලන්න):
  1. Google Cloud Console -> Service Account හදන්න -> JSON key download.
  2. Google Sheet එක එම service-account email එකට share කරන්න (Viewer/Editor).
  3. JSON එක Streamlit secrets එකට `[gcp_service_account]` විදිහට දාන්න.
"""

from __future__ import annotations
import pandas as pd

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _client(service_account_info: dict):
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_info(dict(service_account_info), scopes=SCOPES)
    return gspread.authorize(creds)


def read_sheet(service_account_info: dict, sheet_url_or_key: str,
               worksheet: str | int = 0) -> pd.DataFrame:
    """Google Sheet worksheet එකක් DataFrame එකක් විදිහට කියවනවා (first row = header)."""
    gc = _client(service_account_info)
    sh = (gc.open_by_url(sheet_url_or_key) if sheet_url_or_key.startswith("http")
          else gc.open_by_key(sheet_url_or_key))
    ws = sh.get_worksheet(worksheet) if isinstance(worksheet, int) else sh.worksheet(worksheet)
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()
    header, *rows = values
    # duplicate/blank header guard (gspread 6.x safe)
    seen, clean = {}, []
    for i, h in enumerate(header):
        h = (h or f"col_{i}").strip() or f"col_{i}"
        if h in seen:
            seen[h] += 1
            h = f"{h}_{seen[h]}"
        else:
            seen[h] = 0
        clean.append(h)
    return pd.DataFrame(rows, columns=clean)


def write_sheet(service_account_info: dict, sheet_url_or_key: str,
                worksheet_title: str, df: pd.DataFrame, clear: bool = True):
    """DataFrame එකක් worksheet එකකට ලියනවා (නැත්නම් අලුතෙන් හදනවා)."""
    import gspread
    gc = _client(service_account_info)
    sh = _open(gc, sheet_url_or_key)
    try:
        ws = sh.worksheet(worksheet_title)
        if clear:
            ws.clear()
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_title,
                              rows=max(len(df) + 5, 20), cols=max(len(df.columns) + 2, 10))
    safe = df.fillna("").astype(object).where(pd.notna(df), "")
    payload = [list(map(str, safe.columns))] + safe.astype(str).values.tolist()
    ws.update(payload, value_input_option="USER_ENTERED")
    return ws.url


def _open(gc, sheet_url_or_key: str):
    return (gc.open_by_url(sheet_url_or_key) if sheet_url_or_key.startswith("http")
            else gc.open_by_key(sheet_url_or_key))


# --------------------------------------------------------------------------- #
# Worksheet headers + auto-create
# --------------------------------------------------------------------------- #
SKU_WS = "SKU_MASTER"
SKU_COLS = ["Material code", "Material Desc", "Catergory", "HJ", "SAP"]
LOADID_WS = "LOAD_ID Registry"
LOADID_COLS = ["LOAD ID", "Used At", "Source"]
RUNLOG_WS = "Run Log"
RUNLOG_COLS = ["Timestamp", "Source", "Pick Lines", "Deliveries", "Total Boxes",
               "Total CBM", "Per Minute CBM", "Pick Time", "Cannot Pick Rows", "Rounding"]
CANNOTPICK_WS = "Cannot Pick"
CANNOTPICK_COLS = ["Material", "Material des", "OBD", "Category", "Req Qty", "SAP",
                   "HJ", "Target Pcs", "Pcs/Carton", "Available", "Issue Type", "Detail"]

# Structural worksheets that should exist (with headers) in every data sheet.
INIT_SHEETS = {
    SKU_WS: SKU_COLS,
    LOADID_WS: LOADID_COLS,
    RUNLOG_WS: RUNLOG_COLS,
    CANNOTPICK_WS: CANNOTPICK_COLS,
}


def ensure_worksheet(service_account_info: dict, sheet_url_or_key: str,
                     title: str, header: list) -> bool:
    """Create `title` worksheet with `header` if it doesn't exist (or is empty).

    Returns True if it was created/initialised, False if it already had data.
    """
    import gspread
    gc = _client(service_account_info)
    sh = _open(gc, sheet_url_or_key)
    try:
        ws = sh.worksheet(title)
        if ws.get_all_values():
            return False
        ws.update([header], value_input_option="USER_ENTERED")
        return True
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=200, cols=max(len(header) + 2, 12))
        ws.update([header], value_input_option="USER_ENTERED")
        return True


def init_sheet(service_account_info: dict, sheet_url_or_key: str) -> dict:
    """Auto-create every required worksheet + headers in the data sheet.

    Safe to run on a brand-new empty Google Sheet or an existing one
    (idempotent — existing sheets with data are left untouched).
    Returns {worksheet_title: created_bool, "url": spreadsheet_url}.
    """
    result = {}
    for title, header in INIT_SHEETS.items():
        try:
            result[title] = ensure_worksheet(service_account_info, sheet_url_or_key, title, header)
        except Exception as ex:  # noqa
            result[title] = f"error: {ex}"
    gc = _client(service_account_info)
    result["url"] = _open(gc, sheet_url_or_key).url
    return result


def append_log(service_account_info: dict, sheet_url_or_key: str,
               row: dict, worksheet_title: str = "Run Log"):
    """Append-only history. එක run එකකට එක row එකක් — worksheet එක නැත්නම් header
    එක්ක අලුතෙන් හදනවා, තියෙනවා නම් යටට append කරනවා (clear කරන්නේ නෑ)."""
    import gspread
    gc = _client(service_account_info)
    sh = _open(gc, sheet_url_or_key)
    header = list(row.keys())
    try:
        ws = sh.worksheet(worksheet_title)
        existing = ws.get_all_values()
        if not existing:
            ws.update([header], value_input_option="USER_ENTERED")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=worksheet_title, rows=200, cols=max(len(header) + 2, 10))
        ws.update([header], value_input_option="USER_ENTERED")
    ws.append_row([str(v) for v in row.values()], value_input_option="USER_ENTERED")
    return ws.url


def save_all(service_account_info: dict, sheet_url_or_key: str, res: dict,
             cfg=None, source_label: str = "") -> str:
    """System එක produce කරන හැම data එකක්ම එක Google Sheet එකකට save කරනවා.

    Worksheets (clear + rewrite හැම run එකකම):
      - VIP PICK          : pick table
      - CBM Summary       : CBM / pick-time block
      - LOAD ID QR        : unique LOAD IDs + fixed header codes
      - OutBound MASTER   : INDIA SO master
      - OutBound Detail   : INDIA SO detail
    සහ append-only:
      - Run Log           : එක run එකකට timestamp + counts row එකක්

    Returns: spreadsheet URL.
    """
    from datetime import datetime

    # ensure all structural worksheets + headers exist first
    try:
        init_sheet(service_account_info, sheet_url_or_key)
    except Exception:
        pass

    # --- main data worksheets (overwrite each run) ---
    write_sheet(service_account_info, sheet_url_or_key, "VIP PICK", res["vip_table"])

    summary_df = pd.DataFrame([res["summary"]])
    write_sheet(service_account_info, sheet_url_or_key, "CBM Summary", summary_df)

    header_codes = list(getattr(cfg, "header_qr_codes", []) or [])
    load_df = pd.DataFrame({"LOAD ID": res.get("load_ids", [])})
    load_df["QR Code"] = ""                       # QR images Excel export එකේ විතරයි
    if header_codes:
        load_df["Header QR Codes"] = (header_codes + [""] * len(load_df))[:len(load_df)] \
            if len(load_df) >= len(header_codes) else header_codes
    write_sheet(service_account_info, sheet_url_or_key, "LOAD ID QR", load_df)

    write_sheet(service_account_info, sheet_url_or_key, "OutBound MASTER", res["india_master"])
    write_sheet(service_account_info, sheet_url_or_key, "OutBound Detail", res["india_detail"])

    # --- Cannot Pick (exceptions) ---
    exc = res.get("exceptions")
    if exc is not None and len(exc):
        write_sheet(service_account_info, sheet_url_or_key, "Cannot Pick", exc)

    # --- LOAD_ID registry append (global uniqueness) ---
    try:
        append_load_ids(service_account_info, sheet_url_or_key,
                        res.get("load_ids", []), source=source_label)
    except Exception:
        pass

    # --- monthly pick history append ---
    try:
        append_history(service_account_info, sheet_url_or_key, res, cfg)
    except Exception:
        pass

    # --- run history (append) ---
    pick = res["pick"]
    _exc = res.get("exceptions")
    n_exc = len(_exc) if _exc is not None else 0
    log_row = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Source": source_label,
        "Pick Lines": len(pick),
        "Deliveries": int(pick["OBD"].nunique()) if "OBD" in pick.columns else len(res.get("load_ids", [])),
        "Total Boxes": int(pick["HJ Box Qty"].sum()) if "HJ Box Qty" in pick.columns else "",
        "Total CBM": res["summary"].get("Total CBM Of Pick", ""),
        "Per Minute CBM": res["summary"].get("Per Minute CBM", ""),
        "Pick Time": res["summary"].get("Hours/Minutes/Seconds", ""),
        "Cannot Pick Rows": n_exc,
        "Rounding": getattr(cfg, "rounding", ""),
    }
    url = append_log(service_account_info, sheet_url_or_key, log_row)
    return url


# --------------------------------------------------------------------------- #
# SKU_MASTER store (read + CRUD write-back)
# --------------------------------------------------------------------------- #
def read_sku_master(service_account_info: dict, sheet_url_or_key: str,
                    worksheet: str = SKU_WS) -> pd.DataFrame:
    """Read the SKU_MASTER worksheet as a DataFrame.

    If the worksheet does not exist yet, it is auto-created with the correct
    headers and an empty DataFrame (with SKU_COLS) is returned — so a brand-new
    Google Sheet works straight away.
    """
    import gspread
    try:
        df = read_sheet(service_account_info, sheet_url_or_key, worksheet)
    except gspread.WorksheetNotFound:
        ensure_worksheet(service_account_info, sheet_url_or_key, worksheet, SKU_COLS)
        return pd.DataFrame(columns=SKU_COLS)
    for c in SKU_COLS:
        if c not in df.columns:
            df[c] = ""
    return df[SKU_COLS] if set(SKU_COLS).issubset(set(df.columns)) else df


def save_sku_master(service_account_info: dict, sheet_url_or_key: str,
                    df: pd.DataFrame, worksheet: str = SKU_WS) -> str:
    """SKU_MASTER worksheet එක edit කරපු DataFrame එකෙන් overwrite කරනවා
    (add / update / delete — UI data_editor එකේ වෙනස්කම් save කරන්න)."""
    clean = df.copy().dropna(how="all")
    if "Material code" in clean.columns:
        clean = clean[clean["Material code"].astype(str).str.strip() != ""]
    return write_sheet(service_account_info, sheet_url_or_key, worksheet, clean)


# --------------------------------------------------------------------------- #
# LOAD_ID registry (global uniqueness across all runs)
# --------------------------------------------------------------------------- #
def read_load_id_registry(service_account_info: dict, sheet_url_or_key: str) -> set:
    """දැනට පාවිච්චි කරපු හැම LOAD ID එකක්ම set එකක් (duplicate check එකට)."""
    try:
        df = read_sheet(service_account_info, sheet_url_or_key, LOADID_WS)
    except Exception:
        return set()
    if df.empty:
        return set()
    col = "LOAD ID" if "LOAD ID" in df.columns else df.columns[0]
    return set(str(x).strip() for x in df[col].tolist() if str(x).strip())


def append_load_ids(service_account_info: dict, sheet_url_or_key: str,
                    load_ids: list, source: str = "") -> str:
    """අලුතෙන් පාවිච්චි කරපු LOAD IDs registry එකට append කරනවා (timestamp සමඟ)."""
    from datetime import datetime
    import gspread
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    gc = _client(service_account_info)
    sh = _open(gc, sheet_url_or_key)
    header = ["LOAD ID", "Used At", "Source"]
    try:
        ws = sh.worksheet(LOADID_WS)
        if not ws.get_all_values():
            ws.update([header], value_input_option="USER_ENTERED")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=LOADID_WS, rows=500, cols=6)
        ws.update([header], value_input_option="USER_ENTERED")
    ws.append_rows([[str(x), ts, source] for x in load_ids],
                   value_input_option="USER_ENTERED")
    return ws.url


# --------------------------------------------------------------------------- #
# Monthly pick history
# --------------------------------------------------------------------------- #
def history_ws_name(dt=None) -> str:
    from datetime import datetime
    dt = dt or datetime.now()
    return f"History {dt.strftime('%Y-%m')}"


def append_history(service_account_info: dict, sheet_url_or_key: str, res: dict,
                   cfg=None, when=None) -> str:
    """එක pick run එකක OK pick lines, ඒ මාසෙට අදාල 'History YYYY-MM' tab එකට
    append කරනවා (date stamp සමඟ). මාසෙට අලුත් tab එකක් හැදෙනවා."""
    from datetime import datetime
    import gspread
    when = when or datetime.now()
    ws_name = history_ws_name(when)
    pick = res["pick"].copy()
    keep = [c for c in ["Material", "Material des", "Qty", "OBD",
                        "HJ Box Qty", "HJ Pcs Qty", "Pcs/Box"] if c in pick.columns]
    hist = pick[keep].copy()
    lmap = res.get("load_id_map", {}) or {}
    hist.insert(0, "LOAD ID", [str(lmap.get(o, o)) for o in pick["OBD"]])
    hist.insert(0, "Run Time", when.strftime("%H:%M:%S"))
    hist.insert(0, "Run Date", when.strftime("%Y-%m-%d"))

    gc = _client(service_account_info)
    sh = _open(gc, sheet_url_or_key)
    header = list(hist.columns)
    try:
        ws = sh.worksheet(ws_name)
        if not ws.get_all_values():
            ws.update([header], value_input_option="USER_ENTERED")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=ws_name, rows=max(len(hist) + 50, 100),
                              cols=max(len(header) + 2, 12))
        ws.update([header], value_input_option="USER_ENTERED")
    ws.append_rows(hist.fillna("").astype(str).values.tolist(),
                   value_input_option="USER_ENTERED")
    return ws.url


def list_history_months(service_account_info: dict, sheet_url_or_key: str) -> list:
    """'History YYYY-MM' tabs ඔක්කොම (අලුත්ම මුලින්)."""
    gc = _client(service_account_info)
    sh = _open(gc, sheet_url_or_key)
    months = [ws.title for ws in sh.worksheets() if ws.title.startswith("History ")]
    return sorted(months, reverse=True)


def read_history(service_account_info: dict, sheet_url_or_key: str,
                 ws_name: str) -> pd.DataFrame:
    """Monthly history tab එකක data කියවනවා."""
    return read_sheet(service_account_info, sheet_url_or_key, ws_name)
