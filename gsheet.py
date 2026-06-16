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
    sh = (gc.open_by_url(sheet_url_or_key) if sheet_url_or_key.startswith("http")
          else gc.open_by_key(sheet_url_or_key))
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
