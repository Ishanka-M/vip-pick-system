"""
gsheet.py — Google Sheet backend (ledger + registry + outputs)
==============================================================
Worksheets
  OUTBOUND_MASTER   : WMS master rows (append)
  OUTBOUND_DETAIL   : WMS detail rows (append)
  PALLET_LEDGER     : pallet-level pick — before / picked / balance
  DOC_REGISTRY      : process කරපු Invoice / DC numbers (duplicate gate)
  REJECTED_LOG      : pick කරන්න බැරි වුණ docs + හේතුව
  RUN_LOG           : run summary
"""
from __future__ import annotations

import re
from datetime import datetime
from typing import Any

import pandas as pd

import pick_engine as E

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

WS_MASTER = "OUTBOUND_MASTER"
WS_DETAIL = "OUTBOUND_DETAIL"
WS_LEDGER = "PALLET_LEDGER"
WS_REGISTRY = "DOC_REGISTRY"
WS_REJECT = "REJECTED_LOG"
WS_RUNLOG = "RUN_LOG"
WS_SETTINGS = "APP_SETTINGS"

LEDGER_COLS = E.ALLOC_COLS
REGISTRY_COLS = ["DOC_NUMBER", "DOC_TYPE", "DOC_DATE", "REF_NUMBER", "DOC_CHECK",
                 "LINES", "WMS_LINES",
                 "DOC_QTY", "PICKED_QTY", "PALLETS", "PLANTS", "RUN_ID", "PROCESSED_AT",
                 "SOURCE_FILE"]
REJECT_COLS = ["RUN_ID", "PROCESSED_AT", "DOC_NUMBER", "DOC_TYPE", "REASON", "DETAIL",
               "SOURCE_FILE"]
RUNLOG_COLS = ["RUN_ID", "PROCESSED_AT", "USER_NOTE", "PLANTS", "STRATEGY", "DOCS_OK",
               "DOCS_REJECTED", "MASTER_ROWS", "DETAIL_ROWS", "PALLETS", "TOTAL_QTY"]

_SHEETS = {
    WS_MASTER: ["RUN_ID", "PROCESSED_AT"] + E.MASTER_COLS,
    WS_DETAIL: ["RUN_ID", "PROCESSED_AT"] + E.DETAIL_COLS,
    WS_LEDGER: LEDGER_COLS,
    WS_REGISTRY: REGISTRY_COLS,
    WS_REJECT: REJECT_COLS,
    WS_RUNLOG: RUNLOG_COLS,
    WS_SETTINGS: ["KEY", "VALUE", "UPDATED_AT"],
}


# --------------------------------------------------------------------------- #
# connection
# --------------------------------------------------------------------------- #
def _key(sheet_key: str) -> str:
    s = str(sheet_key).strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", s)
    return m.group(1) if m else s


def open_book(sa_info: dict, sheet_key: str):
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(dict(sa_info), scopes=SCOPES)
    gc = gspread.authorize(creds)
    return gc.open_by_key(_key(sheet_key))


def _ensure(book, title: str, header: list[str]):
    try:
        ws = book.worksheet(title)
    except Exception:
        ws = book.add_worksheet(title=title, rows=1000, cols=max(10, len(header) + 2))
        ws.update(values=[header], range_name="A1", value_input_option="RAW")
        return ws
    first = ws.row_values(1)
    if not first:
        ws.update(values=[header], range_name="A1", value_input_option="RAW")
    return ws


def init_sheet(sa_info: dict, sheet_key: str) -> dict[str, Any]:
    book = open_book(sa_info, sheet_key)
    made = {}
    for title, header in _SHEETS.items():
        existed = title in [w.title for w in book.worksheets()]
        _ensure(book, title, header)
        made[title] = not existed
    made["url"] = book.url
    return made


# --------------------------------------------------------------------------- #
# read
# --------------------------------------------------------------------------- #
def read_ws(sa_info: dict, sheet_key: str, title: str) -> pd.DataFrame:
    book = open_book(sa_info, sheet_key)
    try:
        ws = book.worksheet(title)
    except Exception:
        return pd.DataFrame()
    rows = ws.get_all_values()
    if not rows:
        return pd.DataFrame()
    head, *body = rows
    return pd.DataFrame(body, columns=head)


def read_ledger(sa_info: dict, sheet_key: str) -> pd.DataFrame:
    return read_ws(sa_info, sheet_key, WS_LEDGER)


def read_processed_docs(sa_info: dict, sheet_key: str) -> set[str]:
    df = read_ws(sa_info, sheet_key, WS_REGISTRY)
    if df is None or not len(df) or "DOC_NUMBER" not in df.columns:
        return set()
    return {str(x).strip() for x in df["DOC_NUMBER"] if str(x).strip()}


def read_setting(sa_info: dict, sheet_key: str, key: str, default: str = "") -> str:
    """APP_SETTINGS worksheet එකෙන් එක value එකක් (email book වගේ)."""
    df = read_ws(sa_info, sheet_key, WS_SETTINGS)
    if df is None or not len(df) or "KEY" not in df.columns:
        return default
    hit = df[df["KEY"].astype(str) == str(key)]
    return str(hit.iloc[-1]["VALUE"]) if len(hit) else default


def save_setting(sa_info: dict, sheet_key: str, key: str, value: str) -> None:
    book = open_book(sa_info, sheet_key)
    ws = _ensure(book, WS_SETTINGS, _SHEETS[WS_SETTINGS])
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = ws.get_all_values()
    head = rows[0] if rows else _SHEETS[WS_SETTINGS]
    body = [r for r in rows[1:] if (r[0] if r else "") != str(key)]
    body.append([str(key), str(value), now])
    ws.clear()
    ws.update(values=[head] + body, range_name="A1", value_input_option="RAW")


# --------------------------------------------------------------------------- #
# write
# --------------------------------------------------------------------------- #
def _append(ws, df: pd.DataFrame, header: list[str]) -> int:
    if df is None or not len(df):
        return 0
    d = df.reindex(columns=header).fillna("")
    values = [["" if pd.isna(v) else str(v) for v in row] for row in d.values.tolist()]
    for i in range(0, len(values), 2000):
        ws.append_rows(values[i:i + 2000], value_input_option="RAW",
                       insert_data_option="INSERT_ROWS")
    return len(values)


def save_run(sa_info: dict, sheet_key: str, res: dict, cfg: E.EngineConfig,
             note: str = "") -> dict[str, Any]:
    book = open_book(sa_info, sheet_key)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    run_id = res.get("run_id", "")
    out: dict[str, Any] = {"url": book.url}

    master = res["master"].copy()
    if len(master):
        master.insert(0, "PROCESSED_AT", now)
        master.insert(0, "RUN_ID", run_id)
    detail = res["detail"].copy()
    if len(detail):
        detail.insert(0, "PROCESSED_AT", now)
        detail.insert(0, "RUN_ID", run_id)

    out["master"] = _append(_ensure(book, WS_MASTER, _SHEETS[WS_MASTER]), master,
                            _SHEETS[WS_MASTER])
    out["detail"] = _append(_ensure(book, WS_DETAIL, _SHEETS[WS_DETAIL]), detail,
                            _SHEETS[WS_DETAIL])
    out["ledger"] = _append(_ensure(book, WS_LEDGER, LEDGER_COLS), res["allocations"],
                            LEDGER_COLS)

    reg = res["accepted"].copy()
    if len(reg):
        reg["RUN_ID"] = run_id
        reg["PROCESSED_AT"] = now
    out["registry"] = _append(_ensure(book, WS_REGISTRY, REGISTRY_COLS), reg, REGISTRY_COLS)

    rej = res["rejected"].copy()
    if len(rej):
        rej["RUN_ID"] = run_id
        rej["PROCESSED_AT"] = now
    out["rejected"] = _append(_ensure(book, WS_REJECT, REJECT_COLS), rej, REJECT_COLS)

    alloc = res["allocations"]
    runlog = pd.DataFrame([{
        "RUN_ID": run_id, "PROCESSED_AT": now, "USER_NOTE": note,
        "PLANTS": ", ".join(cfg.plants), "STRATEGY": cfg.strategy,
        "DOCS_OK": len(res["accepted"]), "DOCS_REJECTED": len(res["rejected"]),
        "MASTER_ROWS": len(res["master"]), "DETAIL_ROWS": len(res["detail"]),
        "PALLETS": int(alloc["PALLET"].nunique()) if len(alloc) else 0,
        "TOTAL_QTY": float(pd.to_numeric(alloc["QTY_PICKED"], errors="coerce").sum())
        if len(alloc) else 0,
    }])
    out["runlog"] = _append(_ensure(book, WS_RUNLOG, RUNLOG_COLS), runlog, RUNLOG_COLS)
    return out


def delete_run(sa_info: dict, sheet_key: str, run_id: str) -> dict[str, int]:
    """වැරදුනු run එකක් undo — ලේඩ්ජරයෙන් සහ registry එකෙන් අයින් කරනවා."""
    book = open_book(sa_info, sheet_key)
    removed: dict[str, int] = {}
    for title in (WS_MASTER, WS_DETAIL, WS_LEDGER, WS_REGISTRY, WS_REJECT, WS_RUNLOG):
        try:
            ws = book.worksheet(title)
        except Exception:
            continue
        rows = ws.get_all_values()
        if not rows:
            continue
        head, *body = rows
        if "RUN_ID" not in head:
            continue
        i = head.index("RUN_ID")
        keep = [r for r in body if (r[i] if i < len(r) else "") != run_id]
        removed[title] = len(body) - len(keep)
        ws.clear()
        ws.update(values=[head] + keep, range_name="A1", value_input_option="RAW")
    return removed


RESET_SCOPES = {
    "outputs": [WS_MASTER, WS_DETAIL],
    "ledger": [WS_LEDGER],
    "registry": [WS_REGISTRY],
    "rejected": [WS_REJECT],
    "runlog": [WS_RUNLOG],
    "settings": [WS_SETTINGS],
}


def reset_data(sa_info: dict, sheet_key: str, scope: list[str]) -> list[str]:
    book = open_book(sa_info, sheet_key)
    pick = RESET_SCOPES
    done: list[str] = []
    for s in scope:
        for title in pick.get(s, []):
            try:
                ws = book.worksheet(title)
            except Exception:
                continue
            ws.clear()
            ws.update(values=[_SHEETS[title]], range_name="A1", value_input_option="RAW")
            done.append(title)
    return done


def reset_all(sa_info: dict, sheet_key: str, keep_settings: bool = True) -> dict[str, Any]:
    """
    ⚠️ FULL DB RESET — ledger · registry · outputs · rejected · run log ඔක්කොම clear.
    Header row එක විතරක් ඉතුරු වෙනවා (worksheet delete වෙන්නේ නෑ).
    """
    scope = [k for k in RESET_SCOPES if not (keep_settings and k == "settings")]
    done = reset_data(sa_info, sheet_key, scope)
    return {"cleared": done, "count": len(done), "at":
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
