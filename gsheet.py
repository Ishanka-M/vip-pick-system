"""
gsheet.py — Google Sheet backend + API manager (multi-user safe)
================================================================
Worksheets
  OUTBOUND_MASTER · OUTBOUND_DETAIL · PALLET_LEDGER · DOC_REGISTRY
  REJECTED_LOG · RUN_LOG · SKU_MASTER · APP_SETTINGS · _LOCKS

API management
  * හැම call එකකටම retry + exponential backoff (429 / 5xx)
  * Read cache — TTL එකක් එක්ක (user කීපදෙනෙක් read කරද්දී quota ඉතුරු වෙනවා)
  * Write lock — user දෙන්නෙක් එකවර save කරද්දී corrupt වෙන එක නවත්තනවා
  * Save කරද්දී lock එක ඇතුලේම duplicate re-check එකක්
"""
from __future__ import annotations

import random
import re
import threading
import time
import uuid
from collections import deque
from datetime import datetime, timedelta
from typing import Any, Callable

import pandas as pd

import invoice_register as R
import pick_engine as E

# Bumped whenever this module's public surface changes; app.py refuses to run
# against a stale copy instead of dying with a redacted TypeError.
API = 10

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
WS_SKU = "SKU_MASTER"
WS_INV_SUM = "INVOICE_SUMMARY"
WS_INV_DET = "INVOICE_DETAIL"
WS_LOCK = "_LOCKS"

LEDGER_COLS = E.ALLOC_COLS
REGISTRY_COLS = ["DOC_NUMBER", "DOC_TYPE", "DOC_DATE", "REF_NUMBER", "DOC_CHECK",
                 "LINES", "WMS_LINES", "DOC_QTY", "PICKED_QTY", "WMS_QTY", "VERIFY",
                 "PALLETS", "PLANTS", "RUN_ID", "PROCESSED_AT", "SOURCE_FILE"]
REJECT_COLS = ["RUN_ID", "PROCESSED_AT", "DOC_NUMBER", "DOC_TYPE", "REASON", "DETAIL",
               "SOURCE_FILE"]
RUNLOG_COLS = ["RUN_ID", "PROCESSED_AT", "USER_NOTE", "PLANTS", "STRATEGY", "DOCS_OK",
               "DOCS_REJECTED", "MASTER_ROWS", "DETAIL_ROWS", "PALLETS", "TOTAL_QTY"]
SKU_COLS = ["ITEM_NUMBER", "BASE_ID", "MATCH_KEY", "ITEM_DESCRIPTION",
            "UPDATED_AT", "UPDATED_BY", "SOURCE"]
LOCK_COLS = ["NAME", "OWNER", "ACQUIRED_AT", "EXPIRES_AT"]

_SHEETS = {
    WS_MASTER: ["RUN_ID", "PROCESSED_AT"] + E.MASTER_COLS,
    WS_DETAIL: ["RUN_ID", "PROCESSED_AT"] + E.DETAIL_COLS,
    WS_LEDGER: LEDGER_COLS,
    WS_REGISTRY: REGISTRY_COLS,
    WS_REJECT: REJECT_COLS,
    WS_RUNLOG: RUNLOG_COLS,
    WS_SKU: SKU_COLS,
    WS_INV_SUM: R.SUMMARY_COLS,
    WS_INV_DET: R.DETAIL_COLS,
    WS_SETTINGS: ["KEY", "VALUE", "UPDATED_AT"],
    WS_LOCK: LOCK_COLS,
}

# LOAD_ID තියෙන column එක worksheet එකට අනුව
_LOAD_COL = {
    WS_MASTER: "DISPLAY_ORDER_NUMBER",
    WS_DETAIL: "DISPLAY_ORDER_NUMBER",
    WS_LEDGER: "DOC_NUMBER",
    WS_REGISTRY: "DOC_NUMBER",
    WS_REJECT: "DOC_NUMBER",
}

# --------------------------------------------------------------------------- #
# API manager — retry · cache · stats
# --------------------------------------------------------------------------- #
STATS: dict[str, Any] = {"calls": 0, "retries": 0, "cache_hits": 0, "errors": 0,
                         "batched": 0, "throttled": 0.0, "last_error": "",
                         "last_call": ""}

_CACHE: dict[tuple, tuple[float, Any]] = {}
_CACHE_TTL = 45.0
MAX_RETRY = 5

# Google allows 60 read + 60 write requests per minute per user. Staying a few
# under and *pacing* is far cheaper than getting a 429 and backing off for
# seconds — especially with several people on the app at once.
RATE_LIMIT = 55
_WINDOW: deque = deque()
_RATE_LOCK = threading.Lock()


def _throttle() -> None:
    """Block just long enough to keep the last 60 s under RATE_LIMIT calls."""
    while True:
        with _RATE_LOCK:
            now = time.time()
            while _WINDOW and now - _WINDOW[0] > 60.0:
                _WINDOW.popleft()
            if len(_WINDOW) < RATE_LIMIT:
                _WINDOW.append(now)
                return
            wait = 60.0 - (now - _WINDOW[0]) + 0.05
        STATS["throttled"] = round(STATS["throttled"] + wait, 2)
        time.sleep(min(wait, 5.0))


def quota_used() -> int:
    """Calls in the last 60 seconds."""
    now = time.time()
    with _RATE_LOCK:
        while _WINDOW and now - _WINDOW[0] > 60.0:
            _WINDOW.popleft()
        return len(_WINDOW)


def _key(sheet_key: str) -> str:
    s = str(sheet_key).strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", s)
    return m.group(1) if m else s


def cache_clear(sheet_key: str | None = None) -> None:
    if sheet_key is None:
        _CACHE.clear()
        _TITLES.clear()
        return
    _TITLES.pop(_key(sheet_key), None)
    for k in [k for k in _CACHE if k[0] == _key(sheet_key)]:
        _CACHE.pop(k, None)


def set_cache_ttl(seconds: float) -> None:
    global _CACHE_TTL
    _CACHE_TTL = max(0.0, float(seconds))


def _retryable(ex: Exception) -> bool:
    code = getattr(getattr(ex, "response", None), "status_code", None)
    if code in (429, 500, 502, 503, 504):
        return True
    txt = str(ex)
    return any(s in txt for s in ("429", "500", "502", "503", "504", "Quota exceeded",
                                  "RATE_LIMIT_EXCEEDED", "internal error",
                                  "currently unavailable"))


def api(fn: Callable, *a, **kw):
    """Google API call wrapper — 429 / 5xx වලට exponential backoff + jitter."""
    delay = 1.0
    for attempt in range(MAX_RETRY):
        try:
            _throttle()
            STATS["calls"] += 1
            STATS["last_call"] = datetime.now().strftime("%H:%M:%S")
            return fn(*a, **kw)
        except Exception as ex:                      # noqa: BLE001
            if not _retryable(ex) or attempt == MAX_RETRY - 1:
                STATS["errors"] += 1
                STATS["last_error"] = f"{type(ex).__name__}: {str(ex)[:200]}"
                raise
            STATS["retries"] += 1
            time.sleep(delay + random.uniform(0, 0.6))
            delay = min(delay * 2, 20.0)


# --------------------------------------------------------------------------- #
# connection
# --------------------------------------------------------------------------- #
def open_book(sa_info: dict, sheet_key: str):
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(dict(sa_info), scopes=SCOPES)
    gc = gspread.authorize(creds)
    return api(gc.open_by_key, _key(sheet_key))


def _ensure(book, title: str, header: list[str]):
    try:
        ws = api(book.worksheet, title)
    except Exception:
        ws = api(book.add_worksheet, title=title, rows=1000, cols=max(10, len(header) + 2))
        api(ws.update, values=[header], range_name="A1", value_input_option="RAW")
        return ws
    if not api(ws.row_values, 1):
        api(ws.update, values=[header], range_name="A1", value_input_option="RAW")
    return ws


def sheet_header(ws, expected: list[str]) -> list[str]:
    """
    The header the sheet *actually* has, widened to cover `expected`.

    A release that adds a column used to break the sheet silently: `_ensure`
    only writes a header when the sheet is empty, so an old header stayed put
    while rows were written in the new, longer column order. Every value after
    the insertion point then landed one column to the left of its name — which
    is how DOC_REGISTRY ended up with a plant name under `PROCESSED_AT`.

    So: read the real header, and append anything missing **at the end**, where
    it cannot disturb the alignment of rows already written. Callers then write
    rows in this order rather than blindly in `expected` order.
    """
    have = [h for h in (api(ws.row_values, 1) or []) if str(h).strip()]
    if not have:
        api(ws.update, values=[expected], range_name="A1", value_input_option="RAW")
        return list(expected)
    missing = [c for c in expected if c not in have]
    if missing:
        merged = have + missing
        api(ws.update, values=[merged], range_name="A1", value_input_option="RAW")
        return merged
    return have


def init_sheet(sa_info: dict, sheet_key: str) -> dict[str, Any]:
    book = open_book(sa_info, sheet_key)
    have = _titles(book, sheet_key)
    made: dict[str, Any] = {}
    for title, header in _SHEETS.items():
        _ensure(book, title, header)
        made[title] = title not in have
    made["url"] = book.url
    cache_clear(sheet_key)
    _titles(book, sheet_key)
    return made


def health(sa_info: dict, sheet_key: str) -> dict[str, Any]:
    t0 = time.time()
    book = open_book(sa_info, sheet_key)
    sheets = [w.title for w in api(book.worksheets)]
    return {"ok": True, "ms": int((time.time() - t0) * 1000), "url": book.url,
            "worksheets": sheets, "missing": [t for t in _SHEETS if t not in sheets],
            "cached": len([k for k in _CACHE if k[0] == _key(sheet_key)]),
            "ttl": _CACHE_TTL, "quota_used": quota_used(), "quota_limit": RATE_LIMIT,
            **STATS}


# --------------------------------------------------------------------------- #
# read (cached)
# --------------------------------------------------------------------------- #
def read_ws(sa_info: dict, sheet_key: str, title: str, use_cache: bool = True,
            ttl: float | None = None) -> pd.DataFrame:
    ck = (_key(sheet_key), title)
    life = _CACHE_TTL if ttl is None else ttl
    if use_cache and ck in _CACHE:
        ts, df = _CACHE[ck]
        if time.time() - ts < life:
            STATS["cache_hits"] += 1
            return df.copy()

    book = open_book(sa_info, sheet_key)
    try:
        ws = api(book.worksheet, title)
    except Exception:
        df = pd.DataFrame()
    else:
        df = _frame(api(ws.get_all_values))
    _CACHE[ck] = (time.time(), df.copy())
    return df


_TITLES: dict[str, tuple[float, set[str]]] = {}


def _titles(book, sheet_key: str, ttl: float = 300.0) -> set[str]:
    """Worksheet names, cached — this list changes about once a year."""
    ck = _key(sheet_key)
    hit = _TITLES.get(ck)
    if hit and time.time() - hit[0] < ttl:
        return hit[1]
    names = {w.title for w in api(book.worksheets)}
    _TITLES[ck] = (time.time(), names)
    return names


def _frame(rows: list[list[str]]) -> pd.DataFrame:
    """
    Sheet values -> DataFrame, deliberately **object** dtype.

    A sheet hands back nothing but strings, and pandas 3 now infers the `str`
    dtype for that — a dtype that *rejects* a number being written into it
    ("Invalid value '0.0' for dtype 'str'"). Every routine here writes the odd
    number back into a frame it read from the sheet (deleting a load sets
    PICKED_QTY to 0, the backfill writes quantities), and all of it worked
    until pandas 3 arrived. Keeping these frames as object restores the
    behaviour the rest of the module is written against; values are turned back
    into strings on the way out anyway.
    """
    if not rows:
        return pd.DataFrame()
    head, *body = rows
    head = [h if h else f"COL{i}" for i, h in enumerate(head)]
    body = [r + [""] * (len(head) - len(r)) for r in body]
    return pd.DataFrame([r[:len(head)] for r in body], columns=head, dtype=object)


def read_many(sa_info: dict, sheet_key: str, titles: list[str],
              use_cache: bool = True, ttl: float | None = None
              ) -> dict[str, pd.DataFrame]:
    """
    Several worksheets in a single request.

    Reading six sheets one at a time is six requests against a 60/minute quota;
    values_batch_get makes it one. Anything already cached is not asked for.
    """
    life = _CACHE_TTL if ttl is None else ttl
    out: dict[str, pd.DataFrame] = {}
    want: list[str] = []
    for t in titles:
        ck = (_key(sheet_key), t)
        if use_cache and ck in _CACHE and time.time() - _CACHE[ck][0] < life:
            STATS["cache_hits"] += 1
            out[t] = _CACHE[ck][1].copy()
        else:
            want.append(t)
    if not want:
        return out

    book = open_book(sa_info, sheet_key)
    have = _titles(book, sheet_key)
    ask = [t for t in want if t in have]
    for t in want:
        if t not in have:
            out[t] = pd.DataFrame()
            _CACHE[(_key(sheet_key), t)] = (time.time(), pd.DataFrame())
    if ask:
        try:
            res = api(book.values_batch_get, [f"'{t}'" for t in ask]) or {}
            ranges = res.get("valueRanges", []) if isinstance(res, dict) else []
        except Exception:
            ranges = []                       # fall back to one read per sheet
        if len(ranges) == len(ask):
            STATS["batched"] += max(0, len(ask) - 1)
            for t, vr in zip(ask, ranges):
                df = _frame((vr or {}).get("values") or [])
                out[t] = df
                _CACHE[(_key(sheet_key), t)] = (time.time(), df.copy())
        else:
            for t in ask:
                out[t] = read_ws(sa_info, sheet_key, t, use_cache=False)
    return out


def read_ledger(sa_info: dict, sheet_key: str) -> pd.DataFrame:
    return read_ws(sa_info, sheet_key, WS_LEDGER)


def read_processed_docs(sa_info: dict, sheet_key: str, fresh: bool = False) -> set[str]:
    df = read_ws(sa_info, sheet_key, WS_REGISTRY, use_cache=not fresh)
    if df is None or not len(df) or "DOC_NUMBER" not in df.columns:
        return set()
    return {str(x).strip() for x in df["DOC_NUMBER"] if str(x).strip()}


def read_setting(sa_info: dict, sheet_key: str, key: str, default: str = "") -> str:
    df = read_ws(sa_info, sheet_key, WS_SETTINGS)
    if df is None or not len(df) or "KEY" not in df.columns:
        return default
    hit = df[df["KEY"].astype(str) == str(key)]
    return str(hit.iloc[-1]["VALUE"]) if len(hit) else default


def save_setting(sa_info: dict, sheet_key: str, key: str, value: str) -> None:
    book = open_book(sa_info, sheet_key)
    ws = _ensure(book, WS_SETTINGS, _SHEETS[WS_SETTINGS])
    rows = api(ws.get_all_values)
    head = rows[0] if rows else _SHEETS[WS_SETTINGS]
    body = [r for r in rows[1:] if (r[0] if r else "") != str(key)]
    body.append([str(key), str(value), datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    write_table(book, WS_SETTINGS, head, pd.DataFrame(body, columns=head),
                prev_rows=len(rows) - 1 if rows else 0)
    cache_clear(sheet_key)


# --------------------------------------------------------------------------- #
# write lock (multi-user)
# --------------------------------------------------------------------------- #
class LockBusy(RuntimeError):
    pass


class sheet_lock:
    """
    with sheet_lock(sa, key, "SAVE"):  ... write ...

    `_LOCKS` worksheet එකෙන් soft lock එකක්. තව කෙනෙක් save කරමින් නම් රැඳිලා,
    timeout එකේදී LockBusy. Expire වුණ lock auto-clear වෙනවා.
    """

    def __init__(self, sa_info: dict, sheet_key: str, name: str = "SAVE",
                 owner: str = "", wait: float = 25.0, hold: float = 90.0):
        self.sa, self.key, self.name = sa_info, sheet_key, name
        self.owner = owner or f"user-{uuid.uuid4().hex[:6]}"
        self.wait, self.hold = wait, hold
        self.ws = None
        self.got = False
        self._last = 8

    def _rows(self):
        rows = api(self.ws.get_all_values)
        head = rows[0] if rows else LOCK_COLS
        return head, [r + [""] * (len(head) - len(r)) for r in rows[1:]]

    def _write(self, head, body):
        # one request, padded over the old rows — a lock that needs two requests
        # to take is a lock that two people can both think they hold
        n = max(len(body) + 1, self._last + 1)
        rows = [head] + body + [[""] * len(head) for _ in range(n - len(body) - 1)]
        self._last = len(body)
        api(self.ws.update, values=rows,
            range_name=f"A1:{_a1(len(head))}{len(rows)}", value_input_option="RAW")

    def acquire(self) -> bool:
        book = open_book(self.sa, self.key)
        self.ws = _ensure(book, WS_LOCK, LOCK_COLS)
        deadline = time.time() + self.wait
        while True:
            head, body = self._rows()
            now = datetime.now()
            keep, busy = [], False
            for r in body:
                if not r or not r[0]:
                    continue
                try:
                    exp = datetime.strptime(r[3], "%Y-%m-%d %H:%M:%S")
                except Exception:
                    exp = now - timedelta(seconds=1)
                if exp <= now:
                    continue                          # stale -> drop
                keep.append(r)
                if r[0] == self.name:
                    busy = True
            if not busy:
                keep.append([self.name, self.owner, now.strftime("%Y-%m-%d %H:%M:%S"),
                             (now + timedelta(seconds=self.hold)).strftime(
                                 "%Y-%m-%d %H:%M:%S")])
                self._write(head, keep)
                time.sleep(0.6)                       # settle, then confirm it's ours
                _, chk = self._rows()
                mine = [r for r in chk if r and r[0] == self.name]
                if len(mine) == 1 and mine[0][1] == self.owner:
                    self.got = True
                    return True
            if time.time() >= deadline:
                return False
            time.sleep(1.2 + random.uniform(0, 0.8))

    def release(self) -> None:
        if not self.got or self.ws is None:
            return
        try:
            head, body = self._rows()
            self._write(head, [r for r in body
                               if not (r and r[0] == self.name and r[1] == self.owner)])
        except Exception:
            pass
        finally:
            self.got = False

    def __enter__(self):
        if not self.acquire():
            raise LockBusy("Someone else is saving right now - try again in a minute.")
        return self

    def __exit__(self, *exc):
        self.release()
        return False


def active_locks(sa_info: dict, sheet_key: str) -> pd.DataFrame:
    df = read_ws(sa_info, sheet_key, WS_LOCK, use_cache=False)
    if df is None or not len(df):
        return pd.DataFrame(columns=LOCK_COLS)
    now = datetime.now()
    keep = []
    for _, r in df.iterrows():
        try:
            if datetime.strptime(str(r.get("EXPIRES_AT", "")), "%Y-%m-%d %H:%M:%S") > now:
                keep.append(r)
        except Exception:
            continue
    return pd.DataFrame(keep, columns=df.columns) if keep else pd.DataFrame(columns=LOCK_COLS)


def clear_locks(sa_info: dict, sheet_key: str) -> int:
    book = open_book(sa_info, sheet_key)
    ws = _ensure(book, WS_LOCK, LOCK_COLS)
    n = max(0, len(api(ws.get_all_values)) - 1)
    write_table(book, WS_LOCK, LOCK_COLS, pd.DataFrame(columns=LOCK_COLS),
                prev_rows=n)
    cache_clear(sheet_key)
    return n


# --------------------------------------------------------------------------- #
# write
# --------------------------------------------------------------------------- #
def _a1(n_cols: int) -> str:
    col, x = "", n_cols
    while x:
        x, r = divmod(x - 1, 26)
        col = chr(65 + r) + col
    return col


def write_table(book, title: str, header: list[str], df: pd.DataFrame,
                prev_rows: int | None = None) -> int:
    """
    Replace a worksheet's contents in **one** request.

    `clear()` followed by `update()` is two requests and leaves a window where
    the sheet is empty — if anything fails in between, the data is gone. Writing
    the new rows plus enough blank padding to cover the old ones does the same
    job atomically and halves the request count.
    """
    ws = _ensure(book, title, header)
    rows = [header]
    if df is not None and len(df):
        rows += [["" if pd.isna(v) else str(v) for v in r]
                 for r in df.reindex(columns=header).values.tolist()]
    n_new = len(rows)
    old = prev_rows if prev_rows is not None else max(0, ws.row_count - 1)
    pad = max(0, (old + 1) - n_new)
    blank = [""] * len(header)
    body = rows + [list(blank) for _ in range(pad)]
    if ws.row_count < len(body):
        api(ws.resize, rows=len(body) + 50, cols=max(ws.col_count, len(header)))
    api(ws.update, values=body, range_name=f"A1:{_a1(len(header))}{len(body)}",
        value_input_option="RAW")
    return n_new - 1


def _append(ws, df: pd.DataFrame, header: list[str]) -> int:
    if df is None or not len(df):
        return 0
    # write in the sheet's own column order, not ours — see sheet_header()
    d = df.reindex(columns=sheet_header(ws, header)).fillna("")
    values = [["" if pd.isna(v) else str(v) for v in row] for row in d.values.tolist()]
    for i in range(0, len(values), 2000):
        api(ws.append_rows, values[i:i + 2000], value_input_option="RAW",
            insert_data_option="INSERT_ROWS")
    return len(values)


def save_run(sa_info: dict, sheet_key: str, res: dict, cfg: E.EngineConfig,
             note: str = "", owner: str = "", skip_dupes: bool = True) -> dict[str, Any]:
    """
    Lock එකක් ඇතුලේ save + ඒ ඇතුලේම registry re-check.
    User දෙන්නෙක් එකවර එකම Invoice එක save කරන්න ගියත් එකක් විතරයි යන්නේ.
    """
    out: dict[str, Any] = {}
    with sheet_lock(sa_info, sheet_key, "SAVE", owner=owner) as lk:
        book = open_book(sa_info, sheet_key)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        run_id = res.get("run_id", "")
        out["url"] = book.url
        out["owner"] = lk.owner

        master, detail = res["master"].copy(), res["detail"].copy()
        alloc, reg = res["allocations"].copy(), res["accepted"].copy()
        rej = res["rejected"].copy()

        clash: list[str] = []
        if skip_dupes and len(reg):
            done = read_processed_docs(sa_info, sheet_key, fresh=True)
            clash = [d for d in reg["DOC_NUMBER"].astype(str) if d in done]
            if clash:
                reg = reg[~reg["DOC_NUMBER"].astype(str).isin(clash)]
                master = master[~master["DISPLAY_ORDER_NUMBER"].astype(str).isin(clash)]
                detail = detail[~detail["DISPLAY_ORDER_NUMBER"].astype(str).isin(clash)]
                if len(alloc):
                    alloc = alloc[~alloc["DOC_NUMBER"].astype(str).isin(clash)]
                rej = pd.concat([rej, pd.DataFrame([{
                    "DOC_NUMBER": d, "DOC_TYPE": "", "REASON": "DUPLICATE (other user)",
                    "DETAIL": "Another user saved this document while this run was in progress",
                    "SOURCE_FILE": ""} for d in clash])], ignore_index=True)
        out["skipped"] = clash

        if len(master):
            master.insert(0, "PROCESSED_AT", now)
            master.insert(0, "RUN_ID", run_id)
        if len(detail):
            detail.insert(0, "PROCESSED_AT", now)
            detail.insert(0, "RUN_ID", run_id)
        if len(reg):
            reg["RUN_ID"] = run_id
            reg["PROCESSED_AT"] = now
        if len(rej):
            rej["RUN_ID"] = run_id
            rej["PROCESSED_AT"] = now

        out["master"] = _append(_ensure(book, WS_MASTER, _SHEETS[WS_MASTER]), master,
                                _SHEETS[WS_MASTER])
        out["detail"] = _append(_ensure(book, WS_DETAIL, _SHEETS[WS_DETAIL]), detail,
                                _SHEETS[WS_DETAIL])
        out["ledger"] = _append(_ensure(book, WS_LEDGER, LEDGER_COLS), alloc, LEDGER_COLS)
        out["registry"] = _append(_ensure(book, WS_REGISTRY, REGISTRY_COLS), reg,
                                  REGISTRY_COLS)
        out["rejected"] = _append(_ensure(book, WS_REJECT, REJECT_COLS), rej, REJECT_COLS)

        runlog = pd.DataFrame([{
            "RUN_ID": run_id, "PROCESSED_AT": now, "USER_NOTE": note,
            "PLANTS": ", ".join(cfg.plants), "STRATEGY": cfg.strategy,
            "DOCS_OK": len(reg), "DOCS_REJECTED": len(rej),
            "MASTER_ROWS": len(master), "DETAIL_ROWS": len(detail),
            "PALLETS": int(alloc["PALLET"].nunique()) if len(alloc) else 0,
            "TOTAL_QTY": float(pd.to_numeric(alloc["QTY_PICKED"], errors="coerce").sum())
            if len(alloc) else 0,
        }])
        out["runlog"] = _append(_ensure(book, WS_RUNLOG, RUNLOG_COLS), runlog, RUNLOG_COLS)

    cache_clear(sheet_key)
    return out


# --------------------------------------------------------------------------- #
# LOAD_ID — read / delete
# --------------------------------------------------------------------------- #
def list_loads(sa_info: dict, sheet_key: str, fresh: bool = False) -> pd.DataFrame:
    df = read_ws(sa_info, sheet_key, WS_REGISTRY, use_cache=not fresh)
    if df is None or not len(df):
        return pd.DataFrame(columns=REGISTRY_COLS)
    cols = [c for c in ["DOC_NUMBER", "DOC_TYPE", "DOC_DATE", "LINES", "DOC_QTY",
                        "PICKED_QTY", "VERIFY", "PALLETS", "PLANTS", "RUN_ID",
                        "PROCESSED_AT", "SOURCE_FILE"] if c in df.columns]
    return df[cols].iloc[::-1].reset_index(drop=True)


def read_load(sa_info: dict, sheet_key: str, load_id: str,
              fresh: bool = False) -> dict[str, pd.DataFrame]:
    """LOAD_ID එකකට අදාල ඔක්කොම rows — master · detail · ledger · registry · rejected."""
    lid = str(load_id).strip()
    out: dict[str, pd.DataFrame] = {}
    got = read_many(sa_info, sheet_key, list(_LOAD_COL), use_cache=not fresh)
    for title, col in _LOAD_COL.items():
        df = got.get(title, pd.DataFrame())
        if df is None or not len(df) or col not in df.columns:
            out[title] = pd.DataFrame()
            continue
        out[title] = df[df[col].astype(str).str.strip() == lid].reset_index(drop=True)
    return out


def delete_load(sa_info: dict, sheet_key: str, load_id: str,
                owner: str = "") -> dict[str, int]:
    """
    LOAD_ID එකක් DB එකෙන් සම්පූර්ණයෙන් අයින් —
    ledger එකෙන් යන නිසා **pallet balance ආපහු එනවා**,
    registry එකෙන් යන නිසා **ආපහු pick කරන්නත් පුළුවන්**.
    """
    lid = str(load_id).strip()
    removed: dict[str, int] = {}
    with sheet_lock(sa_info, sheet_key, "SAVE", owner=owner):
        book = open_book(sa_info, sheet_key)
        for title, col in _LOAD_COL.items():
            try:
                ws = api(book.worksheet, title)
            except Exception:
                continue
            rows = api(ws.get_all_values)
            if not rows:
                continue
            head, *body = rows
            if col not in head:
                continue
            i = head.index(col)
            keep = [r for r in body if (r[i].strip() if i < len(r) else "") != lid]
            n = len(body) - len(keep)
            removed[title] = n
            if n:
                write_table(book, title, head, pd.DataFrame(keep, columns=head),
                            prev_rows=len(body))
        # the register must stop claiming this invoice was picked
        cur = read_ws(sa_info, sheet_key, WS_INV_SUM, use_cache=False)
        if cur is not None and len(cur):
            _replace_ws(book, WS_INV_SUM, R.SUMMARY_COLS,
                        R.mark_unpicked(cur, lid, "Load deleted"))
    cache_clear(sheet_key)
    return removed


def delete_run(sa_info: dict, sheet_key: str, run_id: str) -> dict[str, int]:
    """වැරදුනු run එකක් undo."""
    removed: dict[str, int] = {}
    with sheet_lock(sa_info, sheet_key, "SAVE"):
        book = open_book(sa_info, sheet_key)
        for title in (WS_MASTER, WS_DETAIL, WS_LEDGER, WS_REGISTRY, WS_REJECT, WS_RUNLOG):
            try:
                ws = api(book.worksheet, title)
            except Exception:
                continue
            rows = api(ws.get_all_values)
            if not rows:
                continue
            head, *body = rows
            if "RUN_ID" not in head:
                continue
            i = head.index("RUN_ID")
            keep = [r for r in body if (r[i] if i < len(r) else "") != run_id]
            removed[title] = len(body) - len(keep)
            write_table(book, title, head, pd.DataFrame(keep, columns=head),
                        prev_rows=len(body))
    cache_clear(sheet_key)
    return removed


# --------------------------------------------------------------------------- #
# SKU master
# --------------------------------------------------------------------------- #
def read_sku(sa_info: dict, sheet_key: str, fresh: bool = False) -> pd.DataFrame:
    return read_ws(sa_info, sheet_key, WS_SKU, use_cache=not fresh)


def save_sku(sa_info: dict, sheet_key: str, master: pd.DataFrame,
             owner: str = "") -> dict[str, Any]:
    """SKU master full replace (upsert එක app එකේ කරලා එවනවා) — lock එකක් ඇතුලේ."""
    head = list(master.columns) if (master is not None and len(master)) else SKU_COLS
    with sheet_lock(sa_info, sheet_key, "SKU", owner=owner):
        book = open_book(sa_info, sheet_key)
        n = write_table(book, WS_SKU, head, master)
    cache_clear(sheet_key)
    return {"rows": n}


# --------------------------------------------------------------------------- #
# Invoice register
# --------------------------------------------------------------------------- #
def read_register(sa_info: dict, sheet_key: str,
                  fresh: bool = False) -> tuple[pd.DataFrame, pd.DataFrame]:
    got = read_many(sa_info, sheet_key, [WS_INV_SUM, WS_INV_DET], use_cache=not fresh)
    return got.get(WS_INV_SUM, pd.DataFrame()), got.get(WS_INV_DET, pd.DataFrame())


def read_for_run(sa_info: dict, sheet_key: str) -> tuple[pd.DataFrame, set[str]]:
    """Ledger + processed-document list in one request."""
    got = read_many(sa_info, sheet_key, [WS_LEDGER, WS_REGISTRY])
    reg = got.get(WS_REGISTRY, pd.DataFrame())
    done = ({str(x).strip() for x in reg["DOC_NUMBER"] if str(x).strip()}
            if len(reg) and "DOC_NUMBER" in reg.columns else set())
    return got.get(WS_LEDGER, pd.DataFrame()), done


def _replace_ws(book, title: str, header: list[str], df: pd.DataFrame,
                prev_rows: int | None = None) -> int:
    return write_table(book, title, header, df, prev_rows)


def save_register(sa_info: dict, sheet_key: str, summary: pd.DataFrame,
                  details: pd.DataFrame, owner: str = "") -> dict[str, Any]:
    """
    Merge happens inside the lock against a fresh read — two people picking
    different invoices at the same time both end up in the register.
    """
    with sheet_lock(sa_info, sheet_key, "REGISTER", owner=owner):
        book = open_book(sa_info, sheet_key)
        old_s = read_ws(sa_info, sheet_key, WS_INV_SUM, use_cache=False)
        old_d = read_ws(sa_info, sheet_key, WS_INV_DET, use_cache=False)
        merged = R.merge_summary(old_s if len(old_s) else None, summary)
        det = R.merge_details(old_d if len(old_d) else None, details)
        n_s = write_table(book, WS_INV_SUM, R.SUMMARY_COLS, merged["data"],
                          prev_rows=len(old_s))
        n_d = write_table(book, WS_INV_DET, R.DETAIL_COLS, det, prev_rows=len(old_d))
    cache_clear(sheet_key)
    return {"summary_rows": n_s, "detail_rows": n_d, "new": merged["new"],
            "updated": merged["updated"], "picked_now": merged["picked_now"],
            "summary": merged["data"], "details": det}


def register_unpick(sa_info: dict, sheet_key: str, invoice_no: str,
                    remark: str = "Load deleted", owner: str = "") -> int:
    """A deleted load has to show as not picked again."""
    with sheet_lock(sa_info, sheet_key, "REGISTER", owner=owner):
        book = open_book(sa_info, sheet_key)
        cur = read_ws(sa_info, sheet_key, WS_INV_SUM, use_cache=False)
        if cur is None or not len(cur):
            return 0
        upd = R.mark_unpicked(cur, invoice_no, remark)
        n = _replace_ws(book, WS_INV_SUM, R.SUMMARY_COLS, upd)
    cache_clear(sheet_key)
    return n


# --------------------------------------------------------------------------- #
# Picking / Packing / Dispatch — status updates
# --------------------------------------------------------------------------- #
# Each of these reads INVOICE_SUMMARY + INVOICE_DETAIL fresh *inside* the lock,
# applies the update, and writes both back in the same lock — same pattern as
# save_register — so two people uploading a status report / scanning a QR at
# the same time never clobber each other.
def apply_live_status(sa_info: dict, sheet_key: str, live_df,
                      owner: str = "") -> dict[str, Any]:
    with sheet_lock(sa_info, sheet_key, "REGISTER", owner=owner):
        book = open_book(sa_info, sheet_key)
        cur_s = read_ws(sa_info, sheet_key, WS_INV_SUM, use_cache=False)
        cur_d = read_ws(sa_info, sheet_key, WS_INV_DET, use_cache=False)
        res = R.apply_pick_live_status(cur_s, cur_d, live_df)
        if not res.get("error"):
            write_table(book, WS_INV_SUM, R.SUMMARY_COLS, res["summary"],
                        prev_rows=len(cur_s))
            write_table(book, WS_INV_DET, R.DETAIL_COLS, res["details"],
                        prev_rows=len(cur_d))
    cache_clear(sheet_key)
    return res


def scan_status(sa_info: dict, sheet_key: str, load_id: str, column: str = "PACKING",
                owner: str = "") -> dict[str, Any]:
    """
    Mark one LOAD_ID complete for a status column — the Packing and Dispatch
    stations both land here, whether the id was scanned or typed.
    """
    with sheet_lock(sa_info, sheet_key, "REGISTER", owner=owner):
        book = open_book(sa_info, sheet_key)
        got = read_many(sa_info, sheet_key, [WS_INV_SUM, WS_INV_DET], use_cache=False)
        cur_s = got.get(WS_INV_SUM, pd.DataFrame())
        cur_d = got.get(WS_INV_DET, pd.DataFrame())
        res = R.apply_status_scan(cur_s, cur_d, load_id, column=column, user=owner)
        if res["found"] and not res["already"]:
            write_table(book, WS_INV_SUM, R.SUMMARY_COLS, res["summary"],
                        prev_rows=len(cur_s))
            write_table(book, WS_INV_DET, R.DETAIL_COLS, res["details"],
                        prev_rows=len(cur_d))
    cache_clear(sheet_key)
    return res


def scan_packing(sa_info: dict, sheet_key: str, load_id: str,
                 owner: str = "") -> dict[str, Any]:
    return scan_status(sa_info, sheet_key, load_id, "PACKING", owner)


def scan_dispatch(sa_info: dict, sheet_key: str, load_id: str,
                  owner: str = "") -> dict[str, Any]:
    return scan_status(sa_info, sheet_key, load_id, "DISPATCH", owner)


def register_gap(sa_info: dict, sheet_key: str, fresh: bool = False) -> dict[str, Any]:
    """Which picked documents never made it into the invoice register."""
    got = read_many(sa_info, sheet_key, [WS_LEDGER, WS_INV_SUM], use_cache=not fresh)
    led = got.get(WS_LEDGER, pd.DataFrame())
    sm = got.get(WS_INV_SUM, pd.DataFrame())
    picked = set()
    if len(led) and "DOC_NUMBER" in led.columns:
        picked = {R._id_str(x).strip() for x in led["DOC_NUMBER"] if str(x).strip()}
    known = set()
    if len(sm) and "TAX_INVOICE_NO" in sm.columns:
        known = {R._id_str(x).strip() for x in sm["TAX_INVOICE_NO"] if str(x).strip()}
    missing = sorted(picked - known)
    return {"picked": len(picked), "registered": len(known), "missing": missing,
            "n_missing": len(missing)}


def backfill_register(sa_info: dict, sheet_key: str, owner: str = "") -> dict[str, Any]:
    """
    Write register rows for every picked document that has none — rebuilt from
    PALLET_LEDGER inside the same lock the register always uses.
    """
    with sheet_lock(sa_info, sheet_key, "REGISTER", owner=owner):
        book = open_book(sa_info, sheet_key)
        got = read_many(sa_info, sheet_key,
                        [WS_LEDGER, WS_REGISTRY, WS_INV_SUM, WS_INV_DET], use_cache=False)
        cur_s = got.get(WS_INV_SUM, pd.DataFrame())
        cur_d = got.get(WS_INV_DET, pd.DataFrame())
        res = R.backfill_from_history(got.get(WS_LEDGER, pd.DataFrame()),
                                      got.get(WS_REGISTRY, pd.DataFrame()),
                                      cur_s, cur_d, user=owner)
        if res["added"]:
            write_table(book, WS_INV_SUM, R.SUMMARY_COLS, res["summary"],
                        prev_rows=len(cur_s))
            write_table(book, WS_INV_DET, R.DETAIL_COLS, res["details"],
                        prev_rows=len(cur_d))
    cache_clear(sheet_key)
    return res


def save_register_docs(sa_info: dict, sheet_key: str, summary: pd.DataFrame,
                       details: pd.DataFrame, owner: str = "") -> dict[str, Any]:
    """
    Register uploaded documents without a pick run — and, where the ledger
    shows they were picked earlier, file them as picked with their pallets.
    """
    with sheet_lock(sa_info, sheet_key, "REGISTER", owner=owner):
        book = open_book(sa_info, sheet_key)
        got = read_many(sa_info, sheet_key,
                        [WS_LEDGER, WS_INV_SUM, WS_INV_DET], use_cache=False)
        cur_s = got.get(WS_INV_SUM, pd.DataFrame())
        cur_d = got.get(WS_INV_DET, pd.DataFrame())
        sm, det = R.enrich_from_history(summary, details, got.get(WS_LEDGER, pd.DataFrame()))
        merged = R.merge_summary(cur_s if len(cur_s) else None, sm)
        merged_d = R.merge_details(cur_d if len(cur_d) else None, det)
        n_s = write_table(book, WS_INV_SUM, R.SUMMARY_COLS, merged["data"],
                          prev_rows=len(cur_s))
        n_d = write_table(book, WS_INV_DET, R.DETAIL_COLS, merged_d, prev_rows=len(cur_d))
    cache_clear(sheet_key)
    return {"summary_rows": n_s, "detail_rows": n_d, "new": merged["new"],
            "updated": merged["updated"], "picked": int((sm["KORBER_PICK"] == "Yes").sum())
            if len(sm) else 0}


def apply_sales_report(sa_info: dict, sheet_key: str, sales_df,
                       owner: str = "") -> dict[str, Any]:
    """
    Checked against the real WMS output, not the invoice register — reads
    OUTBOUND_MASTER (was this invoice actually picked and pushed to the WMS?)
    and OUTBOUND_DETAIL (item + qty actually sent) alongside the register.
    """
    with sheet_lock(sa_info, sheet_key, "REGISTER", owner=owner):
        book = open_book(sa_info, sheet_key)
        got = read_many(sa_info, sheet_key,
                        [WS_INV_SUM, WS_INV_DET, WS_MASTER, WS_DETAIL], use_cache=False)
        cur_s = got.get(WS_INV_SUM, pd.DataFrame())
        cur_d = got.get(WS_INV_DET, pd.DataFrame())
        wms_m = got.get(WS_MASTER, pd.DataFrame())
        wms_d = got.get(WS_DETAIL, pd.DataFrame())
        res = R.apply_sales_report(cur_s, cur_d, sales_df, wms_master=wms_m, wms_detail=wms_d)
        if not res.get("error"):
            write_table(book, WS_INV_SUM, R.SUMMARY_COLS, res["summary"],
                        prev_rows=len(cur_s))
            write_table(book, WS_INV_DET, R.DETAIL_COLS, res["details"],
                        prev_rows=len(cur_d))
    cache_clear(sheet_key)
    return res


# --------------------------------------------------------------------------- #
# reset
# --------------------------------------------------------------------------- #
RESET_SCOPES = {
    "outputs": [WS_MASTER, WS_DETAIL],
    "ledger": [WS_LEDGER],
    "registry": [WS_REGISTRY],
    "rejected": [WS_REJECT],
    "runlog": [WS_RUNLOG],
    "sku": [WS_SKU],
    "register": [WS_INV_SUM, WS_INV_DET],
    "settings": [WS_SETTINGS],
}


def reset_data(sa_info: dict, sheet_key: str, scope: list[str]) -> list[str]:
    book = open_book(sa_info, sheet_key)
    done: list[str] = []
    for s in scope:
        for title in RESET_SCOPES.get(s, []):
            try:
                ws = api(book.worksheet, title)
            except Exception:
                continue
            api(ws.clear)
            api(ws.update, values=[_SHEETS[title]], range_name="A1",
                value_input_option="RAW")
            _CACHE.pop((_key(sheet_key), title), None)
            done.append(title)
    cache_clear(sheet_key)
    return done


def reset_all(sa_info: dict, sheet_key: str, keep_settings: bool = True,
              keep_sku: bool = True, keep_register: bool = True) -> dict[str, Any]:
    """⚠️ FULL DB RESET — header row එක විතරක් ඉතුරු වෙනවා."""
    skip = set()
    if keep_settings:
        skip.add("settings")
    if keep_sku:
        skip.add("sku")
    if keep_register:
        skip.add("register")
    done = reset_data(sa_info, sheet_key, [k for k in RESET_SCOPES if k not in skip])
    return {"cleared": done, "count": len(done),
            "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
