"""
transactions.py — what the warehouse actually picked
====================================================
The ledger records the pallets the system *chose*. The picker on the floor may
well have taken different ones. Until those two are put side by side, a pallet
the system reserved but nobody touched stays locked out of every later pick, and
a pallet that was really emptied still looks full.

The WMS Transactions History Report is the truth. `Control Number` carries the
load — `INM0DONA-333262712295` is load `333262712295` — and `Starting Hu` is the
pallet the stock actually came off.

Nothing here writes on its own. It produces a reconciliation the user looks at
and applies.
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd

from doc_parser import base_item, clean_item

# Bumped whenever this module's public surface changes; app.py refuses to run
# against a stale copy instead of dying with a redacted TypeError.
API = 1

QTY_TOL = 1e-6

ACTUAL_COLS = ["LOAD_ID", "CONTROL_NUMBER", "PALLET", "ITEM_NUMBER", "BASE_ID",
               "LOT_NUMBER", "QTY", "FROM_LOC", "TO_LOC", "WHEN", "EMPLOYEE"]

RECON_COLS = ["LOAD_ID", "PALLET", "ITEM_NUMBER", "BASE_ID", "SYSTEM_QTY",
              "ACTUAL_QTY", "DIFF", "OUTCOME", "NOTE"]

# what an outcome means for the pallet balance
RELEASE = "RELEASE — system reserved it, the floor never touched it"
CONSUME = "CONSUME — the floor took it, the system does not know"
SHORT_TAKE = "ADJUST — less came off this pallet than the system reserved"
OVER_TAKE = "ADJUST — more came off this pallet than the system reserved"
AGREES = "AGREES"

_ALIASES = {
    "control_number": ["control number", "control no", "control_number"],
    "pallet": ["starting hu", "start hu", "from hu", "pallet", "hu"],
    "to_hu": ["ending hu", "end hu"],
    "item_number": ["item number", "item", "sku"],
    "qty": ["tran qty", "transaction qty", "qty", "quantity"],
    "from_loc": ["starting loc", "start loc", "from location"],
    "to_loc": ["ending loc", "end loc", "to location"],
    "lot_number": ["lot number", "lot"],
    "when": ["end tran date", "start tran date", "tran date", "date"],
    "employee": ["employee id", "employee", "user"],
    "client": ["client code", "client"],
    "wh": ["wh id", "warehouse"],
}


def _col(df: pd.DataFrame, key: str) -> str | None:
    low = {str(c).strip().lower(): c for c in df.columns}
    for want in _ALIASES.get(key, []):
        if want in low:
            return low[want]
    return None


def strip_client(value: Any, client_code: str = "") -> str:
    """
    `INM0DONA-333262712295` -> `333262712295`, `INM0DONA-333/26-27/17` -> `333/26-27/17`.

    The client code is stripped when it is there; anything else is left exactly
    as it is, because a load id is matched character for character later.
    """
    v = str(value or "").strip()
    if not v:
        return ""
    cc = str(client_code or "").strip()
    if cc and v.upper().startswith(cc.upper() + "-"):
        return v[len(cc) + 1:].strip()
    # fall back to a generic "<letters+digits>-" prefix only when what follows
    # still looks like a document number
    m = re.match(r"^[A-Z0-9]{6,12}-(?=.*\d)(.+)$", v, re.I)
    return m.group(1).strip() if m else v


def _collapse_chains(df: pd.DataFrame) -> pd.DataFrame:
    """
    One pick, not one row per leg.

    A pallet leaves the rack, goes to a picker, then to staging, then to the
    dock — four rows, all carrying the same quantity. Summing them multiplies
    the pick by four. Walking each pallet's legs in time order and dropping any
    leg that starts where the pallet was already put collapses the chain back
    to the single movement that actually took stock out of storage. A genuine
    second pick from the same pallet starts somewhere new, so it survives.
    """
    if not len(df):
        return df
    d = df.copy()
    # the same leg written twice is one movement, not two
    d = d.drop_duplicates(subset=["LOAD_ID", "PALLET", "ITEM_NUMBER", "FROM_LOC",
                                  "TO_LOC", "QTY"], keep="first")
    d["_t"] = pd.to_datetime(d["WHEN"], dayfirst=True, errors="coerce")
    d = d.sort_values(["LOAD_ID", "PALLET", "_t"], kind="stable")
    keep: list[bool] = []
    seen: dict[tuple, set] = {}
    for _, r in d.iterrows():
        key = (r["LOAD_ID"], r["PALLET"])
        been = seen.setdefault(key, set())
        first = str(r["FROM_LOC"]).strip()
        keep.append(first not in been)
        been.add(str(r["TO_LOC"]).strip())
    return d[pd.Series(keep, index=d.index)].drop(columns=["_t"])


def normalize(raw: pd.DataFrame, client_code: str = "",
              picks_only: bool = True, collapse: bool = True) -> pd.DataFrame:
    """
    The report as this app needs it.

    `picks_only` keeps the rows that took stock off a pallet — a receipt has no
    Starting Hu, and putting one through a pick reconciliation would be nonsense.
    """
    if raw is None or not len(raw):
        return pd.DataFrame(columns=ACTUAL_COLS)

    df = raw.copy()
    cn = _col(df, "control_number")
    hu = _col(df, "pallet")
    if not cn or not hu:
        missing = "Control Number" if not cn else "Starting Hu"
        raise ValueError(f"This does not look like a Transactions History Report — "
                         f"no '{missing}' column.")

    out = pd.DataFrame()
    out["CONTROL_NUMBER"] = df[cn].astype(str).str.strip()
    out["PALLET"] = df[hu].astype(str).str.strip()
    for key, col in (("ITEM_NUMBER", "item_number"), ("LOT_NUMBER", "lot_number"),
                     ("FROM_LOC", "from_loc"), ("TO_LOC", "to_loc"),
                     ("WHEN", "when"), ("EMPLOYEE", "employee")):
        c = _col(df, col)
        out[key] = df[c].astype(str).str.strip() if c else ""
    qc = _col(df, "qty")
    out["QTY"] = pd.to_numeric(df[qc], errors="coerce").fillna(0.0) if qc else 0.0

    # the item number carries the client code too
    cc = str(client_code or "").strip()
    if not cc:
        c = _col(df, "client")
        vals = df[c].dropna().astype(str).str.strip() if c else pd.Series(dtype=str)
        cc = vals.mode().iat[0] if len(vals) else ""
    out["ITEM_NUMBER"] = out["ITEM_NUMBER"].map(lambda v: strip_client(v, cc))
    out["LOAD_ID"] = out["CONTROL_NUMBER"].map(lambda v: strip_client(v, cc))
    out["BASE_ID"] = out["ITEM_NUMBER"].map(base_item)

    blank = {"nan", "none", "nat", "<na>", ""}
    for c in ("PALLET", "ITEM_NUMBER", "LOT_NUMBER", "FROM_LOC", "TO_LOC",
              "WHEN", "EMPLOYEE", "LOAD_ID", "CONTROL_NUMBER", "BASE_ID"):
        out[c] = out[c].astype(str).str.strip()
        out.loc[out[c].str.lower().isin(blank), c] = ""
    if picks_only:
        out = out[(out["PALLET"] != "") & (out["LOAD_ID"] != "")
                  & (out["QTY"] > QTY_TOL)]
    out = out[ACTUAL_COLS]
    if collapse:
        out = _collapse_chains(out)
    return out.reset_index(drop=True)


def load_ids(actual: pd.DataFrame) -> list[str]:
    if actual is None or not len(actual):
        return []
    return sorted({str(x) for x in actual["LOAD_ID"] if str(x).strip()})


def match_loads(actual: pd.DataFrame, known: list[str]) -> dict[str, str]:
    """
    Report load id -> the load id this system holds.

    A second pick against the same invoice comes back as `…298-A` or `…321A`,
    so an exact match is tried first and a trailing revision suffix after that.
    Nothing else is guessed — a wrong match would move the wrong stock.
    """
    have = {str(k).strip() for k in known if str(k).strip()}
    out: dict[str, str] = {}
    for lid in load_ids(actual):
        if lid in have:
            out[lid] = lid
            continue
        stem = re.sub(r"[-_ ]?[A-Z]$", "", lid, flags=re.I).strip()
        if stem != lid and stem in have:
            out[lid] = stem
    return out


def reconcile(actual: pd.DataFrame, ledger: pd.DataFrame,
              loads: list[str] | None = None,
              client_code: str = "") -> dict[str, Any]:
    """
    Pallet by pallet, for every load in both: what the system reserved against
    what the floor actually took.

    Returns the per-pallet table, plus the two frames that matter — the pallet
    quantity to give back, and the pallet quantity to take away.
    """
    empty = {"rows": pd.DataFrame(columns=RECON_COLS),
             "release": pd.DataFrame(columns=RECON_COLS),
             "consume": pd.DataFrame(columns=RECON_COLS),
             "loads": {}, "unknown": [], "totals": {}}
    if actual is None or not len(actual) or ledger is None or not len(ledger):
        return empty

    led = ledger.copy()
    led["DOC_NUMBER"] = led["DOC_NUMBER"].astype(str).str.strip()
    led["PALLET"] = led["PALLET"].astype(str).str.strip()
    led["QTY_PICKED"] = pd.to_numeric(led["QTY_PICKED"], errors="coerce").fillna(0.0)
    known = sorted(set(led["DOC_NUMBER"]))

    pairs = match_loads(actual, known)
    if loads:
        want = {str(x) for x in loads}
        pairs = {k: v for k, v in pairs.items() if v in want}
    unknown = [l for l in load_ids(actual) if l not in pairs]
    if not pairs:
        return {**empty, "unknown": unknown}

    act = actual[actual["LOAD_ID"].isin(pairs)].copy()
    act["LOAD_ID"] = act["LOAD_ID"].map(pairs)
    a = (act.groupby(["LOAD_ID", "PALLET"], dropna=False)
         .agg(ACTUAL_QTY=("QTY", "sum"), ITEM_NUMBER=("ITEM_NUMBER", "first"),
              BASE_ID=("BASE_ID", "first")).reset_index())

    s = (led[led["DOC_NUMBER"].isin(set(pairs.values()))]
         .groupby(["DOC_NUMBER", "PALLET"], dropna=False)
         .agg(SYSTEM_QTY=("QTY_PICKED", "sum"), ITEM_NUMBER=("ITEM_NUMBER", "first"),
              BASE_ID=("BASE_ID", "first")).reset_index()
         .rename(columns={"DOC_NUMBER": "LOAD_ID"}))

    m = pd.merge(s, a, on=["LOAD_ID", "PALLET"], how="outer",
                 suffixes=("_S", "_A"))
    m["SYSTEM_QTY"] = m["SYSTEM_QTY"].fillna(0.0)
    m["ACTUAL_QTY"] = m["ACTUAL_QTY"].fillna(0.0)
    m["ITEM_NUMBER"] = m["ITEM_NUMBER_S"].fillna(m["ITEM_NUMBER_A"]).fillna("")
    m["BASE_ID"] = m["BASE_ID_S"].fillna(m["BASE_ID_A"]).fillna("")
    m["DIFF"] = m["ACTUAL_QTY"] - m["SYSTEM_QTY"]

    def _outcome(r) -> tuple[str, str]:
        sysq, actq = float(r["SYSTEM_QTY"]), float(r["ACTUAL_QTY"])
        if actq <= QTY_TOL:
            return RELEASE, f"give {sysq:g} back to this pallet"
        if sysq <= QTY_TOL:
            return CONSUME, f"take {actq:g} off this pallet"
        if abs(actq - sysq) <= QTY_TOL:
            return AGREES, ""
        if actq < sysq:
            return SHORT_TAKE, f"give {sysq - actq:g} back"
        return OVER_TAKE, f"take a further {actq - sysq:g}"

    got = m.apply(_outcome, axis=1, result_type="expand")
    m["OUTCOME"], m["NOTE"] = got[0], got[1]
    rows = m[RECON_COLS].sort_values(["LOAD_ID", "OUTCOME", "PALLET"]) \
        .reset_index(drop=True)

    give_back = rows[rows["OUTCOME"].isin([RELEASE, SHORT_TAKE])].copy()
    take_away = rows[rows["OUTCOME"].isin([CONSUME, OVER_TAKE])].copy()
    return {
        "rows": rows, "release": give_back.reset_index(drop=True),
        "consume": take_away.reset_index(drop=True),
        "loads": pairs, "unknown": unknown,
        "totals": {
            "loads": len(set(pairs.values())),
            "pallets": int(len(rows)),
            "agree": int((rows["OUTCOME"] == AGREES).sum()),
            "release_qty": float((give_back["SYSTEM_QTY"]
                                  - give_back["ACTUAL_QTY"]).sum()),
            "consume_qty": float((take_away["ACTUAL_QTY"]
                                  - take_away["SYSTEM_QTY"]).sum()),
        },
    }


def ledger_corrections(recon: dict[str, Any], run_id: str = "",
                       who: str = "") -> pd.DataFrame:
    """
    The ledger rows that put the balance right.

    A correction is an ordinary ledger row with a negative `QTY_PICKED` when
    stock is being given back, so the running balance stays a plain sum and
    nothing downstream needs to know these rows are different.
    """
    from datetime import datetime
    import pick_engine as E

    rows = recon.get("rows")
    if rows is None or not len(rows):
        return pd.DataFrame(columns=E.ALLOC_COLS)
    fix = rows[rows["OUTCOME"] != AGREES]
    if not len(fix):
        return pd.DataFrame(columns=E.ALLOC_COLS)

    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    out = []
    for _, r in fix.iterrows():
        delta = float(r["ACTUAL_QTY"]) - float(r["SYSTEM_QTY"])
        out.append({
            "RUN_ID": run_id or f"RECON-{stamp}", "PICK_DATE": stamp,
            "DOC_TYPE": "RECONCILIATION", "DOC_NUMBER": r["LOAD_ID"], "DOC_LINE": 0,
            "DOC_ITEM_CODE": r["ITEM_NUMBER"], "BASE_ID": r["BASE_ID"],
            "ITEM_NUMBER": r["ITEM_NUMBER"],
            "DESCRIPTION": f"{r['OUTCOME'].split('—')[0].strip()} · {who}".strip(" ·"),
            "PALLET": r["PALLET"], "LOCATION_ID": "", "LOT_NUMBER": "", "PLANT": "",
            "UOM": "", "QTY_BEFORE": 0.0, "QTY_PICKED": delta, "QTY_BALANCE": 0.0,
            "FIFO_DATE": "", "GRN_NUMBER": "", "STORED_ATTRIBUTE_ID": "",
            "ROW_KEY": "", "SOURCE_FILE": "Transactions History Report",
        })
    return pd.DataFrame(out, columns=E.ALLOC_COLS)


def duplicate_pallets(actual: pd.DataFrame,
                      loads: list[str] | None = None) -> pd.DataFrame:
    """
    Pallets the report shows being picked for **more than one load**.

    Over a month of history that is most pallets, and quite right — a pallet
    holds enough for many orders. Pass `loads` to narrow it to the ones being
    reconciled, where the same pallet on two loads is worth a second look.
    """
    if actual is not None and len(actual) and loads:
        actual = actual[actual["LOAD_ID"].isin({str(x) for x in loads})]
    if actual is None or not len(actual):
        return pd.DataFrame(columns=["PALLET", "ITEM_NUMBER", "LOADS", "TIMES",
                                     "TOTAL_QTY", "LOAD_IDS"])
    g = (actual.groupby(["PALLET", "ITEM_NUMBER"], dropna=False)
         .agg(LOADS=("LOAD_ID", "nunique"), TIMES=("LOAD_ID", "size"),
              TOTAL_QTY=("QTY", "sum"),
              LOAD_IDS=("LOAD_ID", lambda s: ", ".join(sorted(set(s))[:6])))
         .reset_index())
    g = g[g["LOADS"] > 1]
    return g.sort_values(["LOADS", "TOTAL_QTY"], ascending=False).reset_index(drop=True)
