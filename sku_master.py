"""
sku_master.py — SKU Master (Item Number + Description + ඕනෑම extra column)
==========================================================================
  * Duplicate නෑ  — ITEM_NUMBER එකට upsert (තියෙන එක update, අලුත් එක insert)
  * Search        — base ID එකෙන්: `07011636` දුන්නම `07011636-000-440` හම්බෙනවා
  * හැම data එකක්ම — upload කරන file එකේ extra columns ඔක්කොම රැකෙනවා
"""
from __future__ import annotations

import io
import re
from datetime import datetime
from typing import Any

import pandas as pd

from doc_parser import base_item, clean_item

# Bumped whenever this module's public surface changes; app.py refuses to run
# against a stale copy instead of dying with a redacted TypeError.
API = 3

CORE = ["ITEM_NUMBER", "BASE_ID", "MATCH_KEY", "ITEM_DESCRIPTION"]
AUDIT = ["UPDATED_AT", "UPDATED_BY", "SOURCE"]

# upload file එකේ column නම් -> canonical
ALIASES = {
    "ITEM_NUMBER": ["item number", "item no", "item code", "itemnumber", "sku",
                    "display item number", "donaldson item", "material"],
    "ITEM_DESCRIPTION": ["item description", "description", "desc", "item desc",
                         "description of goods", "material description"],
}


def _norm(s: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


def _colmap(df: pd.DataFrame) -> dict[str, str]:
    """DataFrame columns -> canonical name (exact match මුලින්, ඊට පස්සේ contains)."""
    out: dict[str, str] = {}
    used: set[str] = set()
    for canon, names in ALIASES.items():
        found = None
        for n in names:
            for c in df.columns:
                if _norm(c) == _norm(n) and c not in used:
                    found = c
                    break
            if found:
                break
        if found is None:
            for n in names:
                for c in df.columns:
                    if _norm(n) in _norm(c) and c not in used:
                        found = c
                        break
                if found:
                    break
        if found is not None:
            used.add(found)
            out[found] = canon
    return out


def normalize(df: pd.DataFrame, source: str = "", user: str = "") -> pd.DataFrame:
    """
    ඕනෑම SKU file එකක් -> canonical frame.
    Core columns rename වෙනවා, extra columns ඒ නමින්ම රැකෙනවා.
    """
    if df is None or not len(df):
        return pd.DataFrame(columns=CORE + AUDIT)

    d = df.copy()
    d.columns = [str(c).strip() for c in d.columns]
    d = d.rename(columns=_colmap(d))

    if "ITEM_NUMBER" not in d.columns:
        raise ValueError("No 'Item Number' column found in this file.")
    if "ITEM_DESCRIPTION" not in d.columns:
        d["ITEM_DESCRIPTION"] = ""

    # ITEM_NUMBER = file එකේ තියෙන හරියටම (upper + trim), '#1301' වගේ ඒවත් රැකෙනවා
    d["ITEM_NUMBER"] = (d["ITEM_NUMBER"].astype(str).str.strip().str.upper()
                        .str.replace(r"\s+", " ", regex=True)
                        .replace({"NAN": "", "NONE": ""}))
    d = d[d["ITEM_NUMBER"] != ""].copy()
    d["MATCH_KEY"] = d["ITEM_NUMBER"].map(clean_item)      # duplicate / lookup key
    d["BASE_ID"] = d["ITEM_NUMBER"].map(base_item)
    d["ITEM_DESCRIPTION"] = (d["ITEM_DESCRIPTION"].astype(str)
                             .replace({"nan": "", "None": ""}).str.strip())

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    d["UPDATED_AT"] = now
    d["UPDATED_BY"] = str(user or "")
    d["SOURCE"] = str(source or "")

    extras = [c for c in d.columns if c not in CORE + AUDIT]
    d = d[CORE + extras + AUDIT]

    # file එකේම duplicate තිබ්බොත් අන්තිම එක ගන්නවා
    d = d.drop_duplicates(subset=["MATCH_KEY"], keep="last").reset_index(drop=True)
    return d


def upsert(existing: pd.DataFrame | None, incoming: pd.DataFrame) -> dict[str, Any]:
    """
    Duplicate නොවෙන්න merge — ITEM_NUMBER එකට:
      * අලුත් නම්      -> insert
      * තියෙනවා නම්    -> value තියෙන fields විතරක් update (හිස් වලින් overwrite නෑ)
    """
    inc = incoming.copy()
    if existing is None or not len(existing):
        inc["_STATUS"] = "NEW"
        return {"data": inc.drop(columns=["_STATUS"]), "new": len(inc), "updated": 0,
                "unchanged": 0, "changes": inc.assign(_STATUS="NEW")}

    old = existing.copy()
    if "MATCH_KEY" not in old.columns:
        old["MATCH_KEY"] = old["ITEM_NUMBER"].map(clean_item)
    old["MATCH_KEY"] = old["MATCH_KEY"].map(clean_item)
    old = old.drop_duplicates(subset=["MATCH_KEY"], keep="last")

    cols = list(dict.fromkeys(list(old.columns) + list(inc.columns)))
    old = old.reindex(columns=cols).fillna("")
    inc2 = inc.reindex(columns=cols).fillna("")

    idx = old.set_index("MATCH_KEY")
    rows: list[dict] = []
    changes: list[dict] = []
    n_new = n_upd = n_same = 0

    for _, r in inc2.iterrows():
        item = r["MATCH_KEY"]
        if item not in idx.index:
            rows.append(r.to_dict())
            changes.append({**r.to_dict(), "_STATUS": "NEW", "_CHANGED": ""})
            n_new += 1
            continue

        cur = idx.loc[item]
        cur = (cur.iloc[0] if isinstance(cur, pd.DataFrame) else cur).to_dict()
        merged = dict(cur)
        diffs: list[str] = []
        for c in cols:
            if c in ("MATCH_KEY",) or c in AUDIT:
                continue
            new_v = str(r.get(c, "") or "").strip()
            old_v = str(cur.get(c, "") or "").strip()
            if new_v and new_v != old_v:
                merged[c] = new_v
                diffs.append(f"{c}: '{old_v}' → '{new_v}'")
        merged["MATCH_KEY"] = item
        merged["ITEM_NUMBER"] = r.get("ITEM_NUMBER") or cur.get("ITEM_NUMBER") or item
        merged["BASE_ID"] = base_item(merged["ITEM_NUMBER"])
        for c in AUDIT:
            if str(r.get(c, "") or "").strip():
                merged[c] = r.get(c)
        rows.append(merged)
        if diffs:
            n_upd += 1
            changes.append({**merged, "_STATUS": "UPDATED", "_CHANGED": " · ".join(diffs)})
        else:
            n_same += 1

    touched = {r["MATCH_KEY"] for r in rows}
    keep = old[~old["MATCH_KEY"].isin(touched)]
    out = pd.concat([keep, pd.DataFrame(rows, columns=cols)], ignore_index=True)
    out = out.drop_duplicates(subset=["MATCH_KEY"], keep="last")
    out = out.sort_values("ITEM_NUMBER").reset_index(drop=True)

    return {"data": out, "new": n_new, "updated": n_upd, "unchanged": n_same,
            "changes": pd.DataFrame(changes)}


def search(master: pd.DataFrame, query: str, base_match: bool = True,
           limit: int = 500) -> pd.DataFrame:
    """
    `07011636` දුන්නොත් `07011636-000-440` හම්බෙනවා (base ID match).
    Description එකෙනුත් හොයනවා. Word කීපයක් = AND.
    """
    if master is None or not len(master) or not str(query).strip():
        return pd.DataFrame(columns=master.columns if master is not None else [])

    d = master.copy()
    item = d["ITEM_NUMBER"].astype(str).str.upper()
    base = (d["BASE_ID"] if "BASE_ID" in d.columns
            else d["ITEM_NUMBER"].map(base_item)).astype(str).str.upper()
    desc = d.get("ITEM_DESCRIPTION", pd.Series("", index=d.index)).astype(str).str.upper()
    mkey = (d["MATCH_KEY"] if "MATCH_KEY" in d.columns
            else d["ITEM_NUMBER"].map(clean_item)).astype(str).str.upper()
    extras = None
    for c in d.columns:
        if c in ("ITEM_NUMBER", "BASE_ID", "MATCH_KEY", "ITEM_DESCRIPTION"):
            continue
        col = d[c].astype(str).str.upper()
        extras = col if extras is None else extras.str.cat(col, sep=" ")

    mask = pd.Series(True, index=d.index)
    for term in str(query).split():
        t = term.strip().upper()
        if not t:
            continue
        tc = clean_item(t)
        tb = base_item(t)
        m = item.str.contains(re.escape(t), na=False) | desc.str.contains(re.escape(t), na=False)
        if tc:
            m |= mkey.str.contains(re.escape(tc), na=False)
        if base_match and tb:
            # 07011636 -> 07011636-000-440   /   07011636-000-440 -> 07011636
            m |= (base == tb) | mkey.str.startswith(tb)
        if extras is not None:
            m |= extras.str.contains(re.escape(t), na=False)
        mask &= m
    return d[mask].head(limit).reset_index(drop=True)


def lookup(master: pd.DataFrame | None) -> dict[str, str]:
    """Pick engine එකට — ITEM / BASE -> description."""
    out: dict[str, str] = {}
    if master is None or not len(master):
        return out
    for _, r in master.iterrows():
        raw = r.get("ITEM_NUMBER") or r.get("MATCH_KEY")
        item = clean_item(r.get("MATCH_KEY") or raw)
        desc = str(r.get("ITEM_DESCRIPTION", "") or "").strip()
        if not item or not desc:
            continue
        out[item] = desc
        # off the raw code — clean_item() has already eaten the space that
        # base_item() splits on, so base_item(item) is not the same base id the
        # rest of the system uses.
        base = str(r.get("BASE_ID", "") or "").strip() or base_item(raw)
        if base:
            out.setdefault(base, desc)
    return out


def to_excel(master: pd.DataFrame, sheet: str = "SKU Master") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        (master if master is not None and len(master)
         else pd.DataFrame(columns=CORE)).to_excel(xw, sheet_name=sheet[:31], index=False)
    return buf.getvalue()


def template_excel() -> bytes:
    demo = pd.DataFrame({
        "Item Number": ["07011636-000-440", "P550945", "100409-101"],
        "Item Description": ["GASKET CLOSED-CELL EPDM 5 MM X 10 MM  10-M PACK",
                             "LUBE FILTER, SPIN-ON FULL FLOW",
                             "SWITCH - DELTA-P, 9 INCH"],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        demo.to_excel(xw, sheet_name="SKU Master", index=False)
    return buf.getvalue()
