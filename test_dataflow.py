"""
test_dataflow.py — the invariants the whole pipeline has to keep
================================================================

Every check here failed at least once in a real deploy, or guards a boundary
where two modules have to agree on the shape of the same value:

    PDF  ->  DocLine  ->  pick engine  ->  WMS output
                      ->  invoice register  ->  Google Sheet  ->  dashboard

Run:  python test_dataflow.py     (or)  python -m pytest test_dataflow.py
"""
from __future__ import annotations

import pandas as pd

import gsheet as G
import invoice_register as R
import pick_engine as PE
import sku_master as SKU
from doc_parser import ParsedDoc, DocLine, base_item, clean_item, tidy_item

# Codes that have all bitten at some point. The space-separated ones are the
# ERP's other spelling of a hyphenated suffix ("X770132 003710" = X770132-003-710);
# -INL is the same part again; 05-47174 and DFO-312-KIT must NOT be collapsed.
CODES = ["P601560 710", "P951413 000710", "X770132 003710", "1C072323-INL",
         "P162400-000-140", "05-47174", "DFO-312-KIT"]


def _inv(codes=CODES, qty=100, pick_id="0"):
    return pd.DataFrame([{"Item Number": c, "Lot Number": "L", "Pallet ID": f"PAL{i}",
                          "Location Id": "A", "Actual Qty": qty, "Plant": "PL1",
                          "Status": "Available", "Pick Id": pick_id, "UOM": "EA",
                          "Description": "d"} for i, c in enumerate(codes)])


def _doc(num, lines, dtype="INVOICE", date="12-AUG-2026"):
    return ParsedDoc(
        doc_type=dtype, doc_number=num, doc_date=date, ref_number="AR" + num,
        customer="ACME", customer_code="C1", source_file=num + ".pdf",
        lines=[DocLine(line_no=i + 1, item_code=tidy_item(c), description="d",
                       qty=q, uom="EA", unit_price=10.0, line_total=10.0 * q)
               for i, (c, q) in enumerate(lines)],
        total_incl_tax=sum(q for _, q in lines) * 10.0)


def _to_sheet(df, cols):
    """What write_table sends and _frame reads back — everything is a string."""
    d = df.reindex(columns=cols).fillna("")
    rows = [cols] + [[("" if v is None else str(v)) for v in r] for r in d.values.tolist()]
    return G._frame(rows)


# --------------------------------------------------------------------------- #
# base id — the one key every layer matches on
# --------------------------------------------------------------------------- #
def test_base_id_is_the_same_in_every_layer():
    ninv = PE.normalize_inventory(_inv())
    sku = SKU.normalize(pd.DataFrame({"Item Number": CODES, "Description": ["d"] * len(CODES)}))
    for c in CODES:
        want = base_item(c)
        assert ninv.loc[ninv["item_number"] == clean_item(c), "base_id"].iloc[0] == want, c
        assert sku.loc[sku["MATCH_KEY"] == clean_item(c), "BASE_ID"].iloc[0] == want, c
        assert DocLine(line_no=1, item_code=tidy_item(c), qty=1).base == want, c


def test_base_id_splits_a_space_but_not_a_real_distinction():
    assert base_item("X770132 003710") == "X770132"
    assert base_item("P951413 000710") == "P951413"
    assert base_item("1C072323-INL") == "1C072323"
    assert base_item("05-47174") == "05-47174"          # 5 digits, not a suffix
    assert base_item("DFO-312-KIT") == "DFO-312-KIT"    # 312/324 are two sizes


def test_the_document_finds_its_own_stock():
    doc = _doc("I1", [(c, 2) for c in CODES])
    res = PE.run_pick([doc], _inv(), PE.EngineConfig())
    assert len(res["accepted"]) == 1, res["rejected"].to_dict("records")
    assert res["accepted"].iloc[0]["PICKED_QTY"] == 2 * len(CODES)


def test_an_unsplittable_code_still_matches_itself():
    """"P601560710" on the document, "P601560 710" in the inventory."""
    doc = _doc("I2", [("P601560710", 2)])
    res = PE.run_pick([doc], _inv(["P601560 710"]), PE.EngineConfig())
    assert len(res["accepted"]) == 1, res["rejected"].to_dict("records")


# --------------------------------------------------------------------------- #
# allocation
# --------------------------------------------------------------------------- #
def test_quantities_are_conserved_end_to_end():
    docs = [_doc("I1", [("P601560 710", 10), ("1C072323-INL", 12)])]
    res = PE.run_pick(docs, _inv(), PE.EngineConfig())
    a = res["accepted"].iloc[0]
    assert a["PICKED_QTY"] == a["DOC_QTY"] == a["WMS_QTY"] == 22
    assert res["allocations"]["QTY_PICKED"].sum() == 22
    assert res["detail"]["QTY"].astype(float).sum() == 22
    assert set(res["detail"]["DISPLAY_ORDER_NUMBER"]) <= set(res["master"]["DISPLAY_ORDER_NUMBER"])
    assert (res["master"]["LOAD_ID"] == res["master"]["DISPLAY_ORDER_NUMBER"]).all()


def test_a_pallet_is_never_over_picked():
    res = PE.run_pick([_doc("I1", [("P601560 710", 6)])], _inv(["P601560 710"], qty=10),
                      PE.EngineConfig())
    led = res["allocations"]
    assert float(led.iloc[0]["QTY_BALANCE"]) == 4.0
    # the same stale inventory, plus the ledger: only 4 left
    again = PE.run_pick([_doc("I2", [("P601560 710", 6)])], _inv(["P601560 710"], qty=10),
                        PE.EngineConfig(), ledger=led)
    assert len(again["accepted"]) == 0
    assert float(again["shortage"].iloc[0]["AVAILABLE"]) == 4.0


def test_a_fresh_wms_count_overrides_the_ledger():
    res = PE.run_pick([_doc("I1", [("P601560 710", 6)])], _inv(["P601560 710"], qty=10),
                      PE.EngineConfig())
    r2 = PE.run_pick([_doc("I2", [("P601560 710", 9)])], _inv(["P601560 710"], qty=9),
                     PE.EngineConfig(), ledger=res["allocations"])
    assert len(r2["accepted"]) == 1, r2["rejected"].to_dict("records")
    assert set(r2["basis"]["MODE"]) == {"NEW BASELINE"}


def test_stock_on_another_pick_task_is_left_alone():
    locked = _inv(["P601560 710"], pick_id="PK9")
    res = PE.run_pick([_doc("I5", [("P601560 710", 2)])], locked, PE.EngineConfig())
    assert len(res["accepted"]) == 0
    assert "Pick Id" in str(res["shortage"].iloc[0]["REASON"])
    freed = PE.run_pick([_doc("I6", [("P601560 710", 2)])], locked,
                        PE.EngineConfig(release_locked={"I6": ["PK9"]}))
    assert len(freed["accepted"]) == 1


def test_a_document_is_all_or_nothing():
    res = PE.run_pick([_doc("I7", [("P601560 710", 2), ("NOSUCH", 1)])],
                      _inv(["P601560 710"]), PE.EngineConfig())
    assert len(res["accepted"]) == 0
    assert len(res["allocations"]) == 0


# --------------------------------------------------------------------------- #
# register
# --------------------------------------------------------------------------- #
def _built():
    docs = [_doc("30426013174", [("P601560 710", 10)]),
            _doc("30426013175", [("P601560 710", 5)])]
    res = PE.run_pick(docs, _inv(["P601560 710"], qty=50), PE.EngineConfig())
    return docs, res, R.build(docs, res, R.DEFAULT_MRP, user="t", plant="PL1")


def test_register_columns_and_totals():
    _docs, _res, (s, d) = _built()
    assert list(s.columns) == R.SUMMARY_COLS
    assert list(d.columns) == R.DETAIL_COLS
    for num in ("30426013174", "30426013175"):
        row = s[s["TAX_INVOICE_NO"] == num].iloc[0]
        dd = d[d["TAX_INVOICE_NO"] == num]
        assert row["KORBER_PICK"] == "Yes"
        assert row["PICKED_QTY"] == row["QTY"]
        assert dd["PICKED_QTY"].sum() == row["PICKED_QTY"]
        assert dd["DOC_QTY"].sum() == row["QTY"]
        assert row["LINES"] == len(dd)
    assert (s[R.STATUS_COLS] == R.STATUS_PENDING).all().all()
    assert (d[R.STATUS_COLS] == R.STATUS_PENDING).all().all()


def test_register_keeps_the_item_code_as_written():
    docs = [_doc("I1", [(c, 1) for c in CODES])]
    res = PE.run_pick(docs, _inv(), PE.EngineConfig())
    _s, d = R.build(docs, res, R.DEFAULT_MRP)
    for c in CODES:
        row = d[d["ITEM_CODE"] == tidy_item(c)]
        assert len(row) == 1, (c, list(d["ITEM_CODE"]))
        assert row.iloc[0]["BASE_ID"] == base_item(c)


def test_the_sheet_round_trip_changes_no_number():
    _docs, _res, (s, d) = _built()
    S, D = _to_sheet(s, R.SUMMARY_COLS), _to_sheet(d, R.DETAIL_COLS)
    assert list(S.columns) == R.SUMMARY_COLS and list(D.columns) == R.DETAIL_COLS
    a, b = R.dashboard(s)["kpi"], R.dashboard(S)["kpi"]
    for k in ("total", "picked", "pending", "qty", "qty_picked", "qty_pending"):
        assert a[k] == b[k], (k, a[k], b[k])


def test_a_repeated_sheet_header_still_reads_as_columns():
    f = G._frame([["A", "B", "A", ""], ["1", "2", "3", "4"]])
    assert list(f.columns) == ["A", "B", "A_2", "COL3"]
    assert isinstance(f["A"], pd.Series)


def test_a_re_upload_never_loses_a_pick_or_a_floor_status():
    _docs, _res, (s, d) = _built()
    S, D = _to_sheet(s, R.SUMMARY_COLS), _to_sheet(d, R.DETAIL_COLS)
    pk = R.apply_packing_scan(S, D, "30426013174", user="pk")
    cur_s, cur_d = pk["summary"], pk["details"]
    # the same invoices again, this time with no stock at all
    docs = [_doc("30426013174", [("P601560 710", 10)]),
            _doc("30426013175", [("P601560 710", 5)])]
    res2 = PE.run_pick(docs, _inv(["P601560 710"], qty=0), PE.EngineConfig())
    s2, d2 = R.build(docs, res2, R.DEFAULT_MRP)
    m, md = R.merge_summary(cur_s, s2), R.merge_details(cur_d, d2)
    row = m["data"][m["data"]["TAX_INVOICE_NO"] == "30426013174"].iloc[0]
    lines = md[md["TAX_INVOICE_NO"] == "30426013174"]
    assert row["KORBER_PICK"] == "Yes" and float(row["PICKED_QTY"]) == 10.0
    assert row["PACKING"] == R.STATUS_DONE
    assert (lines["PACKING"] == R.STATUS_DONE).all()
    assert set(lines["KORBER_PICK"]) == {"Yes"}          # summary and detail agree


def test_deleting_a_load_releases_the_detail_lines_too():
    _docs, _res, (s, d) = _built()
    us = R.mark_unpicked(s, "30426013174")
    ud = R.mark_unpicked_details(d, "30426013174")
    lines = ud[ud["TAX_INVOICE_NO"] == "30426013174"]
    assert us[us["TAX_INVOICE_NO"] == "30426013174"].iloc[0]["KORBER_PICK"] == "No"
    assert (lines["KORBER_PICK"] == "No").all()
    assert (lines["PICKED_QTY"].astype(float) == 0).all()
    assert (lines["PALLETS"].astype(str) == "").all()


def test_one_invoice_number_however_it_is_spelled():
    """Excel hands an all-digit number back as 30426013174.0."""
    _docs, _res, (s, d) = _built()
    s2 = s.copy(); s2["TAX_INVOICE_NO"] = [float(x) for x in s2["TAX_INVOICE_NO"]]
    d2 = d.copy(); d2["TAX_INVOICE_NO"] = [float(x) for x in d2["TAX_INVOICE_NO"]]
    m = R.merge_summary(s, s2)
    assert len(m["data"]) == len(s) and m["new"] == 0
    assert len(R.merge_details(d, d2)) == len(d)
    dash = R.dashboard(m["data"])
    assert len(R.details_for(d, dash["invoices"])) == len(d)


# --------------------------------------------------------------------------- #
# status columns — Pending -> Completed, never back
# --------------------------------------------------------------------------- #
def _live(load, total, open_, ship):
    return pd.DataFrame([{"Load Id": load, "Total Pick": total,
                          "Open Pick": open_, "Shipped Pick": ship}])


def test_pick_live_status_rules():
    _docs, _res, (s, d) = _built()
    lv = R.apply_pick_live_status(s, d, pd.concat([
        _live(30426013174.0, 10, 0, 0),      # open 0 -> picking done
        _live("30426013175", 5, 2, 5),       # shipped != 0 -> dispatch done
        _live("99999999999", 1, 0, 1),       # not ours
    ], ignore_index=True))
    ls = lv["summary"]
    assert ls.loc[ls["TAX_INVOICE_NO"] == "30426013174", "PICKING"].iloc[0] == R.STATUS_DONE
    assert ls.loc[ls["TAX_INVOICE_NO"] == "30426013175", "PICKING"].iloc[0] == R.STATUS_PENDING
    assert ls.loc[ls["TAX_INVOICE_NO"] == "30426013175", "DISPATCH"].iloc[0] == R.STATUS_DONE
    assert "99999999999" in lv["unmatched"]


def test_a_stale_report_never_undoes_the_floor():
    _docs, _res, (s, d) = _built()
    done = R.apply_pick_live_status(s, d, _live("30426013174", 10, 0, 0))
    back = R.apply_pick_live_status(done["summary"], done["details"],
                                    _live("30426013174", 10, 10, 0))
    ls = back["summary"]
    assert ls.loc[ls["TAX_INVOICE_NO"] == "30426013174", "PICKING"].iloc[0] == R.STATUS_DONE


def test_a_scan_that_only_the_details_still_need_is_not_skipped():
    _docs, _res, (s, d) = _built()
    s.loc[s["TAX_INVOICE_NO"] == "30426013174", "PACKING"] = R.STATUS_DONE
    res = R.apply_status_scan(s, d, "30426013174", column="PACKING")
    assert res["found"] and res["already"] and res["changed"]
    again = R.apply_status_scan(res["summary"], res["details"], "30426013174",
                                column="PACKING")
    assert not again["changed"]


# --------------------------------------------------------------------------- #
# dates
# --------------------------------------------------------------------------- #
def test_every_date_format_reads_the_same_way_alone_or_in_a_column():
    cases = {"12-AUG-2026": "2026-08-12", "01-Aug-2026": "2026-08-01",
             "2026-08-01": "2026-08-01", "01/08/2026": "2026-08-01",
             "2026/08/01": "2026-08-01", "2026-08-01 13:45:00": "2026-08-01",
             "31/12/2025": "2025-12-31"}
    col = R.parse_dates(pd.Series(list(cases)))
    for (raw, want), got in zip(cases.items(), col):
        assert str(got)[:10] == want, (raw, got)
        assert str(R.parse_date(raw))[:10] == want, raw
    assert R.parse_dates(pd.Series(["", "junk", None])).isna().all()


# --------------------------------------------------------------------------- #
# backfill / reconciliation
# --------------------------------------------------------------------------- #
LEDGER = pd.DataFrame([
    {"RUN_ID": "R1", "PICK_DATE": "2026-08-13 09:00:00", "DOC_TYPE": "INVOICE",
     "DOC_NUMBER": "30426013174", "DOC_LINE": 1, "DOC_ITEM_CODE": "P601560 710",
     "ITEM_NUMBER": "P601560 710", "PALLET": "PAL1", "LOCATION_ID": "A1",
     "LOT_NUMBER": "L1", "PLANT": "PL1", "UOM": "EA", "QTY_PICKED": 6,
     "DESCRIPTION": "d", "SOURCE_FILE": "a.pdf"},
    {"RUN_ID": "R1", "PICK_DATE": "2026-08-13 09:00:00", "DOC_TYPE": "INVOICE",
     "DOC_NUMBER": "30426013174", "DOC_LINE": 1, "DOC_ITEM_CODE": "P601560 710",
     "ITEM_NUMBER": "P601560 710", "PALLET": "PAL2", "LOCATION_ID": "A2",
     "LOT_NUMBER": "L2", "PLANT": "PL1", "UOM": "EA", "QTY_PICKED": 4,
     "DESCRIPTION": "d", "SOURCE_FILE": "a.pdf"},
    {"RUN_ID": "R2", "PICK_DATE": "2026-08-14 10:00:00", "DOC_TYPE": "DELIVERY CHALLAN",
     "DOC_NUMBER": "DC900", "DOC_LINE": 1, "DOC_ITEM_CODE": "X770132 003710",
     "ITEM_NUMBER": "X770132 003710", "PALLET": "PAL9", "LOCATION_ID": "A9",
     "LOT_NUMBER": "L9", "PLANT": "PL1", "UOM": "EA", "QTY_PICKED": 3,
     "DESCRIPTION": "d", "SOURCE_FILE": "b.pdf"},
])
REGISTRY = pd.DataFrame([
    {"DOC_NUMBER": "30426013174", "DOC_TYPE": "INVOICE", "DOC_DATE": "13-AUG-2026",
     "REF_NUMBER": "AR1"},
    {"DOC_NUMBER": "DC900", "DOC_TYPE": "DELIVERY CHALLAN", "DOC_DATE": "14-AUG-2026",
     "REF_NUMBER": "OR9"},
])


def test_backfill_rebuilds_the_register_from_the_ledger():
    b = R.backfill_from_history(LEDGER, REGISTRY, None, None, user="admin")
    s, d = b["summary"], b["details"]
    assert b["added"] == 2 and len(d) == 2
    assert list(s.columns) == R.SUMMARY_COLS and list(d.columns) == R.DETAIL_COLS
    row = s[s["TAX_INVOICE_NO"] == "30426013174"].iloc[0]
    line = d[d["TAX_INVOICE_NO"] == "30426013174"].iloc[0]
    assert float(row["QTY"]) == float(row["PICKED_QTY"]) == 10.0    # both pallets
    assert set(str(line["PALLETS"]).split(", ")) == {"PAL1", "PAL2"}
    assert line["BASE_ID"] == "P601560"
    assert d[d["TAX_INVOICE_NO"] == "DC900"].iloc[0]["BASE_ID"] == "X770132"


def test_backfill_is_idempotent_and_does_not_fight_the_floor():
    b = R.backfill_from_history(LEDGER, REGISTRY, None, None)
    scanned = R.apply_packing_scan(b["summary"], b["details"], "30426013174")
    again = R.backfill_from_history(LEDGER, REGISTRY, scanned["summary"],
                                    scanned["details"])
    assert again["added"] == 0
    row = again["summary"][again["summary"]["TAX_INVOICE_NO"] == "30426013174"].iloc[0]
    assert row["PACKING"] == R.STATUS_DONE


def test_sales_report_reconciles_against_the_wms():
    b = R.backfill_from_history(LEDGER, REGISTRY, None, None)
    master = pd.DataFrame([{"LOAD_ID": "30426013174",
                            "DISPLAY_ORDER_NUMBER": "30426013174"}])
    detail = pd.DataFrame([{"DISPLAY_ORDER_NUMBER": "30426013174",
                            "DISPLAY_ITEM_NUMBER": "P601560 710", "QTY": "10"}])
    sales = pd.DataFrame([
        {"Tax Invoice No.": 30426013174.0, "Customer Item": "P601560",
         "Item Code": "P601560 710", "QTY": 10},
        {"Tax Invoice No.": "30426013174", "Customer Item": "ZZZZ",
         "Item Code": "ZZZZ", "QTY": 1},
        {"Tax Invoice No.": "DC900", "Customer Item": "X770132",
         "Item Code": "X770132 003710", "QTY": 3},
    ])
    res = R.apply_sales_report(b["summary"], b["details"], sales, master, detail)
    rep = res["report"]
    assert res["used_wms"] and len(rep) == len(sales)
    assert rep.iloc[0]["STATUS"] == "Matched"
    assert "OUTBOUND_DETAIL" in rep.iloc[1]["STATUS"]
    assert "OUTBOUND_MASTER" in rep.iloc[2]["STATUS"]
    assert len(R.sales_reconciliation_excel(rep)) > 0


if __name__ == "__main__":
    import sys
    fails = 0
    for name, fn in sorted(globals().items()):
        if not (name.startswith("test_") and callable(fn)):
            continue
        try:
            fn()
            print(f"  ok   {name}")
        except Exception as ex:                                   # noqa: BLE001
            fails += 1
            print(f"  FAIL {name}: {type(ex).__name__}: {ex}")
    print(f"\n{fails} failed")
    sys.exit(1 if fails else 0)
