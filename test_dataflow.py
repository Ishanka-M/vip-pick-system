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

from datetime import datetime, timedelta

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




# --------------------------------------------------------------------------- #
# partial pick — "we cannot wait, send what we have"
# --------------------------------------------------------------------------- #
def _short_case():
    """AAA has plenty, BBB is 4 short of the 12 the document asks for."""
    inv = pd.DataFrame([
        {"Item Number": "AAA", "Lot Number": "L", "Pallet ID": "PAL1",
         "Location Id": "A", "Actual Qty": 100, "Plant": "PL1", "Status": "Available",
         "Pick Id": "0", "UOM": "EA", "Description": "d"},
        {"Item Number": "BBB", "Lot Number": "L", "Pallet ID": "PAL2",
         "Location Id": "A", "Actual Qty": 4, "Plant": "PL1", "Status": "Available",
         "Pick Id": "0", "UOM": "EA", "Description": "d"},
    ])
    return inv, _doc("I1", [("AAA", 10), ("BBB", 12)])


def test_all_or_nothing_is_still_the_default():
    inv, doc = _short_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig())
    assert len(res["accepted"]) == 0
    assert len(res["allocations"]) == 0
    assert "STOCK SHORT" in res["rejected"].iloc[0]["REASON"]


def test_partialable_offers_only_what_can_actually_ship():
    inv, doc = _short_case()                    # AAA 10 ok, BBB 12 asked / 4 there
    res = PE.run_pick([doc], inv, PE.EngineConfig())
    offer = PE.partialable(res)
    assert list(offer["DOC_NUMBER"]) == ["I1"]
    row = offer.iloc[0]
    # the whole document, and what a partial pick would really load: the full
    # line in full (10) plus whatever the short line has (4)
    assert row["DOC_QTY"] == 22
    assert row["CAN_PICK_NOW"] == 14
    assert row["STILL_SHORT"] == 8
    assert row["LINES"] == 2 and row["SHORT_LINES"] == 1
    # a document with nothing at all is not on offer — it is named separately
    nothing = PE.run_pick([_doc("I2", [("ZZZ", 5)])], inv, PE.EngineConfig())
    assert len(PE.partialable(nothing)) == 0
    assert list(PE.no_partial(nothing)["DOC_NUMBER"]) == ["I2"]


def test_a_confirmed_partial_picks_what_is_there():
    inv, doc = _short_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    assert res["partial"] == ["I1"]
    a = res["accepted"].iloc[0]
    assert a["PICK_STATUS"] == "PARTIAL"
    assert a["DOC_QTY"] == 22 and a["PICKED_QTY"] == 14 and a["SHORT_QTY"] == 8
    # the WMS files carry what is really being picked, not the invoice quantity
    assert res["detail"]["QTY"].astype(float).sum() == 14
    assert res["allocations"]["QTY_PICKED"].sum() == 14
    assert "PARTIAL PICK" in a["DOC_CHECK"]


def test_a_partial_never_takes_more_than_the_document_asks():
    inv, doc = _short_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["*"]))
    per_line = res["allocations"].groupby("DOC_LINE")["QTY_PICKED"].sum()
    for ln in doc.lines:
        assert per_line.get(ln.line_no, 0.0) <= ln.qty + 1e-9


def test_nothing_on_the_floor_is_not_a_partial_pick():
    inv, _ = _short_case()
    res = PE.run_pick([_doc("I2", [("ZZZ", 5)])], inv, PE.EngineConfig(partial_docs=["*"]))
    assert len(res["accepted"]) == 0
    assert "STOCK SHORT" in res["rejected"].iloc[0]["REASON"]


def _three_line_case():
    """AAA and CCC are fine, BBB is 8 short of the 12 asked for."""
    inv = pd.DataFrame([
        {"Item Number": c, "Lot Number": "L", "Pallet ID": f"P{i}", "Location Id": "A",
         "Actual Qty": q, "Plant": "PL1", "Status": "Available", "Pick Id": "0",
         "UOM": "EA", "Description": "d"}
        for i, (c, q) in enumerate([("AAA", 100), ("BBB", 4), ("CCC", 50)])])
    return inv, _doc("I1", [("AAA", 10), ("BBB", 12), ("CCC", 6)])


def test_the_offer_shows_both_ways_of_sending_it():
    inv, doc = _three_line_case()
    row = PE.partialable(PE.run_pick([doc], inv, PE.EngineConfig())).iloc[0]
    assert row["DOC_QTY"] == 28
    assert row["CAN_PICK_NOW"] == 20        # 10 + 4 on the floor + 6
    assert row["WHOLE_LINES_ONLY"] == 16    # the two complete lines only
    assert row["COMPLETE_LINES"] == 2 and row["SHORT_LINES"] == 1


def test_whole_mode_leaves_the_short_item_off_the_load():
    inv, doc = _three_line_case()
    res = PE.run_pick([doc], inv,
                      PE.EngineConfig(partial_docs=["I1"], partial_mode="whole"))
    per_line = res["allocations"].groupby("DOC_LINE")["QTY_PICKED"].sum()
    assert 2 not in per_line                 # nothing of the short item goes out
    assert per_line[1] == 10 and per_line[3] == 6
    a = res["accepted"].iloc[0]
    assert a["PICK_STATUS"] == "PARTIAL" and a["PICKED_QTY"] == 16
    assert res["detail"]["QTY"].astype(float).sum() == 16
    # the whole of the short line is still owed, not the balance of a split
    owed = res["shortage"].set_index("DOC_LINE")["SHORT"]
    assert owed[2] == 12


def test_a_line_held_back_by_hand_stays_owed():
    inv, doc = _three_line_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"],
                                                  skip_lines={"I1": [3]}))
    per_line = res["allocations"].groupby("DOC_LINE")["QTY_PICKED"].sum()
    assert 3 not in per_line
    held = res["shortage"][res["shortage"]["DOC_LINE"] == 3].iloc[0]
    assert held["SHORT"] == 6
    assert "Left out" in held["REASON"]
    assert res["accepted"].iloc[0]["PICK_STATUS"] == "PARTIAL"


def test_holding_a_line_back_is_itself_a_partial_pick():
    """No partial_docs confirmation, but a held line still must not fail the run."""
    inv, doc = _three_line_case()
    doc.lines[1].qty = 4                     # BBB now fits, so nothing is short
    res = PE.run_pick([doc], inv, PE.EngineConfig(skip_lines={"I1": [1]}))
    a = res["accepted"].iloc[0]
    assert a["PICK_STATUS"] == "PARTIAL"
    assert a["PICKED_QTY"] == 10             # BBB 4 + CCC 6, AAA held back
    assert len(res["rejected"]) == 0


def test_holding_every_line_back_sends_nothing():
    inv, doc = _three_line_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"],
                                                  skip_lines={"I1": [1, 2, 3]}))
    assert len(res["accepted"]) == 0
    assert len(res["allocations"]) == 0


def test_the_balance_is_picked_later_and_only_the_balance():
    inv, doc = _short_case()
    first = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    prev = G.picked_lines_from(first["allocations"])
    assert prev == {"I1": {1: 10.0, 2: 4.0}}
    # BBB has been restocked
    inv2 = inv.copy(); inv2.loc[inv2["Item Number"] == "BBB", "Actual Qty"] = 50
    second = PE.run_pick([doc], inv2, PE.EngineConfig(), picked_before=prev)
    a = second["accepted"].iloc[0]
    assert a["PICK_STATUS"] == "FULL"
    assert a["PICKED_QTY"] == 8 and a["PREV_QTY"] == 14 and a["TOTAL_PICKED"] == 22
    assert a["SHORT_QTY"] == 0
    assert second["detail"]["QTY"].astype(float).sum() == 8      # the balance only
    assert set(second["allocations"]["DOC_LINE"]) == {2}         # line 1 was done


def test_a_partly_picked_document_stays_open_for_its_balance():
    reg = pd.DataFrame([
        {"DOC_NUMBER": "FULL1", "PICK_STATUS": "FULL"},
        {"DOC_NUMBER": "PART1", "PICK_STATUS": "PARTIAL"},
        {"DOC_NUMBER": "PART2", "PICK_STATUS": "PARTIAL"},
        {"DOC_NUMBER": "PART2", "PICK_STATUS": "FULL"},      # balance went out
    ])
    assert G.open_docs(reg) == {"FULL1", "PART2"}
    # rows written before partial picks existed have no status and were all full
    assert G.open_docs(pd.DataFrame([{"DOC_NUMBER": "OLD1"}])) == {"OLD1"}


def test_the_register_calls_a_partial_document_partial():
    inv, doc = _short_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    s, d = R.build([doc], res, R.DEFAULT_MRP, user="t", plant="PL1")
    row = s.iloc[0]
    assert row["KORBER_PICK"] == R.PICK_PART
    assert row["QTY"] == 22 and row["PICKED_QTY"] == 14
    assert "still owed" in str(row["REMARK"])
    # line by line: the full line is Yes, the short line is Partial
    assert d.loc[d["LINE"] == 1, "KORBER_PICK"].iloc[0] == R.PICK_YES
    assert d.loc[d["LINE"] == 2, "KORBER_PICK"].iloc[0] == R.PICK_PART
    assert d.loc[d["LINE"] == 2, "PICKED_QTY"].iloc[0] == 4


def test_the_register_closes_the_document_when_the_balance_goes_out():
    inv, doc = _short_case()
    first = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    s1, d1 = R.build([doc], first, R.DEFAULT_MRP)
    inv2 = inv.copy(); inv2.loc[inv2["Item Number"] == "BBB", "Actual Qty"] = 50
    second = PE.run_pick([doc], inv2, PE.EngineConfig(),
                         picked_before=G.picked_lines_from(first["allocations"]))
    s2, d2 = R.build([doc], second, R.DEFAULT_MRP)
    assert s2.iloc[0]["KORBER_PICK"] == R.PICK_YES
    assert s2.iloc[0]["PICKED_QTY"] == 22          # both runs together
    merged = R.merge_summary(s1, s2)["data"]
    assert merged.iloc[0]["KORBER_PICK"] == R.PICK_YES
    md = R.merge_details(d1, d2)
    assert set(md["KORBER_PICK"]) == {R.PICK_YES}


def test_a_partial_is_never_walked_back_to_not_picked():
    inv, doc = _short_case()
    part = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    s1, d1 = R.build([doc], part, R.DEFAULT_MRP)
    # a later run with no stock at all takes nothing
    nothing = PE.run_pick([doc], inv.assign(**{"Actual Qty": [0, 0]}), PE.EngineConfig())
    s2, d2 = R.build([doc], nothing, R.DEFAULT_MRP)
    merged = R.merge_summary(s1, s2)["data"]
    assert merged.iloc[0]["KORBER_PICK"] == R.PICK_PART
    assert float(merged.iloc[0]["PICKED_QTY"]) == 14.0
    md = R.merge_details(d1, d2)
    assert md.loc[md["LINE"] == 1, "KORBER_PICK"].iloc[0] == R.PICK_YES
    assert md.loc[md["LINE"] == 2, "PICKED_QTY"].iloc[0] == 4


def test_the_dashboard_counts_a_partial_as_still_owed():
    inv, doc = _short_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    s, _d = R.build([doc], res, R.DEFAULT_MRP)
    k = R.dashboard(s)["kpi"]
    assert k["total"] == 1 and k["picked"] == 0 and k["partial"] == 1
    assert k["pending"] == 1                    # not finished
    assert k["qty_owed"] == 8                   # but only the balance is left
    assert R.dashboard(s)["pending"].iloc[0]["REASON"] == "Partially picked"


def test_the_pick_sheet_says_it_is_partial():
    import pick_pdf as PP
    inv, doc = _short_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    b = PE.doc_bundle(res, "I1")
    assert b["info"]["PICK_STATUS"] == "PARTIAL"
    assert b["info"]["SHORT_QTY"] == 8
    pdf = PP.build_pick_sheet(b["info"], b["allocations"], b["verify"])
    assert pdf[:4] == b"%PDF"
    assert len(pdf) > 1000



# --------------------------------------------------------------------------- #
# manual entry — a pick with no PDF in hand
# --------------------------------------------------------------------------- #
def _typed(rows):
    return pd.DataFrame(rows, columns=["Item Code", "Description", "Qty", "Doc UOM"])


def test_a_typed_document_is_a_document():
    from doc_parser import manual_doc, manual_frame
    d = manual_doc("30426013174",
                   _typed([{"Item Code": "P601560 710", "Description": "filter",
                            "Qty": 10, "Doc UOM": "EA"},
                           {"Item Code": "", "Description": "", "Qty": None,
                            "Doc UOM": "EA"},          # blank grid row
                           {"Item Code": "1c072323-inl", "Description": "",
                            "Qty": 4, "Doc UOM": ""}]),
                   doc_date="21-AUG-2026", customer="ACME", ref_number="AR9")
    ok, problems = d.completeness()
    assert ok, problems
    assert len(d.lines) == 2                       # the blank row is not a line
    assert [ln.line_no for ln in d.lines] == [1, 2]
    assert d.lines[0].item_code == "P601560 710"   # separators kept
    assert d.lines[0].base == "P601560"
    assert d.lines[1].base == "1C072323"
    assert d.lines[1].uom == "EA"                  # blank UOM defaults
    assert d.source_file == "manual entry"
    assert d.declared_qty is None                  # nothing to cross-check against
    # an empty grid is not a document
    assert not manual_doc("X", manual_frame()).completeness()[0]


def test_a_typed_document_picks_exactly_like_a_parsed_one():
    from doc_parser import manual_doc
    typed = manual_doc("30426013174",
                       _typed([{"Item Code": "P601560", "Qty": 10, "Doc UOM": "EA",
                                "Description": ""}]),
                       doc_date="21-AUG-2026", customer="ACME")
    parsed = _doc("30426013174", [("P601560 710", 10)])
    a = PE.run_pick([typed], _inv(["P601560 710"], qty=50), PE.EngineConfig())
    b = PE.run_pick([parsed], _inv(["P601560 710"], qty=50), PE.EngineConfig())
    assert len(a["accepted"]) == len(b["accepted"]) == 1
    assert a["accepted"].iloc[0]["PICKED_QTY"] == b["accepted"].iloc[0]["PICKED_QTY"]
    assert list(a["detail"]["QTY"]) == list(b["detail"]["QTY"])
    assert list(a["detail"]["DISPLAY_ITEM_NUMBER"]) == list(b["detail"]["DISPLAY_ITEM_NUMBER"])
    # and it reaches the register the same way
    s, d = R.build([typed], a, R.DEFAULT_MRP, user="t", plant="PL1")
    assert s.iloc[0]["KORBER_PICK"] == R.PICK_YES
    assert s.iloc[0]["TAX_INVOICE_NO"] == "30426013174"
    assert float(s.iloc[0]["QTY"]) == 10.0
    assert d.iloc[0]["BASE_ID"] == "P601560"


def test_a_typed_document_is_short_checked_like_any_other():
    from doc_parser import manual_doc
    typed = manual_doc("I9", _typed([{"Item Code": "P601560", "Qty": 999,
                                      "Doc UOM": "EA", "Description": ""}]))
    res = PE.run_pick([typed], _inv(["P601560 710"], qty=50), PE.EngineConfig())
    assert len(res["accepted"]) == 0
    assert "STOCK SHORT" in res["rejected"].iloc[0]["REASON"]
    assert len(PE.partialable(res)) == 1          # partial pick is offered


def test_the_review_table_round_trips_a_typed_document():
    from doc_parser import manual_doc, docs_to_frame, frame_to_docs
    d = manual_doc("I1", _typed([{"Item Code": "X770132 003710", "Qty": 3,
                                  "Doc UOM": "EA", "Description": "d"}]),
                   customer="ACME", ref_number="AR1")
    back = frame_to_docs(docs_to_frame([d]), [d])[0]
    assert back.doc_number == "I1"
    assert back.customer == "ACME" and back.ref_number == "AR1"   # not dropped
    assert back.lines[0].item_code == "X770132 003710"            # separators kept
    assert back.lines[0].base == "X770132"


# --------------------------------------------------------------------------- #
# period filter — today / a range / everything
# --------------------------------------------------------------------------- #
def _period_register():
    """Invoices dated today, yesterday, 10 days ago, and one with no date."""
    today = datetime.now().date()
    rows = []
    for tag, days, picked in (("TODAY", 0, True), ("YDAY", 1, True),
                              ("OLD", 10, False)):
        d = today - timedelta(days=days)
        r = {c: "" for c in R.SUMMARY_COLS}
        r.update({"TAX_INVOICE_NO": tag, "CUSTOMER_NAME": "ACME",
                  "DOC_TYPE": "INVOICE", "QTY": 10, "LINES": 1, "MRP": "No",
                  "TAX_INVOICE_DATE": d.strftime("%d-%b-%Y").upper(),
                  "PICKED_AT": d.strftime("%Y-%m-%d 09:00:00") if picked else "",
                  "UPDATED_AT": d.strftime("%Y-%m-%d 09:00:00"),
                  "KORBER_PICK": R.PICK_YES if picked else R.PICK_NO,
                  "PICKED_QTY": 10 if picked else 0})
        for c in R.STATUS_COLS:
            r[c] = R.STATUS_PENDING
        rows.append(r)
    blank = {c: "" for c in R.SUMMARY_COLS}
    blank.update({"TAX_INVOICE_NO": "NODATE", "CUSTOMER_NAME": "ACME",
                  "DOC_TYPE": "INVOICE", "QTY": 5, "LINES": 1, "MRP": "No",
                  "KORBER_PICK": R.PICK_NO, "PICKED_QTY": 0})
    for c in R.STATUS_COLS:
        blank[c] = R.STATUS_PENDING
    rows.append(blank)
    return pd.DataFrame(rows, columns=R.SUMMARY_COLS)


def test_the_presets_pick_the_right_days():
    t = datetime(2026, 8, 21)
    assert R.date_preset("Today", t) == (t.date(), t.date())
    assert R.date_preset("Yesterday", t) == (datetime(2026, 8, 20).date(),) * 2
    assert R.date_preset("Last 7 days", t) == (datetime(2026, 8, 15).date(), t.date())
    assert R.date_preset("Last 30 days", t) == (datetime(2026, 7, 23).date(), t.date())
    assert R.date_preset("This month", t) == (datetime(2026, 8, 1).date(), t.date())
    assert R.date_preset("All time", t) == (None, None)
    assert R.date_preset("Custom range", t) == (None, None)


def test_today_means_today_on_both_tabs():
    s = _period_register()
    f, t = R.date_preset("Today")
    # invoice date: today's invoice, plus the one whose date could not be read —
    # a missing invoice date must not hide outstanding work
    by_inv = R.filter_by_date(s, f, t, "TAX_INVOICE_DATE")
    assert set(by_inv["TAX_INVOICE_NO"]) == {"TODAY", "NODATE"}
    # picked date: only what was actually picked today
    by_pick = R.filter_by_date(s, f, t, "PICKED_AT")
    assert set(by_pick["TAX_INVOICE_NO"]) == {"TODAY"}
    # the dashboard reads the same rule, so the two tabs agree
    assert R.dashboard(s, f, t, None, date_col="TAX_INVOICE_DATE")["kpi"]["total"] == 2
    assert R.dashboard(s, f, t, None, date_col="PICKED_AT")["kpi"]["total"] == 1


def test_a_range_and_the_full_register():
    s = _period_register()
    f, t = R.date_preset("Last 7 days")
    assert set(R.filter_by_date(s, f, t, "TAX_INVOICE_DATE")["TAX_INVOICE_NO"]) == {
        "TODAY", "YDAY", "NODATE"}
    # All time is every row, undated included
    assert len(R.filter_by_date(s, None, None, "TAX_INVOICE_DATE")) == len(s)
    assert R.dashboard(s)["kpi"]["total"] == len(s)
    # a custom single day is inclusive of that whole day
    y = (datetime.now().date() - timedelta(days=1))
    assert set(R.filter_by_date(s, y, y, "PICKED_AT")["TAX_INVOICE_NO"]) == {"YDAY"}


def test_the_period_filter_keeps_the_summary_and_the_details_in_step():
    s = _period_register()
    f, t = R.date_preset("Today")
    dash = R.dashboard(s, f, t, None, date_col="TAX_INVOICE_DATE")
    det = pd.DataFrame([{**{c: "" for c in R.DETAIL_COLS},
                         "TAX_INVOICE_NO": n, "LINE": 1, "DOC_QTY": 10}
                        for n in ("TODAY", "YDAY", "OLD", "NODATE")],
                       columns=R.DETAIL_COLS)
    kept = R.details_for(det, dash["invoices"])
    assert set(kept["TAX_INVOICE_NO"]) == {"TODAY", "NODATE"}



def test_a_dead_line_does_not_hide_what_the_rest_of_the_document_can_send():
    """
    One line with no stock at all, three lines with plenty.

    The offer used to be worked out from the shortage table, which holds only
    the *short* lines — so this document showed 0 available and was dropped off
    the partial list, even though three quarters of it was sitting on the floor
    ready to go.
    """
    inv = pd.concat([_inv(["AAA"], qty=100), _inv(["BBB"], qty=100),
                     _inv(["CCC"], qty=100)], ignore_index=True)
    doc = _doc("I1", [("AAA", 5), ("BBB", 5), ("CCC", 5), ("R010077", 3)])
    res = PE.run_pick([doc], inv, PE.EngineConfig())
    assert len(res["accepted"]) == 0                       # all-or-nothing default
    offer = PE.partialable(res)
    assert list(offer["DOC_NUMBER"]) == ["I1"], offer.to_dict("records")
    row = offer.iloc[0]
    assert row["DOC_QTY"] == 18 and row["CAN_PICK_NOW"] == 15 and row["STILL_SHORT"] == 3
    assert row["LINES"] == 4 and row["SHORT_LINES"] == 1
    assert len(PE.no_partial(res)) == 0
    # and confirming it really does send those 15
    go = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    assert go["partial"] == ["I1"]
    assert go["detail"]["QTY"].astype(float).sum() == 15
    assert go["accepted"].iloc[0]["SHORT_QTY"] == 3


def test_a_document_with_nothing_at_all_is_named_not_hidden():
    """The whole document has no stock — there is nothing to offer, and the
    screen has to say so rather than leave the user hunting for a section."""
    doc = _doc("333262712447", [("R010077", 3)])
    res = PE.run_pick([doc], _inv(["AAA"], qty=100), PE.EngineConfig())
    assert len(PE.partialable(res)) == 0                   # nothing to send
    dead = PE.no_partial(res)
    assert list(dead["DOC_NUMBER"]) == ["333262712447"]
    assert dead.iloc[0]["CAN_PICK_NOW"] == 0
    assert dead.iloc[0]["STILL_SHORT"] == 3
    # and forcing a partial anyway still refuses it
    forced = PE.run_pick([doc], _inv(["AAA"], qty=100), PE.EngineConfig(partial_docs=["*"]))
    assert len(forced["accepted"]) == 0


def test_the_offer_counts_what_an_earlier_partial_already_sent():
    inv, doc = _short_case()
    first = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    prev = G.picked_lines_from(first["allocations"])
    # still short on the balance run, and still nothing new on the floor
    second = PE.run_pick([doc], inv.assign(**{"Actual Qty": [100, 0]}),
                         PE.EngineConfig(), picked_before=prev)
    offer = PE.partialable(second)
    if len(offer):
        row = offer.iloc[0]
        assert row["ALREADY_SENT"] == 14
        assert row["DOC_QTY"] == 22
    else:
        assert PE.no_partial(second).iloc[0]["ALREADY_SENT"] == 14



# --------------------------------------------------------------------------- #
# the pick email — document qty vs what is actually being picked
# --------------------------------------------------------------------------- #
def _mail(res, load_id):
    import pick_pdf as PP
    b = PE.doc_bundle(res, load_id)
    return PP.pick_email_text([b["info"]], b["allocations"], "Thanks,",
                              verify=b["verify"])


def test_the_email_never_calls_the_document_qty_picked():
    """A full pick: the two happen to be equal, and both are stated."""
    inv, doc = _short_case()
    inv.loc[inv["Item Number"] == "BBB", "Actual Qty"] = 100
    res = PE.run_pick([doc], inv, PE.EngineConfig())
    subj, body, html = _mail(res, "I1")
    assert not subj.startswith("PARTIAL")
    assert "Document qty  : 22 pcs over 2 lines" in body
    assert "Picking now   : 22 pcs" in body
    assert "Total document qty: 22" in body
    assert "Total picked qty  : 22" in body
    assert "Still owed" not in body


def test_the_partial_email_says_what_is_short_line_by_line():
    inv, doc = _short_case()                 # AAA 10 ok, BBB 12 asked / 4 there
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    subj, body, html = _mail(res, "I1")
    assert subj.startswith("PARTIAL ")
    assert "PARTIAL PICK — I1. 14 of 22 pcs are being picked; 8 pcs are still owed" in body
    assert "[PARTIAL]" in body
    assert "Document qty  : 22 pcs over 2 lines" in body
    assert "Picking now   : 14 pcs" in body
    assert "Still owed    : 8 pcs" in body
    # the line table carries doc qty / picked / short for every line
    assert "Line summary — document qty vs picked" in body
    rows = [l for l in body.splitlines() if l.startswith(("1 ", "2 "))]
    assert any(r.split() == ["1", "AAA", "AAA", "10", "10", "0"] for r in rows), rows
    assert any(r.split() == ["2", "BBB", "BBB", "12", "4", "8"] for r in rows), rows
    assert "Total document qty: 22" in body
    assert "Total picked qty  : 14" in body
    assert "Still owed        : 8" in body
    # and the HTML part says the same
    assert "PARTIAL PICK" in html and "Still owed" in html


def test_the_balance_email_shows_what_went_out_earlier():
    inv, doc = _short_case()
    first = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    inv2 = inv.copy(); inv2.loc[inv2["Item Number"] == "BBB", "Actual Qty"] = 50
    second = PE.run_pick([doc], inv2, PE.EngineConfig(),
                         picked_before=G.picked_lines_from(first["allocations"]))
    subj, body, _html = _mail(second, "I1")
    assert not subj.startswith("PARTIAL")          # this run completes it
    assert "Picking now   : 8 pcs" in body
    assert "Already sent  : 14 pcs (earlier pick)" in body
    assert "Total picked qty  : 8" in body
    assert "Already sent      : 14 (earlier pick)" in body
    assert "Delivered in all  : 22" in body
    assert "Still owed" not in body
    # every line is now complete against the document
    assert "12      12      0" in body or "12  12  0" in body.replace("  ", "  ")


def test_the_email_totals_add_up_across_several_documents():
    import pick_pdf as PP
    inv = pd.concat([_inv(["AAA"], qty=100), _inv(["BBB"], qty=4)], ignore_index=True)
    docs = [_doc("I1", [("AAA", 10)]), _doc("I2", [("AAA", 5), ("BBB", 12)])]
    res = PE.run_pick(docs, inv, PE.EngineConfig(partial_docs=["I2"]))
    infos = [PE.doc_bundle(res, l)["info"] for l in PE.load_ids(res)]
    _s, body, _h = PP.pick_email_text(infos, res["allocations"], "Thanks,",
                                      verify=res["verify"])
    doc_qty = sum(float(i["TOTAL_QTY"]) for i in infos)
    picked = sum(float(i["PICKED_QTY"]) for i in infos)
    assert f"Total document qty: {int(doc_qty)}" in body      # 27
    assert f"Total picked qty  : {int(picked)}" in body       # 19
    assert picked == res["allocations"]["QTY_PICKED"].sum()
    assert "Still owed        : 8" in body


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


# =========================================================================== #
# Transactions History Report — what the floor actually picked
# =========================================================================== #
import transactions as TX


def _legs():
    """One pallet leaving the rack, then moving on twice — three rows, one pick."""
    return pd.DataFrame([
        {"Control Number": "CL01-INV1", "Starting Hu": "PAL-A", "Ending Hu": "PAL-A",
         "Item Number": "CL01-AAA", "Tran Qty": "10", "Starting Loc": "RACK-1",
         "Ending Loc": "PICKER", "End Tran Date": "12-08-2026 10:00:00",
         "Employee Id": "PICKER", "Client Code": "CL01", "Lot Number": "L"},
        {"Control Number": "CL01-INV1", "Starting Hu": "PAL-A", "Ending Hu": "PAL-A",
         "Item Number": "CL01-AAA", "Tran Qty": "10", "Starting Loc": "PICKER",
         "Ending Loc": "STAGE", "End Tran Date": "12-08-2026 11:00:00",
         "Employee Id": "PICKER", "Client Code": "CL01", "Lot Number": "L"},
        {"Control Number": "CL01-INV1", "Starting Hu": "PAL-A", "Ending Hu": "PAL-A",
         "Item Number": "CL01-AAA", "Tran Qty": "10", "Starting Loc": "STAGE",
         "Ending Loc": "DOCK", "End Tran Date": "12-08-2026 12:00:00",
         "Employee Id": "DOCK", "Client Code": "CL01", "Lot Number": "L"},
    ])


def test_the_load_id_comes_out_of_the_control_number():
    assert TX.strip_client("INM0DONA-333262712295", "INM0DONA") == "333262712295"
    assert TX.strip_client("INM0DONA-333/26-27/17", "INM0DONA") == "333/26-27/17"
    assert TX.strip_client("333262712295", "INM0DONA") == "333262712295"


def test_a_movement_chain_is_one_pick_not_three():
    act = TX.normalize(_legs(), "CL01")
    assert len(act) == 1
    assert float(act.iloc[0]["QTY"]) == 10
    assert act.iloc[0]["LOAD_ID"] == "INV1"
    assert act.iloc[0]["ITEM_NUMBER"] == "AAA"     # the client code is stripped


def test_a_second_pick_off_the_same_pallet_survives_the_collapse():
    legs = pd.concat([_legs(), pd.DataFrame([{
        "Control Number": "CL01-INV1", "Starting Hu": "PAL-A", "Ending Hu": "PAL-A",
        "Item Number": "CL01-AAA", "Tran Qty": "4", "Starting Loc": "RACK-2",
        "Ending Loc": "PICKER", "End Tran Date": "13-08-2026 09:00:00",
        "Employee Id": "PICKER", "Client Code": "CL01", "Lot Number": "L"}])],
        ignore_index=True)
    act = TX.normalize(legs, "CL01")
    assert float(act["QTY"].sum()) == 14


def test_reconcile_releases_a_pallet_the_floor_never_touched():
    act = TX.normalize(_legs(), "CL01")
    led = pd.DataFrame([
        {"DOC_NUMBER": "INV1", "PALLET": "PAL-A", "QTY_PICKED": 10,
         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA"},
        {"DOC_NUMBER": "INV1", "PALLET": "PAL-GHOST", "QTY_PICKED": 6,
         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA"}])
    rec = TX.reconcile(act, led)
    assert rec["totals"]["agree"] == 1
    assert rec["totals"]["release_qty"] == 6
    corr = TX.ledger_corrections(rec)
    ghost = corr[corr["PALLET"] == "PAL-GHOST"].iloc[0]
    assert float(ghost["QTY_PICKED"]) == -6      # negative = handed back


def test_reconcile_consumes_a_pallet_the_system_never_chose():
    act = TX.normalize(_legs(), "CL01")
    led = pd.DataFrame([{"DOC_NUMBER": "INV1", "PALLET": "PAL-OTHER",
                         "QTY_PICKED": 10, "ITEM_NUMBER": "AAA", "BASE_ID": "AAA"}])
    rec = TX.reconcile(act, led)
    assert rec["totals"]["consume_qty"] == 10
    assert rec["totals"]["release_qty"] == 10
    corr = TX.ledger_corrections(rec)
    assert float(corr.loc[corr["PALLET"] == "PAL-A", "QTY_PICKED"].iat[0]) == 10
    assert float(corr.loc[corr["PALLET"] == "PAL-OTHER", "QTY_PICKED"].iat[0]) == -10


def test_a_load_not_in_this_system_is_left_alone():
    act = TX.normalize(_legs(), "CL01")
    led = pd.DataFrame([{"DOC_NUMBER": "SOMETHING-ELSE", "PALLET": "P", "QTY_PICKED": 1,
                         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA"}])
    rec = TX.reconcile(act, led)
    assert rec["unknown"] == ["INV1"]
    assert not len(rec["rows"])


def test_a_revision_suffix_still_matches_its_load():
    legs = _legs().assign(**{"Control Number": "CL01-INV1-A"})
    act = TX.normalize(legs, "CL01")
    led = pd.DataFrame([{"DOC_NUMBER": "INV1", "PALLET": "PAL-A", "QTY_PICKED": 10,
                         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA"}])
    assert TX.reconcile(act, led)["loads"] == {"INV1-A": "INV1"}


def test_the_real_report_reads_and_reconciles():
    raw = pd.read_excel("/mnt/user-data/uploads/Transactions_History_Report.xlsx",
                        dtype=str)
    act = TX.normalize(raw, "INM0DONA")
    assert len(act) and (act["PALLET"] != "").all()
    one = act[act["LOAD_ID"] == "333262712295"]
    assert float(one["QTY"].sum()) == 130        # 467 before the chains collapse


def test_a_correction_actually_moves_the_pallet_balance():
    """The whole point: a release has to reach `ledger_state`, not sit beside it."""
    act = pd.DataFrame([{"LOAD_ID": "L1", "CONTROL_NUMBER": "C", "PALLET": "P1",
                         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA", "LOT_NUMBER": "",
                         "QTY": 15.0, "FROM_LOC": "R", "TO_LOC": "S", "WHEN": "",
                         "EMPLOYEE": ""}])
    led = pd.DataFrame([{"DOC_NUMBER": "L1", "PALLET": "P1", "ITEM_NUMBER": "AAA",
                         "LOT_NUMBER": "LOT9", "QTY_BEFORE": 20, "QTY_PICKED": 20,
                         "QTY_BALANCE": 0, "ROW_KEY": "P1|LOCA|AAA|LOT9|",
                         "BASE_ID": "AAA", "LOCATION_ID": "LOCA", "PLANT": "PL1",
                         "UOM": "EA"}])
    corr = TX.ledger_corrections(TX.reconcile(act, led))
    assert corr.iloc[0]["ROW_KEY"] == "P1|LOCA|AAA|LOT9|"   # binds to the same pallet
    assert float(corr.iloc[0]["QTY_BEFORE"]) == 20          # baseline is not dragged to 0
    before = PE.ledger_state(led)[0]["P1|LOCA|AAA|LOT9|"]
    after = PE.ledger_state(pd.concat([led, corr], ignore_index=True))[0][
        "P1|LOCA|AAA|LOT9|"]
    assert before["balance"] == 0 and after["balance"] == 5


def test_released_stock_can_be_picked_again_on_the_next_run():
    """End to end: reserve it, find out the floor never took it, pick it again."""
    inv = pd.DataFrame([{"Item Number": "AAA", "Lot Number": "L", "Pallet ID": "P1",
                         "Location Id": "A", "Actual Qty": 20, "Plant": "PL1",
                         "Status": "Available", "Pick Id": "0", "UOM": "EA",
                         "Description": "d"}])
    r1 = PE.run_pick([_doc("L1", [("AAA", 20)])], inv, PE.EngineConfig())
    led = r1["allocations"]
    assert float(led["QTY_PICKED"].sum()) == 20

    # the same inventory export is uploaded again, so the pallet still reads 20
    blocked = PE.run_pick([_doc("L2", [("AAA", 5)])], inv, PE.EngineConfig(),
                          ledger=led)
    assert len(blocked["rejected"]) == 1          # ledger says the pallet is spent

    # the report shows the floor never touched it
    act = pd.DataFrame([{"LOAD_ID": "L1", "CONTROL_NUMBER": "C", "PALLET": "ELSEWHERE",
                         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA", "LOT_NUMBER": "",
                         "QTY": 20.0, "FROM_LOC": "R", "TO_LOC": "S", "WHEN": "",
                         "EMPLOYEE": ""}])
    corr = TX.ledger_corrections(TX.reconcile(act, led))
    fixed = pd.concat([led, corr], ignore_index=True)
    ok = PE.run_pick([_doc("L3", [("AAA", 5)])], inv, PE.EngineConfig(), ledger=fixed)
    assert len(ok["rejected"]) == 0
    assert float(ok["allocations"]["QTY_PICKED"].sum()) == 5


def test_keys_are_borrowed_for_a_pallet_the_system_never_chose():
    act = pd.DataFrame([{"LOAD_ID": "L1", "CONTROL_NUMBER": "C", "PALLET": "P9",
                         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA", "LOT_NUMBER": "",
                         "QTY": 4.0, "FROM_LOC": "R", "TO_LOC": "S", "WHEN": "",
                         "EMPLOYEE": ""}])
    led = pd.DataFrame([{"DOC_NUMBER": "L1", "PALLET": "P1", "ITEM_NUMBER": "AAA",
                         "LOT_NUMBER": "LOT1", "QTY_BEFORE": 10, "QTY_PICKED": 4,
                         "QTY_BALANCE": 6, "ROW_KEY": "P1|A|AAA|LOT1|",
                         "BASE_ID": "AAA", "LOCATION_ID": "A", "PLANT": "PL1",
                         "UOM": "EA"}])
    keys = pd.DataFrame([{"PALLET": "P9", "ROW_KEY": "P9|B|AAA|LOT7|",
                          "LOT_NUMBER": "LOT7", "LOCATION_ID": "B", "PLANT": "PL1",
                          "UOM": "EA", "QTY_BEFORE": 9}])
    corr = TX.ledger_corrections(TX.reconcile(act, led, keys=keys))
    p9 = corr[corr["PALLET"] == "P9"].iloc[0]
    assert p9["ROW_KEY"] == "P9|B|AAA|LOT7|" and p9["LOT_NUMBER"] == "LOT7"


def test_a_correction_cannot_be_applied_twice():
    """Apply, then complete the load with the same frame — the stock must not move again."""
    led = pd.DataFrame([{"DOC_NUMBER": "L1", "PALLET": "P1", "ITEM_NUMBER": "AAA",
                         "LOT_NUMBER": "LOT9", "QTY_BEFORE": 20, "QTY_PICKED": 20,
                         "QTY_BALANCE": 0, "ROW_KEY": "P1|A|AAA|LOT9|",
                         "BASE_ID": "AAA", "LOCATION_ID": "A", "PLANT": "PL1",
                         "UOM": "EA", "RUN_ID": "R1"}])
    act = pd.DataFrame([{"LOAD_ID": "L1", "CONTROL_NUMBER": "C", "PALLET": "P1",
                         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA", "LOT_NUMBER": "",
                         "QTY": 15.0, "FROM_LOC": "R", "TO_LOC": "S", "WHEN": "",
                         "EMPLOYEE": ""}])
    corr = TX.ledger_corrections(TX.reconcile(act, led))
    one = pd.concat([led, corr], ignore_index=True)
    assert PE.ledger_state(one)[0]["P1|A|AAA|LOT9|"]["balance"] == 5
    assert not len(TX.already_applied(corr, one))          # nothing left to write
    two = pd.concat([one, TX.already_applied(corr, one)], ignore_index=True)
    assert PE.ledger_state(two)[0]["P1|A|AAA|LOT9|"]["balance"] == 5


def test_re_comparing_after_applying_finds_nothing_to_do():
    led = pd.DataFrame([{"DOC_NUMBER": "L1", "PALLET": "P1", "ITEM_NUMBER": "AAA",
                         "LOT_NUMBER": "LOT9", "QTY_BEFORE": 20, "QTY_PICKED": 20,
                         "QTY_BALANCE": 0, "ROW_KEY": "P1|A|AAA|LOT9|",
                         "BASE_ID": "AAA", "LOCATION_ID": "A", "PLANT": "PL1",
                         "UOM": "EA"}])
    act = pd.DataFrame([{"LOAD_ID": "L1", "CONTROL_NUMBER": "C", "PALLET": "P1",
                         "ITEM_NUMBER": "AAA", "BASE_ID": "AAA", "LOT_NUMBER": "",
                         "QTY": 15.0, "FROM_LOC": "R", "TO_LOC": "S", "WHEN": "",
                         "EMPLOYEE": ""}])
    fixed = pd.concat([led, TX.ledger_corrections(TX.reconcile(act, led))],
                      ignore_index=True)
    again = TX.reconcile(act, fixed)
    assert again["rows"].iloc[0]["OUTCOME"] == TX.AGREES
    assert not len(TX.ledger_corrections(again))


def test_every_written_frame_fits_its_worksheet_header():
    """A column the header does not know is silently dropped on write."""
    import gsheet as G
    inv, doc = _three_line_case()
    res = PE.run_pick([doc], inv, PE.EngineConfig(partial_docs=["I1"]))
    s_df, d_df = R.build([doc], res, R.load_contacts(None))
    for title, df in ((G.WS_MASTER, res["master"]), (G.WS_DETAIL, res["detail"]),
                      (G.WS_LEDGER, res["allocations"]),
                      (G.WS_INV_SUM, s_df), (G.WS_INV_DET, d_df)):
        extra = set(df.columns) - set(G._SHEETS[title])
        assert not extra, f"{title} would drop {sorted(extra)}"
