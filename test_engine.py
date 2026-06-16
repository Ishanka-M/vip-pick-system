"""
Validation harness for the VIP pick engine (new HJ/SAP logic).

Pick formula:  target_pcs = Req Qty * HJ / SAP
  - LOOSE: SAP == HJ -> target = Req Qty
  - SET  : HJ/SAP is the set multiplier
A carton (inventory Actual Qty) can't be split, so lines whose target is not a
whole multiple of the carton size are routed to a separate 'Cannot Pick' report,
as are lines where Req exceeds available stock.

Run:  python test_engine.py   (or)   python -m pytest test_engine.py
"""
import os
import io
import pandas as pd
import pick_engine as pe
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(HERE, "sample_data")
REQ = os.path.join(SAMPLE, "1781617499524_Requament.xlsx")
SKU = os.path.join(SAMPLE, "1781617499524_SKU_MASTER.xlsx")
INV = os.path.join(SAMPLE, "1781617499523_Inventory_Report.xlsx")


def _run(existing=None):
    req = pd.read_excel(REQ)
    sku = pd.read_excel(SKU)
    inv = pd.read_excel(INV)
    return pe.run_pipeline(req, sku, inv, pe.EngineConfig(), existing_load_ids=existing)


def test_formula_example():
    """Req=2, SAP=1, HJ=2 -> pick 4 (the spec's worked example)."""
    req = pd.DataFrame({"Material": ["X1"], "Material Description": ["d"],
                        "Req Qty.": [2], "Delivery No": [1000]})
    sku = pd.DataFrame({"Material code": ["X1"], "Material Desc": ["d"],
                        "Catergory": ["SET"], "HJ": [2], "SAP": [1]})
    inv = pd.DataFrame({"Item Number": ["X1", "X1", "X1"], "Actual Qty": [4, 4, 4],
                        "Cbm": [0.1, 0.1, 0.1]})
    res = pe.run_pipeline(req, sku, inv, pe.EngineConfig())
    assert int(res["pick"]["HJ Pcs Qty"].iloc[0]) == 4, "expected pick 4"


def test_loose_equals_req():
    """LOOSE (SAP==HJ) picks exactly Req Qty pieces."""
    req = pd.DataFrame({"Material": ["L1"], "Material Description": ["d"],
                        "Req Qty.": [6], "Delivery No": [1]})
    sku = pd.DataFrame({"Material code": ["L1"], "Material Desc": ["d"],
                        "Catergory": ["LOOSE"], "HJ": [3], "SAP": [3]})
    inv = pd.DataFrame({"Item Number": ["L1", "L1"], "Actual Qty": [3, 3], "Cbm": [0.1, 0.1]})
    res = pe.run_pipeline(req, sku, inv, pe.EngineConfig())
    assert int(res["pick"]["HJ Pcs Qty"].iloc[0]) == 6


def test_carton_split_exception():
    """Target not divisible by carton size -> pick whole-carton portion, show variance."""
    req = pd.DataFrame({"Material": ["C1"], "Material Description": ["d"],
                        "Req Qty.": [23], "Delivery No": [1]})
    sku = pd.DataFrame({"Material code": ["C1"], "Material Desc": ["d"],
                        "Catergory": ["LOOSE"], "HJ": [1], "SAP": [1]})
    inv = pd.DataFrame({"Item Number": ["C1"] * 12, "Actual Qty": [2] * 12, "Cbm": [0.1] * 12})
    res = pe.run_pipeline(req, sku, inv, pe.EngineConfig())
    # 22 picked (11 cartons of 2), in main pick
    assert int(res["pick"]["HJ Pcs Qty"].iloc[0]) == 22
    # 1 variance recorded in Cannot Pick
    assert len(res["exceptions"]) == 1
    assert int(res["exceptions"]["Picked Pcs"].iloc[0]) == 22
    assert int(res["exceptions"]["Variance"].iloc[0]) == 1


def test_short_stock_exception():
    """Req > available stock -> Cannot Pick report (Inventory මදි)."""
    req = pd.DataFrame({"Material": ["S1"], "Material Description": ["d"],
                        "Req Qty.": [100], "Delivery No": [1]})
    sku = pd.DataFrame({"Material code": ["S1"], "Material Desc": ["d"],
                        "Catergory": ["LOOSE"], "HJ": [1], "SAP": [1]})
    inv = pd.DataFrame({"Item Number": ["S1", "S1"], "Actual Qty": [2, 2], "Cbm": [0.1, 0.1]})
    res = pe.run_pipeline(req, sku, inv, pe.EngineConfig())
    assert len(res["exceptions"]) == 1
    assert "SHORT_STOCK" in res["full_pick"]["_status"].tolist()


def test_sample_classification():
    res = _run()
    counts = res["full_pick"]["_status"].value_counts().to_dict()
    assert counts.get("OK", 0) == 152
    assert counts.get("CARTON_SPLIT", 0) == 17
    # 152 full + 14 partial-carton lines that pick at least one carton = 166
    assert len(res["pick"]) == 166
    assert len(res["exceptions"]) == 17
    # every exception carries a positive variance
    assert (res["exceptions"]["Variance"].astype(int) > 0).all()


def test_files_generate():
    res = _run()
    assert len(res["vip_pick_bytes"]) > 5000
    assert len(res["india_so_bytes"]) > 5000
    wb = load_workbook(io.BytesIO(res["vip_pick_bytes"]))
    assert "Summary" in wb.sheetnames
    assert "Cannot Pick" in wb.sheetnames
    assert "LOAD ID QR" in wb.sheetnames


def test_india_detail_equals_pcs():
    res = _run()
    detail = res["india_detail"]
    pick = res["pick"]
    qty_col = [c for c in detail.columns if c.upper() == "QTY"][0]
    detail_qty = pd.to_numeric(detail[qty_col], errors="coerce").fillna(0).astype(int).tolist()
    pick_pcs = pd.to_numeric(pick["HJ Pcs Qty"], errors="coerce").fillna(0).astype(int).tolist()
    assert detail_qty == pick_pcs


def test_load_id_qr_sheet():
    res = _run()
    wb = load_workbook(io.BytesIO(res["vip_pick_bytes"]))
    ws = wb["LOAD ID QR"]
    assert ws["A1"].value == "LOAD ID"
    assert ws["B1"].value == "QR Code"
    assert ws["D1"].value == "INM0VIP"
    assert ws["G1"].value == "PKINM0"
    assert ws["J1"].value == "IMSA05"


def test_load_id_dedup():
    """A second run against the first run's registry must suffix every LOAD ID."""
    res1 = _run()
    existing = set(str(x) for x in res1["load_ids"])
    res2 = _run(existing=existing)
    assert all("-" in str(x) for x in res2["load_ids"]), "duplicates must be suffixed"
    assert not (set(str(x) for x in res2["load_ids"]) & existing)


if __name__ == "__main__":
    test_formula_example()
    test_loose_equals_req()
    test_carton_split_exception()
    test_short_stock_exception()
    test_sample_classification()
    test_files_generate()
    test_india_detail_equals_pcs()
    test_load_id_qr_sheet()
    test_load_id_dedup()
    print("All validation checks passed.")
