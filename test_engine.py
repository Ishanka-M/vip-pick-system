"""
Validation harness for the VIP pick engine.

Runs the full pipeline against the bundled sample_data and asserts that
the computed HJ Pcs Qty matches the reference VIP_PICK.xlsx for at least
166 of 169 lines.

The 3 known mismatches are Qty == 1 partial-carton manual loose picks
that no deterministic carton-rounding rule reproduces.

Run:  python -m pytest test_engine.py   (or)   python test_engine.py
"""
import os
import pandas as pd
import pick_engine as pe

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(HERE, "sample_data")

REQ = os.path.join(SAMPLE, "1781617499524_Requament.xlsx")
SKU = os.path.join(SAMPLE, "1781617499524_SKU_MASTER.xlsx")
INV = os.path.join(SAMPLE, "1781617499523_Inventory_Report.xlsx")
REF_VIP = os.path.join(SAMPLE, "1781617499522_VIP_PICK.xlsx")


def _run():
    req = pd.read_excel(REQ)
    sku = pd.read_excel(SKU)
    inv = pd.read_excel(INV)
    return pe.run_pipeline(req, sku, inv, pe.EngineConfig())


def test_row_counts():
    res = _run()
    assert len(res["pick"]) == 169, "expected 169 requirement lines"
    assert len(res["india_master"]) == 25, "expected 25 unique deliveries"
    assert len(res["india_detail"]) == 169, "detail must match requirement lines"


def test_files_generate():
    res = _run()
    assert len(res["vip_pick_bytes"]) > 5000
    assert len(res["india_so_bytes"]) > 5000


def test_pick_accuracy():
    res = _run()
    pick = res["pick"]
    ref = pd.read_excel(REF_VIP, sheet_name="Summary", header=3)
    ref.columns = [str(c).strip() for c in ref.columns]
    ref = ref.dropna(subset=["Material"]).reset_index(drop=True)

    n = min(len(pick), len(ref))
    match = 0
    for i in range(n):
        try:
            if int(pick["HJ Pcs Qty"].iloc[i]) == int(ref["HJ Pcs Qty"].iloc[i]):
                match += 1
        except (ValueError, TypeError):
            pass
    print(f"HJ Pcs Qty match: {match}/{n}")
    assert match >= 166, f"accuracy regressed: {match}/{n} (expected >= 166)"


def test_india_detail_equals_pcs():
    """INDIA SO Detail QTY must equal HJ Pcs Qty for every line."""
    res = _run()
    detail = res["india_detail"]
    pick = res["pick"]
    qty_col = "QTY" if "QTY" in detail.columns else [c for c in detail.columns if c.upper() == "QTY"][0]
    detail_qty = pd.to_numeric(detail[qty_col], errors="coerce").fillna(0).astype(int).tolist()
    pick_pcs = pd.to_numeric(pick["HJ Pcs Qty"], errors="coerce").fillna(0).astype(int).tolist()
    assert detail_qty == pick_pcs, "INDIA detail QTY drifted from HJ Pcs Qty"


def test_load_id_qr_sheet():
    """VIP PICK workbook must contain a 'LOAD ID QR' sheet listing every
    unique delivery (25 in the sample)."""
    import io
    from openpyxl import load_workbook
    res = _run()
    assert len(res["load_ids"]) == 25, "expected 25 unique LOAD IDs"
    wb = load_workbook(io.BytesIO(res["vip_pick_bytes"]))
    assert "LOAD ID QR" in wb.sheetnames, "LOAD ID QR sheet missing"
    ws = wb["LOAD ID QR"]
    assert ws["A1"].value == "LOAD ID"
    assert ws["B1"].value == "QR Code"
    # fixed header QR labels
    assert ws["D1"].value == "INM0VIP"
    assert ws["G1"].value == "PKINM0"
    assert ws["J1"].value == "IMSA05"
    # QR images embedded: 25 load + 3 header = 28
    assert len(ws._images) == 28, "expected 25 load + 3 header QR images"


if __name__ == "__main__":
    test_row_counts()
    test_files_generate()
    test_pick_accuracy()
    test_india_detail_equals_pcs()
    test_load_id_qr_sheet()
    print("All validation checks passed.")
