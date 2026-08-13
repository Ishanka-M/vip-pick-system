"""
doc_parser.py — Donaldson Invoice / Delivery Challan PDF -> structured lines
===========================================================================
Supports
  * TAX INVOICE (Domestic)  -> coordinate (column-bucket) parser
  * Delivery Challan        -> ruled-table parser (multi-copy safe)

Every parsed document carries its own completeness check, because the rule is:
"Invoice/DC complete නැත්නම් pick කරන්න එපා."
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Any

import pandas as pd
import pdfplumber

# --------------------------------------------------------------------------- #
# Item-code helpers
# --------------------------------------------------------------------------- #
_ITEM_CLEAN = re.compile(r"[^A-Z0-9\-/]")


def clean_item(code: Any) -> str:
    """'  p550576-016-140. ' -> 'P550576-016-140'"""
    s = str(code or "").strip().upper()
    s = s.replace("\n", "").replace(" ", "")
    s = s.rstrip(".")
    s = _ITEM_CLEAN.sub("", s)
    return s


def base_item(code: Any) -> str:
    """
    Matching key — Donaldson suffix එක අයින් කරලා base ID එක විතරක්.
        P162400-000-140  -> P162400
        P550576-016-140. -> P550576
        07011636-000-440 -> 07011636
        100409-101       -> 100409
        P550945          -> P550945
    Suffix එක කියලා ගන්නේ **3-digit කෑලි විතරක්** නම් — එහෙම නැත්නම් මුළු code එකම.
        05-47174 -> 05-47174   (05 නෙවෙයි — වැරදි match වළක්වන්න)
    """
    s = clean_item(code)
    if not s:
        return ""
    parts = s.split("-")
    if len(parts) > 1 and all(re.fullmatch(r"\d{3}", p) for p in parts[1:]):
        return parts[0]
    return s


def _num(x: Any) -> float | None:
    if x is None:
        return None
    s = str(x).replace(",", "").replace("\n", " ").strip()
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


# --------------------------------------------------------------------------- #
# Data classes
# --------------------------------------------------------------------------- #
@dataclass
class DocLine:
    line_no: int
    item_code: str
    description: str = ""
    qty: float = 0.0
    uom: str = ""
    customer_po: str = ""
    sales_order: str = ""
    so_line: str = ""
    line_amount: float | None = None

    @property
    def base(self) -> str:
        return base_item(self.item_code)


@dataclass
class ParsedDoc:
    doc_type: str = "INVOICE"          # INVOICE | DELIVERY CHALLAN
    doc_number: str = ""
    doc_date: str = ""
    ref_number: str = ""               # AR Invoice No / Order No
    delivery_number: str = ""
    customer_po: str = ""
    source_file: str = ""
    lines: list[DocLine] = field(default_factory=list)
    declared_qty: float | None = None      # "Total Quantity"
    declared_amount: float | None = None   # "Sub Total" / "Grand Total"
    notes: list[str] = field(default_factory=list)

    # ---------------- completeness ---------------- #
    def completeness(self, tol: float = 0.75) -> tuple[bool, list[str]]:
        """(ok, problems) — problem එකක් තිබ්බොත් මේ doc එක pick කරන්නේ නෑ."""
        p: list[str] = []
        if not self.doc_number:
            p.append("Document number කියවගන්න බැරි වුණා")
        if not self.lines:
            p.append("Line items හම්බුණේ නෑ")

        for ln in self.lines:
            if not ln.item_code:
                p.append(f"Line {ln.line_no}: Item code නෑ")
            if not ln.qty or ln.qty <= 0:
                p.append(f"Line {ln.line_no}: Qty නෑ / 0")

        nos = [ln.line_no for ln in self.lines]
        if nos and sorted(nos) != list(range(1, len(nos) + 1)):
            p.append(f"S.No sequence කැඩිලා ({sorted(nos)}) — line එකක් missing වෙන්න පුළුවන්")

        if self.declared_qty is not None:
            tot = sum(ln.qty for ln in self.lines)
            if abs(tot - self.declared_qty) > 0.01:
                p.append(f"Qty total mismatch: lines={tot:g} vs document={self.declared_qty:g}")

        if self.declared_amount is not None:
            amts = [ln.line_amount for ln in self.lines if ln.line_amount is not None]
            if len(amts) == len(self.lines) and self.lines:
                tot = sum(amts)
                if abs(tot - self.declared_amount) > tol:
                    p.append(
                        f"Amount total mismatch: lines={tot:,.2f} vs document={self.declared_amount:,.2f}"
                    )
        return (len(p) == 0), p


# --------------------------------------------------------------------------- #
# Low level word helpers
# --------------------------------------------------------------------------- #
def _cluster_rows(words: list[dict], tol: float = 2.6) -> list[list[dict]]:
    out: list[list[dict]] = []
    cur: list[dict] = []
    top = None
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if top is None or abs(w["top"] - top) <= tol:
            cur.append(w)
            top = w["top"] if top is None else top
        else:
            out.append(sorted(cur, key=lambda x: x["x0"]))
            cur, top = [w], w["top"]
    if cur:
        out.append(sorted(cur, key=lambda x: x["x0"]))
    return out


def _header_groups(row: list[dict], gap: float = 8.0) -> list[dict]:
    """Header row එකේ words -> column groups (label, x0, x1)."""
    groups: list[dict] = []
    for w in row:
        if groups and (w["x0"] - groups[-1]["x1"]) <= gap:
            groups[-1]["x1"] = max(groups[-1]["x1"], w["x1"])
            groups[-1]["label"] += " " + w["text"]
        else:
            groups.append({"label": w["text"], "x0": w["x0"], "x1": w["x1"]})
    return groups


def _bounds(groups: list[dict]) -> list[float]:
    return [(groups[i]["x1"] + groups[i + 1]["x0"]) / 2 for i in range(len(groups) - 1)]


def _bucket(row: list[dict], bounds: list[float]) -> list[str]:
    cells = [""] * (len(bounds) + 1)
    for w in row:
        c = (w["x0"] + w["x1"]) / 2
        idx = 0
        while idx < len(bounds) and c > bounds[idx]:
            idx += 1
        cells[idx] = (cells[idx] + " " + w["text"]).strip()
    return cells


def _col(labels: list[str], *keys: str) -> int | None:
    """label list එකෙන් keyword ගැලපෙන column index එක."""
    norm = [re.sub(r"[^a-z]", "", l.lower()) for l in labels]
    for k in keys:
        kk = re.sub(r"[^a-z]", "", k.lower())
        for i, n in enumerate(norm):
            if n == kk:
                return i
    for k in keys:
        kk = re.sub(r"[^a-z]", "", k.lower())
        for i, n in enumerate(norm):
            if kk and kk in n:
                return i
    return None


# --------------------------------------------------------------------------- #
# INVOICE parser (coordinate buckets)
# --------------------------------------------------------------------------- #
def _parse_invoice(pdf: pdfplumber.PDF, filename: str) -> ParsedDoc:
    doc = ParsedDoc(doc_type="INVOICE", source_file=filename)

    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    txt_lines = [l.strip() for l in full_text.split("\n")]

    for l in txt_lines:
        if not doc.doc_number:
            m = re.match(r"^Invoice No\.?\s*:?\s*([A-Za-z0-9\-/]+)", l)
            if m:
                doc.doc_number = m.group(1).strip()
        if not doc.ref_number:
            m = re.match(r"^AR Invoice No\.?\s*:?\s*([A-Za-z0-9\-/]+)", l)
            if m:
                doc.ref_number = m.group(1).strip()
        if not doc.doc_date:
            m = re.match(r"^Invoice Date\s*:?\s*([0-9A-Za-z\-/]+)", l)
            if m:
                doc.doc_date = m.group(1).strip()
        if not doc.delivery_number:
            m = re.search(r"Delivery Number:\s*([A-Za-z0-9\-/]+)", l)
            if m:
                doc.delivery_number = m.group(1).strip()

    m = re.search(r"Total Quantity\s+([\d,]+(?:\.\d+)?)", full_text)
    if m:
        doc.declared_qty = _num(m.group(1))
    m = re.search(r"Sub Total:\s*(?:INR)?\s*([\d,]+\.\d{2})", full_text)
    if m:
        doc.declared_amount = _num(m.group(1))

    bounds: list[float] = []
    labels: list[str] = []
    cur: DocLine | None = None
    seen: set[tuple] = set()

    for page in pdf.pages:
        words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        rows = _cluster_rows(words)

        started = False
        for row in rows:
            joined = " ".join(w["text"] for w in row)

            # ---- header row ----
            if ("S.No" in joined and "Qty" in joined and "UOM" in joined):
                g = _header_groups(row)
                labels = [x["label"] for x in g]
                bounds = _bounds(g)
                started = True
                continue

            if not bounds:
                continue
            if not started and not doc.lines:
                continue
            if re.search(r"Sub Total|Remittance Address|Total Quantity|Total Amount", joined):
                started = False
                continue

            cells = _bucket(row, bounds)
            i_sno = _col(labels, "S.No.", "SNo") or 0
            i_so = _col(labels, "Sales Order")
            i_sol = _col(labels, "SO Line No")
            i_po = _col(labels, "Customer PO")
            i_item = _col(labels, "Donaldson Item", "Item Code")
            i_desc = _col(labels, "Item Description", "Description")
            i_qty = _col(labels, "Qty", "Quantity")
            i_uom = _col(labels, "UOM")
            i_ext = _col(labels, "Ext Price")

            def cell(idx: int | None) -> str:
                return cells[idx].strip() if (idx is not None and idx < len(cells)) else ""

            sno = cell(i_sno)
            qty = _num(cell(i_qty))
            item = clean_item(cell(i_item))

            if re.fullmatch(r"\d{1,3}", sno) and qty is not None and item:
                key = (sno, item, qty)
                if key in seen:            # duplicate copy of the same page
                    cur = None
                    continue
                seen.add(key)
                cur = DocLine(
                    line_no=int(sno),
                    item_code=item,
                    description=cell(i_desc),
                    qty=float(qty),
                    uom=cell(i_uom) or "EA",
                    customer_po=cell(i_po),
                    sales_order=cell(i_so),
                    so_line=cell(i_sol),
                    line_amount=_num(cell(i_ext)),
                )
                doc.lines.append(cur)
                continue

            # ---- wrap / continuation row ----
            if cur is not None and not sno:
                extra_desc = cell(i_desc)
                extra_po = cell(i_po)
                if extra_desc and not re.fullmatch(r"[\d.,%@() ]+", extra_desc):
                    cur.description = (cur.description + " " + extra_desc).strip()
                if extra_po and not re.fullmatch(r"\d{6,}", extra_po):
                    cur.customer_po = (cur.customer_po + " " + extra_po).strip()

    if doc.lines:
        doc.customer_po = doc.lines[0].customer_po
    doc.lines.sort(key=lambda l: l.line_no)
    return doc


# --------------------------------------------------------------------------- #
# DELIVERY CHALLAN parser (ruled table)
# --------------------------------------------------------------------------- #
_TS_LINES = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
_TS_TEXT = {"vertical_strategy": "text", "horizontal_strategy": "text"}


def _flat(cell: Any, sep: str = "") -> str:
    return sep.join(str(cell or "").split("\n")).strip()


def _parse_challan(pdf: pdfplumber.PDF, filename: str) -> ParsedDoc:
    doc = ParsedDoc(doc_type="DELIVERY CHALLAN", source_file=filename)

    full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    m = re.search(r"Delivery Challan no\.?\s*([A-Za-z0-9\-/]+)", full_text)
    if m:
        doc.doc_number = m.group(1).strip()
    m = re.search(r"Delivery Challan Date\.?\s*([0-9A-Za-z\-/]+)", full_text)
    if m:
        doc.doc_date = m.group(1).strip()
    m = re.search(r"Order No\.?\s*([A-Za-z0-9\-/]+)", full_text)
    if m:
        doc.ref_number = m.group(1).strip()
    m = re.search(r"Grand Total[^0-9]*([\d,]+\.\d{2})", full_text)
    if m:
        doc.declared_amount = _num(m.group(1))

    seen: set[tuple] = set()
    for page in pdf.pages:
        tables = page.extract_tables(_TS_LINES) or page.extract_tables(_TS_TEXT) or []
        for tbl in tables:
            if not tbl or len(tbl) < 2:
                continue
            head = [_flat(c, " ") for c in tbl[0]]
            i_sno = _col(head, "S.No", "SNo")
            i_item = _col(head, "Item Code")
            i_desc = _col(head, "Description of Goods", "Description")
            i_qty = _col(head, "Quantity", "Qty")
            i_uom = _col(head, "UOM")
            i_tot = _col(head, "Total Amount (incl. tax)", "Taxable Amount", "Total Amount")
            if i_item is None or i_qty is None:
                continue

            for raw in tbl[1:]:
                def cell(idx: int | None, sep: str = "") -> str:
                    if idx is None or idx >= len(raw):
                        return ""
                    return _flat(raw[idx], sep)

                sno = cell(i_sno)
                item = clean_item(cell(i_item))
                qty = _num(cell(i_qty))
                if not (re.fullmatch(r"\d{1,3}", sno) and item and qty):
                    continue
                key = (sno, item, qty)
                if key in seen:                    # Original / Duplicate / Triplicate copies
                    continue
                seen.add(key)

                amt = None
                if i_tot is not None:
                    amt = _num(cell(i_tot))
                    if amt is None:
                        for j in range(len(raw) - 1, -1, -1):
                            amt = _num(_flat(raw[j]))
                            if amt:
                                break
                doc.lines.append(
                    DocLine(
                        line_no=int(sno),
                        item_code=item,
                        description=cell(i_desc, " "),
                        qty=float(qty),
                        uom=cell(i_uom) or "EA",
                        line_amount=amt,
                    )
                )

    doc.lines.sort(key=lambda l: l.line_no)
    return doc


# --------------------------------------------------------------------------- #
# Public API
# --------------------------------------------------------------------------- #
def parse_pdf(data: bytes, filename: str = "") -> ParsedDoc:
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        head = (pdf.pages[0].extract_text() or "")[:3000].upper()
        if "DELIVERY CHALLAN" in head:
            doc = _parse_challan(pdf, filename)
        elif "TAX INVOICE" in head or "INVOICE NO" in head:
            doc = _parse_invoice(pdf, filename)
        else:
            doc = _parse_challan(pdf, filename)
            if not doc.lines:
                doc = _parse_invoice(pdf, filename)
            doc.notes.append("Document type auto-detect වුණේ නෑ — fallback parser use කළා")
    if not doc.lines:
        doc.notes.append("Line items parse කරගන්න බැරි වුණා — review table එකේ manually add කරන්න")
    return doc


def docs_to_frame(docs: list[ParsedDoc]) -> pd.DataFrame:
    """Review / edit table (st.data_editor)."""
    rows = []
    for d in docs:
        for ln in d.lines:
            rows.append(
                {
                    "Use": True,
                    "Doc Type": d.doc_type,
                    "Doc Number": d.doc_number,
                    "Doc Date": d.doc_date,
                    "Line": ln.line_no,
                    "Item Code": ln.item_code,
                    "Base ID": ln.base,
                    "Description": ln.description,
                    "Qty": ln.qty,
                    "Doc UOM": ln.uom,
                    "Customer PO": ln.customer_po,
                    "Sales Order": ln.sales_order,
                    "SO Line": ln.so_line,
                    "Source File": d.source_file,
                }
            )
    cols = ["Use", "Doc Type", "Doc Number", "Doc Date", "Line", "Item Code", "Base ID",
            "Description", "Qty", "Doc UOM", "Customer PO", "Sales Order", "SO Line",
            "Source File"]
    return pd.DataFrame(rows, columns=cols)


def frame_to_docs(df: pd.DataFrame, originals: list[ParsedDoc] | None = None) -> list[ParsedDoc]:
    """Edit කරපු table එකෙන් ආපහු ParsedDoc list එකක්."""
    meta = {d.doc_number: d for d in (originals or [])}
    out: dict[str, ParsedDoc] = {}
    df = df.copy()
    if "Use" in df.columns:
        df = df[df["Use"].fillna(False).astype(bool)]

    for _, r in df.iterrows():
        num = str(r.get("Doc Number", "")).strip()
        if not num:
            continue
        if num not in out:
            src = meta.get(num)
            out[num] = ParsedDoc(
                doc_type=str(r.get("Doc Type", "INVOICE")).strip() or "INVOICE",
                doc_number=num,
                doc_date=str(r.get("Doc Date", "") or ""),
                ref_number=src.ref_number if src else "",
                delivery_number=src.delivery_number if src else "",
                source_file=str(r.get("Source File", "") or ""),
                declared_qty=src.declared_qty if src else None,
                declared_amount=None,       # manual edit -> amount check skip
            )
        item = clean_item(r.get("Item Code"))
        qty = _num(r.get("Qty")) or 0.0
        if not item:
            continue
        out[num].lines.append(
            DocLine(
                line_no=int(_num(r.get("Line")) or (len(out[num].lines) + 1)),
                item_code=item,
                description=str(r.get("Description", "") or ""),
                qty=float(qty),
                uom=str(r.get("Doc UOM", "") or "EA"),
                customer_po=str(r.get("Customer PO", "") or ""),
                sales_order=str(r.get("Sales Order", "") or ""),
                so_line=str(r.get("SO Line", "") or ""),
            )
        )

    docs = list(out.values())
    for d in docs:
        d.lines.sort(key=lambda l: l.line_no)
        for i, ln in enumerate(d.lines, start=1):
            ln.line_no = i
    return docs
