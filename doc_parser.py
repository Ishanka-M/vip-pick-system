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

# Bumped whenever this module's public surface changes; app.py refuses to run
# against a stale copy instead of dying with a redacted TypeError.
API = 4

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
    line_amount: float | None = None       # Ext Price (pre-tax)
    unit_price: float | None = None
    line_total: float | None = None        # "Total" column — incl. tax, per line

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
    customer: str = ""            # Ship To / Consignee  ·  Name of Consignee(Shipped To)
    customer_code: str = ""
    contact_person: str = ""
    contact_email: str = ""
    source_file: str = ""
    lines: list[DocLine] = field(default_factory=list)
    declared_qty: float | None = None      # "Total Quantity"
    declared_amount: float | None = None   # "Sub Total" / "Grand Total"
    total_incl_tax: float | None = None    # "Total Amount (Incl. Tax)" — invoice grand total
    notes: list[str] = field(default_factory=list)

    # ---------------- completeness ---------------- #
    def completeness(self, tol: float = 0.75) -> tuple[bool, list[str]]:
        """(ok, problems) — problem එකක් තිබ්බොත් මේ doc එක pick කරන්නේ නෑ."""
        p: list[str] = []
        if not self.doc_number:
            p.append("Document number could not be read")
        if not self.lines:
            p.append("No line items found")

        for ln in self.lines:
            if not ln.item_code:
                p.append(f"Line {ln.line_no}: no item code")
            if not ln.qty or ln.qty <= 0:
                p.append(f"Line {ln.line_no}: qty missing or zero")

        nos = [ln.line_no for ln in self.lines]
        if nos and sorted(nos) != list(range(1, len(nos) + 1)):
            p.append(f"S.No sequence is broken ({sorted(nos)}) - a line may be missing)")

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


def _band_lines(rows: list[list[dict]], x0: float, x1: float, top_after: float,
                limit: int = 14) -> list[str]:
    """
    Text of one header column, read straight down.

    The header of a Donaldson invoice is five columns side by side, so the flat
    text mixes them ("438567 438549 Email: ..."). Reading by x band keeps
    Ship To / Consignee separate from Bill To and from Contact Person.
    """
    out: list[str] = []
    for row in rows:
        top = min(w["top"] for w in row)
        if top <= top_after + 0.5:
            continue
        words = [w for w in row if x0 - 0.5 <= (w["x0"] + w["x1"]) / 2 < x1]
        if words:
            out.append(" ".join(w["text"] for w in sorted(words, key=lambda w: w["x0"])))
        if len(out) >= limit:
            break
    return out


def _band_of(row: list[dict], label_words: tuple[str, ...], page_width: float,
             gap: float = 18.0) -> tuple[float, float] | None:
    """
    x range of the column a label sits in.

    The edge is the first *wide* gap on the row, not simply the next word —
    "Contact Person: Sharma, Rahul" is all one column, and cutting at the next
    word would have returned "Sharma," on its own.
    """
    for i, w in enumerate(row):
        if w["text"] in label_words:
            j, start = i, w["x0"]
            while j > 0 and row[j]["x0"] - row[j - 1]["x1"] < 12:   # rest of the label
                j -= 1
                start = row[j]["x0"]
            k = i
            while k + 1 < len(row) and row[k + 1]["x0"] - row[k]["x1"] < gap:
                k += 1
            return start, (row[k + 1]["x0"] if k + 1 < len(row) else page_width)
    return None


_LABELS = re.compile(r"^(Phone|Email|Sales Contact|Cnee Contact|Contact Person|Fax)\b",
                     re.I)


def _clean_name(lines: list[str]) -> tuple[str, str]:
    """(code, name) — the consignee block starts with an account code."""
    code = ""
    for ln in lines:
        t = ln.strip()
        if not t:
            continue
        if not code and re.fullmatch(r"\d{3,10}", t):
            code = t
            continue
        if _LABELS.match(t) or re.fullmatch(r"[\d,\s.:/-]+", t):
            continue
        return code, t
    return code, ""


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
    m = re.search(r"Total Amount\s*\(Incl\.?\s*Tax\)\s*:?\s*(?:INR)?\s*([\d,]+\.\d{2})",
                  full_text, re.I)
    if m:
        doc.total_incl_tax = _num(m.group(1))

    bounds: list[float] = []
    labels: list[str] = []
    cur: DocLine | None = None
    seen: set[tuple] = set()

    for page in pdf.pages:
        words = page.extract_words(keep_blank_chars=False, use_text_flow=False)
        rows = _cluster_rows(words)

        if not doc.customer:
            _read_invoice_header(doc, rows, float(page.width))

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
            i_unit = _col(labels, "Unit Price")
            i_total = _col(labels, "Total")

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
                    unit_price=_num(cell(i_unit)),
                    line_total=_num(cell(i_total)),
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


def _read_invoice_header(doc: ParsedDoc, rows: list[list[dict]], width: float) -> None:
    """Ship To / Consignee  +  Contact Person / Email, by column band."""
    anchor = None
    for row in rows:
        texts = [w["text"] for w in row]
        if "Consignee:" in texts and ("Contact" in texts or "Person:" in texts):
            anchor = row
            break
    if anchor is None:
        return
    top = min(w["top"] for w in anchor)

    band = _band_of(anchor, ("Consignee:",), width)
    if band:
        doc.customer_code, doc.customer = _clean_name(
            _band_lines(rows, band[0], band[1], top))

    cband = _band_of(anchor, ("Person:",), width)
    if not cband:
        return
    x0, x1 = cband
    same = [w for w in anchor if x0 - 0.5 <= (w["x0"] + w["x1"]) / 2 < x1]
    txt = " ".join(w["text"] for w in sorted(same, key=lambda w: w["x0"]))
    m = re.search(r"Person:\s*(.+)$", txt)
    if m:
        doc.contact_person = m.group(1).strip()

    # The email wraps mid-address ("rahul.sharma1@donaldso" / "n.com"), and the
    # second Email: in this column belongs to Sales Contact — stop before it.
    buf: list[str] = []
    started = False
    for ln in _band_lines(rows, x0, x1, top):
        t = ln.strip()
        if re.match(r"^Sales Contact\b|^Cnee\b", t, re.I):
            break
        if not started:
            m = re.match(r"^Email:\s*(.*)$", t, re.I)
            if m:
                started = True
                buf.append(m.group(1))
            continue
        if re.match(r"^(Phone|Email|Fax)\b", t, re.I):
            break
        buf.append(t)
        if re.search(r"@[\w.-]+\.\w{2,}", "".join(buf)):
            break
    joined = "".join(x.replace(" ", "") for x in buf)
    m = re.search(r"[\w.+-]+@[\w.-]+\.\w{2,}", joined)
    if m:
        doc.contact_email = m.group(0)


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
    m = re.search(r"Total Amount\s*\(Incl\.?\s*Tax\)\s*:?\s*(?:INR)?\s*([\d,]+\.\d{2})",
                  full_text, re.I)
    doc.total_incl_tax = _num(m.group(1)) if m else doc.declared_amount

    seen: set[tuple] = set()
    for page in pdf.pages:
        if not doc.customer:
            rows = _cluster_rows(page.extract_words(keep_blank_chars=False))
            for row in rows:
                joined = " ".join(w["text"] for w in row)
                if "Consignee(Shipped" in joined:
                    band = _band_of(row, ("Consignee(Shipped", "To)"), float(page.width))
                    if band:
                        top = min(w["top"] for w in row)
                        doc.customer_code, doc.customer = _clean_name(
                            _band_lines(rows, band[0], band[1], top))
                    break

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
                        line_total=amt,
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
            doc.notes.append("Document type not detected - used the fallback parser")
    if not doc.lines:
        doc.notes.append("Could not parse line items - add them in the review table")
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
