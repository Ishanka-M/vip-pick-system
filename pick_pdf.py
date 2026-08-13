"""
pick_pdf.py — Pick Sheet PDF (LOAD_ID QR) + original document merge + Email
==========================================================================
  build_doc_pdf()  : Pick sheet (QR) + upload කරපු Invoice / DC PDF එකම එකට
  build_eml()      : Default mail app එකෙන් open වෙන .eml (attachment එක්ක)
  mailto_link()    : Quick mailto: link (attachment නෑ)
"""
from __future__ import annotations

import io
import re
from datetime import datetime
from email.message import EmailMessage
from typing import Any
from urllib.parse import quote

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt          # noqa: E402
import pandas as pd
import qrcode
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (Image, KeepTogether, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)

# --------------------------------------------------------------------------- #
# palette
# --------------------------------------------------------------------------- #
INK = colors.HexColor("#0F1F33")
LINE = colors.HexColor("#B9C6D6")
SOFT = colors.HexColor("#EEF3F9")
ACC = colors.HexColor("#FF365B")
OKC = colors.HexColor("#0E8F5E")

_SS = getSampleStyleSheet()
P_TITLE = ParagraphStyle("t", parent=_SS["Title"], fontSize=15, leading=18,
                         textColor=INK, alignment=0, spaceAfter=0)
P_SUB = ParagraphStyle("s", parent=_SS["Normal"], fontSize=7.6, leading=10,
                       textColor=colors.HexColor("#5B6C82"))
P_H = ParagraphStyle("h", parent=_SS["Normal"], fontSize=9.5, leading=12,
                     textColor=INK, spaceBefore=4, spaceAfter=3, fontName="Helvetica-Bold")
P_CELL = ParagraphStyle("c", parent=_SS["Normal"], fontSize=7.2, leading=8.8,
                        textColor=INK)
P_CELLH = ParagraphStyle("ch", parent=_SS["Normal"], fontSize=7.2, leading=8.8,
                         textColor=colors.white, fontName="Helvetica-Bold")
P_LOAD = ParagraphStyle("l", parent=_SS["Normal"], fontSize=12.5, leading=14,
                        textColor=INK, alignment=TA_CENTER, fontName="Helvetica-Bold")
P_LOADK = ParagraphStyle("lk", parent=_SS["Normal"], fontSize=6.5, leading=8,
                         textColor=colors.HexColor("#5B6C82"), alignment=TA_CENTER)


_EMOJI = {"✅": "", "❌": "*", "⚠️": "!", "⚠": "!", "✔": "", "✓": ""}


def _txt(v: Any) -> str:
    """reportlab-safe — emoji අයින් කරලා."""
    s = str("" if v is None else v)
    for k, r in _EMOJI.items():
        s = s.replace(k, r)
    return "".join(ch if ord(ch) < 0x2500 else " " for ch in s).strip()


def safe_name(text: Any, fallback: str = "DOC") -> str:
    """'333/26-27/62' -> '333-26-27-62'  (filename safe)"""
    s = re.sub(r"[\\/:*?\"<>|\s]+", "-", str(text or "").strip())
    s = re.sub(r"-{2,}", "-", s).strip("-.")
    return s or fallback


# --------------------------------------------------------------------------- #
# QR
# --------------------------------------------------------------------------- #
def qr_png(text: str, box: int = 10, border: int = 2) -> bytes:
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M,
                       box_size=box, border=border)
    qr.add_data(str(text))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# --------------------------------------------------------------------------- #
# Pick sheet
# --------------------------------------------------------------------------- #
def _kv_table(pairs: list[tuple[str, str]], cols: int, width: float) -> Table:
    rows: list[list] = []
    cells: list = []
    for k, v in pairs:
        cells.append(Paragraph(f"<font color='#5B6C82' size=6.4>{_txt(k)}</font><br/>"
                               f"<b>{_txt(v) or '-'}</b>", P_CELL))
        if len(cells) == cols:
            rows.append(cells)
            cells = []
    if cells:
        cells += [""] * (cols - len(cells))
        rows.append(cells)
    t = Table(rows, colWidths=[width / cols] * cols)
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
        ("BACKGROUND", (0, 0), (-1, -1), SOFT),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t


def _grid(header: list[str], rows: list[list[str]], widths: list[float],
          aligns: dict[int, str] | None = None) -> Table:
    data = [[Paragraph(_txt(h), P_CELLH) for h in header]]
    for r in rows:
        data.append([Paragraph(_txt(c), P_CELL) for c in r])
    t = Table(data, colWidths=widths, repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), INK),
        ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.35, LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, SOFT]),
    ]
    for i, a in (aligns or {}).items():
        style.append(("ALIGN", (i, 0), (i, -1), a))
    t.setStyle(TableStyle(style))
    return t


def build_pick_sheet(info: dict[str, Any], alloc: pd.DataFrame,
                     verify: pd.DataFrame | None = None) -> bytes:
    """Pick sheet — LOAD_ID QR + pallet pick details + qty verification."""
    buf = io.BytesIO()
    page = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=page,
                            leftMargin=11 * mm, rightMargin=11 * mm,
                            topMargin=9 * mm, bottomMargin=10 * mm,
                            title=f"Pick Sheet {info.get('LOAD_ID','')}",
                            author="EFL · Donaldson OutBound")
    W = page[0] - 22 * mm
    load_id = str(info.get("LOAD_ID", ""))
    story: list = []

    # ---------------- header ----------------
    qr_img = Image(io.BytesIO(qr_png(load_id, box=8, border=1)), width=30 * mm, height=30 * mm)
    left = [
        [Paragraph("OUTBOUND PICK SHEET", P_TITLE)],
        [Paragraph(f"{info.get('CLIENT','INM0DONA')} &nbsp;·&nbsp; "
                   f"WH {info.get('WH_ID','')} &nbsp;·&nbsp; "
                   f"{info.get('DOC_TYPE','')} &nbsp;<b>{info.get('DOC_NUMBER','')}</b>", P_SUB)],
        [Paragraph(f"Printed {datetime.now():%d-%b-%Y %H:%M} &nbsp;·&nbsp; "
                   f"RUN {info.get('RUN_ID','')}", P_SUB)],
    ]
    lt = Table(left, colWidths=[W - 42 * mm])
    lt.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                            ("TOPPADDING", (0, 0), (-1, -1), 0)]))
    qr_block = Table([[qr_img], [Paragraph(load_id, P_LOAD)], [Paragraph("LOAD ID", P_LOADK)]],
                     colWidths=[40 * mm])
    qr_block.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("BOX", (0, 0), (-1, -1), 0.8, INK),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    head = Table([[lt, qr_block]], colWidths=[W - 42 * mm, 42 * mm])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (0, 0), "TOP"),
                              ("VALIGN", (1, 0), (1, 0), "TOP"),
                              ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story += [head, Spacer(1, 5)]

    # ---------------- doc info ----------------
    story.append(_kv_table([
        ("DOCUMENT", f"{info.get('DOC_TYPE','')} {info.get('DOC_NUMBER','')}"),
        ("DOC DATE", info.get("DOC_DATE", "")),
        ("REF / AR NO", info.get("REF_NUMBER", "")),
        ("PLANT", info.get("PLANT", "")),
        ("PICK DATE", info.get("PICK_DATE", "")),
        ("STRATEGY", info.get("STRATEGY", "")),
        ("DOC LINES", _n(info.get("LINES", ""))),
        ("TOTAL QTY", _n(info.get("TOTAL_QTY", ""))),
        ("PALLETS", _n(info.get("PALLETS", ""))),
        ("QTY CHECK", info.get("VERIFY", "")),
    ], cols=5, width=W))
    story.append(Spacer(1, 7))

    # ---------------- pick details ----------------
    story.append(Paragraph("PICK DETAILS — pallet level", P_H))
    rows: list[list[str]] = []
    if alloc is not None and len(alloc):
        for _, r in alloc.iterrows():
            rows.append([
                str(r.get("DOC_LINE", "")), str(r.get("ITEM_NUMBER", "")),
                str(r.get("DESCRIPTION", ""))[:48], str(r.get("LOT_NUMBER", "")),
                str(r.get("PALLET", "")), str(r.get("LOCATION_ID", "")),
                str(r.get("UOM", "")), _n(r.get("QTY_BEFORE")), _n(r.get("QTY_PICKED")),
                _n(r.get("QTY_BALANCE")), "",
            ])
    widths = [11, 30, 56, 16, 40, 32, 12, 18, 18, 20, 22]
    scale = W / sum(widths)
    story.append(_grid(
        ["Ln", "Item Number", "Description", "Lot", "Pallet", "Location", "UOM",
         "Stock", "PICK QTY", "Balance", "Picked  [  ]"],
        rows, [w * scale for w in widths],
        aligns={0: "CENTER", 6: "CENTER", 7: "RIGHT", 8: "RIGHT", 9: "RIGHT", 10: "CENTER"},
    ))
    story.append(Spacer(1, 7))

    # ---------------- verification ----------------
    if verify is not None and len(verify):
        vr = [[str(r.get("LINE", "")), str(r.get("ITEM_CODE", "")),
               str(r.get("ITEM_NUMBER", "")), _n(r.get("DOC_QTY")),
               _n(r.get("PICKED_QTY")), _n(r.get("DIFF")), str(r.get("STATUS", ""))]
              for _, r in verify.iterrows()]
        vw = [12, 34, 34, 22, 22, 18, 22]
        s2 = (W * 0.62) / sum(vw)
        vtab = _grid(["Ln", "Doc Item Code", "Picked Item Number", "Doc Qty",
                      "Picked Qty", "Diff", "Status"], vr, [w * s2 for w in vw],
                     aligns={0: "CENTER", 3: "RIGHT", 4: "RIGHT", 5: "RIGHT", 6: "CENTER"})
        story.append(KeepTogether([Paragraph("QUANTITY VERIFICATION — Doc Qty vs Picked Qty",
                                             P_H), vtab]))
        story.append(Spacer(1, 8))

    # ---------------- sign off ----------------
    sign = Table([[Paragraph("<b>Picked by</b><br/><br/>______________________<br/>"
                             "<font size=6>Name / Date / Time</font>", P_CELL),
                   Paragraph("<b>Checked by</b><br/><br/>______________________<br/>"
                             "<font size=6>Name / Date / Time</font>", P_CELL),
                   Paragraph("<b>Loaded by</b><br/><br/>______________________<br/>"
                             "<font size=6>Vehicle No / Seal No</font>", P_CELL),
                   Paragraph("<b>Remarks</b><br/><br/>______________________________", P_CELL)]],
                 colWidths=[W / 4] * 4)
    sign.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.6, LINE),
                              ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
                              ("TOPPADDING", (0, 0), (-1, -1), 7),
                              ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                              ("LEFTPADDING", (0, 0), (-1, -1), 6)]))
    story.append(KeepTogether(sign))

    def _footer(canv, _doc):
        canv.saveState()
        canv.setFont("Helvetica", 6.6)
        canv.setFillColor(colors.HexColor("#7C90AB"))
        canv.drawString(11 * mm, 5.5 * mm,
                        f"LOAD ID {load_id} · {info.get('DOC_TYPE','')} "
                        f"{info.get('DOC_NUMBER','')} · generated by EFL OutBound Pick Generator")
        canv.drawRightString(page[0] - 11 * mm, 5.5 * mm, f"Page {canv.getPageNumber()}")
        canv.setStrokeColor(ACC)
        canv.setLineWidth(1.2)
        canv.line(11 * mm, page[1] - 6 * mm, page[0] - 11 * mm, page[1] - 6 * mm)
        canv.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


def _n(v: Any) -> str:
    try:
        f = float(v)
    except (TypeError, ValueError):
        return "" if v is None else str(v)
    return str(int(round(f))) if abs(f - round(f)) < 1e-9 else f"{f:g}"


# --------------------------------------------------------------------------- #
# merge with the uploaded document
# --------------------------------------------------------------------------- #
def merge_pdfs(parts: list[bytes]) -> bytes:
    w = PdfWriter()
    for p in parts:
        if not p:
            continue
        try:
            for page in PdfReader(io.BytesIO(p)).pages:
                w.add_page(page)
        except Exception:
            continue
    out = io.BytesIO()
    w.write(out)
    return out.getvalue()


def build_doc_pdf(info: dict[str, Any], alloc: pd.DataFrame,
                  verify: pd.DataFrame | None = None,
                  source_pdf: bytes | None = None,
                  attach_source: bool = True) -> bytes:
    """Pick sheet + (upload කරපු Invoice / DC PDF එකම එකට)."""
    sheet = build_pick_sheet(info, alloc, verify)
    if attach_source and source_pdf:
        return merge_pdfs([sheet, source_pdf])
    return sheet


# --------------------------------------------------------------------------- #
# Charts — email / PDF එකට item details chart එකක්
# --------------------------------------------------------------------------- #
CHART_INK = "#0F1F33"
CHART_ACC = "#FF365B"
CHART_OK = "#0E8F5E"
CHART_WARN = "#F2B33D"


# on-screen preview needs a dark canvas; email + PDF always stay on white paper
_DARK = {"page": "#141C26", "ink": "#E6EDF5", "soft": "#93A3B5", "grid": "#22303F",
         "axis": "#3A4B5D", "bar": "#7FA8CF"}


def _fig_png(fig, dpi: int = 130, face: str = "white") -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=face, edgecolor="none")
    plt.close(fig)
    return buf.getvalue()


def _palette(dark: bool) -> dict:
    if not dark:
        return {"page": "white", "ink": CHART_INK, "soft": "#5B6C82",
                "grid": "#EEF3F9", "axis": "#D6DEE8", "bar": CHART_INK}
    return _DARK


def _short_label(v: Any, n: int = 22) -> str:
    s = str(v or "")
    return s if len(s) <= n else s[: n - 1] + "…"


def pick_chart_png(alloc: pd.DataFrame, title: str = "Picked qty by item",
                   top: int = 12, dark: bool = False) -> bytes | None:
    """Item එකකට කීයද pick කරේ — horizontal bar chart (email එකට)."""
    if alloc is None or not len(alloc):
        return None
    d = alloc.copy()
    d["QTY_PICKED"] = pd.to_numeric(d["QTY_PICKED"], errors="coerce").fillna(0)
    g = (d.groupby("ITEM_NUMBER", dropna=False)
           .agg(qty=("QTY_PICKED", "sum"), pallets=("PALLET", "nunique"))
           .sort_values("qty", ascending=True))
    if not len(g):
        return None
    g = g.tail(top)

    pal = _palette(dark)
    h = max(2.0, 0.42 * len(g) + 1.1)
    fig, ax = plt.subplots(figsize=(7.6, h))
    ax.set_facecolor(pal["page"])
    labels = [_short_label(i, 26) for i in g.index]
    bars = ax.barh(labels, g["qty"], color=pal["bar"], height=0.62)
    top_i = int(g["qty"].values.argmax())
    bars[top_i].set_color(CHART_ACC)

    for b, (q, p) in zip(bars, zip(g["qty"], g["pallets"])):
        ax.text(b.get_width() + max(g["qty"]) * 0.015,
                b.get_y() + b.get_height() / 2,
                f"{_n(q)}  ({int(p)} plt)", va="center", fontsize=8, color=pal["soft"])

    ax.set_title(title, fontsize=11, color=pal["ink"], weight="bold", loc="left", pad=10)
    ax.set_xlabel("Qty", fontsize=8.5, color=pal["soft"])
    ax.set_xlim(0, float(g["qty"].max()) * 1.22)
    ax.tick_params(axis="y", labelsize=8.5, colors=pal["ink"], length=0)
    ax.tick_params(axis="x", labelsize=8, colors=pal["soft"])
    for sp in ("top", "right", "left"):
        ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color(pal["axis"])
    ax.grid(axis="x", color=pal["grid"], linewidth=0.9)
    ax.set_axisbelow(True)
    return _fig_png(fig, face=pal["page"])


def shortage_chart_png(short: pd.DataFrame, title: str = "Shortage by item",
                       top: int = 12, dark: bool = False) -> bytes | None:
    """Required vs Available vs Short — grouped bars."""
    if short is None or not len(short):
        return None
    d = short.copy()
    for c in ("REQUIRED", "AVAILABLE", "SHORT"):
        d[c] = pd.to_numeric(d.get(c), errors="coerce").fillna(0)
    key = "DOC_ITEM_CODE" if "DOC_ITEM_CODE" in d.columns else d.columns[0]
    g = (d.groupby(key).agg(REQUIRED=("REQUIRED", "sum"), AVAILABLE=("AVAILABLE", "sum"),
                            SHORT=("SHORT", "sum"))
           .sort_values("SHORT", ascending=True).tail(top))
    if not len(g):
        return None

    pal = _palette(dark)
    idx = range(len(g))
    h = max(2.2, 0.62 * len(g) + 1.2)
    fig, ax = plt.subplots(figsize=(7.6, h))
    ax.set_facecolor(pal["page"])
    bh = 0.26
    ax.barh([i + bh for i in idx], g["REQUIRED"], height=bh, color=pal["bar"],
            label="Required")
    ax.barh(list(idx), g["AVAILABLE"], height=bh, color=CHART_OK, label="Available")
    ax.barh([i - bh for i in idx], g["SHORT"], height=bh, color=CHART_ACC, label="Short")

    ax.set_yticks(list(idx))
    ax.set_yticklabels([_short_label(i, 26) for i in g.index], fontsize=8.5,
                       color=pal["ink"])
    for i, v in zip(idx, g["SHORT"]):
        if v > 0:
            ax.text(v + float(g["REQUIRED"].max()) * 0.015, i - bh, f"-{_n(v)}",
                    va="center", fontsize=8, color=CHART_ACC, weight="bold")
    ax.set_title(title, fontsize=11, color=pal["ink"], weight="bold", loc="left", pad=10)
    ax.set_xlabel("Qty", fontsize=8.5, color=pal["soft"])
    leg = ax.legend(fontsize=8, frameon=False, loc="lower right", ncols=3)
    for t in leg.get_texts():
        t.set_color(pal["ink"])
    ax.tick_params(axis="x", labelsize=8, colors=pal["soft"])
    ax.tick_params(axis="y", length=0)
    for sp in ("top", "right", "left"):
        ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color(pal["axis"])
    ax.grid(axis="x", color=pal["grid"], linewidth=0.9)
    ax.set_axisbelow(True)
    return _fig_png(fig, face=pal["page"])


# --------------------------------------------------------------------------- #
# Shortage PDF (invoice එකත් එක්කම)
# --------------------------------------------------------------------------- #
def build_shortage_sheet(info: dict[str, Any], short: pd.DataFrame,
                         doc_lines: pd.DataFrame | None = None,
                         chart: bytes | None = None) -> bytes:
    """Stock මදි නිසා pick කරන්න බැරි වුණ document එකට shortage sheet එකක්."""
    buf = io.BytesIO()
    page = landscape(A4)
    doc = SimpleDocTemplate(buf, pagesize=page, leftMargin=11 * mm, rightMargin=11 * mm,
                            topMargin=9 * mm, bottomMargin=10 * mm,
                            title=f"Shortage {info.get('DOC_NUMBER','')}")
    W = page[0] - 22 * mm
    story: list = []
    num = str(info.get("DOC_NUMBER", ""))

    qr_img = Image(io.BytesIO(qr_png(num, box=8, border=1)), width=30 * mm, height=30 * mm)
    left = Table([[Paragraph("STOCK SHORTAGE  -  NOT PICKED", P_TITLE)],
                  [Paragraph(f"{_txt(info.get('CLIENT','INM0DONA'))} &nbsp;·&nbsp; "
                             f"WH {_txt(info.get('WH_ID',''))} &nbsp;·&nbsp; "
                             f"{_txt(info.get('DOC_TYPE',''))} <b>{_txt(num)}</b>", P_SUB)],
                  [Paragraph(f"Printed {datetime.now():%d-%b-%Y %H:%M} &nbsp;·&nbsp; "
                             f"RUN {_txt(info.get('RUN_ID',''))}", P_SUB)]],
                 colWidths=[W - 42 * mm])
    left.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("TOPPADDING", (0, 0), (-1, -1), 0),
                              ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    qb = Table([[qr_img], [Paragraph(_txt(num), P_LOAD)], [Paragraph("DOCUMENT", P_LOADK)]],
               colWidths=[40 * mm])
    qb.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"),
                            ("BOX", (0, 0), (-1, -1), 0.8, ACC),
                            ("TOPPADDING", (0, 0), (-1, -1), 3),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    head = Table([[left, qb]], colWidths=[W - 42 * mm, 42 * mm])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, 0), "TOP"),
                              ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story += [head, Spacer(1, 5)]

    tot_req = float(pd.to_numeric(short.get("REQUIRED"), errors="coerce").sum() or 0)
    tot_sh = float(pd.to_numeric(short.get("SHORT"), errors="coerce").sum() or 0)
    story.append(_kv_table([
        ("DOCUMENT", f"{info.get('DOC_TYPE','')} {num}"),
        ("DOC DATE", info.get("DOC_DATE", "")),
        ("PLANT", info.get("PLANT", "")),
        ("SHORT LINES", str(len(short))),
        ("SHORT QTY", _n(tot_sh)),
        ("REQUIRED QTY", _n(tot_req)),
    ], cols=6, width=W))
    story.append(Spacer(1, 7))

    story.append(Paragraph("SHORT LINES", P_H))
    rows = []
    for _, r in short.iterrows():
        rows.append([str(r.get("DOC_LINE", "")), str(r.get("DOC_ITEM_CODE", "")),
                     str(r.get("BASE_ID", "")), str(r.get("DESCRIPTION", ""))[:46],
                     _n(r.get("REQUIRED")), _n(r.get("AVAILABLE")), _n(r.get("SHORT")),
                     str(r.get("REASON", ""))])
    widths = [11, 32, 26, 60, 20, 20, 18, 34]
    sc = W / sum(widths)
    story.append(_grid(["Ln", "Doc Item Code", "Base ID", "Description", "Required",
                        "Available", "SHORT", "Reason"], rows, [w * sc for w in widths],
                       aligns={0: "CENTER", 4: "RIGHT", 5: "RIGHT", 6: "RIGHT"}))
    story.append(Spacer(1, 8))

    if chart:
        story.append(KeepTogether([Paragraph("SHORTAGE CHART", P_H),
                                   Image(io.BytesIO(chart), width=W * 0.62,
                                         height=W * 0.62 * _ratio(chart))]))
        story.append(Spacer(1, 8))

    if doc_lines is not None and len(doc_lines):
        rows2 = [[str(r.get("Line", r.get("DOC_LINE", ""))),
                  str(r.get("Item Code", r.get("DOC_ITEM_CODE", ""))),
                  str(r.get("Description", ""))[:52], _n(r.get("Qty", r.get("DOC_QTY"))),
                  str(r.get("Doc UOM", r.get("UOM", "")))]
                 for _, r in doc_lines.iterrows()]
        w2 = [12, 34, 70, 18, 14]
        s2 = (W * 0.72) / sum(w2)
        story.append(KeepTogether([Paragraph("DOCUMENT LINES (full)", P_H),
                                   _grid(["Ln", "Item Code", "Description", "Qty", "UOM"],
                                         rows2, [w * s2 for w in w2],
                                         aligns={0: "CENTER", 3: "RIGHT"})]))
        story.append(Spacer(1, 8))

    sign = Table([[Paragraph("<b>Raised by</b><br/><br/>______________________", P_CELL),
                   Paragraph("<b>Stock checked by</b><br/><br/>______________________",
                             P_CELL),
                   Paragraph("<b>Action</b><br/><br/>____________________________________",
                             P_CELL)]], colWidths=[W / 3] * 3)
    sign.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.6, LINE),
                              ("INNERGRID", (0, 0), (-1, -1), 0.4, LINE),
                              ("TOPPADDING", (0, 0), (-1, -1), 7),
                              ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                              ("LEFTPADDING", (0, 0), (-1, -1), 6)]))
    story.append(KeepTogether(sign))

    def _footer(canv, _d):
        canv.saveState()
        canv.setFont("Helvetica", 6.6)
        canv.setFillColor(colors.HexColor("#7C90AB"))
        canv.drawString(11 * mm, 5.5 * mm,
                        f"SHORTAGE · {_txt(info.get('DOC_TYPE',''))} {_txt(num)} · "
                        f"not picked - stock short · EFL OutBound Pick Generator")
        canv.drawRightString(page[0] - 11 * mm, 5.5 * mm, f"Page {canv.getPageNumber()}")
        canv.setStrokeColor(ACC)
        canv.setLineWidth(1.6)
        canv.line(11 * mm, page[1] - 6 * mm, page[0] - 11 * mm, page[1] - 6 * mm)
        canv.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()


def _ratio(png: bytes) -> float:
    try:
        from PIL import Image as PILImage
        w, h = PILImage.open(io.BytesIO(png)).size
        return h / w
    except Exception:
        return 0.5


def build_shortage_pdf(info: dict[str, Any], short: pd.DataFrame,
                       doc_lines: pd.DataFrame | None = None,
                       source_pdf: bytes | None = None,
                       attach_source: bool = True,
                       chart: bytes | None = None) -> bytes:
    """Shortage sheet + upload කරපු Invoice / DC PDF එකම එකට."""
    sheet = build_shortage_sheet(info, short, doc_lines, chart)
    if attach_source and source_pdf:
        return merge_pdfs([sheet, source_pdf])
    return sheet


# --------------------------------------------------------------------------- #
# Email
# --------------------------------------------------------------------------- #
CID_CHART = "itemchart@efl"


def _addr_list(raw: Any) -> list[str]:
    if raw is None:
        return []
    if isinstance(raw, (list, tuple, set)):
        items = list(raw)
    else:
        items = re.split(r"[,;\s]+", str(raw))
    out, seen = [], set()
    for a in items:
        a = a.strip()
        if a and "@" in a and a.lower() not in seen:
            seen.add(a.lower())
            out.append(a)
    return out


def pick_email_text(docs_info: list[dict[str, Any]], alloc: pd.DataFrame,
                    signature: str = "") -> tuple[str, str, str]:
    """(subject, plain body, html body)"""
    loads = [d.get("LOAD_ID", "") for d in docs_info]
    total = sum(float(d.get("TOTAL_QTY", 0) or 0) for d in docs_info)
    subject = ("OutBound Pick · LOAD ID " + ", ".join(loads[:3])
               + (f" +{len(loads) - 3}" if len(loads) > 3 else ""))

    lines = ["Hi,", "", "Below pick details for the OutBound order(s):", ""]
    html = [
        "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#0F1F33'>",
        "<p>Hi,</p><p>Below pick details for the OutBound order(s):</p>",
    ]

    for d in docs_info:
        lid = d.get("LOAD_ID", "")
        lines += [
            f"LOAD ID       : {lid}",
            f"Document      : {d.get('DOC_TYPE','')} {d.get('DOC_NUMBER','')}"
            f"  ({d.get('DOC_DATE','')})",
            f"Plant         : {d.get('PLANT','')}",
            f"Lines / Qty   : {d.get('LINES','')} lines · {_n(d.get('TOTAL_QTY'))} pcs",
            f"Pallets       : {d.get('PALLETS','')}",
            f"Qty check     : {d.get('VERIFY','')}",
            "",
        ]
        html.append(
            "<table cellpadding='5' cellspacing='0' style='border-collapse:collapse;"
            "border:1px solid #B9C6D6;margin-bottom:8px;font-size:12.5px'>"
            f"<tr><td style='background:#0F1F33;color:#fff' colspan='2'><b>LOAD ID {lid}</b></td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Document</td>"
            f"<td style='border:1px solid #B9C6D6'>{d.get('DOC_TYPE','')} "
            f"{d.get('DOC_NUMBER','')} ({d.get('DOC_DATE','')})</td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Plant</td>"
            f"<td style='border:1px solid #B9C6D6'>{d.get('PLANT','')}</td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Lines / Qty</td>"
            f"<td style='border:1px solid #B9C6D6'>{d.get('LINES','')} lines · "
            f"{_n(d.get('TOTAL_QTY'))} pcs</td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Pallets</td>"
            f"<td style='border:1px solid #B9C6D6'>{d.get('PALLETS','')}</td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Qty check</td>"
            f"<td style='border:1px solid #B9C6D6'>{d.get('VERIFY','')}</td></tr></table>"
        )

    if alloc is not None and len(alloc):
        lines += ["Pick details:",
                  f"{'Ln':<4}{'Item Number':<20}{'Pallet':<24}{'Location':<18}"
                  f"{'Qty':>6}  Balance"]
        html.append("<table cellpadding='5' cellspacing='0' style='border-collapse:collapse;"
                    "border:1px solid #B9C6D6;font-size:12px'><tr style='background:#0F1F33;"
                    "color:#fff'><th>Doc</th><th>Ln</th><th>Item Number</th><th>Pallet</th>"
                    "<th>Location</th><th>Qty</th><th>Balance</th></tr>")
        for _, r in alloc.iterrows():
            lines.append(
                f"{str(r.get('DOC_LINE','')):<4}{str(r.get('ITEM_NUMBER','')):<20}"
                f"{str(r.get('PALLET','')):<24}{str(r.get('LOCATION_ID','')):<18}"
                f"{_n(r.get('QTY_PICKED')):>6}  {_n(r.get('QTY_BALANCE'))}"
            )
            html.append(
                "<tr>"
                f"<td style='border:1px solid #B9C6D6'>{r.get('DOC_NUMBER','')}</td>"
                f"<td style='border:1px solid #B9C6D6' align='center'>{r.get('DOC_LINE','')}</td>"
                f"<td style='border:1px solid #B9C6D6'>{r.get('ITEM_NUMBER','')}</td>"
                f"<td style='border:1px solid #B9C6D6'>{r.get('PALLET','')}</td>"
                f"<td style='border:1px solid #B9C6D6'>{r.get('LOCATION_ID','')}</td>"
                f"<td style='border:1px solid #B9C6D6' align='right'>{_n(r.get('QTY_PICKED'))}</td>"
                f"<td style='border:1px solid #B9C6D6' align='right'>{_n(r.get('QTY_BALANCE'))}</td>"
                "</tr>"
            )
        html.append("</table>")

    lines += ["", f"Total picked qty: {_n(total)}", "",
              signature or "Thanks & regards,", ""]
    html.append("<p style='margin:14px 0 4px'><b>Item details</b></p>"
                f"<img src='cid:{CID_CHART}' alt='Picked qty by item' "
                "style='max-width:640px;width:100%;border:1px solid #B9C6D6;"
                "border-radius:6px'/>")
    html.append(f"<p><b>Total picked qty: {_n(total)}</b></p>"
                f"<p>{(signature or 'Thanks &amp; regards,')}</p></div>")
    return subject, "\n".join(lines), "".join(html)


def shortage_email_text(docs_info: list[dict[str, Any]], short: pd.DataFrame,
                        signature: str = "") -> tuple[str, str, str]:
    """Stock මදි නිසා pick කරන්න බැරි වුණ document වලට email එකක්."""
    nums = [str(d.get("DOC_NUMBER", "")) for d in docs_info]
    tot_short = float(pd.to_numeric(short.get("SHORT"), errors="coerce").sum() or 0) \
        if short is not None and len(short) else 0.0

    subject = ("⚠ Stock Shortage · " + ", ".join(nums[:3])
               + (f" +{len(nums) - 3}" if len(nums) > 3 else ""))

    lines = ["Hi,", "",
             "Below document(s) could NOT be picked - stock is short.",
             "Full document qty නැති නිසා pick කරලා නෑ (partial pick කරන්නේ නෑ).", ""]
    html = ["<div style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
            "color:#0F1F33'><p>Hi,</p><p><b style='color:#FF365B'>Stock shortage</b> - "
            "below document(s) could not be picked. Full document qty නැති නිසා "
            "pick කරලා නෑ.</p>"]

    for d in docs_info:
        num = d.get("DOC_NUMBER", "")
        lines += [f"Document   : {d.get('DOC_TYPE','')} {num}  ({d.get('DOC_DATE','')})",
                  f"Plant      : {d.get('PLANT','')}",
                  f"Doc qty    : {_n(d.get('TOTAL_QTY'))}",
                  f"Reason     : {d.get('REASON','Stock short')}", ""]
        html.append(
            "<table cellpadding='5' cellspacing='0' style='border-collapse:collapse;"
            "border:1px solid #B9C6D6;margin-bottom:8px;font-size:12.5px'>"
            f"<tr><td style='background:#FF365B;color:#fff' colspan='2'>"
            f"<b>{d.get('DOC_TYPE','')} {num}</b></td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Doc date</td>"
            f"<td style='border:1px solid #B9C6D6'>{d.get('DOC_DATE','')}</td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Plant</td>"
            f"<td style='border:1px solid #B9C6D6'>{d.get('PLANT','')}</td></tr>"
            f"<tr><td style='border:1px solid #B9C6D6'>Doc qty</td>"
            f"<td style='border:1px solid #B9C6D6'>{_n(d.get('TOTAL_QTY'))}</td></tr>"
            "</table>")

    if short is not None and len(short):
        lines += ["Short lines:",
                  f"{'Doc':<16}{'Ln':<4}{'Item':<20}{'Req':>7}{'Avail':>8}{'Short':>8}"]
        html.append("<table cellpadding='5' cellspacing='0' style='border-collapse:"
                    "collapse;border:1px solid #B9C6D6;font-size:12px'>"
                    "<tr style='background:#0F1F33;color:#fff'><th>Document</th><th>Ln</th>"
                    "<th>Item Code</th><th>Description</th><th>Required</th>"
                    "<th>Available</th><th>Short</th></tr>")
        for _, r in short.iterrows():
            lines.append(
                f"{str(r.get('DOC_NUMBER','')):<16}{str(r.get('DOC_LINE','')):<4}"
                f"{str(r.get('DOC_ITEM_CODE','')):<20}{_n(r.get('REQUIRED')):>7}"
                f"{_n(r.get('AVAILABLE')):>8}{_n(r.get('SHORT')):>8}")
            html.append(
                "<tr>"
                f"<td style='border:1px solid #B9C6D6'>{r.get('DOC_NUMBER','')}</td>"
                f"<td style='border:1px solid #B9C6D6' align='center'>"
                f"{r.get('DOC_LINE','')}</td>"
                f"<td style='border:1px solid #B9C6D6'>{r.get('DOC_ITEM_CODE','')}</td>"
                f"<td style='border:1px solid #B9C6D6'>"
                f"{str(r.get('DESCRIPTION',''))[:40]}</td>"
                f"<td style='border:1px solid #B9C6D6' align='right'>"
                f"{_n(r.get('REQUIRED'))}</td>"
                f"<td style='border:1px solid #B9C6D6' align='right'>"
                f"{_n(r.get('AVAILABLE'))}</td>"
                f"<td style='border:1px solid #B9C6D6;color:#FF365B' align='right'>"
                f"<b>{_n(r.get('SHORT'))}</b></td></tr>")
        html.append("</table>")

    html.append("<p style='margin:14px 0 4px'><b>Item details</b></p>"
                f"<img src='cid:{CID_CHART}' alt='Shortage by item' "
                "style='max-width:640px;width:100%;border:1px solid #B9C6D6;"
                "border-radius:6px'/>")
    lines += ["", f"Total short qty: {_n(tot_short)}",
              "Please arrange stock / confirm short shipment.", "",
              signature or "Thanks & regards,", ""]
    html.append(f"<p><b>Total short qty: {_n(tot_short)}</b></p>"
                "<p>Please arrange stock / confirm short shipment.</p>"
                f"<p>{(signature or 'Thanks &amp; regards,')}</p></div>")
    return subject, "\n".join(lines), "".join(html)


def mailto_link(to: Any, subject: str, body: str, cc: Any = None,
                max_body: int = 1600) -> str:
    """Default mail app එක open කරන mailto: link (attachment support නෑ)."""
    tos = ",".join(_addr_list(to))
    ccs = ",".join(_addr_list(cc))
    b = body if len(body) <= max_body else body[:max_body] + "\n\n... (full details attached)"
    q = f"subject={quote(subject)}&body={quote(b)}"
    if ccs:
        q += f"&cc={quote(ccs)}"
    return f"mailto:{quote(tos)}?{q}"


def build_eml(to: Any, subject: str, body: str, html: str = "", cc: Any = None,
              sender: str = "", attachments: list[tuple[str, bytes, str]] | None = None,
              inline_png: bytes | None = None, cid: str = CID_CHART) -> bytes:
    """
    .eml file — double-click කරාම default mail app එකේ **draft** එකක් විදිහට open වෙනවා
    (attachment ඔක්කොම ඇතුලේ). Outlook එකට 'X-Unsent: 1' header එක ඕන.
    """
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["To"] = ", ".join(_addr_list(to))
    if _addr_list(cc):
        msg["Cc"] = ", ".join(_addr_list(cc))
    if sender:
        msg["From"] = sender
    msg["X-Unsent"] = "1"
    msg["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S %z") or ""
    msg.set_content(body)
    if html:
        msg.add_alternative(html, subtype="html")
        if inline_png:
            # HTML part එකට chart එක inline (cid:) විදිහට
            html_part = msg.get_payload()[-1]
            html_part.add_related(inline_png, maintype="image", subtype="png",
                                  cid=f"<{cid}>", filename="item_chart.png")

    for name, data, mime in (attachments or []):
        if not data:
            continue
        maintype, _, subtype = mime.partition("/")
        msg.add_attachment(data, maintype=maintype or "application",
                           subtype=subtype or "octet-stream", filename=name)
    return msg.as_bytes()
