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
from matplotlib.ticker import MaxNLocator  # noqa: E402
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



def _line_labels(keys: list[str], items: list[str], n_max: int = 14) -> tuple[list[str], int]:
    """Tick labels — line number, plus the item number when there is room."""
    if len(keys) <= n_max:
        return [f"{k}\n{_short_label(i, 17)}" for k, i in zip(keys, items)], (
            0 if len(keys) <= 8 else 20)
    return list(keys), (0 if len(keys) <= 24 else 60)


def pick_line_chart_png(alloc: pd.DataFrame, title: str = "Pick details by line",
                        dark: bool = False, max_docs: int = 6) -> bytes | None:
    """
    Line chart over the document lines — this is what the pick actually is:
    line 1, 2, 3 … in order, with the qty taken and the stock left behind.

    One document  -> Pick qty + Balance left
    Several       -> one line per document
    """
    if alloc is None or not len(alloc):
        return None
    d = alloc.copy()
    for c in ("QTY_PICKED", "QTY_BALANCE"):
        d[c] = pd.to_numeric(d.get(c), errors="coerce").fillna(0)
    d["DOC_LINE"] = pd.to_numeric(d["DOC_LINE"], errors="coerce").fillna(0).astype(int)

    docs = list(dict.fromkeys(d["DOC_NUMBER"].astype(str)))
    # a line needs at least two points — one line item is a bar chart
    if len(docs) == 1 and d["DOC_LINE"].nunique() < 2:
        return pick_chart_png(alloc, "Picked qty by item", dark=dark)

    pal = _palette(dark)
    fig, ax = plt.subplots(figsize=(7.8, 3.5))
    ax.set_facecolor(pal["page"])

    ax2 = None
    if len(docs) == 1:
        g = (d.groupby("DOC_LINE")
               .agg(qty=("QTY_PICKED", "sum"), bal=("QTY_BALANCE", "sum"),
                    plt_n=("PALLET", "nunique"), item=("ITEM_NUMBER", "first"))
               .sort_index())
        x = list(range(len(g)))
        # Balance left is often 100x the pick qty (a full pallet behind a 2-unit
        # pick), so it goes on its own axis or the pick line flattens to nothing.
        ax2 = ax.twinx()
        ax2.set_facecolor("none")
        ax2.plot(x, g["bal"], marker="s", markersize=4.5, linewidth=1.5, linestyle="--",
                 color=CHART_OK, label="Balance left", zorder=2)
        ax.plot(x, g["qty"], marker="o", markersize=6, linewidth=2.2,
                color=CHART_ACC, label="Pick qty", zorder=3)
        top = max(float(g["qty"].max()), 1.0)
        for i, (q, n) in enumerate(zip(g["qty"], g["plt_n"])):
            ax.annotate(f"{_n(q)}" + (f"  ({int(n)} plt)" if n > 1 else ""),
                        (i, q), textcoords="offset points", xytext=(0, 10),
                        ha="center", fontsize=7.6, color=CHART_ACC, weight="bold",
                        zorder=5)
        ax2.set_ylim(0, max(float(g["bal"].max()), 1.0) * 1.35)
        ax2.yaxis.set_major_locator(MaxNLocator(nbins=6, integer=True))
        ax2.set_ylabel("Balance left on pallets", fontsize=8.5, color=CHART_OK)
        ax2.tick_params(axis="y", labelsize=8, colors=CHART_OK)
        for sp in ("top", "left"):
            ax2.spines[sp].set_visible(False)
        ax2.spines["right"].set_color(CHART_OK)
        h1, l1 = ax.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        leg = ax.legend(h1 + h2, l1 + l2, fontsize=8, frameon=False,
                        loc="upper left", ncols=2,
                        bbox_to_anchor=(0, 1.02))
        for t in leg.get_texts():
            t.set_color(pal["ink"])
        labels, rot = _line_labels([str(i) for i in g.index],
                                   [str(v) for v in g["item"]])
        ax.set_xticks(x)
        ax.set_xticklabels(labels, fontsize=7.4, color=pal["ink"], rotation=rot,
                           ha="center" if rot == 0 else "right")
        ax.set_title(f"{title}  ·  {docs[0]}", fontsize=11, color=pal["ink"],
                     weight="bold", loc="left", pad=16)
    else:
        top = 1.0
        for k, doc in enumerate(docs[:max_docs]):
            one = (d[d["DOC_NUMBER"].astype(str) == doc].groupby("DOC_LINE")["QTY_PICKED"]
                   .sum().sort_index())
            top = max(top, float(one.max()))
            ax.plot(list(one.index), one.values, marker="o", markersize=5,
                    linewidth=1.8, label=_short_label(doc, 18),
                    color=[CHART_ACC, CHART_INK if not dark else pal["bar"], CHART_OK,
                           CHART_WARN, "#7B5EA7", "#2F8FA8"][k % 6], zorder=3)
        ax.set_xlabel("Document line", fontsize=8.5, color=pal["soft"])
        ax.xaxis.set_major_locator(MaxNLocator(integer=True))   # lines are 1,2,3…
        ax.tick_params(axis="x", labelsize=8, colors=pal["soft"])
        ax.set_title(f"{title}  ·  {len(docs)} documents", fontsize=11, color=pal["ink"],
                     weight="bold", loc="left", pad=12)

    ax.set_ylabel("Pick qty", fontsize=8.5,
                  color=CHART_ACC if ax2 is not None else pal["soft"])
    ax.set_ylim(0, top * 1.34)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=7, integer=True))
    ax.tick_params(axis="y", labelsize=8,
                   colors=CHART_ACC if ax2 is not None else pal["soft"])
    if ax2 is None:
        leg = ax.legend(fontsize=8, frameon=False, loc="upper right", ncols=3)
        for t in leg.get_texts():
            t.set_color(pal["ink"])
    ax.spines["top"].set_visible(False)
    if ax2 is None:
        ax.spines["right"].set_visible(False)
    ax.spines["bottom"].set_color(pal["axis"])
    ax.spines["left"].set_color(CHART_ACC if ax2 is not None else pal["axis"])
    ax.grid(axis="y", color=pal["grid"], linewidth=0.9)
    ax.set_axisbelow(True)
    return _fig_png(fig, face=pal["page"])


def shortage_line_chart_png(short: pd.DataFrame, title: str = "Shortage by line",
                            dark: bool = False) -> bytes | None:
    """Required / free / short across the short lines, in document order."""
    if short is None or not len(short):
        return None
    d = short.copy()
    for c in ("REQUIRED", "AVAILABLE", "SHORT", "ON_PICK_TASK"):
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0)
        else:
            d[c] = 0.0
    d["DOC_LINE"] = pd.to_numeric(d.get("DOC_LINE"), errors="coerce").fillna(0).astype(int)
    d = d.sort_values(["DOC_NUMBER", "DOC_LINE"]).reset_index(drop=True)
    if len(d) < 2:                       # one short line — bars read better
        return shortage_chart_png(short, "Shortage by item", dark=dark)

    pal = _palette(dark)
    fig, ax = plt.subplots(figsize=(7.8, 3.5))
    ax.set_facecolor(pal["page"])
    x = range(len(d))
    ax.plot(x, d["REQUIRED"], marker="o", markersize=5.5, linewidth=1.9,
            color=CHART_INK if not dark else pal["bar"], label="Required", zorder=3)
    ax.plot(x, d["AVAILABLE"], marker="s", markersize=4.8, linewidth=1.7,
            color=CHART_OK, label="Free stock", zorder=3)
    ax.plot(x, d["SHORT"], marker="v", markersize=5.5, linewidth=1.9,
            color=CHART_ACC, label="Short", zorder=4)
    if float(d["ON_PICK_TASK"].sum()) > 0:
        ax.plot(x, d["ON_PICK_TASK"], marker="d", markersize=4.5, linewidth=1.4,
                linestyle=":", color=CHART_WARN, label="On pick task", zorder=2)

    for i, v in enumerate(d["SHORT"]):
        if v > 0:
            ax.annotate(f"-{_n(v)}", (i, v), textcoords="offset points", xytext=(0, 10),
                        ha="center", fontsize=7.8, color=CHART_ACC, weight="bold")

    labels, rot = _line_labels([str(v) for v in d["DOC_LINE"]],
                               [str(v) for v in d.get("DOC_ITEM_CODE", "")])
    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=7.4, color=pal["ink"], rotation=rot,
                       ha="center" if rot == 0 else "right")
    ax.set_title(title, fontsize=11, color=pal["ink"], weight="bold", loc="left", pad=12)
    ax.set_ylabel("Qty", fontsize=8.5, color=pal["soft"])
    ax.set_ylim(0, max(1.0, float(d["REQUIRED"].max())) * 1.3)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=7, integer=True))
    ax.tick_params(axis="y", labelsize=8, colors=pal["soft"])
    leg = ax.legend(fontsize=8, frameon=False, loc="upper right", ncols=4)
    for t in leg.get_texts():
        t.set_color(pal["ink"])
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    for sp in ("bottom", "left"):
        ax.spines[sp].set_color(pal["axis"])
    ax.grid(axis="y", color=pal["grid"], linewidth=0.9)
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
                     str(r.get("BASE_ID", "")), str(r.get("DESCRIPTION", ""))[:42],
                     _n(r.get("REQUIRED")), _n(r.get("AVAILABLE")),
                     _n(r.get("ON_PICK_TASK")), _n(r.get("SHORT")),
                     str(r.get("REASON", ""))])
    widths = [11, 32, 26, 54, 19, 19, 20, 17, 40]
    sc = W / sum(widths)
    story.append(_grid(["Ln", "Doc Item Code", "Base ID", "Description", "Required",
                        "Free", "On pick task", "SHORT", "Reason"], rows,
                       [w * sc for w in widths],
                       aligns={0: "CENTER", 4: "RIGHT", 5: "RIGHT", 6: "RIGHT",
                               7: "RIGHT"}))
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


def _ascii_table(headers: list[str], rows: list[list[str]],
                 right: set[int] | None = None, gap: str = "  ") -> list[str]:
    """
    Plain-text table where the header always sits over its own column.

    Column widths come from the data, not from guessed constants — a long doc
    number or pallet id used to push every following column out of line.
    Mail clients that render text/plain in a proportional font will still drift,
    which is why the HTML part carries a real <table>; this is the fallback.
    """
    right = right or set()
    cells = [[str("" if c is None else c) for c in r] for r in rows]
    n = len(headers)
    width = [max([len(headers[i])] + [len(r[i]) for r in cells]) if cells
             else len(headers[i]) for i in range(n)]

    def fmt(vals: list[str]) -> str:
        out = []
        for i, v in enumerate(vals):
            out.append(v.rjust(width[i]) if i in right else v.ljust(width[i]))
        return gap.join(out).rstrip()

    rule = gap.join("-" * w for w in width)
    return [fmt(headers), rule] + [fmt(r) for r in cells]


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
        lines.append("Pick details:")
        # line numbers restart per document, so name the document when there
        # is more than one — otherwise "Ln 1" twice is ambiguous
        _multi = alloc["DOC_NUMBER"].astype(str).nunique() > 1
        _rows: list[list[str]] = []
        html.append("<table cellpadding='5' cellspacing='0' style='border-collapse:collapse;"
                    "border:1px solid #B9C6D6;font-size:12px'><tr style='background:#0F1F33;"
                    "color:#fff'><th>Doc</th><th>Ln</th><th>Item Number</th><th>Pallet</th>"
                    "<th>Location</th><th>Qty</th><th>Balance</th></tr>")
        for _, r in alloc.iterrows():
            _row = [str(r.get("DOC_LINE", "")), str(r.get("ITEM_NUMBER", "")),
                    str(r.get("PALLET", "")), str(r.get("LOCATION_ID", "")),
                    _n(r.get("QTY_PICKED")), _n(r.get("QTY_BALANCE"))]
            _rows.append(([str(r.get("DOC_NUMBER", ""))] + _row) if _multi else _row)
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
        _head = ["Ln", "Item Number", "Pallet", "Location", "Qty", "Balance"]
        lines += (_ascii_table(["Document"] + _head, _rows, right={1, 5, 6}) if _multi
                  else _ascii_table(_head, _rows, right={0, 4, 5}))
        html.append("</table>")

    lines += ["", f"Total picked qty: {_n(total)}", "",
              signature or "Thanks & regards,", ""]
    html.append("<p style='margin:14px 0 4px'><b>Item details</b></p>"
                f"<img src='cid:{CID_CHART}' alt='Pick details by line' "
                "style='max-width:660px;width:100%;border:1px solid #B9C6D6;"
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
             "The document(s) below could NOT be picked - stock is short.",
             "Nothing was picked for them: an order is only picked when the full "
             "document quantity is available (no partial picking).", ""]
    html = ["<div style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;"
            "color:#0F1F33'><p>Hi,</p><p><b style='color:#FF365B'>Stock shortage</b> - "
            "the document(s) below could not be picked. An order is only picked when "
            "the full document quantity is available, so nothing was picked.</p>"]

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
        lines.append("Short lines:")
        _multi = short["DOC_NUMBER"].astype(str).nunique() > 1
        _rows: list[list[str]] = []
        html.append("<table cellpadding='5' cellspacing='0' style='border-collapse:"
                    "collapse;border:1px solid #B9C6D6;font-size:12px'>"
                    "<tr style='background:#0F1F33;color:#fff'><th>Document</th><th>Ln</th>"
                    "<th>Item Code</th><th>Description</th><th>Required</th>"
                    "<th>Available</th><th>Short</th></tr>")
        for _, r in short.iterrows():
            _row = [str(r.get("DOC_LINE", "")), str(r.get("DOC_ITEM_CODE", "")),
                    _n(r.get("REQUIRED")), _n(r.get("AVAILABLE")),
                    _n(r.get("ON_PICK_TASK")), _n(r.get("SHORT"))]
            _rows.append(([str(r.get("DOC_NUMBER", ""))] + _row) if _multi else _row)
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
        _head = ["Ln", "Item", "Req", "Free", "On pick", "Short"]
        lines += (_ascii_table(["Document"] + _head, _rows, right={1, 3, 4, 5, 6})
                  if _multi else _ascii_table(_head, _rows, right={0, 2, 3, 4, 5}))
        html.append("</table>")

    html.append("<p style='margin:14px 0 4px'><b>Item details</b></p>"
                f"<img src='cid:{CID_CHART}' alt='Shortage by line' "
                "style='max-width:660px;width:100%;border:1px solid #B9C6D6;"
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
    # Some clients (Outlook in particular) render text/plain in a proportional
    # font, which pulls fixed-width tables apart. Ask for a monospace face —
    # clients that honour format=flowed/fixed will keep the columns straight.
    msg.set_content(body)
    try:
        msg.replace_header("Content-Type",
                           'text/plain; charset="utf-8"; format=fixed')
    except KeyError:
        pass
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
