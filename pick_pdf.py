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
# Email
# --------------------------------------------------------------------------- #
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
    html.append(f"<p><b>Total picked qty: {_n(total)}</b></p>"
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
              sender: str = "", attachments: list[tuple[str, bytes, str]] | None = None) -> bytes:
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

    for name, data, mime in (attachments or []):
        if not data:
            continue
        maintype, _, subtype = mime.partition("/")
        msg.add_attachment(data, maintype=maintype or "application",
                           subtype=subtype or "octet-stream", filename=name)
    return msg.as_bytes()
