#!/usr/bin/env python3
"""Build a Cisco Enterprise Agreement proposal PPTX from an EAMP price estimate."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from parse_eamp import apply_partner_margin, parse_eamp_workbook, suite_rollups

DEFAULT_MARGIN_POINTS = 20

# Widescreen 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

NAVY = RGBColor(0x04, 0x1C, 0x3C)
NAVY_MID = RGBColor(0x0B, 0x3A, 0x60)
CISCO = RGBColor(0x00, 0x9E, 0xDB)
TEAL = RGBColor(0x00, 0xB3, 0xA4)
GREEN = RGBColor(0x6B, 0xB7, 0x46)
GOLD = RGBColor(0xF7, 0xB5, 0x00)
ORANGE = RGBColor(0xED, 0x72, 0x2F)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xF5, 0xF7, 0xFA)
ROW_ALT = RGBColor(0xF3, 0xF7, 0xFB)
INK = RGBColor(0x12, 0x1C, 0x2B)
MUTED = RGBColor(0x5B, 0x67, 0x75)
LINE = RGBColor(0xC9, 0xD3, 0xDE)
TOTAL_BG = RGBColor(0x0B, 0x3A, 0x60)

FONT = "Calibri"
HERE = Path(__file__).resolve().parent
LOGO = HERE / "assets" / "cisco-logo.png"

PORTFOLIO_COLORS = {
    "Networking Infrastructure": CISCO,
    "Security": ORANGE,
    "Services": TEAL,
    "Collaboration": GOLD,
    "Applications Infrastructure": GREEN,
    "Hybrid": RGBColor(0x7A, 0x5C, 0xF5),
    "Provider Connectivity": RGBColor(0x8A, 0x96, 0xA3),
    "Cisco Professional Services": RGBColor(0x2D, 0x6A, 0x4F),
}


def money(value, digits=0) -> str:
    if value is None or value == "":
        return "—"
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(number) < 0.005:
        return "$0" if digits == 0 else "$0.00"
    if digits == 0:
        return f"${number:,.0f}"
    return f"${number:,.{digits}f}"


def qty(value) -> str:
    if value is None or value == "":
        return "—"
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if number.is_integer():
        return f"{int(number):,}"
    return f"{number:,.1f}"


def pct(value, digits=0) -> str:
    if value is None or value == "":
        return "—"
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(number) < 0.05:
        return "0%"
    return f"{number:.{digits}f}%"


def clean_text(value) -> str:
    if value is None:
        return "—"
    text = str(value).strip()
    if text in ("", "-", "N/A"):
        return "—"
    return text.replace("  ", " ")


def duration_label(months) -> str:
    try:
        months_n = int(float(months))
    except (TypeError, ValueError):
        return clean_text(months)
    years = months_n / 12
    if years.is_integer():
        return f"{int(years)} years ({months_n} months)"
    return f"{months_n} months"


def annual_from_tcv(tcv, months) -> float:
    try:
        m = float(months)
    except (TypeError, ValueError):
        m = 60.0
    if m <= 0:
        m = 60.0
    return float(tcv or 0) * 12.0 / m


def find_total_row(rows: list[dict]) -> dict:
    for row in rows:
        if str(row.get("portfolio", "")).strip().upper() == "TOTAL":
            return row
    return {}


def nonempty_portfolios(data: dict) -> list[dict]:
    return [p for p in data["portfolios"] if p.get("populated")]


def set_run(run, size=14, bold=False, color=INK, font=FONT, italic=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = font


def add_text_box(slide, l, t, w, h, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, font=FONT, italic=False, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(anchor, "t"))
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font, italic=italic)
    return box


def add_rect(slide, l, t, w, h, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def add_round(slide, l, t, w, h, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.adjustments[0] = 0.08
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1)
    shape.shadow.inherit = False
    return shape


def _set_cell_border(cell, color_hex="C9D3DE", width="6350"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for edge in ("lnL", "lnR", "lnT", "lnB"):
        existing = tcPr.find(qn(f"a:{edge}"))
        if existing is not None:
            tcPr.remove(existing)
        ln = etree.SubElement(tcPr, qn(f"a:{edge}"), w=width, cap="flat", cmpd="sng", algn="ctr")
        srgb = etree.SubElement(etree.SubElement(ln, qn("a:solidFill")), qn("a:srgbClr"))
        srgb.set("val", color_hex)
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
        etree.SubElement(ln, qn("a:round"))


def set_cell(cell, text, size=11, bold=False, color=INK, fill=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE):
    cell.text = ""
    cell.vertical_anchor = anchor
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = "" if text is None else str(text)
    set_run(run, size=size, bold=bold, color=color)
    _set_cell_border(cell)
    # tighten margins
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    mar = {
        "marL": "72000",
        "marR": "72000",
        "marT": "40000",
        "marB": "40000",
    }
    for attr, val in mar.items():
        tcPr.set(attr, val)


def is_licensed(item: dict) -> bool:
    return bool(float(item.get("tcv") or 0) or float(item.get("list_price") or 0) or float(item.get("unit_list_month") or 0))


def licensed_qty(items, field) -> str:
    return qty(sum(float(i.get(field) or 0) for i in items if is_licensed(i)))


def add_table(slide, left, top, width, height, headers, rows, col_widths=None, font_size=11, numeric_cols=None):
    numeric_cols = set(numeric_cols or [])
    n_rows = len(rows) + 1
    if height is None:
        height = min(Inches(0.38) * n_rows, Inches(6.05))
    table_shape = slide.shapes.add_table(n_rows, len(headers), left, top, width, height)
    table = table_shape.table
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = w
    for i, header in enumerate(headers):
        set_cell(table.cell(0, i), header, size=font_size, bold=True, color=WHITE, fill=NAVY, align=PP_ALIGN.CENTER if i in numeric_cols else PP_ALIGN.LEFT)
    for r, row in enumerate(rows, start=1):
        is_total = str(row[0]).strip().upper() in {"TOTAL", "PORTFOLIO TOTAL", "SOFTWARE TOTAL", "SERVICES TOTAL"}
        fill = TOTAL_BG if is_total else (ROW_ALT if r % 2 == 0 else WHITE)
        color = WHITE if is_total else INK
        bold = is_total
        for c, value in enumerate(row):
            align = PP_ALIGN.RIGHT if c in numeric_cols else PP_ALIGN.LEFT
            if c == 0 and not is_total:
                align = PP_ALIGN.LEFT
            set_cell(table.cell(r, c), value, size=font_size, bold=bold, color=color, fill=fill, align=align)
    return table


def footer(slide, page, total, confidential=True):
    add_rect(slide, 0, Inches(7.22), SLIDE_W, Inches(0.28), NAVY)
    left = "Cisco Confidential  |  Customer pricing includes 20 pts partner margin  |  Not an approved Cisco quote"
    add_text_box(slide, Inches(0.4), Inches(7.24), Inches(10.2), Inches(0.24), left, size=10, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_text_box(slide, Inches(11.2), Inches(7.24), Inches(1.7), Inches(0.24), f"{page}  /  {total}", size=10, color=WHITE, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def header_bar(slide, title, subtitle=None, logo=True):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.78), NAVY)
    add_rect(slide, 0, Inches(0.78), SLIDE_W, Inches(0.06), CISCO)
    if logo and LOGO.exists():
        slide.shapes.add_picture(str(LOGO), Inches(0.32), Inches(0.16), height=Inches(0.46))
    add_text_box(slide, Inches(1.7), Inches(0.10), Inches(10.9), Inches(0.38), title, size=22, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text_box(slide, Inches(1.7), Inches(0.44), Inches(10.9), Inches(0.28), subtitle, size=12, color=CISCO)


def kpi_card(slide, l, t, w, h, label, value, caption=None, accent=CISCO):
    add_round(slide, l, t, w, h, WHITE, LINE)
    add_rect(slide, l, t, Inches(0.08), h, accent)
    add_text_box(slide, l + Inches(0.22), t + Inches(0.12), w - Inches(0.32), Inches(0.28), label.upper(), size=11, bold=True, color=MUTED)
    add_text_box(slide, l + Inches(0.22), t + Inches(0.38), w - Inches(0.32), Inches(0.46), value, size=26, bold=True, color=NAVY)
    if caption:
        add_text_box(slide, l + Inches(0.22), t + Inches(0.86), w - Inches(0.32), Inches(0.32), caption, size=11, color=MUTED)


def paginate(items, size):
    if not items:
        return [[]]
    return [items[i : i + size] for i in range(0, len(items), size)]


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build_presentation(data: dict, output: Path) -> Path:
    meta = data["meta"]
    financial = data["financial_summary"]
    total = find_total_row(financial)
    duration = None
    billing = None
    rsd = None
    for p in data["portfolios"]:
        duration = duration or p["header"].get("duration_months")
        billing = billing or p["header"].get("billing_model")
        rsd = rsd or p["header"].get("requested_ship_date")
    duration = duration or 60
    billing = billing or "Annual"
    years = int(float(duration) / 12) if duration else 5
    tcv = float(total.get("total") or 0)
    list_price = float(total.get("list_price") or 0)
    discount = float(total.get("discount_amount") or 0)
    otd = float(total.get("one_time_discount") or 0)
    sw = float(total.get("software") or 0)
    svcs = float(total.get("services") or 0)
    eamp_tcv = float(total.get("eamp_total") or tcv)
    eamp_discount = float(total.get("eamp_discount_amount") or discount)
    partner_margin = float(total.get("partner_margin") or 0)
    margin_points = int(data.get("margin_points") or DEFAULT_MARGIN_POINTS)
    annual = annual_from_tcv(tcv, duration)
    discount_pct = (list_price - tcv) / list_price * 100 if list_price else 0
    pricing_note = f"Customer pricing includes {margin_points} pts partner margin"
    customer_name = clean_text(meta.get("smart_account")) or "Customer"
    partner = clean_text(meta.get("partner"))
    proposal_id = clean_text(meta.get("proposal_id"))
    deal_id = clean_text(meta.get("deal_id"))
    quote_id = clean_text(meta.get("quote_id"))
    populated = nonempty_portfolios(data)

    # Pre-count slides so footers can show X / N. We build a plan then render.
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slides_built: list = []

    def _is_full_bleed(slide) -> bool:
        if not slide.shapes:
            return False
        shape = slide.shapes[0]
        return shape.width >= SLIDE_W - 20000 and shape.height >= SLIDE_H - 20000

    def stamp_footers():
        total_pages = len(prs.slides)
        for idx, slide in enumerate(prs.slides, start=1):
            if _is_full_bleed(slide):
                continue
            footer(slide, idx, total_pages)

    # ----- 1. Cover -----
    s = blank_slide(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, CISCO)
    if LOGO.exists():
        s.shapes.add_picture(str(LOGO), Inches(0.7), Inches(0.45), height=Inches(0.62))
    add_text_box(s, Inches(0.7), Inches(1.35), Inches(11.5), Inches(0.35), "CISCO ENTERPRISE AGREEMENT 3.0", size=14, bold=True, color=CISCO)
    add_text_box(s, Inches(0.7), Inches(1.75), Inches(12.0), Inches(1.1), customer_name, size=40, bold=True, color=WHITE)
    add_rect(s, Inches(0.7), Inches(2.95), Inches(2.4), Inches(0.07), CISCO)
    add_text_box(
        s,
        Inches(0.7),
        Inches(3.2),
        Inches(11.5),
        Inches(0.7),
        f"{years}-Year Enterprise Agreement Proposal  ·  {margin_points} pts partner margin",
        size=24,
        color=WHITE,
    )
    add_text_box(
        s,
        Inches(0.7),
        Inches(3.85),
        Inches(11.5),
        Inches(0.4),
        clean_text(meta.get("proposal_name")),
        size=14,
        color=RGBColor(0xB8, 0xC7, 0xD6),
    )

    facts = [
        ("Total contract value", money(tcv)),
        ("Term / billing", f"{years} years  ·  {billing}"),
        ("Expected book date", clean_text(meta.get("expected_book_date"))),
        ("Partner", partner),
    ]
    card_w = Inches(2.85)
    for i, (label, value) in enumerate(facts):
        left = Inches(0.7) + i * Inches(3.05)
        add_round(s, left, Inches(4.55), card_w, Inches(1.15), NAVY_MID)
        add_text_box(s, left + Inches(0.16), Inches(4.65), card_w - Inches(0.28), Inches(0.28), label.upper(), size=11, bold=True, color=CISCO)
        add_text_box(s, left + Inches(0.16), Inches(4.95), card_w - Inches(0.28), Inches(0.55), value, size=16, bold=True, color=WHITE)

    add_text_box(
        s,
        Inches(0.7),
        Inches(6.05),
        Inches(12),
        Inches(0.55),
        f"Proposal ID {proposal_id}   ·   Deal ID {deal_id}   ·   Quote ID {quote_id}   ·   Buying Program {clean_text(meta.get('buying_program_id'))}\n"
        f"Report created {clean_text(meta.get('report_created_on'))}   ·   Cisco Account Manager {clean_text(meta.get('account_manager'))}",
        size=12,
        color=RGBColor(0x9A, 0xB0, 0xC3),
    )
    add_text_box(
        s,
        Inches(0.7),
        Inches(6.85),
        Inches(12),
        Inches(0.35),
        f"CISCO CONFIDENTIAL  ·  Indicative customer pricing includes {margin_points} points of partner margin on EAMP net. Not an approved Cisco quote.",
        size=11,
        color=RGBColor(0x7A, 0x90, 0xA4),
    )

    # ----- 2. Agenda -----
    s = blank_slide(prs)
    header_bar(s, "Agenda", "Hallmark CX EA  ·  proposal walkthrough")
    agenda = [
        ("01", "Opportunity snapshot", "Customer, partner, commercial identifiers, and term."),
        ("02", "Why an Enterprise Agreement", "Price protection, True Forward, and a single enterprise term."),
        ("03", "Investment summary", "Customer TCV with 20 pts partner margin, list vs. net, annual view."),
        ("04", "Portfolio detail", "Networking Infrastructure, Security (ISE), and CX services."),
        ("05", "Commercial terms & next steps", "Billing, True Forward, partner path, and recommended actions."),
        ("06", "Appendix", "SKU-level software, CX services, and how to read the EAMP columns."),
    ]
    for i, (num, title, blurb) in enumerate(agenda):
        top = Inches(1.1) + i * Inches(0.92)
        add_round(s, Inches(0.55), top, Inches(12.2), Inches(0.82), WHITE, LINE)
        add_round(s, Inches(0.75), top + Inches(0.16), Inches(0.55), Inches(0.50), CISCO)
        add_text_box(s, Inches(0.75), top + Inches(0.22), Inches(0.55), Inches(0.40), num, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, Inches(1.55), top + Inches(0.12), Inches(10.8), Inches(0.32), title, size=18, bold=True, color=NAVY)
        add_text_box(s, Inches(1.55), top + Inches(0.44), Inches(10.8), Inches(0.28), blurb, size=13, color=MUTED)

    # ----- 3. Opportunity snapshot -----
    s = blank_slide(prs)
    header_bar(s, "Opportunity snapshot", "From EAMP Price Estimate ID " + proposal_id)
    left_items = [
        ("End customer", customer_name),
        ("Account", clean_text(meta.get("end_customer_account")) or customer_name),
        ("Location", (meta.get("end_customer_address") or "—").replace("\n", ", ")),
        ("Smart Account", customer_name),
        ("Cisco Account Manager", clean_text(meta.get("account_manager"))),
    ]
    right_items = [
        ("Partner", partner),
        ("Project", clean_text(meta.get("project_name"))),
        ("EA motion", clean_text(meta.get("ea_motion"))),
        ("Price list", clean_text(meta.get("price_list"))),
        ("Requested ship / book", f"{clean_text(rsd)}  ·  book {clean_text(meta.get('expected_book_date'))}"),
    ]
    for i, (label, value) in enumerate(left_items):
        top = Inches(1.08) + i * Inches(0.72)
        add_round(s, Inches(0.5), top, Inches(6.05), Inches(0.64), WHITE, LINE)
        add_text_box(s, Inches(0.7), top + Inches(0.06), Inches(5.7), Inches(0.22), label.upper(), size=10, bold=True, color=MUTED)
        add_text_box(s, Inches(0.7), top + Inches(0.28), Inches(5.7), Inches(0.30), value, size=14, bold=True, color=NAVY)
    for i, (label, value) in enumerate(right_items):
        top = Inches(1.08) + i * Inches(0.72)
        add_round(s, Inches(6.75), top, Inches(6.05), Inches(0.64), WHITE, LINE)
        add_text_box(s, Inches(6.95), top + Inches(0.06), Inches(5.7), Inches(0.22), label.upper(), size=10, bold=True, color=MUTED)
        add_text_box(s, Inches(6.95), top + Inches(0.28), Inches(5.7), Inches(0.30), value, size=14, bold=True, color=NAVY)

    id_row = [
        ("Deal ID", deal_id),
        ("Quote ID", quote_id),
        ("Proposal ID", proposal_id),
        ("Buying Program", clean_text(meta.get("buying_program_id"))),
    ]
    for i, (label, value) in enumerate(id_row):
        left = Inches(0.5) + i * Inches(3.15)
        add_round(s, left, Inches(4.78), Inches(3.0), Inches(0.95), OFF_WHITE, LINE)
        add_text_box(s, left + Inches(0.16), Inches(4.88), Inches(2.7), Inches(0.22), label.upper(), size=10, bold=True, color=MUTED)
        add_text_box(s, left + Inches(0.16), Inches(5.12), Inches(2.7), Inches(0.42), value, size=16, bold=True, color=NAVY)

    add_text_box(
        s,
        Inches(0.5),
        Inches(5.9),
        Inches(12.3),
        Inches(1.1),
        f"EAMP proposal: {clean_text(meta.get('proposal_link'))}\n"
        f"Reseller services bill-to: World Wide Technology, Inc. (1001363275)   ·   Duration {duration_label(duration)}   ·   Billing model {billing}",
        size=12,
        color=MUTED,
    )

    # ----- 4. Why EA -----
    s = blank_slide(prs)
    header_bar(s, "Why a Cisco Enterprise Agreement", "What this structure does for Hallmark")
    pillars = [
        (CISCO, "One enterprise term", "A single 5-year agreement covering Networking Infrastructure, Security, and Cisco CX services — one Smart Account, one book date, one True Forward calendar."),
        (TEAL, "Price protection", "Full Commit suites carry not-to-exceed unit pricing for the remaining suite term. Partial Commit and add-on suites keep a fixed discount against then-current list."),
        (GREEN, "True Forward, not true-up", "Growth is measured on the EA schedule and billed going forward only. There is no retroactive catch-up invoice for licenses already in use."),
        (ORANGE, "Room to grow", "ISE (Security) includes a 15% growth allowance after the first six months. Networking Full Commit suites true-forward annually rather than quarterly."),
        (GOLD, "Value shift", "Unused license value can be moved within a suite — and between eligible suites — at True Forward, so the estate can be right-sized without buying twice."),
        (NAVY_MID, "Software + CX together", "Lifecycle / Smart Net-style coverage is enrolled as EA services suites against the underlying software, with CX PIDs co-termed to the same 60-month term."),
    ]
    for i, (color, title, blurb) in enumerate(pillars):
        col = i % 3
        row = i // 3
        left = Inches(0.45) + col * Inches(4.2)
        top = Inches(1.1) + row * Inches(2.85)
        add_round(s, left, top, Inches(4.0), Inches(2.65), WHITE, LINE)
        add_rect(s, left, top, Inches(4.0), Inches(0.10), color)
        add_text_box(s, left + Inches(0.22), top + Inches(0.28), Inches(3.55), Inches(0.5), title, size=18, bold=True, color=NAVY)
        add_text_box(s, left + Inches(0.22), top + Inches(0.85), Inches(3.55), Inches(1.55), blurb, size=13, color=INK)

    # ----- 5. Investment at a glance -----
    s = blank_slide(prs)
    header_bar(s, "Investment at a glance", f"{pricing_note}  ·  {years}-year term  ·  {billing.lower()} billing")
    cards = [
        ("Total contract value", money(tcv), f"{money(annual)} per year", CISCO),
        ("Software TCV", money(sw), f"{money(annual_from_tcv(sw, duration))} / year", NAVY_MID),
        ("Services TCV", money(svcs), f"{money(annual_from_tcv(svcs, duration))} / year", TEAL),
        ("Effective savings vs list", pct(discount_pct, 1), f"{money(list_price - tcv)} below list", GREEN),
    ]
    for i, (label, value, caption, accent) in enumerate(cards):
        kpi_card(s, Inches(0.4) + i * Inches(3.2), Inches(1.1), Inches(3.05), Inches(1.28), label, value, caption, accent)

    # waterfall-style comparison
    add_text_box(s, Inches(0.5), Inches(2.55), Inches(12), Inches(0.35), "From list price to customer TCV", size=16, bold=True, color=NAVY)
    headers = ["", "Amount (USD)"]
    rows = [
        ["Total list price", money(list_price, 2)],
        ["Cisco program / subscription discounts", "− " + money(eamp_discount, 2)],
        ["One-time discounts (OTD / IB / ramp / migration)", "− " + money(otd, 2)],
        ["EAMP net (Cisco)", money(eamp_tcv, 2)],
        [f"Partner margin ({margin_points} points)", "+ " + money(partner_margin, 2)],
        ["Customer TCV", money(tcv, 2)],
    ]
    add_table(s, Inches(0.5), Inches(2.88), Inches(7.6), Inches(3.05), headers, rows, col_widths=[Inches(5.4), Inches(2.2)], numeric_cols={1}, font_size=12)

    add_round(s, Inches(8.35), Inches(2.88), Inches(4.45), Inches(3.05), NAVY)
    add_text_box(s, Inches(8.55), Inches(3.05), Inches(4.05), Inches(0.28), "ANNUAL CASH VIEW", size=11, bold=True, color=CISCO)
    add_text_box(s, Inches(8.55), Inches(3.35), Inches(4.05), Inches(0.5), money(annual), size=26, bold=True, color=WHITE)
    add_text_box(
        s,
        Inches(8.55),
        Inches(3.9),
        Inches(4.05),
        Inches(1.8),
        f"Five equal {billing.lower()} payments over {int(duration)} months, before any future True Forward.\n\n"
        f"Software {money(annual_from_tcv(sw, duration))}/yr\nServices {money(annual_from_tcv(svcs, duration))}/yr\n"
        f"Includes {margin_points} pts partner margin ({money(partner_margin)} on the 5-year TCV).",
        size=12,
        color=RGBColor(0xD5, 0xE4, 0xF0),
    )

    add_text_box(
        s,
        Inches(0.5),
        Inches(6.05),
        Inches(12.3),
        Inches(1.05),
        f"Customer TCV replaces EAMP net and includes {margin_points} points of partner margin (20% gross margin; sell = EAMP net ÷ 0.80). "
        "Cisco list, subscription %, and OTDs are unchanged. Networking software is 10% programmatic + 70% subscription; ISE is 68%; CX is ~37% plus migration OTDs.",
        size=13,
        color=MUTED,
    )

    # ----- 6. Financial summary by portfolio -----
    s = blank_slide(prs)
    header_bar(s, "Financial summary by portfolio", f"{pricing_note}  ·  Cisco list and OTDs unchanged")
    headers = ["Portfolio", "Software", "Services", "Customer TCV", "List price", "Discount", "OTD", "Net vs list"]
    table_rows = []
    chart_cats = []
    chart_vals = []
    chart_colors = []
    for row in financial:
        name = row["portfolio"]
        vs_list = (1 - row["total"] / row["list_price"]) * 100 if row["list_price"] else 0
        table_rows.append(
            [
                name,
                money(row["software"], 2),
                money(row["services"], 2),
                money(row["total"], 2),
                money(row["list_price"], 2),
                money(row["discount_amount"], 2),
                money(row["one_time_discount"], 2),
                pct(vs_list, 1),
            ]
        )
        if name.upper() != "TOTAL" and row["total"] > 0:
            chart_cats.append(name)
            chart_vals.append(row["total"])
            chart_colors.append(PORTFOLIO_COLORS.get(name, CISCO))
    add_table(
        s,
        Inches(0.35),
        Inches(1.05),
        Inches(12.6),
        Inches(2.35),
        headers,
        table_rows,
        col_widths=[Inches(2.5), Inches(1.45), Inches(1.45), Inches(1.5), Inches(1.55), Inches(1.4), Inches(1.3), Inches(1.45)],
        numeric_cols={1, 2, 3, 4, 5, 6, 7},
        font_size=11,
    )

    chart_data = CategoryChartData()
    chart_data.categories = chart_cats
    chart_data.add_series("Net TCV", chart_vals)
    chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.35), Inches(3.55), Inches(7.7), Inches(3.45), chart_data).chart
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.number_format = '"$"#,##0'
    plot.data_labels.font.size = Pt(10)
    plot.data_labels.font.bold = True
    plot.data_labels.font.color.rgb = NAVY
    series = chart.series[0]
    for i, color in enumerate(chart_colors):
        try:
            series.points[i].format.fill.solid()
            series.points[i].format.fill.fore_color.rgb = color
        except Exception:
            pass
    chart.category_axis.tick_labels.font.size = Pt(11)
    chart.category_axis.tick_labels.font.color.rgb = INK
    chart.value_axis.tick_labels.font.size = Pt(10)
    chart.value_axis.tick_labels.number_format = '"$"#,##0'
    chart.value_axis.has_major_gridlines = True

    add_round(s, Inches(8.2), Inches(3.55), Inches(4.7), Inches(3.45), WHITE, LINE)
    add_text_box(s, Inches(8.4), Inches(3.7), Inches(4.3), Inches(0.35), "Where the net TCV sits", size=16, bold=True, color=NAVY)
    notes = [
        f"Networking software is {pct(sw / tcv * 100 if tcv else 0, 0)} of TCV — Catalyst DNA, Nexus, Meraki, Spaces, and Cisco Networking suites.",
        f"Cisco CX / lifecycle services are {pct(svcs / tcv * 100 if tcv else 0, 0)} of TCV, covering Nexus, Catalyst switching, and Collaboration installed base.",
        f"Security (ISE Advantage + Essentials) is {money(next((r['total'] for r in financial if r['portfolio']=='Security'), 0))} net for 4,550 desired seats.",
    ]
    for i, note in enumerate(notes):
        add_text_box(s, Inches(8.4), Inches(4.15) + i * Inches(0.85), Inches(4.3), Inches(0.8), "•  " + note, size=13, color=INK)

    # ----- 7. Included portfolios -----
    s = blank_slide(prs)
    header_bar(s, "What's in this Enterprise Agreement", "Portfolios with committed suites in proposal " + proposal_id)
    tiles = [
        ("Networking Infrastructure", True, "Meraki, Catalyst / DNA, Nexus, Cisco Networking, Spaces"),
        ("Security", True, "ISE Advantage + Essentials (Zero Trust), 1,000 SCU"),
        ("Services (CX)", True, "Nexus, Catalyst switching, Collaboration lifecycle"),
        ("Collaboration software", False, "No software suite in this estimate; CX coverage only"),
        ("Applications Infrastructure", False, "Not enrolled in this proposal"),
        ("Hybrid / Provider / CPS", False, "Not enrolled in this proposal"),
    ]
    for i, (title, included, blurb) in enumerate(tiles):
        col = i % 3
        row = i // 3
        left = Inches(0.45) + col * Inches(4.2)
        top = Inches(1.15) + row * Inches(2.75)
        add_round(s, left, top, Inches(4.0), Inches(2.5), WHITE, LINE)
        badge = "INCLUDED" if included else "NOT IN THIS ESTIMATE"
        badge_color = GREEN if included else MUTED
        add_round(s, left + Inches(0.22), top + Inches(0.22), Inches(2.15), Inches(0.32), badge_color)
        add_text_box(s, left + Inches(0.22), top + Inches(0.24), Inches(2.15), Inches(0.28), badge, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, left + Inches(0.22), top + Inches(0.7), Inches(3.55), Inches(0.7), title, size=18, bold=True, color=NAVY)
        add_text_box(s, left + Inches(0.22), top + Inches(1.4), Inches(3.55), Inches(0.85), blurb, size=13, color=INK)

    # ----- 8. Networking suite rollup -----
    net = next((p for p in data["portfolios"] if p["name"] == "Networking Infrastructure"), None)
    if net and net["software"]:
        s = blank_slide(prs)
        header_bar(s, "Networking Infrastructure — software suites", f"Portfolio software TCV  {money(net['header'].get('software_tcv') or net['header'].get('software_line_tcv'))}")
        rollups = suite_rollups(net["software"])
        headers = ["Solution", "Suite", "Commit", "Lines", "Qty IB", "Qty desired", "List price", "Net TCV"]
        rows = []
        for r in rollups:
            if r["tcv"] == 0 and r["list_price"] == 0:
                continue
            rows.append(
                [
                    r["solution"],
                    r["suite"],
                    r["commitment"] or "—",
                    str(r["lines"]),
                    qty(r["qty_ib"]),
                    qty(r["qty_desired"]),
                    money(r["list_price"]),
                    money(r["tcv"]),
                ]
            )
        rows.append(
            [
                "TOTAL",
                "",
                "",
                str(sum(r["lines"] for r in rollups)),
                "—",
                "—",
                money(sum(r["list_price"] for r in rollups)),
                money(sum(r["tcv"] for r in rollups)),
            ]
        )
        add_table(
            s,
            Inches(0.3),
            Inches(1.05),
            Inches(12.7),
            None,
            headers,
            rows,
            col_widths=[Inches(2.15), Inches(3.15), Inches(1.0), Inches(0.8), Inches(1.05), Inches(1.25), Inches(1.65), Inches(1.65)],
            numeric_cols={3, 4, 5, 6, 7},
            font_size=11,
        )
        add_text_box(
            s,
            Inches(0.35),
            Inches(5.55),
            Inches(12.5),
            Inches(0.55),
            "Quantities are licensed assets only (zero-dollar support-embed PIDs are excluded). Suite quantity totals are not summed across mixed unit types (APs, switches, cameras, MX).",
            size=12,
            color=MUTED,
        )

        # Highlight cards for biggest suites
        s = blank_slide(prs)
        header_bar(s, "Networking — where software spend concentrates", f"Customer TCV of the four largest suites  ·  {margin_points} pts margin")
        top4 = sorted([r for r in rollups if r["tcv"] > 0], key=lambda r: r["tcv"], reverse=True)[:4]
        palettes = [CISCO, NAVY_MID, TEAL, ORANGE]
        for i, r in enumerate(top4):
            left = Inches(0.4) + i * Inches(3.2)
            add_round(s, left, Inches(1.15), Inches(3.05), Inches(2.55), WHITE, LINE)
            add_rect(s, left, Inches(1.15), Inches(3.05), Inches(0.1), palettes[i])
            add_text_box(s, left + Inches(0.18), Inches(1.38), Inches(2.7), Inches(0.55), r["suite"], size=14, bold=True, color=NAVY)
            add_text_box(s, left + Inches(0.18), Inches(1.95), Inches(2.7), Inches(0.5), money(r["tcv"]), size=24, bold=True, color=palettes[i])
            share = r["tcv"] / sw * 100 if sw else 0
            add_text_box(
                s,
                left + Inches(0.18),
                Inches(2.5),
                Inches(2.7),
                Inches(0.95),
                f"{pct(share, 0)} of software TCV\n{r['commitment']} commit  ·  {r['lines']} SKUs\nDesired qty {qty(r['qty_desired'])}  (IB {qty(r['qty_ib'])})",
                size=12,
                color=MUTED,
            )
        c9500 = next((i for i in net["software"] if i.get("part_number") == "E3N-C95005-A"), {})
        callouts = [
            ("Catalyst DNA Switching", f"Core campus/core switching software. C9500 48Y4C Advantage is the largest single SKU in the proposal at {money(c9500.get('tcv'))} customer TCV / {qty(c9500.get('qty_desired'))} desired."),
            ("Nexus Switching", "Data-center fabric: 30 × Nexus 9300 XF Advantage, 2 × XF2 Advantage, 4 × XF Essential. Full commit with Standard DCN support attached."),
            ("DNA Wireless + Spaces", "Wireless entitlement steps up from 132 IB to 525 desired APs, plus 350 Cisco Spaces ACT add-on licenses (partial commit)."),
            ("Meraki + Cisco Networking", "Meraki MX/MR/MS/MV plus Cisco Networking Switching Access Tier 1 (49 large / 42 medium Essential) for the evolving campus edge."),
        ]
        for i, (title, blurb) in enumerate(callouts):
            top = Inches(3.9) + (i % 2) * Inches(1.45)
            left = Inches(0.4) + (i // 2) * Inches(6.4)
            add_round(s, left, top, Inches(6.2), Inches(1.32), OFF_WHITE, LINE)
            add_text_box(s, left + Inches(0.2), top + Inches(0.12), Inches(5.8), Inches(0.32), title, size=14, bold=True, color=NAVY)
            add_text_box(s, left + Inches(0.2), top + Inches(0.46), Inches(5.8), Inches(0.74), blurb, size=12, color=INK)

    def sku_table_slides(title, subtitle, items):
        headers = ["Part number", "Description", "Commit", "Tier", "Qty IB", "Qty", "List", "Disc %", "TCV"]
        prepared = []
        for item in items:
            if not item.get("part_number"):
                continue
            prepared.append(
                [
                    str(item.get("part_number") or ""),
                    str(item.get("description") or ""),
                    str(item.get("commitment") or "—"),
                    str(item.get("license_tier") or item.get("service_tier") or "—"),
                    qty(item.get("qty_ib")),
                    qty(item.get("qty_desired")),
                    money(item.get("list_price")),
                    pct(item.get("subscription_discount_pct"), 0),
                    money(item.get("tcv")),
                ]
            )
        widths = [Inches(1.7), Inches(3.55), Inches(0.85), Inches(1.05), Inches(0.9), Inches(0.85), Inches(1.25), Inches(0.85), Inches(1.35)]
        page_size = 14
        for page_i, chunk in enumerate(paginate(prepared, page_size), start=1):
            sl = blank_slide(prs)
            page_count = max(1, (len(prepared) + page_size - 1) // page_size)
            suffix = f"  ·  {page_i} of {page_count}" if page_count > 1 else ""
            header_bar(sl, title, subtitle + suffix)
            rows = chunk
            if page_i == page_count:
                rows = chunk + [
                    [
                        "TOTAL",
                        "",
                        "",
                        "",
                        licensed_qty(items, "qty_ib"),
                        licensed_qty(items, "qty_desired"),
                        money(sum(float(i.get("list_price") or 0) for i in items)),
                        "",
                        money(sum(float(i.get("tcv") or 0) for i in items)),
                    ]
                ]
            add_table(
                sl,
                Inches(0.25),
                Inches(1.0),
                Inches(12.8),
                None,
                headers,
                rows,
                col_widths=widths,
                numeric_cols={4, 5, 6, 7, 8},
                font_size=10,
            )

    def emit_sku_appendix():
        if not net:
            return
        meraki = [i for i in net["software"] if "meraki" in str(i.get("suite") or "").lower()]
        if meraki:
            sku_table_slides(
                "Appendix — Meraki network & cameras",
                "Full commit on Network Infrastructure  ·  Partial commit on Camera Systems",
                meraki,
            )
        cisco_net = [
            i
            for i in net["software"]
            if str(i.get("suite") or "") in ("Cisco Networking Wireless", "Cisco Networking Switching")
        ]
        if cisco_net:
            sku_table_slides(
                "Appendix — Cisco Networking switching & wireless",
                "Full commit  ·  Access Tier 1 Essential plus wireless Advantage/Essential",
                cisco_net,
            )
        nexus = [i for i in net["software"] if "nexus" in str(i.get("suite") or "").lower()]
        if nexus:
            sku_table_slides(
                "Appendix — Nexus switching",
                "Full commit  ·  60-month term with Standard Data Center Networking support",
                nexus,
            )
        dna_sw = [i for i in net["software"] if str(i.get("suite") or "") == "Cisco DNA Switching"]
        if dna_sw:
            sku_table_slides(
                "Appendix — Cisco DNA Switching",
                "Full commit  ·  Catalyst 9200 / 9300 / 9300X / 9500 families",
                dna_sw,
            )
        dna_wl = [
            i
            for i in net["software"]
            if str(i.get("suite") or "") in ("Cisco DNA Wireless", "Cisco Spaces Add on for Wireless")
        ]
        if dna_wl:
            sku_table_slides(
                "Appendix — Cisco DNA Wireless & Spaces",
                "DNA Wireless is Full commit  ·  Spaces ACT is Partial / add-on",
                dna_wl,
            )

        if net["services"]:
            s = blank_slide(prs)
            header_bar(
                s,
                "Appendix — Networking CX / lifecycle services",
                f"Portfolio services TCV  {money(net['header'].get('services_tcv') or net['header'].get('services_line_tcv'))}",
            )
            headers = ["Suite", "Part number", "Description", "IB", "Disc %", "List price", "Migration OTD", "Net TCV"]
            rows = []
            for item in net["services"]:
                if not item.get("part_number"):
                    continue
                rows.append(
                    [
                        str(item.get("suite") or ""),
                        str(item.get("part_number") or ""),
                        str(item.get("description") or ""),
                        qty(item.get("qty_ib")),
                        pct(item.get("subscription_discount_pct"), 1),
                        money(item.get("list_price")),
                        money(item.get("otd_program_migration")),
                        money(item.get("tcv")),
                    ]
                )
            rows.append(
                [
                    "TOTAL",
                    "",
                    "",
                    "",
                    "",
                    money(sum(float(i.get("list_price") or 0) for i in net["services"])),
                    money(sum(float(i.get("otd_program_migration") or 0) for i in net["services"])),
                    money(sum(float(i.get("tcv") or 0) for i in net["services"])),
                ]
            )
            add_table(
                s,
                Inches(0.25),
                Inches(1.0),
                Inches(12.8),
                None,
                headers,
                rows,
                col_widths=[Inches(2.55), Inches(1.85), Inches(3.15), Inches(0.7), Inches(0.85), Inches(1.25), Inches(1.2), Inches(1.25)],
                numeric_cols={3, 4, 5, 6, 7},
                font_size=10,
            )
            cx_enbd = next((i for i in net["services"] if i.get("part_number") == "E3-CX-CS-ENBD"), {})
            cx_t14 = next((i for i in net["services"] if i.get("part_number") == "E3-CX-DCN-T14HR"), {})
            otd_sum = sum(float(i.get("otd_program_migration") or 0) for i in net["services"])
            add_text_box(
                s,
                Inches(0.35),
                Inches(6.4),
                Inches(12.5),
                Inches(0.65),
                f"Largest CX lines: Catalyst switching 8x5xNBD ({money(cx_enbd.get('tcv'))} customer TCV, {qty(cx_enbd.get('qty_ib'))} IB) and Nexus 24x7x4 ({money(cx_t14.get('tcv'))} customer TCV, {qty(cx_t14.get('qty_ib'))} IB). "
                f"Program-migration OTDs ({money(otd_sum)} combined) credit the move onto EA services. Zero-dollar wireless/switching SW-only PIDs enroll coverage without incremental TCV.",
                size=12,
                color=MUTED,
            )

    # ----- Security -----
    sec = next((p for p in data["portfolios"] if p["name"] == "Security"), None)
    if sec and sec["software"]:
        s = blank_slide(prs)
        header_bar(s, "Security — Zero Trust / ISE", f"Software TCV  {money(sec['header'].get('software_tcv') or sec['header'].get('software_line_tcv'))}   ·   SCU count  {clean_text(sec['header'].get('scu_count'))}")
        ise = [i for i in sec["software"] if i.get("part_number")]
        # KPI row
        adv = next((i for i in ise if i.get("part_number") == "E3S-ISE-ADV"), {})
        ess = next((i for i in ise if i.get("part_number") == "E3S-ISE-ESS"), {})
        kpis = [
            ("ISE Advantage", qty(adv.get("qty_desired")), f"IB {qty(adv.get('qty_ib'))}  ·  {money(adv.get('tcv'))} TCV"),
            ("ISE Essentials", qty(ess.get("qty_desired")), f"IB {qty(ess.get('qty_ib'))}  ·  {money(ess.get('tcv'))} TCV"),
            ("Subscription discount", "68%", "Full commit  ·  no programmatic %"),
            ("SCU count", clean_text(sec["header"].get("scu_count")), "Security Content Users on this EA"),
        ]
        for i, (label, value, caption) in enumerate(kpis):
            kpi_card(s, Inches(0.4) + i * Inches(3.2), Inches(1.1), Inches(3.05), Inches(1.28), label, value, caption, ORANGE if i < 2 else CISCO)

        headers = ["Part number", "Description", "Commit", "Qty IB", "Qty desired", "Unit list / mo", "List price", "OTD (IB)", "Net TCV"]
        rows = []
        for item in ise:
            rows.append(
                [
                    str(item.get("part_number") or ""),
                    str(item.get("description") or ""),
                    str(item.get("commitment") or "—"),
                    qty(item.get("qty_ib")),
                    qty(item.get("qty_desired")),
                    money(item.get("unit_list_month"), 2),
                    money(item.get("list_price"), 2),
                    money(item.get("otd_ib"), 2),
                    money(item.get("tcv"), 2),
                ]
            )
        rows.append(
            [
                "TOTAL",
                "",
                "",
                licensed_qty(ise, "qty_ib"),
                licensed_qty(ise, "qty_desired"),
                "",
                money(sum(float(i.get("list_price") or 0) for i in ise), 2),
                money(sum(float(i.get("otd_ib") or 0) for i in ise), 2),
                money(sum(float(i.get("tcv") or 0) for i in ise), 2),
            ]
        )
        add_table(
            s,
            Inches(0.3),
            Inches(2.6),
            Inches(12.7),
            None,
            headers,
            rows,
            col_widths=[Inches(1.55), Inches(2.35), Inches(0.9), Inches(1.0), Inches(1.25), Inches(1.4), Inches(1.45), Inches(1.3), Inches(1.5)],
            numeric_cols={3, 4, 5, 6, 7, 8},
            font_size=11,
        )
        add_round(s, Inches(0.3), Inches(5.3), Inches(12.7), Inches(1.7), OFF_WHITE, LINE)
        add_text_box(s, Inches(0.5), Inches(5.45), Inches(12.3), Inches(0.3), "What this enrollment does", size=14, bold=True, color=NAVY)
        add_text_box(
            s,
            Inches(0.5),
            Inches(5.8),
            Inches(12.3),
            Inches(1.05),
            "Hallmark is stepping ISE Advantage from 400 installed-base seats to 3,450 desired, and Essentials from 300 to 1,100. "
            "IB adjustment OTDs ($1,719 + $270) prevent double-pay on the overlapping installed base. "
            "As a Full Commit Security suite, ISE receives not-to-exceed pricing for the remaining term and a 15% growth allowance after the first six months. "
            "SVS-E3S-ISE-B enrolls Cisco Support Basic for ISE at $0 incremental TCV.",
            size=13,
            color=INK,
        )

    # ----- Collaboration services -----
    colab = next((p for p in data["portfolios"] if p["name"] == "Collaboration"), None)
    if colab and (colab["software"] or colab["services"]):
        s = blank_slide(prs)
        header_bar(s, "Collaboration — CX coverage only", "No Collaboration software suite in this estimate")
        headers = ["Part number", "Description", "IB", "Qty", "Disc %", "List price", "Migration OTD", "Net TCV"]
        rows = []
        for item in colab["services"]:
            if not item.get("part_number"):
                continue
            rows.append(
                [
                    str(item.get("part_number") or ""),
                    str(item.get("description") or ""),
                    qty(item.get("qty_ib")),
                    qty(item.get("qty_desired")),
                    pct(item.get("subscription_discount_pct"), 1),
                    money(item.get("list_price"), 2),
                    money(item.get("otd_program_migration"), 2),
                    money(item.get("tcv"), 2),
                ]
            )
        rows.append(
            [
                "TOTAL",
                "",
                "",
                "",
                "",
                money(sum(float(i.get("list_price") or 0) for i in colab["services"]), 2),
                money(sum(float(i.get("otd_program_migration") or 0) for i in colab["services"]), 2),
                money(sum(float(i.get("tcv") or 0) for i in colab["services"]), 2),
            ]
        )
        add_table(
            s,
            Inches(0.3),
            Inches(1.05),
            Inches(12.7),
            None,
            headers,
            rows,
            col_widths=[Inches(1.9), Inches(3.7), Inches(0.8), Inches(0.7), Inches(1.0), Inches(1.55), Inches(1.55), Inches(1.5)],
            numeric_cols={2, 3, 4, 5, 6, 7},
            font_size=12,
        )
        add_round(s, Inches(0.3), Inches(4.0), Inches(12.7), Inches(2.9), WHITE, LINE)
        add_text_box(s, Inches(0.55), Inches(4.2), Inches(12.2), Inches(0.35), "How to read this", size=16, bold=True, color=NAVY)
        add_text_box(
            s,
            Inches(0.55),
            Inches(4.65),
            Inches(12.2),
            Inches(2.0),
            f"Collaboration software TCV is $0 in this export. The {money(next((i.get('tcv') for i in colab['services'] if i.get('part_number')=='E3-CX-COL-ENBD'), 0))} customer TCV is Cisco Support Standard 8x5xNBD for Collaboration "
            f"(E3-CX-COL-ENBD) against 145 installed-base assets, plus a {money(next((i.get('tcv') for i in colab['services'] if i.get('part_number')=='E3-CX-COL-LAPS-ST'), 0), 2)} prior-period lapse charge. "
            f"A 36.9% subscription discount and a {money(sum(float(i.get('otd_program_migration') or 0) for i in colab['services']))} program-migration OTD are already applied. "
            "8x7xNCD is enrolled for 25 IB assets at $0 TCV. If Hallmark later adds Webex / Calling / CX suites, they can attach to this same EA term.",
            size=14,
            color=INK,
        )

    # ----- Commercial terms -----
    s = blank_slide(prs)
    header_bar(s, "Commercial terms", "Term, billing, True Forward, and buying path")
    terms = [
        ("Term", duration_label(duration), "Requested ship / expected book 15 Dec 2026. All enrolled suites co-term to this EA."),
        ("Billing model", str(billing), "Annual invoices for software and services TCV. Capital financing is not used on this estimate."),
        ("True Forward", "Annual on Full Commit", "Partial Commit (Meraki cameras, Spaces ACT) follows the EA TF calendar. Growth is billed going forward only."),
        ("Buying path", "2-tier via WWT", "World Wide Technology, Inc. is the EA partner (bill-to 1001363275). Price list: Global Price List US Availability."),
        ("Smart Account", customer_name, f"Buying Program {clean_text(meta.get('buying_program_id'))}. Motion: {clean_text(meta.get('ea_motion'))}."),
        ("Pricing status", f"{margin_points} pts margin", "Customer TCV replaces EAMP net at 20% gross margin (sell = EAMP ÷ 0.80). Indicative only — convert in CCW with the EA Smart Account before order."),
    ]
    for i, (kicker, title, blurb) in enumerate(terms):
        col = i % 3
        row = i // 3
        left = Inches(0.4) + col * Inches(4.25)
        top = Inches(1.1) + row * Inches(2.8)
        add_round(s, left, top, Inches(4.05), Inches(2.55), WHITE, LINE)
        add_text_box(s, left + Inches(0.22), top + Inches(0.18), Inches(3.6), Inches(0.25), kicker.upper(), size=11, bold=True, color=CISCO)
        add_text_box(s, left + Inches(0.22), top + Inches(0.48), Inches(3.6), Inches(0.55), title, size=18, bold=True, color=NAVY)
        add_text_box(s, left + Inches(0.22), top + Inches(1.15), Inches(3.6), Inches(1.15), blurb, size=13, color=INK)

    # ----- Next steps -----
    s = blank_slide(prs)
    header_bar(s, "Recommended next steps", "Path from this EAMP estimate to a bookable EA")
    steps = [
        ("1", "Validate installed base", "Confirm IB quantities that drive DNA Switching, Nexus, ISE, and CX lines. Several Catalyst SKUs show IB above desired (right-sizing); ISE and DNA Wireless show the reverse (growth)."),
        ("2", "Lock suite mix", "Keep Full Commit on the networking and ISE suites that qualify Hallmark for NTE pricing. Treat Meraki cameras and Spaces as intentional Partial / add-on."),
        ("3", "Convert EAMP → CCW", "Build the CCW quote from proposal 9343374, assign Smart Account HALLMARK CARDS, INCORPORATED and the EA virtual account, and apply Buying Program Awareness on follow-on hardware."),
        ("4", "Legal / EA terms", "Execute EA Program Terms (or confirm they already apply) so subsequent suite purchases inherit this term, TF date, and price protection."),
        ("5", "Book by 15 Dec 2026", "Align WWT, Cisco AM Robin Randolph, and Hallmark procurement to the expected book date and 15 Dec requested ship date."),
        ("6", "Operate the EA", "Use EAMP for consumption visibility, annual True Forward, and value-shift decisions rather than one-off quotes for in-suite growth."),
    ]
    for i, (num, title, blurb) in enumerate(steps):
        top = Inches(1.05) + i * Inches(0.95)
        add_round(s, Inches(0.45), top, Inches(12.4), Inches(0.85), WHITE, LINE)
        add_round(s, Inches(0.65), top + Inches(0.17), Inches(0.52), Inches(0.52), CISCO)
        add_text_box(s, Inches(0.65), top + Inches(0.24), Inches(0.52), Inches(0.4), num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, Inches(1.4), top + Inches(0.08), Inches(11.2), Inches(0.28), title, size=16, bold=True, color=NAVY)
        add_text_box(s, Inches(1.4), top + Inches(0.38), Inches(11.2), Inches(0.4), blurb, size=12, color=INK)

    # ----- Appendix divider -----
    s = blank_slide(prs)
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, CISCO)
    add_text_box(s, Inches(0.8), Inches(2.6), Inches(11.5), Inches(0.4), "APPENDIX", size=16, bold=True, color=CISCO)
    add_text_box(s, Inches(0.8), Inches(3.05), Inches(11.5), Inches(0.8), "SKU-level detail from the EAMP export", size=32, bold=True, color=WHITE)
    add_text_box(
        s,
        Inches(0.8),
        Inches(4.0),
        Inches(11.5),
        Inches(0.8),
        "Figures match the 21 Nov 2026 Hallmark EAMP Price Estimate (ID 9343374). Indicative only.",
        size=16,
        color=RGBColor(0xB8, 0xC7, 0xD6),
    )

    emit_sku_appendix()

    s = blank_slide(prs)
    header_bar(s, "Appendix — how to read the numbers", "Column definitions from the EAMP export")
    defs = [
        ("Qty Found (IB)", "Installed-base licenses Cisco currently associates to the Hallmark Smart Account."),
        ("Qty Desired", "Entitlement requested for this 60-month EA. Growth vs IB is new paid consumption; IB below desired is expansion; IB above desired is right-sizing."),
        ("Unit list / month", "Monthly list before discounts, × 60 months × qty = Total list price."),
        ("Programmatic %", "EA multi-suite / offer discount (10% on most Networking Full Commit lines; 0% on ISE and Spaces)."),
        ("Subscription %", "Primary EA software discount (70% Networking, 68% ISE) or CX services discount (~36.9%)."),
        ("OTD (IB adjustment)", "One-time credit so overlapping installed-base software is not paid twice."),
        ("OTD (Ramp)", "Phased-quantity credit on selected Cisco Networking wireless/switching lines."),
        ("OTD (Program migration)", "CX credit for moving coverage onto the EA services suite."),
        ("TCV", "Customer 60-month contract value: EAMP net with 20 points of partner margin (sell = EAMP ÷ 0.80). Annual invoice ≈ TCV ÷ 5."),
        ("Full vs Partial commit", "Full Commit: NTE unit price and annual TF. Partial / add-on: fixed discount, typically quarterly TF. Cameras and Spaces are Partial in this file."),
    ]
    for i, (title, blurb) in enumerate(defs):
        col = i % 2
        row = i // 2
        left = Inches(0.4) + col * Inches(6.4)
        top = Inches(1.05) + row * Inches(1.1)
        add_round(s, left, top, Inches(6.2), Inches(1.0), WHITE, LINE)
        add_text_box(s, left + Inches(0.18), top + Inches(0.08), Inches(5.85), Inches(0.26), title, size=13, bold=True, color=NAVY)
        add_text_box(s, left + Inches(0.18), top + Inches(0.36), Inches(5.85), Inches(0.56), blurb, size=12, color=INK)

    # Disclaimer
    s = blank_slide(prs)
    header_bar(s, "Disclaimer", "Cisco Confidential")
    add_round(s, Inches(0.5), Inches(1.15), Inches(12.3), Inches(5.7), WHITE, LINE)
    add_text_box(s, Inches(0.8), Inches(1.4), Inches(11.7), Inches(0.4), "Indicative pricing only", size=22, bold=True, color=NAVY)
    add_text_box(
        s,
        Inches(0.8),
        Inches(1.95),
        Inches(11.7),
        Inches(4.5),
        "This presentation is derived from the Cisco EAMP Price Estimate export for HALLMARK CARDS, INCORPORATED "
        f"(Proposal ID {proposal_id}, Deal ID {deal_id}, Quote ID {quote_id}), generated {clean_text(meta.get('report_created_on'))}.\n\n"
        f"Displayed software, services, and TCV figures replace EAMP net with customer pricing that includes {margin_points} points of partner margin "
        f"(20% gross margin; sell = EAMP net ÷ 0.80). Cisco list prices, subscription percentages, quantities, and OTDs are unchanged from the EAMP export. "
        f"Partner margin on this quote is {money(partner_margin)} over the 5-year term.\n\n"
        "This is NOT an approved Cisco quote. Pricing is indicative as of the date generated and is subject to change, "
        "to Cisco EA 3.0 Program Terms, and to CCW conversion.\n\n"
        "Installed-base quantities, desired quantities, and OTDs should be validated before order. "
        "True Forward charges for future growth are not included in the TCV shown here.\n\n"
        "Cisco Confidential. Do not distribute outside Hallmark Cards, World Wide Technology, and the Cisco account team "
        "without authorization.\n\n"
        f"Cisco Account Manager: {clean_text(meta.get('account_manager'))}",
        size=15,
        color=INK,
    )

    stamp_footers()
    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output))
    return output


def main(argv=None):
    parser = argparse.ArgumentParser(description="Build an Enterprise Agreement PPTX from an EAMP xlsx.")
    parser.add_argument(
        "--xlsx",
        default=str(HERE / "source" / "HALLMARK_CARDS_INCORPORATED_EAMP_Price_Estimate_ID_9343374.xlsx"),
    )
    parser.add_argument(
        "--out",
        default=str(HERE / "HALLMARK_CARDS_INCORPORATED_Enterprise_Agreement_Proposal.pptx"),
    )
    parser.add_argument(
        "--margin",
        type=float,
        default=DEFAULT_MARGIN_POINTS,
        help="Partner gross-margin points applied to EAMP net (default: 20).",
    )
    args = parser.parse_args(argv)
    data = apply_partner_margin(parse_eamp_workbook(args.xlsx), points=args.margin)
    path = build_presentation(data, Path(args.out))
    print(f"Wrote {path} ({path.stat().st_size:,} bytes)")
    total = find_total_row(data["financial_summary"])
    print(f"Customer: {data['meta'].get('smart_account')}")
    print(f"Proposal: {data['meta'].get('proposal_id')}  Deal: {data['meta'].get('deal_id')}")
    print(f"EAMP net: {money(total.get('eamp_total'), 2)}")
    print(f"Partner margin ({int(args.margin)} pts): {money(total.get('partner_margin'), 2)}")
    print(f"Customer TCV: {money(total.get('total'), 2)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
