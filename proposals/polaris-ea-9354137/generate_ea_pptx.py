#!/usr/bin/env python3
"""Build a Cisco Enterprise Agreement proposal PPTX from an EAMP export.

Customer price = Cisco net (Total Contract Value) + 20 points of margin,
i.e. Customer Price = Cisco Net x 1.20.
"""

from __future__ import annotations

from collections import defaultdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

HERE = Path(__file__).resolve().parent
XLSX = HERE / "POLARIS_INC_-_EAMP_Price_Estimate_-_ID_9354137.xlsx"
OUT = HERE / "POLARIS_INC_Enterprise_Agreement_Proposal.pptx"

MARGIN_POINTS = Decimal("20")
MARGIN_FACTOR = Decimal("1") + (MARGIN_POINTS / Decimal("100"))  # 1.20

# Widescreen 16:9
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

NAVY = RGBColor(0x0B, 0x1F, 0x3A)
NAVY_MID = RGBColor(0x14, 0x36, 0x5C)
TEAL = RGBColor(0x00, 0x9B, 0xD8)
TEAL_DK = RGBColor(0x00, 0x7A, 0xAB)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xF5, 0xF7, 0xFA)
ICE = RGBColor(0xE8, 0xF1, 0xF6)
ROW_ALT = RGBColor(0xF0, 0xF5, 0xF9)
GRAY = RGBColor(0x5B, 0x67, 0x75)
SLATE = RGBColor(0x2C, 0x3A, 0x4A)
MUTED = RGBColor(0x8A, 0x95, 0xA1)
RED_ACCENT = RGBColor(0xC0, 0x4A, 0x2A)
GREEN = RGBColor(0x1F, 0x7A, 0x4D)
HEADER_ROW = RGBColor(0x0B, 0x1F, 0x3A)
TOTAL_ROW = RGBColor(0x14, 0x36, 0x5C)
BORDER = "C5CDD6"


def D(value) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def money(value: Decimal) -> Decimal:
    return D(value).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def customer_of(net: Decimal) -> Decimal:
    return money(D(net) * MARGIN_FACTOR)


def margin_of(net: Decimal) -> Decimal:
    return money(customer_of(net) - money(net))


def usd(value: Decimal, cents: bool = True) -> str:
    n = money(value)
    if not cents and abs(n) >= Decimal("1000"):
        return f"${n:,.0f}"
    return f"${n:,.2f}"


def usd_short(value: Decimal) -> str:
    n = money(value)
    if abs(n) >= Decimal("1000000"):
        return f"${n / Decimal('1000000'):,.2f}M"
    if abs(n) >= Decimal("1000"):
        return f"${n / Decimal('1000'):,.1f}K"
    return f"${n:,.2f}"


def num(value) -> Decimal:
    return D(value)


def set_run(paragraph, text, size=12, bold=False, color=SLATE, name="Calibri", italic=False):
    paragraph.clear()
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = name
    return run


def add_text_box(slide, left, top, width, height, text, size=12, bold=False, color=SLATE,
                 align=PP_ALIGN.LEFT, font="Calibri", italic=False, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr",
                                         MSO_ANCHOR.BOTTOM: "b"}.get(anchor, "t"))
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p, text, size=size, bold=bold, color=color, name=font, italic=italic)
    return box


def add_rect(slide, left, top, width, height, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    shape.shadow.inherit = False
    return shape


def add_round(slide, left, top, width, height, fill, adjust=0.08):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    shape.shadow.inherit = False
    try:
        shape.adjustments[0] = adjust
    except Exception:
        pass
    return shape


def _set_cell_border(cell, color=BORDER, width="6350"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag == qn("a:tcBdr"):
            tcPr.remove(child)
    tcBdr = parse_xml(
        f"""<a:tcBdr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:lnL w="{width}" cap="flat" cmpd="sng" algn="ctr">
                <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                <a:prstDash val="solid"/>
            </a:lnL>
            <a:lnR w="{width}" cap="flat" cmpd="sng" algn="ctr">
                <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                <a:prstDash val="solid"/>
            </a:lnR>
            <a:lnT w="{width}" cap="flat" cmpd="sng" algn="ctr">
                <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                <a:prstDash val="solid"/>
            </a:lnT>
            <a:lnB w="{width}" cap="flat" cmpd="sng" algn="ctr">
                <a:solidFill><a:srgbClr val="{color}"/></a:solidFill>
                <a:prstDash val="solid"/>
            </a:lnB>
        </a:tcBdr>"""
    )
    tcPr.append(tcBdr)


def style_cell(cell, text, size=10, bold=False, color=SLATE, fill=None, align=PP_ALIGN.LEFT,
               font="Calibri"):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p, str(text), size=size, bold=bold, color=color, name=font)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.margin_left = Inches(0.06)
    cell.margin_right = Inches(0.06)
    cell.margin_top = Inches(0.03)
    cell.margin_bottom = Inches(0.03)
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.background()
    _set_cell_border(cell)


def add_table(slide, rows, cols, left, top, width, height):
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table
    table.horz_banding = False
    table.vert_banding = False
    return table


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def add_header(slide, title, subtitle=None):
    add_rect(slide, 0, 0, SLIDE_W, Inches(0.92), NAVY)
    add_rect(slide, 0, Inches(0.92), SLIDE_W, Inches(0.06), TEAL)
    add_text_box(slide, Inches(0.4), Inches(0.16), Inches(10.6), Inches(0.42),
                 title, size=22, bold=True, color=WHITE, font="Calibri")
    if subtitle:
        add_text_box(slide, Inches(0.4), Inches(0.54), Inches(10.6), Inches(0.32),
                     subtitle, size=12, color=TEAL, font="Calibri")
    add_text_box(slide, Inches(10.7), Inches(0.22), Inches(2.2), Inches(0.5),
                 "POLARIS INC  |  EA", size=10, bold=True, color=GOLD,
                 align=PP_ALIGN.RIGHT)


def add_footer(slide, page, total, confidential=True):
    add_rect(slide, 0, Inches(7.22), SLIDE_W, Inches(0.28), NAVY)
    left = "CISCO CONFIDENTIAL  ·  Indicative pricing as of 25 Nov 2026  ·  Not an approved Cisco quote"
    if not confidential:
        left = "Polaris Inc  ·  Cisco Enterprise Agreement  ·  Prepared by Trace3"
    add_text_box(slide, Inches(0.35), Inches(7.24), Inches(10.4), Inches(0.24),
                 left, size=9, color=RGBColor(0xC5, 0xD0, 0xDA))
    add_text_box(slide, Inches(11.4), Inches(7.24), Inches(1.55), Inches(0.24),
                 f"{page}  /  {total}", size=9, color=WHITE, align=PP_ALIGN.RIGHT)


def add_card(slide, left, top, width, height, label, value, caption=None, value_size=22):
    add_round(slide, left, top, width, height, WHITE, adjust=0.08)
    add_rect(slide, left, top, Inches(0.08), height, TEAL)
    add_text_box(slide, left + Inches(0.22), top + Inches(0.12), width - Inches(0.32), Inches(0.28),
                 label.upper(), size=10, bold=True, color=TEAL_DK)
    add_text_box(slide, left + Inches(0.22), top + Inches(0.38), width - Inches(0.32), Inches(0.46),
                 value, size=value_size, bold=True, color=NAVY)
    if caption:
        add_text_box(slide, left + Inches(0.22), top + Inches(0.86), width - Inches(0.32), Inches(0.28),
                     caption, size=10, color=GRAY)


def paint_content_bg(slide):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)


def parse_workbook(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    summary = wb["Proposal Summary"]
    ni = wb["Networking Infrastructure"]
    apps = wb["Applications Infrastructure"]

    meta = {
        "report_date": str(summary["B8"].value or ""),
        "deal_id": str(summary["B9"].value or ""),
        "quote_id": str(summary["B10"].value or ""),
        "project_name": str(summary["B11"].value or ""),
        "buying_program_id": str(summary["B12"].value or ""),
        "smart_account": str(summary["B13"].value or ""),
        "proposal_name": str(summary["B14"].value or ""),
        "proposal_id": str(int(summary["B15"].value)) if summary["B15"].value else "",
        "expected_book": str(summary["B16"].value or ""),
        "price_list": str(summary["B17"].value or ""),
        "cam": str(summary["G17"].value or "").strip(),
        "motion": str(summary["B18"].value or ""),
        "customer_block": str(summary["E9"].value or ""),
        "partner": str(summary["I9"].value or ""),
        "disti": str(summary["M9"].value or ""),
        "rsd": str(ni["B9"].value or ""),
        "billing": str(ni["B10"].value or ""),
        "duration_months": D(ni["B12"].value),
        "disti_bill_to": str(ni["B13"].value or ""),
        "reseller_bill_to": str(ni["B14"].value or ""),
    }
    meta["duration_years"] = meta["duration_months"] / Decimal("12")

    portfolios = []
    for row in range(25, 28):
        name = summary[f"A{row}"].value
        if not name:
            continue
        portfolios.append({
            "name": str(name),
            "software": money(D(summary[f"B{row}"].value)),
            "services": money(D(summary[f"C{row}"].value)),
            "total": money(D(summary[f"D{row}"].value)),
            "list": money(D(summary[f"E{row}"].value)),
            "discount": money(D(summary[f"F{row}"].value)),
            "otd": money(D(summary[f"G{row}"].value)),
        })
    totals = {
        "software": money(D(summary["B28"].value)),
        "services": money(D(summary["C28"].value)),
        "total": money(D(summary["D28"].value)),
        "list": money(D(summary["E28"].value)),
        "discount": money(D(summary["F28"].value)),
        "otd": money(D(summary["G28"].value)),
    }

    def parse_software(ws, start, end, tcv_col="U"):
        items = []
        for r in range(start, end + 1):
            pn = ws[f"D{r}"].value
            if not pn:
                continue
            item = {
                "solution": str(ws[f"A{r}"].value or ""),
                "suite": str(ws[f"B{r}"].value or ""),
                "commitment": str(ws[f"C{r}"].value or ""),
                "part": str(pn),
                "desc": str(ws[f"E{r}"].value or ""),
                "tier": str(ws[f"F{r}"].value or ""),
                "qty_ib": D(ws[f"G{r}"].value),
                "qty": D(ws[f"H{r}"].value),
                "prog_disc": D(ws[f"I{r}"].value),
                "unit_list_mo": D(ws[f"J{r}"].value),
                "list": money(D(ws[f"K{r}"].value)),
                "sub_disc": D(ws[f"L{r}"].value),
                "net": money(D(ws[f"{tcv_col}{r}"].value)),
            }
            item["customer"] = customer_of(item["net"])
            item["margin"] = margin_of(item["net"])
            items.append(item)
        return items

    def parse_services(ws, start, end):
        items = []
        for r in range(start, end + 1):
            pn = ws[f"D{r}"].value
            if not pn:
                continue
            item = {
                "solution": str(ws[f"A{r}"].value or ""),
                "suite": str(ws[f"B{r}"].value or ""),
                "commitment": str(ws[f"C{r}"].value or ""),
                "part": str(pn),
                "desc": str(ws[f"E{r}"].value or ""),
                "tier": str(ws[f"F{r}"].value or ""),
                "qty_ib": D(ws[f"G{r}"].value),
                "qty": D(ws[f"H{r}"].value),
                "list": money(D(ws[f"K{r}"].value)),
                "sub_disc": D(ws[f"L{r}"].value),
                "net": money(D(ws[f"R{r}"].value)),
            }
            item["customer"] = customer_of(item["net"])
            item["margin"] = margin_of(item["net"])
            items.append(item)
        return items

    ni_sw = parse_software(ni, 21, 59)
    ni_svc = parse_services(ni, 64, 88)
    app_sw = parse_software(apps, 20, 24)

    def rollup(items):
        out = defaultdict(lambda: {
            "list": Decimal("0"), "net": Decimal("0"), "customer": Decimal("0"),
            "margin": Decimal("0"), "qty": Decimal("0"), "count": 0,
        })
        for it in items:
            b = out[it["suite"]]
            b["list"] += it["list"]
            b["net"] += it["net"]
            b["customer"] += it["customer"]
            b["margin"] += it["margin"]
            b["qty"] += it["qty"]
            b["count"] += 1
        return dict(out)

    data = {
        "meta": meta,
        "portfolios": portfolios,
        "totals": totals,
        "ni_sw": ni_sw,
        "ni_svc": ni_svc,
        "app_sw": app_sw,
        "ni_sw_suites": rollup(ni_sw),
        "ni_svc_suites": rollup(ni_svc),
        "app_suites": rollup(app_sw),
    }

    # Customer-level portfolio numbers from line items so cents reconcile.
    data["cust"] = {
        "ni_sw_net": money(sum((i["net"] for i in ni_sw), Decimal("0"))),
        "ni_sw_list": money(sum((i["list"] for i in ni_sw), Decimal("0"))),
        "ni_sw_cust": money(sum((i["customer"] for i in ni_sw), Decimal("0"))),
        "ni_svc_net": money(sum((i["net"] for i in ni_svc), Decimal("0"))),
        "ni_svc_list": money(sum((i["list"] for i in ni_svc), Decimal("0"))),
        "ni_svc_cust": money(sum((i["customer"] for i in ni_svc), Decimal("0"))),
        "app_net": money(sum((i["net"] for i in app_sw), Decimal("0"))),
        "app_list": money(sum((i["list"] for i in app_sw), Decimal("0"))),
        "app_cust": money(sum((i["customer"] for i in app_sw), Decimal("0"))),
    }
    c = data["cust"]
    c["net"] = money(c["ni_sw_net"] + c["ni_svc_net"] + c["app_net"])
    c["list"] = money(c["ni_sw_list"] + c["ni_svc_list"] + c["app_list"])
    c["customer"] = money(c["ni_sw_cust"] + c["ni_svc_cust"] + c["app_cust"])
    c["margin"] = money(c["customer"] - c["net"])
    years = meta["duration_years"]
    c["annual_net"] = money(c["net"] / years)
    c["annual_cust"] = money(c["customer"] / years)
    c["ni_sw_annual"] = money(c["ni_sw_cust"] / years)
    c["ni_svc_annual"] = money(c["ni_svc_cust"] / years)
    c["app_annual"] = money(c["app_cust"] / years)
    return data


def chunk(seq, n):
    buf = []
    for item in seq:
        buf.append(item)
        if len(buf) == n:
            yield buf
            buf = []
    if buf:
        yield buf


def build_presentation(data: dict) -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    slides = []

    def slide():
        s = blank_slide(prs)
        slides.append(s)
        return s

    meta = data["meta"]
    c = data["cust"]
    years = meta["duration_years"]
    months = meta["duration_months"]
    term_label = f"{months.quantize(Decimal('0.1'))} months  (~{years.quantize(Decimal('0.01'))} years remaining)"

    # ------------------------------------------------------------------
    # 1. Cover
    # ------------------------------------------------------------------
    s = slide()
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, TEAL)
    add_rect(s, 0, Inches(6.55), SLIDE_W, Inches(0.08), GOLD)
    add_text_box(s, Inches(0.7), Inches(0.55), Inches(11.5), Inches(0.32),
                 "TRACE3  ·  CISCO PARTNER PROPOSAL  ·  CONFIDENTIAL",
                 size=12, bold=True, color=GOLD)
    add_text_box(s, Inches(0.7), Inches(1.35), Inches(11.8), Inches(0.7),
                 "Cisco Enterprise Agreement", size=36, bold=True, color=WHITE)
    add_text_box(s, Inches(0.7), Inches(2.05), Inches(11.8), Inches(0.55),
                 "Polaris Inc  —  EA Adder Pricing Proposal", size=24, color=TEAL)
    add_text_box(s, Inches(0.7), Inches(2.7), Inches(11.5), Inches(0.4),
                 meta["project_name"], size=16, color=RGBColor(0xC5, 0xD0, 0xDA))

    facts = [
        ("Proposal ID", meta["proposal_id"]),
        ("Deal / Quote", f"{meta['deal_id']}  /  {meta['quote_id']}"),
        ("Buying Program", meta["buying_program_id"]),
        ("Expected Book Date", "December 15, 2026"),
        ("Remaining Term", term_label),
        ("Billing Model", f"{meta['billing']}  ·  Global Price List US"),
    ]
    y = Inches(3.35)
    for i, (lab, val) in enumerate(facts):
        col = i % 2
        row = i // 2
        x = Inches(0.7) + col * Inches(6.1)
        yy = y + row * Inches(0.7)
        add_text_box(s, x, yy, Inches(5.6), Inches(0.22), lab.upper(), size=10, bold=True, color=TEAL)
        add_text_box(s, x, yy + Inches(0.22), Inches(5.6), Inches(0.32), val, size=14, color=WHITE)

    add_text_box(s, Inches(0.7), Inches(6.75), Inches(8.5), Inches(0.4),
                 "Prepared for Polaris Inc  ·  Partner: Trace3, LLC  ·  Distributor: Ingram Micro\n"
                 "Customer pricing = Cisco EAMP net + 20 points of margin",
                 size=12, color=RGBColor(0xC5, 0xD0, 0xDA))
    add_text_box(s, Inches(9.6), Inches(6.85), Inches(3.2), Inches(0.35),
                 "25 November 2026", size=13, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)

    # ------------------------------------------------------------------
    # 2. Agenda
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Agenda", "How this Enterprise Agreement adder is presented")
    items = [
        ("01", "Deal snapshot", "Parties, remaining term, billing, and the motion (EA adder)."),
        ("02", "Investment summary", "Cisco net, 20-point margin, and customer contract value."),
        ("03", "Pricing methodology", "How 20 points are applied to every net-cost line."),
        ("04", "Networking software", "Meraki, Catalyst DNA, Nexus, and Cisco Networking SKUs."),
        ("05", "CX / lifecycle services", "Signature and Standard support across the install base."),
        ("06", "Applications", "ThousandEyes observability adder (units, users, Insights, TAM)."),
        ("07", "Billing view & next steps", "Indicative annual payments, assumptions, and close plan."),
    ]
    for i, (num, title, blurb) in enumerate(items):
        y = Inches(1.2) + i * Inches(0.78)
        add_round(s, Inches(0.45), y, Inches(12.4), Inches(0.7), WHITE, 0.12)
        add_round(s, Inches(0.62), y + Inches(0.14), Inches(0.48), Inches(0.42), NAVY, 0.2)
        add_text_box(s, Inches(0.62), y + Inches(0.18), Inches(0.48), Inches(0.36),
                     num, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, Inches(1.3), y + Inches(0.1), Inches(10.8), Inches(0.28),
                     title, size=16, bold=True, color=NAVY)
        add_text_box(s, Inches(1.3), y + Inches(0.38), Inches(10.8), Inches(0.26),
                     blurb, size=12, color=GRAY)

    # ------------------------------------------------------------------
    # 3. Deal snapshot
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Deal snapshot", "EA modification — adder for new suites against existing program EA59697")

    cards = [
        ("End customer", "Polaris Inc", "2100 Highway 55, Hamel, MN 55340"),
        ("Smart Account", "POLARIS INDUSTRIES INC.", "Account: POLARIS INC CORE US-26612"),
        ("Partner / Disti", "Trace3  ·  Ingram Micro", meta["disti_bill_to"]),
        ("Cisco AM", meta["cam"].split("(")[0].strip(), "kbarstad@cisco.com"),
    ]
    for i, (lab, val, cap) in enumerate(cards):
        add_card(s, Inches(0.4) + i * Inches(3.2), Inches(1.2), Inches(3.05), Inches(1.28),
                 lab, val, cap, value_size=13)

    # Detail table
    rows = [
        ("Proposal name", meta["proposal_name"]),
        ("Project / SR", meta["project_name"]),
        ("Proposal ID / Deal ID / Quote ID", f"{meta['proposal_id']}  ·  {meta['deal_id']}  ·  {meta['quote_id']}"),
        ("Buying Program ID", meta["buying_program_id"]),
        ("Enterprise Agreement motion", "EA  —  modification / adder (not a net-new EA)"),
        ("Requested ship / expected book", "December 15, 2026"),
        ("Remaining duration", term_label),
        ("Billing model", "Annual  ·  Capital financing: none"),
        ("Price list", meta["price_list"]),
        ("Reseller services bill-to", meta["reseller_bill_to"]),
    ]
    table = add_table(s, len(rows), 2, Inches(0.4), Inches(2.7), Inches(12.5), Inches(4.15))
    table.columns[0].width = Inches(3.4)
    table.columns[1].width = Inches(9.1)
    for i, (k, v) in enumerate(rows):
        fill = WHITE if i % 2 == 0 else ROW_ALT
        style_cell(table.cell(i, 0), k, size=11, bold=True, color=NAVY, fill=fill)
        style_cell(table.cell(i, 1), v, size=11, color=SLATE, fill=fill)

    # ------------------------------------------------------------------
    # 4. Investment at a glance
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Investment at a glance",
               "Customer pricing applies 20 points of margin to Cisco EAMP net cost")

    add_card(s, Inches(0.4), Inches(1.2), Inches(3.05), Inches(1.35),
             "Cisco net (EAMP TCV)", usd(c["net"]),
             "Partner cost from price estimate", 16)
    add_card(s, Inches(3.6), Inches(1.2), Inches(3.05), Inches(1.35),
             "Margin (20 points)", usd(c["margin"]),
             "20% of Cisco net, added to cost", 16)
    add_card(s, Inches(6.8), Inches(1.2), Inches(3.05), Inches(1.35),
             "Customer TCV", usd(c["customer"]),
             "Cisco net × 1.20  ·  full remaining term", 16)
    add_card(s, Inches(10.0), Inches(1.2), Inches(2.9), Inches(1.35),
             "Indicative annual", usd(c["annual_cust"]),
             f"Customer TCV ÷ {years.quantize(Decimal('0.01'))} years", 16)

    # Portfolio table
    headers = ["Portfolio", "List (USD)", "Cisco Net (USD)", "20-pt Margin", "Customer TCV", "Annual"]
    body = [
        ["Networking software", c["ni_sw_list"], c["ni_sw_net"],
         money(c["ni_sw_cust"] - c["ni_sw_net"]), c["ni_sw_cust"], c["ni_sw_annual"]],
        ["CX / lifecycle services", c["ni_svc_list"], c["ni_svc_net"],
         money(c["ni_svc_cust"] - c["ni_svc_net"]), c["ni_svc_cust"], c["ni_svc_annual"]],
        ["Applications (ThousandEyes)", c["app_list"], c["app_net"],
         money(c["app_cust"] - c["app_net"]), c["app_cust"], c["app_annual"]],
        ["TOTAL", c["list"], c["net"], c["margin"], c["customer"], c["annual_cust"]],
    ]
    table = add_table(s, 1 + len(body), 6, Inches(0.4), Inches(2.75), Inches(12.5), Inches(2.55))
    widths = [3.15, 1.85, 1.9, 1.75, 1.95, 1.9]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, size=10, bold=True, color=WHITE, fill=HEADER_ROW,
                   align=PP_ALIGN.RIGHT if j else PP_ALIGN.LEFT)
    for i, row in enumerate(body):
        last = i == len(body) - 1
        fill = TOTAL_ROW if last else (WHITE if i % 2 == 0 else ROW_ALT)
        fg = WHITE if last else SLATE
        fg0 = WHITE if last else NAVY
        style_cell(table.cell(i + 1, 0), row[0], size=11, bold=True, color=fg0, fill=fill)
        for j in range(1, 6):
            style_cell(table.cell(i + 1, j), usd(row[j]), size=11, bold=last, color=fg, fill=fill,
                       align=PP_ALIGN.RIGHT)

    add_text_box(s, Inches(0.45), Inches(5.5), Inches(7.4), Inches(1.5),
                 "Notes\n"
                 "• Cisco Net is Total Contract Value from EAMP proposal 9354137 (indicative, not an approved quote).\n"
                 "• Customer TCV = Cisco Net + 20 points (Customer TCV = Cisco Net × 1.20).\n"
                 "• Annual is Customer TCV spread evenly over the remaining 42.07-month term.\n"
                 "• Security, Collaboration, Hybrid, Provider Connectivity, and Cisco PS are not in this adder.",
                 size=11, color=GRAY)

    # Mini mix bars via shapes
    mix = [
        ("Software", c["ni_sw_cust"] + c["app_cust"], TEAL),
        ("Services", c["ni_svc_cust"], GOLD),
    ]
    total_cust = c["customer"]
    add_text_box(s, Inches(8.15), Inches(5.45), Inches(4.7), Inches(0.28),
                 "CUSTOMER TCV MIX", size=10, bold=True, color=TEAL_DK)
    y = Inches(5.8)
    for label, val, color in mix:
        pct = (val / total_cust) * Decimal("100")
        add_text_box(s, Inches(8.15), y, Inches(1.7), Inches(0.28),
                     f"{label}  {pct.quantize(Decimal('0.1'))}%", size=11, color=NAVY)
        add_round(s, Inches(9.9), y + Inches(0.05), Inches(2.85), Inches(0.18), ICE, 0.5)
        bar_w = Inches(float(2.85 * float(val / total_cust)))
        add_round(s, Inches(9.9), y + Inches(0.05), bar_w, Inches(0.18), color, 0.5)
        y += Inches(0.42)

    # ------------------------------------------------------------------
    # 5. Pricing methodology
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Pricing methodology", "20 points of margin added to Cisco net on every priced line")

    steps = [
        ("1", "Start with EAMP net",
         "Cisco Total Contract Value on each SKU is the partner net cost for the remaining EA term."),
        ("2", "Add 20 points",
         "Customer Price = Cisco Net × 1.20. Twenty points means twenty percent of net cost, not of list."),
        ("3", "Hold the term",
         "The remaining duration (42.07 months) and annual billing model are unchanged from EAMP."),
        ("4", "Present TCV + annual",
         "We show full remaining-term customer TCV and an even annualized amount for budgeting."),
    ]
    for i, (n, title, body) in enumerate(steps):
        x = Inches(0.4) + i * Inches(3.2)
        add_round(s, x, Inches(1.25), Inches(3.05), Inches(2.35), WHITE, 0.08)
        add_round(s, x + Inches(0.18), Inches(1.42), Inches(0.42), Inches(0.42), TEAL, 0.2)
        add_text_box(s, x + Inches(0.18), Inches(1.46), Inches(0.42), Inches(0.36),
                     n, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text_box(s, x + Inches(0.18), Inches(1.95), Inches(2.7), Inches(0.4),
                     title, size=15, bold=True, color=NAVY)
        add_text_box(s, x + Inches(0.18), Inches(2.4), Inches(2.7), Inches(1.0),
                     body, size=12, color=GRAY)

    # Example math
    add_round(s, Inches(0.4), Inches(3.8), Inches(12.5), Inches(3.15), WHITE, 0.06)
    add_text_box(s, Inches(0.65), Inches(3.95), Inches(12.0), Inches(0.32),
                 "Worked example — proposal total", size=16, bold=True, color=NAVY)
    example = [
        ["", "Cisco Net (EAMP)", "× 1.20", "Customer Price", "Margin $"],
        ["Networking software", usd(c["ni_sw_net"]), "20 pts", usd(c["ni_sw_cust"]),
         usd(c["ni_sw_cust"] - c["ni_sw_net"])],
        ["CX / lifecycle services", usd(c["ni_svc_net"]), "20 pts", usd(c["ni_svc_cust"]),
         usd(c["ni_svc_cust"] - c["ni_svc_net"])],
        ["ThousandEyes / apps", usd(c["app_net"]), "20 pts", usd(c["app_cust"]),
         usd(c["app_cust"] - c["app_net"])],
        ["Grand total", usd(c["net"]), "20 pts", usd(c["customer"]), usd(c["margin"])],
    ]
    table = add_table(s, 5, 5, Inches(0.65), Inches(4.4), Inches(12.05), Inches(2.3))
    widths = [3.1, 2.35, 1.5, 2.55, 2.55]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for i, row in enumerate(example):
        last = i == 4
        fill = HEADER_ROW if i == 0 else (TOTAL_ROW if last else (WHITE if i % 2 else ROW_ALT))
        fg = WHITE if i in (0, 4) else SLATE
        for j, val in enumerate(row):
            align = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            style_cell(table.cell(i, j), val, size=12, bold=(i in (0, 4) or j == 0),
                       color=fg, fill=fill, align=align)

    # ------------------------------------------------------------------
    # 6. Commercial terms
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Commercial terms", "Taken from the EAMP proposal export; customer prices include 20-point margin")

    left_items = [
        ("Buying program", "Cisco Enterprise Agreement (EA) 3.x adder"),
        ("Motion", "Modify existing EA — SR112879 — new suites"),
        ("Term remaining", term_label),
        ("Billing", "Annual"),
        ("Currency / price list", "USD  ·  Global Price List US Availability"),
        ("True Forward", "Standard EA True Forward applies to committed suites"),
        ("Pre-EA assets", "Actively licensed IB is treated as consumed at EA activation"),
        ("Financing", "No Cisco Capital frequency on this estimate"),
    ]
    right_items = [
        ("End customer", "Polaris Inc\n2100 Highway 55, Hamel, MN 55340, USA"),
        ("Smart Account", "POLARIS INDUSTRIES INC."),
        ("Partner of record", "Trace3, LLC"),
        ("Distributor", "Ingram Micro (Bill-to 403766138)"),
        ("Reseller services bill-to", "Calian Corp (1061230934)"),
        ("Cisco Account Manager", meta["cam"]),
        ("Expected book / RSD", "15 December 2026"),
        ("Validity", "Indicative as of 25 November 2026; subject to Cisco quote approval"),
    ]

    def term_col(items, x):
        for i, (k, v) in enumerate(items):
            y = Inches(1.2) + i * Inches(0.7)
            add_round(s, x, y, Inches(6.15), Inches(0.64), WHITE, 0.1)
            add_text_box(s, x + Inches(0.18), y + Inches(0.06), Inches(5.8), Inches(0.2),
                         k.upper(), size=9, bold=True, color=TEAL_DK)
            add_text_box(s, x + Inches(0.18), y + Inches(0.26), Inches(5.8), Inches(0.32),
                         v.replace("\n", "  ·  "), size=12, color=NAVY)

    term_col(left_items, Inches(0.4))
    term_col(right_items, Inches(6.75))

    # ------------------------------------------------------------------
    # 7. Scope
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Scope of this adder", "Portfolios priced in EAMP 9354137 versus empty EA families")

    included = [
        ("Networking Infrastructure — software",
         "Meraki Network Infrastructure & Cameras, Catalyst DNA Switching & Wireless, Nexus Switching, Cisco Networking Switching/Wireless."),
        ("Networking Infrastructure — CX services",
         "Cisco Support Signature (ENTWAN, Nexus) and Support Standard (Switching, Wireless, Meraki, Cisco Switching)."),
        ("Applications Infrastructure",
         "ThousandEyes Cloud/Enterprise Agents (2,200 units), End User Monitoring (110), Internet Insights, TAM 5 hr/week."),
    ]
    excluded = [
        "Security — no SKUs on this estimate",
        "Collaboration — no SKUs on this estimate",
        "Hybrid Cloud — no SKUs on this estimate",
        "Provider Connectivity — no SKUs",
        "Cisco Professional Services — no SKUs",
        "Hardware / non-EA product (out of file)",
    ]
    add_text_box(s, Inches(0.45), Inches(1.18), Inches(7.4), Inches(0.3),
                 "IN THIS PROPOSAL", size=12, bold=True, color=GREEN)
    for i, (t, b) in enumerate(included):
        y = Inches(1.52) + i * Inches(1.35)
        add_round(s, Inches(0.4), y, Inches(7.6), Inches(1.22), WHITE, 0.08)
        add_rect(s, Inches(0.4), y, Inches(0.1), Inches(1.22), GREEN)
        add_text_box(s, Inches(0.7), y + Inches(0.14), Inches(7.1), Inches(0.32),
                     t, size=14, bold=True, color=NAVY)
        add_text_box(s, Inches(0.7), y + Inches(0.48), Inches(7.1), Inches(0.62),
                     b, size=12, color=GRAY)

    add_text_box(s, Inches(8.3), Inches(1.18), Inches(4.6), Inches(0.3),
                 "NOT PRICED IN THIS ADDER", size=12, bold=True, color=RED_ACCENT)
    add_round(s, Inches(8.25), Inches(1.52), Inches(4.65), Inches(5.35), WHITE, 0.08)
    for i, line in enumerate(excluded):
        add_text_box(s, Inches(8.5), Inches(1.75) + i * Inches(0.55), Inches(4.2), Inches(0.5),
                     "▸  " + line, size=13, color=SLATE)

    # ------------------------------------------------------------------
    # 8. Networking software by suite
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Networking software — by suite",
               f"Cisco net {usd(c['ni_sw_net'])}  →  customer {usd(c['ni_sw_cust'])}  after 20 points")

    suite_order = [
        "Cisco DNA Switching",
        "Nexus Switching",
        "Cisco DNA Wireless",
        "Meraki - Network Infrastructure",
        "Cisco Networking Switching",
        "Meraki - Camera Systems",
        "Cisco Networking Wireless",
    ]
    headers = ["Suite", "SKUs", "List TCV", "Cisco Net", "Customer (20 pts)", "Annual"]
    rows = []
    for name in suite_order:
        b = data["ni_sw_suites"][name]
        rows.append([name, str(b["count"]), b["list"], b["net"], b["customer"],
                     money(b["customer"] / years)])
    rows.append(["Networking software total", str(sum(b["count"] for b in data["ni_sw_suites"].values())),
                 c["ni_sw_list"], c["ni_sw_net"], c["ni_sw_cust"], c["ni_sw_annual"]])

    table = add_table(s, 1 + len(rows), 6, Inches(0.4), Inches(1.2), Inches(12.5), Inches(4.6))
    widths = [3.7, 0.8, 1.9, 1.95, 2.2, 1.95]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, size=11, bold=True, color=WHITE, fill=HEADER_ROW,
                   align=PP_ALIGN.RIGHT if j else PP_ALIGN.LEFT)
    for i, row in enumerate(rows):
        last = i == len(rows) - 1
        fill = TOTAL_ROW if last else (WHITE if i % 2 == 0 else ROW_ALT)
        fg = WHITE if last else SLATE
        style_cell(table.cell(i + 1, 0), row[0], size=12, bold=True, color=WHITE if last else NAVY, fill=fill)
        style_cell(table.cell(i + 1, 1), row[1], size=12, bold=last, color=fg, fill=fill, align=PP_ALIGN.RIGHT)
        for j in range(2, 6):
            style_cell(table.cell(i + 1, j), usd(row[j]), size=12, bold=last, color=fg, fill=fill,
                       align=PP_ALIGN.RIGHT)

    add_text_box(s, Inches(0.45), Inches(6.0), Inches(12.4), Inches(0.95),
                 "Cisco DNA Switching is the largest software suite in this adder, driven by Catalyst 9300/9500 EA licenses "
                 "on the existing install base. Nexus 9300 Essential licenses (XF / XF2 / XF3) are the primary data-center "
                 "software commitment. Meraki MX/MR/MS Essentials and a partial Meraki MV camera suite are included. "
                 "Zero-dollar embedded support PIDs (SVS-E3S-*, E3N-*-S) are omitted from SKU slides but sit in the EAMP file.",
                 size=12, color=GRAY)

    # ------------------------------------------------------------------
    # 9+ Networking software SKUs
    # ------------------------------------------------------------------
    sw_priced = [i for i in data["ni_sw"] if i["net"] > 0]
    sw_chunks = list(chunk(sw_priced, 12))
    for page_i, group in enumerate(sw_chunks, 1):
        s = slide()
        paint_content_bg(s)
        add_header(s, f"Networking software SKUs  ({page_i} of {len(sw_chunks)})",
                   "Customer Price = Cisco Net × 1.20   ·   TCV is remaining-term contract value")
        headers = ["Part number", "Description", "Qty", "List", "Cisco Net", "Customer"]
        table = add_table(s, 1 + len(group), 6, Inches(0.3), Inches(1.15), Inches(12.7), Inches(5.7))
        widths = [2.15, 4.55, 0.85, 1.7, 1.7, 1.75]
        for i, w in enumerate(widths):
            table.columns[i].width = Inches(w)
        for j, h in enumerate(headers):
            style_cell(table.cell(0, j), h, size=10, bold=True, color=WHITE, fill=HEADER_ROW,
                       align=PP_ALIGN.RIGHT if j >= 2 else PP_ALIGN.LEFT)
        for i, it in enumerate(group):
            fill = WHITE if i % 2 == 0 else ROW_ALT
            desc = it["desc"]
            if len(desc) > 58:
                desc = desc[:55] + "…"
            style_cell(table.cell(i + 1, 0), it["part"], size=10, bold=True, color=NAVY, fill=fill)
            style_cell(table.cell(i + 1, 1), desc, size=10, color=SLATE, fill=fill)
            style_cell(table.cell(i + 1, 2), f"{int(it['qty'])}", size=10, color=SLATE, fill=fill,
                       align=PP_ALIGN.RIGHT)
            for j, val in enumerate((it["list"], it["net"], it["customer"]), 3):
                style_cell(table.cell(i + 1, j), usd(val), size=10, color=SLATE, fill=fill,
                           align=PP_ALIGN.RIGHT)

    # ------------------------------------------------------------------
    # CX services by suite
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "CX / lifecycle services — by suite",
               f"Cisco net {usd(c['ni_svc_net'])}  →  customer {usd(c['ni_svc_cust'])}  after 20 points")

    svc_order = [
        "Services: Cisco DNA Switching",
        "Services: Cisco DNA SD-WAN & Routing",
        "Services: Cisco DNA Wireless",
        "Services: Nexus Switching",
        "Services: Meraki - Network Infrastructure",
        "Services: Cisco Networking Switching",
    ]
    headers = ["Service suite", "List TCV", "Cisco Net", "Customer (20 pts)", "Annual"]
    rows = []
    for name in svc_order:
        b = data["ni_svc_suites"][name]
        if b["net"] == 0 and b["list"] == 0:
            continue
        label = name.replace("Services: ", "")
        rows.append([label, b["list"], b["net"], b["customer"], money(b["customer"] / years)])
    rows.append(["CX / lifecycle services total", c["ni_svc_list"], c["ni_svc_net"],
                 c["ni_svc_cust"], c["ni_svc_annual"]])

    table = add_table(s, 1 + len(rows), 5, Inches(0.4), Inches(1.2), Inches(12.5), Inches(4.0))
    widths = [4.4, 2.0, 2.05, 2.15, 1.9]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, size=11, bold=True, color=WHITE, fill=HEADER_ROW,
                   align=PP_ALIGN.RIGHT if j else PP_ALIGN.LEFT)
    for i, row in enumerate(rows):
        last = i == len(rows) - 1
        fill = TOTAL_ROW if last else (WHITE if i % 2 == 0 else ROW_ALT)
        fg = WHITE if last else SLATE
        style_cell(table.cell(i + 1, 0), row[0], size=12, bold=True, color=WHITE if last else NAVY, fill=fill)
        for j in range(1, 5):
            style_cell(table.cell(i + 1, j), usd(row[j]), size=12, bold=last, color=fg, fill=fill,
                       align=PP_ALIGN.RIGHT)

    add_round(s, Inches(0.4), Inches(5.4), Inches(12.5), Inches(1.55), WHITE, 0.08)
    add_text_box(s, Inches(0.65), Inches(5.55), Inches(12.0), Inches(0.3),
                 "Service-level notes", size=14, bold=True, color=NAVY)
    add_text_box(s, Inches(0.65), Inches(5.9), Inches(12.0), Inches(0.9),
                 "Switching and wireless CX are Cisco Support Standard (8x5xNBD / 8x7xNCD / SW). ENTWAN and Nexus use "
                 "Cisco Support Signature (including 24x7x4 on a subset of Nexus and ENTWAN assets). Quantity-desired is 1 "
                 "on CX PIDs because EA services are priced as a coverage envelope over the install base, not per-device SKUs. "
                 "Prior-period lapse charges ($5.20 and $32.94 net) are carried through at the same 20-point markup.",
                 size=12, color=GRAY)

    # ------------------------------------------------------------------
    # CX SKUs
    # ------------------------------------------------------------------
    svc_priced = [i for i in data["ni_svc"] if i["net"] > 0]
    svc_chunks = list(chunk(svc_priced, 8))
    for page_i, group in enumerate(svc_chunks, 1):
        s = slide()
        paint_content_bg(s)
        add_header(s, f"CX / lifecycle service SKUs  ({page_i} of {len(svc_chunks)})",
                   "Customer Price = Cisco Net × 1.20")
        headers = ["Part number", "Description", "List", "Cisco Net", "Customer"]
        table = add_table(s, 1 + len(group), 5, Inches(0.3), Inches(1.15), Inches(12.7), Inches(5.7))
        widths = [2.4, 4.9, 1.8, 1.8, 1.8]
        for i, w in enumerate(widths):
            table.columns[i].width = Inches(w)
        for j, h in enumerate(headers):
            style_cell(table.cell(0, j), h, size=10, bold=True, color=WHITE, fill=HEADER_ROW,
                       align=PP_ALIGN.RIGHT if j >= 2 else PP_ALIGN.LEFT)
        for i, it in enumerate(group):
            fill = WHITE if i % 2 == 0 else ROW_ALT
            desc = it["desc"]
            if len(desc) > 62:
                desc = desc[:59] + "…"
            style_cell(table.cell(i + 1, 0), it["part"], size=10, bold=True, color=NAVY, fill=fill)
            style_cell(table.cell(i + 1, 1), desc, size=10, color=SLATE, fill=fill)
            for j, val in enumerate((it["list"], it["net"], it["customer"]), 2):
                style_cell(table.cell(i + 1, j), usd(val), size=10, color=SLATE, fill=fill,
                           align=PP_ALIGN.RIGHT)

    # ------------------------------------------------------------------
    # Applications / ThousandEyes
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Applications — ThousandEyes",
               f"Cisco net {usd(c['app_net'])}  →  customer {usd(c['app_cust'])}  after 20 points")

    add_card(s, Inches(0.4), Inches(1.18), Inches(3.05), Inches(1.2),
             "Cloud & Enterprise Agents", "2,200 units", "E3A-TE-UNITS  ·  Full commitment")
    add_card(s, Inches(3.6), Inches(1.18), Inches(3.05), Inches(1.2),
             "End-user monitoring", "110 users", "E3A-TE-USERS  ·  Advantage")
    add_card(s, Inches(6.8), Inches(1.18), Inches(3.05), Inches(1.2),
             "Internet Insights", "1 package", "E3A-TE-INSIGHTS")
    add_card(s, Inches(10.0), Inches(1.18), Inches(2.9), Inches(1.2),
             "TAM support", "5 hrs / week", "E3A-TE-TAM-5WK  ·  Partial suite")

    headers = ["Part number", "Description", "Qty", "List", "Cisco Net", "Customer", "Annual"]
    apps_rows = [i for i in data["app_sw"] if i["net"] > 0]
    app_desc = {
        "E3A-TE-TAM-5WK": "TAM 1/8 support (5 hours/week)",
        "E3A-TE-UNITS": "Cloud & Enterprise Agents (per unit)",
        "E3A-TE-USERS": "End User Monitoring Advantage",
        "E3A-TE-INSIGHTS": "Internet Insights (per package)",
    }
    table = add_table(s, 1 + len(apps_rows) + 1, 7, Inches(0.3), Inches(2.6), Inches(12.7), Inches(2.7))
    widths = [2.05, 3.55, 0.8, 1.55, 1.55, 1.6, 1.6]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, size=10, bold=True, color=WHITE, fill=HEADER_ROW,
                   align=PP_ALIGN.RIGHT if j >= 2 else PP_ALIGN.LEFT)
    for i, it in enumerate(apps_rows):
        fill = WHITE if i % 2 == 0 else ROW_ALT
        style_cell(table.cell(i + 1, 0), it["part"], size=11, bold=True, color=NAVY, fill=fill)
        style_cell(table.cell(i + 1, 1), app_desc.get(it["part"], it["desc"]), size=11, color=SLATE, fill=fill)
        style_cell(table.cell(i + 1, 2), f"{int(it['qty'])}", size=11, color=SLATE, fill=fill,
                   align=PP_ALIGN.RIGHT)
        vals = (it["list"], it["net"], it["customer"], money(it["customer"] / years))
        for j, val in enumerate(vals, 3):
            style_cell(table.cell(i + 1, j), usd(val), size=11, color=SLATE, fill=fill,
                       align=PP_ALIGN.RIGHT)
    last = 1 + len(apps_rows)
    style_cell(table.cell(last, 0), "Applications total", size=11, bold=True, color=WHITE, fill=TOTAL_ROW)
    style_cell(table.cell(last, 1), "ThousandEyes suite + TAM adder", size=11, bold=True, color=WHITE, fill=TOTAL_ROW)
    style_cell(table.cell(last, 2), "", size=11, color=WHITE, fill=TOTAL_ROW)
    for j, val in enumerate((c["app_list"], c["app_net"], c["app_cust"], c["app_annual"]), 3):
        style_cell(table.cell(last, j), usd(val), size=11, bold=True, color=WHITE, fill=TOTAL_ROW,
                   align=PP_ALIGN.RIGHT)

    add_round(s, Inches(0.4), Inches(5.5), Inches(12.5), Inches(1.45), WHITE, 0.08)
    add_text_box(s, Inches(0.65), Inches(5.65), Inches(12.0), Inches(1.15),
                 "The ThousandEyes commitment is Full on the core suite (units, users, Insights) and Partial on the TAM "
                 "add-on (Technical Account Manager, 5 hours per week). Embedded Cisco Support Basic (E3A-TE-S) is $0 TCV "
                 "and is omitted from the table. Subscription discounts in EAMP are 30% on TE software and 55% on the TAM PID; "
                 "those Cisco discounts are already inside the net cost before Trace3’s 20-point margin is applied.",
                 size=13, color=SLATE)

    # ------------------------------------------------------------------
    # Annual billing view
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Indicative annual billing view",
               "Even annualization of customer TCV over the remaining 42.07-month EA term")

    stub_years = years - Decimal("3")
    annual = c["annual_cust"]
    p1 = annual
    p2 = annual
    p3 = annual
    p4 = money(c["customer"] - (p1 + p2 + p3))  # residual so it ties to TCV

    headers = ["Period", "Timing (indicative)", "Software", "Services", "Apps", "Customer amount"]
    sw_a = c["ni_sw_annual"]
    svc_a = c["ni_svc_annual"]
    app_a = c["app_annual"]
    sw_stub = money(c["ni_sw_cust"] - sw_a * 3)
    svc_stub = money(c["ni_svc_cust"] - svc_a * 3)
    app_stub = money(c["app_cust"] - app_a * 3)
    body = [
        ["Year 1", "At book  ·  15 Dec 2026", sw_a, svc_a, app_a, p1],
        ["Year 2", "15 Dec 2027", sw_a, svc_a, app_a, p2],
        ["Year 3", "15 Dec 2028", sw_a, svc_a, app_a, p3],
        [f"Stub ({stub_years.quantize(Decimal('0.01'))} yr)", "15 Dec 2029 → ~Jun 2030",
         sw_stub, svc_stub, app_stub, p4],
        ["Remaining-term TCV", "Full adder", c["ni_sw_cust"], c["ni_svc_cust"], c["app_cust"], c["customer"]],
    ]
    table = add_table(s, 1 + len(body), 6, Inches(0.35), Inches(1.2), Inches(12.6), Inches(3.35))
    widths = [2.15, 2.55, 1.95, 1.95, 1.85, 2.15]
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, size=11, bold=True, color=WHITE, fill=HEADER_ROW,
                   align=PP_ALIGN.RIGHT if j >= 2 else PP_ALIGN.LEFT)
    for i, row in enumerate(body):
        last = i == len(body) - 1
        fill = TOTAL_ROW if last else (WHITE if i % 2 == 0 else ROW_ALT)
        fg = WHITE if last else SLATE
        style_cell(table.cell(i + 1, 0), row[0], size=12, bold=True, color=WHITE if last else NAVY, fill=fill)
        style_cell(table.cell(i + 1, 1), row[1], size=11, color=fg, fill=fill)
        for j in range(2, 6):
            style_cell(table.cell(i + 1, j), usd(row[j]), size=12, bold=last, color=fg, fill=fill,
                       align=PP_ALIGN.RIGHT)

    add_round(s, Inches(0.4), Inches(4.8), Inches(12.5), Inches(2.15), WHITE, 0.08)
    add_text_box(s, Inches(0.65), Inches(4.95), Inches(12.0), Inches(0.3),
                 "How to read this schedule", size=14, bold=True, color=NAVY)
    add_text_box(s, Inches(0.65), Inches(5.3), Inches(12.0), Inches(1.45),
                 "EAMP lists Billing Model = Annual and Duration = 42.0667 months. This slide annualizes customer TCV "
                 f"({usd(c['customer'])}) evenly, then drops the leftover ~6.07 months into a stub so the four rows equal "
                 "remaining-term TCV. It is a budget view, not a Cisco invoice preview. Actual EA invoices follow CCW / "
                 "distributor bill-to (Ingram Micro) and may differ on start-date proration, tax, True Forward, and any "
                 "co-term true-up. No Cisco Capital financing is attached to this estimate.",
                 size=13, color=SLATE)

    # ------------------------------------------------------------------
    # Value vs list
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Value versus list", "EA discounts are already in Cisco net; 20 points sit on top of that net")

    list_p = c["list"]
    net_p = c["net"]
    cust_p = c["customer"]
    cisco_off = money(list_p - net_p)
    cisco_pct = (cisco_off / list_p * Decimal("100")).quantize(Decimal("0.1"))
    cust_off = money(list_p - cust_p)
    cust_pct = (cust_off / list_p * Decimal("100")).quantize(Decimal("0.1"))

    add_card(s, Inches(0.4), Inches(1.2), Inches(4.05), Inches(1.4),
             "Total list (EAMP)", usd(list_p), "Software + services remaining-term list", 18)
    add_card(s, Inches(4.65), Inches(1.2), Inches(4.05), Inches(1.4),
             "Cisco discount vs list", f"{cisco_pct}%", f"{usd(cisco_off)} already in EAMP net")
    add_card(s, Inches(8.9), Inches(1.2), Inches(4.0), Inches(1.4),
             "Customer vs list", f"{cust_pct}% off", f"{usd(cust_off)} after 20-point margin")

    # Waterfall-style bars
    add_round(s, Inches(0.4), Inches(2.85), Inches(12.5), Inches(4.1), WHITE, 0.06)
    add_text_box(s, Inches(0.65), Inches(3.0), Inches(12.0), Inches(0.32),
                 "Price build-up", size=16, bold=True, color=NAVY)

    steps_wf = [
        ("List TCV", list_p, NAVY_MID, usd(list_p)),
        ("Cisco EA / IB discounts", cisco_off, TEAL, f"− {usd(cisco_off)}"),
        ("Cisco net (EAMP)", net_p, NAVY, usd(net_p)),
        ("20 points of margin", c["margin"], GOLD, f"+ {usd(c['margin'])}"),
        ("Customer TCV", cust_p, GREEN, usd(cust_p)),
    ]
    max_v = list_p
    y = Inches(3.45)
    for label, val, color, caption in steps_wf:
        add_text_box(s, Inches(0.7), y, Inches(3.4), Inches(0.38), label, size=13, bold=True, color=NAVY)
        add_round(s, Inches(4.2), y + Inches(0.08), Inches(5.7), Inches(0.26), ICE, 0.5)
        bar_w = max(Inches(0.12), Inches(float(5.7 * float(val / max_v))))
        add_round(s, Inches(4.2), y + Inches(0.08), bar_w, Inches(0.26), color, 0.5)
        add_text_box(s, Inches(10.05), y, Inches(2.55), Inches(0.38), caption, size=13, bold=True,
                     color=NAVY, align=PP_ALIGN.RIGHT)
        y += Inches(0.62)

    # ------------------------------------------------------------------
    # Next steps
    # ------------------------------------------------------------------
    s = slide()
    paint_content_bg(s)
    add_header(s, "Assumptions, exclusions, and next steps",
               "What has to happen between this estimate and a bookable Cisco quote")

    col_a = [
        ("Assumptions baked into the numbers",
         "• Remaining term 42.07 months from the 15 Dec 2026 book date.\n"
         "• Annual billing; no Cisco Capital on this estimate.\n"
         "• 20 points of margin on every non-zero Cisco net line.\n"
         "• Quantities, IB credits, ramp, and OTD follow the EAMP file.\n"
         "• USD; Global Price List US Availability."),
        ("Exclusions",
         "• Tax, freight, and distributor surcharge (if any).\n"
         "• Hardware not on this EA estimate.\n"
         "• Security, Collaboration, Hybrid, Provider Connectivity, Cisco PS.\n"
         "• Future True Forward growth beyond committed quantities.\n"
         "• FX and price-list changes after 25 Nov 2026."),
    ]
    col_b = [
        ("Recommended close path",
         "1. Confirm suite mix and quantities with Polaris stakeholders.\n"
         "2. Convert EAMP estimate to an approved CCW / DSA quote.\n"
         "3. Align Ingram Micro bill-to and Calian services bill-to.\n"
         "4. Refresh EUIF / Smart Account mapping (POLARIS INDUSTRIES INC.).\n"
         "5. Target book date 15 December 2026."),
        ("Open items",
         "• Cisco quote approval — this file is explicitly not an approved quote.\n"
         "• Reseller services bill-to shows Calian Corp; confirm if still correct for Trace3.\n"
         "• Partial Meraki MV and ThousandEyes TAM commitments vs full suites.\n"
         "• Any additional suites 12–17 beyond what EAMP priced here."),
    ]
    for i, (title, body) in enumerate(col_a):
        y = Inches(1.2) + i * Inches(2.8)
        add_round(s, Inches(0.4), y, Inches(6.2), Inches(2.6), WHITE, 0.08)
        add_text_box(s, Inches(0.6), y + Inches(0.12), Inches(5.85), Inches(0.32),
                     title, size=14, bold=True, color=NAVY)
        add_text_box(s, Inches(0.6), y + Inches(0.48), Inches(5.85), Inches(2.0),
                     body, size=12, color=SLATE)
    for i, (title, body) in enumerate(col_b):
        y = Inches(1.2) + i * Inches(2.8)
        add_round(s, Inches(6.75), y, Inches(6.15), Inches(2.6), WHITE, 0.08)
        add_text_box(s, Inches(6.95), y + Inches(0.12), Inches(5.8), Inches(0.32),
                     title, size=14, bold=True, color=NAVY)
        add_text_box(s, Inches(6.95), y + Inches(0.48), Inches(5.8), Inches(2.0),
                     body, size=12, color=SLATE)

    # ------------------------------------------------------------------
    # Closing / disclaimer
    # ------------------------------------------------------------------
    s = slide()
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, TEAL)
    add_rect(s, 0, Inches(6.55), SLIDE_W, Inches(0.08), GOLD)
    add_text_box(s, Inches(0.7), Inches(1.3), Inches(12.0), Inches(0.5),
                 "Polaris Inc  ·  Cisco Enterprise Agreement", size=16, bold=True, color=GOLD)
    add_text_box(s, Inches(0.7), Inches(1.85), Inches(12.0), Inches(0.7),
                 "Customer investment", size=32, bold=True, color=WHITE)
    add_text_box(s, Inches(0.7), Inches(2.6), Inches(12.0), Inches(0.55),
                 f"{usd(c['customer'])}  remaining-term TCV    ·    {usd(c['annual_cust'])}  indicative annual",
                 size=20, color=TEAL)
    add_text_box(s, Inches(0.7), Inches(3.35), Inches(12.0), Inches(1.4),
                 "Built from EAMP Price Estimate ID 9354137 (25 Nov 2026).\n"
                 "Twenty points of margin have been added to Cisco net cost (Customer Price = Cisco Net × 1.20).\n"
                 "This presentation contains indicative pricing only and is not an approved Cisco quote.",
                 size=15, color=RGBColor(0xC5, 0xD0, 0xDA))
    add_text_box(s, Inches(0.7), Inches(5.0), Inches(12.0), Inches(1.2),
                 "Trace3, LLC  ·  Ingram Micro  ·  Cisco Account Manager: Katlyn Barstad\n"
                 "Target book date: 15 December 2026  ·  Buying Program EA59697  ·  Deal 93067714",
                 size=14, color=WHITE)
    add_text_box(s, Inches(0.7), Inches(6.75), Inches(12.0), Inches(0.4),
                 "CISCO CONFIDENTIAL  ·  For Polaris Inc and Trace3 use  ·  25 November 2026",
                 size=12, color=GOLD)

    # Footers on content slides (skip cover = 0 and closing = last)
    total = len(slides)
    for i, sl in enumerate(slides, 1):
        if i == 1 or i == total:
            continue
        add_footer(sl, i, total)

    return prs


def main():
    data = parse_workbook(XLSX)
    prs = build_presentation(data)
    prs.save(OUT)
    c = data["cust"]
    print(f"Wrote {OUT}")
    print(f"  Cisco net:     {usd(c['net'])}")
    print(f"  20-pt margin:  {usd(c['margin'])}")
    print(f"  Customer TCV:  {usd(c['customer'])}")
    print(f"  Annual:        {usd(c['annual_cust'])}")
    print(f"  Slides:        {len(prs.slides)}")


if __name__ == "__main__":
    main()
