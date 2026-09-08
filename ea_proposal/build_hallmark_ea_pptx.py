#!/usr/bin/env python3
"""Build the Hallmark Cards Cisco Enterprise Agreement proposal PPTX.

Pricing is sourced from EAMP Price Estimate ID 9325586 (indicative only;
not an approved Cisco quote).
"""

from __future__ import annotations

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# --- Brand ---
NAVY = RGBColor(0x05, 0x2B, 0x4E)
NAVY_HEX = "052B4E"
DEEP = RGBColor(0x0A, 0x3A, 0x66)
BLUE = RGBColor(0x04, 0x9F, 0xD9)
CYAN = RGBColor(0x00, 0xBC, 0xEB)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
GOLD_HEX = "C4A35A"
GREEN = RGBColor(0x2E, 0xA8, 0x4E)
GREEN_HEX = "2EA84E"
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
WHITE_HEX = "FFFFFF"
OFF_WHITE = RGBColor(0xF4, 0xF7, 0xFB)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5A, 0x64, 0x72)
ROW_ALT = "EEF3F8"
ROW_WHITE = "FFFFFF"
RED_SOFT = RGBColor(0xB3, 0x3A, 0x3A)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def money(n: float, cents: bool = True) -> str:
    if cents:
        return f"${n:,.2f}"
    return f"${n:,.0f}"


def money_short(n: float) -> str:
    if abs(n) >= 1_000_000:
        return f"${n / 1_000_000:.2f}M"
    if abs(n) >= 1_000:
        return f"${n / 1_000:.0f}K"
    return money(n)


def set_run(run, text, size=14, bold=False, color=INK, font="Calibri"):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_textbox(slide, l, t, w, h, text, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, font="Calibri", anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(_ensure_run(p), text, size, bold, color, font)
    return box


def _ensure_run(p):
    if p.runs:
        return p.runs[0]
    return p.add_run()


def add_rect(slide, l, t, w, h, fill, line=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
    shp.shadow.inherit = False
    return shp


def add_round(slide, l, t, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    shp.adjustments[0] = 0.08
    return shp


def set_cell_fill(cell, hex_color: str):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag == qn("a:solidFill"):
            tcPr.remove(child)
    solid = etree.SubElement(tcPr, qn("a:solidFill"))
    srgb = etree.SubElement(solid, "{%s}srgbClr" % A_NS)
    srgb.set("val", hex_color)


def set_cell_margins(cell, l=60000, t=40000, r=60000, b=40000):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.set("marL", str(l))
    tcPr.set("marR", str(r))
    tcPr.set("marT", str(t))
    tcPr.set("marB", str(b))


def style_cell(cell, text, size=11, bold=False, color=INK, fill=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE):
    cell.text = ""
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    set_run(run, text, size, bold, color)
    if fill:
        set_cell_fill(cell, fill)
    set_cell_margins(cell)
    try:
        cell.vertical_anchor = anchor
    except Exception:
        pass


def add_table(slide, rows, cols, l, t, w, h):
    table_shape = slide.shapes.add_table(rows, cols, l, t, w, h)
    table = table_shape.table
    table.horz_banding = False
    table.vert_banding = False
    return table


def fill_table(table, headers, data, col_align=None, header_fill=NAVY_HEX, sizes=None, header_size=11, body_size=10):
    n_cols = len(headers)
    col_align = col_align or [PP_ALIGN.LEFT] * n_cols
    sizes = sizes or [body_size] * n_cols
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, size=header_size, bold=True, color=WHITE, fill=header_fill, align=col_align[j])
    for i, row in enumerate(data, 1):
        fill = ROW_ALT if i % 2 == 0 else ROW_WHITE
        for j, val in enumerate(row):
            label = str(row[0]).upper()
            is_last = i == len(data) and (label.startswith("TOTAL") or label.startswith("3-YEAR"))
            style_cell(
                table.cell(i, j),
                str(val),
                size=sizes[j] if j < len(sizes) else body_size,
                bold=is_last or (j == 0 and i > 0 and False),
                color=WHITE if is_last else INK,
                fill=NAVY_HEX if is_last else fill,
                align=col_align[j],
            )
            if is_last:
                table.cell(i, j).text_frame.paragraphs[0].font.bold = True


def set_col_widths(table, widths):
    for i, w in enumerate(widths):
        table.columns[i].width = w


def footer(slide, page, total, confidential=True):
    add_rect(slide, Inches(0), Inches(7.28), SLIDE_W, Inches(0.22), NAVY)
    add_textbox(
        slide,
        Inches(0.35),
        Inches(7.28),
        Inches(8.5),
        Inches(0.22),
        "CISCO CONFIDENTIAL  |  Indicative pricing as of 08-Sep-2026  |  Not an approved Cisco quote",
        size=9,
        color=WHITE,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        Inches(10.6),
        Inches(7.28),
        Inches(2.4),
        Inches(0.22),
        f"Hallmark CX EA  ·  {page} / {total}",
        size=9,
        color=WHITE,
        align=PP_ALIGN.RIGHT,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def header_bar(slide, eyebrow, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.18), NAVY)
    add_rect(slide, Inches(0), Inches(0), Inches(0.12), Inches(1.18), GOLD)
    add_textbox(slide, Inches(0.4), Inches(0.12), Inches(12.4), Inches(0.28), eyebrow.upper(), size=11, bold=True, color=GOLD)
    add_textbox(slide, Inches(0.4), Inches(0.36), Inches(12.4), Inches(0.42), title, size=24, bold=True, color=WHITE)
    if subtitle:
        add_textbox(slide, Inches(0.4), Inches(0.78), Inches(12.4), Inches(0.28), subtitle, size=12, color=CYAN)


def kpi_card(slide, l, t, w, h, label, value, caption=None, accent=GOLD):
    add_round(slide, l, t, w, h, WHITE)
    # left accent
    add_rect(slide, l, t, Inches(0.08), h, accent)
    add_textbox(slide, l + Inches(0.18), t + Inches(0.10), w - Inches(0.28), Inches(0.26), label.upper(), size=10, bold=True, color=MUTED)
    add_textbox(slide, l + Inches(0.18), t + Inches(0.34), w - Inches(0.28), Inches(0.42), value, size=22, bold=True, color=NAVY)
    if caption:
        add_textbox(slide, l + Inches(0.18), t + Inches(0.78), w - Inches(0.28), Inches(0.32), caption, size=10, color=MUTED)


def bullet_block(slide, l, t, w, h, items, size=14, color=INK, spacing=True):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.level = 0
        p.space_after = Pt(8 if spacing else 4)
        set_run(p.add_run(), "•  " + item, size, False, color)
    return box


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

TOTAL_SLIDES = 18


def s01_title(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, GOLD)
    add_textbox(s, Inches(0.7), Inches(1.35), Inches(12), Inches(0.35), "CISCO ENTERPRISE AGREEMENT  ·  PROPOSAL ID 9325586", 13, True, GOLD)
    add_textbox(s, Inches(0.7), Inches(1.80), Inches(12), Inches(1.1), "Hallmark Cards, Incorporated", 36, True, WHITE)
    add_textbox(s, Inches(0.7), Inches(2.70), Inches(12), Inches(0.5), "3-Year Cisco Enterprise Agreement  |  Hallmark CX EA", 22, False, CYAN)
    add_rect(s, Inches(0.7), Inches(3.40), Inches(2.2), Inches(0.06), GOLD)

    facts = [
        ("TERM", "36 months"),
        ("BILLING", "Annual"),
        ("NET TCV", "$3.15M"),
        ("BOOK DATE", "16 Oct 2026"),
    ]
    for i, (k, v) in enumerate(facts):
        x = Inches(0.7) + Inches(i * 3.05)
        add_textbox(s, x, Inches(3.70), Inches(2.8), Inches(0.24), k, 11, True, GOLD)
        add_textbox(s, x, Inches(3.94), Inches(2.8), Inches(0.40), v, 20, True, WHITE)

    add_textbox(
        s,
        Inches(0.7),
        Inches(5.00),
        Inches(11.5),
        Inches(0.70),
        "Partner: World Wide Technology, Inc    ·    Cisco AM: Robin Randolph    ·    Buying Program: EA85986\nSmart Account: HALLMARK CARDS, INCORPORATED    ·    Deal ID 85356428",
        13,
        False,
        RGBColor(0xC5, 0xD4, 0xE3),
    )
    add_textbox(
        s,
        Inches(0.7),
        Inches(6.55),
        Inches(12),
        Inches(0.45),
        "CISCO CONFIDENTIAL  ·  Indicative pricing only as of 08-Sep-2026  ·  This is not an approved Cisco quote",
        11,
        False,
        GOLD,
    )
    return s


def s02_agenda(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Hallmark CX EA", "Agenda")
    items = [
        ("01", "Deal snapshot & commercial terms", "Who, what, when, and how this EA is structured"),
        ("02", "Investment summary", "Net TCV, software vs. services, and savings versus list"),
        ("03", "Portfolio architecture", "Networking, Applications, Security, and Collaboration CX"),
        ("04", "Bill of materials", "Suite commitments, quantities, and contract value"),
        ("05", "EA operating model & next steps", "True-forward, annual billing, and path to book"),
    ]
    for i, (num, title, desc) in enumerate(items):
        y = Inches(1.50) + Inches(i * 1.05)
        add_round(s, Inches(0.45), y, Inches(12.4), Inches(0.92), WHITE)
        add_rect(s, Inches(0.45), y, Inches(0.12), Inches(0.92), GOLD if i % 2 == 0 else BLUE)
        add_textbox(s, Inches(0.80), y + Inches(0.16), Inches(1.0), Inches(0.60), num, 22, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(1.90), y + Inches(0.14), Inches(10.5), Inches(0.36), title, 18, True, NAVY)
        add_textbox(s, Inches(1.90), y + Inches(0.48), Inches(10.5), Inches(0.32), desc, 13, False, MUTED)
    footer(s, 2, TOTAL_SLIDES)
    return s


def s03_snapshot(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Deal snapshot", "Hallmark CX EA at a glance", "Proposal 8-31-26 1Hallmark 3 year 10-16-26")

    rows = [
        ("End customer", "Hallmark Cards, Incorporated"),
        ("Sold-to address", "2501 McGee Traffic Way, Kansas City, MO 64141, US"),
        ("Account", "HALLMARK CARDS INC US"),
        ("Smart Account", "HALLMARK CARDS, INCORPORATED"),
        ("Partner", "World Wide Technology, Inc  (Bill-to 1001363275)"),
        ("Cisco Account Manager", "Robin Randolph  ·  robrando@cisco.com"),
        ("Project / motion", "Hallmark CX EA  ·  Enterprise Agreement (EA)"),
        ("Buying Program ID", "EA85986"),
        ("Deal ID / Proposal ID", "85356428  /  9325586"),
        ("Price list", "Global Price List US Availability"),
        ("Requested ship / book", "16 October 2026"),
        ("Duration / billing", "36 months  ·  Annual billing  ·  No capital financing"),
    ]
    table = add_table(s, len(rows) + 1, 2, Inches(0.45), Inches(1.42), Inches(12.4), Inches(5.55))
    set_col_widths(table, [Inches(3.4), Inches(9.0)])
    style_cell(table.cell(0, 0), "FIELD", 11, True, WHITE, NAVY_HEX)
    style_cell(table.cell(0, 1), "DETAIL", 11, True, WHITE, NAVY_HEX)
    for i, (k, v) in enumerate(rows, 1):
        fill = ROW_ALT if i % 2 == 0 else ROW_WHITE
        style_cell(table.cell(i, 0), k, 12, True, NAVY, fill)
        style_cell(table.cell(i, 1), v, 12, False, INK, fill)
    footer(s, 3, TOTAL_SLIDES)
    return s


def s04_why_ea(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Why this structure", "Cisco Enterprise Agreement for Hallmark", "One buying program across Networking, Applications, Security, and CX")

    cards = [
        ("Predictable 3-year spend", "Fixed EA rates for the 36-month term with annual true-forward. Growth is licensed at contracted EA pricing rather than one-off quotes."),
        ("Suite coverage, not SKU sprawl", "Full-commitment suites (Catalyst switching/wireless, Nexus, Meraki, ISE, Splunk, ThousandEyes) plus targeted partial suites (Spaces, Meraki cameras)."),
        ("Software + CX aligned", "EA software subscriptions and Cisco Support Standard / Enhanced services ride together so coverage, SLAs, and software entitlement stay in lockstep."),
        ("Installed-base credit", "In-use assets (IB) generate one-time discounts and uncovered-asset credits, reducing net TCV versus buying the same footprint a la carte."),
        ("Partner-delivered", "World Wide Technology is the reseller of record, with Cisco AM sponsorship and Hallmark’s Smart Account as the entitlement home."),
        ("Operational simplicity", "Annual billing, a single buying program (EA85986), and one expected book date (16 Oct 2026) in place of staggered renewals."),
    ]
    for i, (title, body) in enumerate(cards):
        col = i % 3
        row = i // 3
        x = Inches(0.40) + Inches(col * 4.25)
        y = Inches(1.45) + Inches(row * 2.70)
        add_round(s, x, y, Inches(4.05), Inches(2.50), WHITE)
        add_rect(s, x, y, Inches(4.05), Inches(0.08), GOLD if row == 0 else BLUE)
        add_textbox(s, x + Inches(0.20), y + Inches(0.25), Inches(3.65), Inches(0.55), title, 16, True, NAVY)
        add_textbox(s, x + Inches(0.20), y + Inches(0.85), Inches(3.65), Inches(1.45), body, 13, False, MUTED)
    footer(s, 4, TOTAL_SLIDES)
    return s


def s05_investment(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Investment summary", "Net 3-year contract value", "Financial Summary (Net Pricing) from EAMP export — USD")

    kpis = [
        ("3-YEAR NET TCV", "$3.15M", "USD 3,146,589.96"),
        ("SOFTWARE", "$1.54M", "49% of net TCV"),
        ("CISCO CX SERVICES", "$1.60M", "51% of net TCV"),
        ("VS. LIST", "55% off", "List $6.99M  ·  Save $3.84M"),
    ]
    for i, (lab, val, cap) in enumerate(kpis):
        kpi_card(s, Inches(0.40) + Inches(i * 3.20), Inches(1.45), Inches(3.05), Inches(1.22), lab, val, cap, GOLD if i == 0 else BLUE)

    # Two columns of narrative
    add_round(s, Inches(0.40), Inches(2.90), Inches(6.20), Inches(4.05), WHITE)
    add_textbox(s, Inches(0.65), Inches(3.05), Inches(5.8), Inches(0.35), "How the investment is built", 16, True, NAVY)
    bullet_block(
        s,
        Inches(0.65),
        Inches(3.45),
        Inches(5.8),
        Inches(3.30),
        [
            "Networking software is the largest software block at $1.03M net (Catalyst DNA, Nexus, Meraki, Spaces).",
            "Applications add Splunk Cloud (50 GB/day) and ThousandEyes (1,500 agents + 1,500 EUM users) for $437K.",
            "Security is ISE Advantage (3,450) and Essentials (1,100) at $77K software plus $115K CX.",
            "CX services ($1.60M) cover switching, wireless, Nexus, ISE, and Collaboration hardware support.",
        ],
        size=13,
    )

    add_round(s, Inches(6.80), Inches(2.90), Inches(6.10), Inches(4.05), WHITE)
    add_textbox(s, Inches(7.05), Inches(3.05), Inches(5.7), Inches(0.35), "Commercial mechanics", 16, True, NAVY)
    rows = [
        ("Annual payment (even)", "$1,048,863"),
        ("Implied monthly", "$87,405"),
        ("Subscription discount", "Up to 68% (Networking)"),
        ("Programmatic discount", "5–10% by suite"),
        ("One-time discounts", "$475,060"),
        ("IB / uncovered credits", "Applied on CX + software"),
        ("Billing model", "Annual, 36 months"),
        ("Financing", "None (cash / standard)"),
    ]
    table = add_table(s, len(rows) + 1, 2, Inches(7.05), Inches(3.45), Inches(5.60), Inches(3.25))
    set_col_widths(table, [Inches(2.7), Inches(2.9)])
    style_cell(table.cell(0, 0), "TERM", 10, True, WHITE, NAVY_HEX)
    style_cell(table.cell(0, 1), "VALUE", 10, True, WHITE, NAVY_HEX, align=PP_ALIGN.RIGHT)
    for i, (k, v) in enumerate(rows, 1):
        fill = ROW_ALT if i % 2 == 0 else ROW_WHITE
        style_cell(table.cell(i, 0), k, 11, False, INK, fill)
        style_cell(table.cell(i, 1), v, 11, True, NAVY, fill, align=PP_ALIGN.RIGHT)
    footer(s, 5, TOTAL_SLIDES)
    return s


def s06_financials(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Financial summary", "Portfolio net pricing (USD)", "Software and services as reported on the Proposal Summary sheet")

    headers = ["PORTFOLIO", "SOFTWARE", "SERVICES", "NET TCV", "LIST PRICE", "DISCOUNT", "ONE-TIME"]
    data = [
        ["Networking Infrastructure", "$1,030,043.52", "$0.00*", "$1,030,043.52", "$3,252,177.72", "$2,211,543.72", "$10,590.48"],
        ["Applications Infrastructure", "$437,058.00", "$0.00", "$437,058.00", "$650,862.00", "$213,804.00", "$0.00"],
        ["Security", "$76,726.68", "$0.00*", "$76,726.68", "$136,243.50", "$57,226.50", "$2,290.32"],
        ["Services (CX rollup)", "$0.00", "$1,602,761.76", "$1,602,761.76", "$2,948,542.56", "$883,601.88", "$462,179.16"],
        ["TOTAL", "$1,543,828.20", "$1,602,761.76", "$3,146,589.96", "$6,987,825.78", "$3,366,176.10", "$475,059.96"],
    ]
    table = add_table(s, 6, 7, Inches(0.30), Inches(1.45), Inches(12.7), Inches(3.15))
    widths = [Inches(2.70), Inches(1.70), Inches(1.70), Inches(1.75), Inches(1.70), Inches(1.60), Inches(1.55)]
    set_col_widths(table, widths)
    aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 6
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)

    add_textbox(
        s,
        Inches(0.40),
        Inches(4.75),
        Inches(12.5),
        Inches(0.35),
        "* Proposal Summary shows CX under the Services row. Portfolio sheets break those services out: Networking CX $1,127,847  ·  Security CX $115,233  ·  Collaboration CX $359,682.",
        11,
        False,
        MUTED,
    )

    # Three insight chips
    chips = [
        ("Effective discount", "54.97%", "List $6.99M → net $3.15M"),
        ("Software mix", "49 / 51", "Software $1.54M  ·  CX $1.60M"),
        ("Largest lever", "Networking", "Catalyst + Nexus + Meraki + CX"),
    ]
    for i, (lab, val, cap) in enumerate(chips):
        x = Inches(0.40) + Inches(i * 4.25)
        add_round(s, x, Inches(5.20), Inches(4.05), Inches(1.75), WHITE)
        add_textbox(s, x + Inches(0.20), Inches(5.35), Inches(3.65), Inches(0.28), lab.upper(), 11, True, MUTED)
        add_textbox(s, x + Inches(0.20), Inches(5.62), Inches(3.65), Inches(0.50), val, 26, True, NAVY)
        add_textbox(s, x + Inches(0.20), Inches(6.18), Inches(3.65), Inches(0.45), cap, 12, False, MUTED)
    footer(s, 6, TOTAL_SLIDES)
    return s


def s07_mix(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Where the investment goes", "Net TCV by portfolio (software + attached CX)", "CX from each portfolio sheet is shown with its parent technology")

    # Combined portfolio TCV including attached services
    blocks = [
        ("Networking", 1030043.52 + 1127846.52, "Software $1.03M\nCX $1.13M", NAVY),
        ("Collaboration CX", 359681.76, "Support only\nno new collab SW", BLUE),
        ("Applications", 437058.00, "Splunk + ThousandEyes", CYAN),
        ("Security", 76726.68 + 115233.48, "ISE software\n+ ISE CX", GOLD),
    ]
    total = 3146589.96
    # visual bars
    max_w = Inches(8.6)
    add_textbox(s, Inches(0.45), Inches(1.40), Inches(12), Inches(0.30), "Relative scale of 3-year net contract value", 12, True, MUTED)
    for i, (name, val, cap, color) in enumerate(blocks):
        y = Inches(1.80) + Inches(i * 1.15)
        add_textbox(s, Inches(0.45), y, Inches(2.35), Inches(0.85), name, 14, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        bar_w = int(max_w * (val / total))
        add_round(s, Inches(2.90), y + Inches(0.12), bar_w, Inches(0.48), color)
        add_textbox(s, Inches(2.90) + bar_w + Inches(0.12), y, Inches(2.4), Inches(0.48), money_short(val), 16, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(2.90), y + Inches(0.58), Inches(8.5), Inches(0.35), cap.replace("\n", "  ·  "), 11, False, MUTED)

    # right side total card
    add_round(s, Inches(10.55), Inches(1.80), Inches(2.40), Inches(4.70), NAVY)
    add_textbox(s, Inches(10.70), Inches(2.10), Inches(2.10), Inches(0.30), "NET TCV", 11, True, GOLD)
    add_textbox(s, Inches(10.70), Inches(2.45), Inches(2.10), Inches(0.90), "$3.15M", 26, True, WHITE)
    add_textbox(
        s,
        Inches(10.70),
        Inches(3.45),
        Inches(2.10),
        Inches(2.60),
        "Networking  69%\nCollab CX  11%\nApplications  14%\nSecurity  6%",
        13,
        False,
        CYAN,
    )
    footer(s, 7, TOTAL_SLIDES)
    return s


def s08_networking_overview(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking Infrastructure", "Campus, data center, Meraki, and Spaces", "Software TCV $1,030,044  ·  CX TCV $1,127,847  ·  36 months  ·  Full commitment unless noted")

    headers = ["SUITE", "COMMIT", "KEY ENTITLEMENT", "QTY", "SW TCV"]
    data = [
        ["Meraki – Network Infrastructure", "Full", "MX XL Essentials, MR, MS100/MS300", "216", "$45,702"],
        ["Meraki – Camera Systems", "Partial", "MV Large Essentials", "2", "$415"],
        ["Cisco Networking Wireless", "Full", "CW Advantage + Essential", "2", "$380"],
        ["Cisco Networking Switching", "Full", "Access T1 Large/Medium Essential", "91", "$57,330"],
        ["Nexus Switching", "Full", "N9300 XF/XF2 Advantage + Essential", "36", "$304,670"],
        ["Cisco Spaces (wireless add-on)", "Partial", "Spaces ACT subscription", "350", "$64,638"],
        ["Cisco DNA Switching", "Full", "C3850/C9300/C9500/C9200 DNA", "298", "$449,547"],
        ["Cisco DNA Wireless", "Full", "DNA Wireless Advantage", "525", "$107,361"],
        ["TOTAL SOFTWARE", "", "", "", "$1,030,044"],
    ]
    table = add_table(s, 10, 5, Inches(0.35), Inches(1.42), Inches(12.6), Inches(5.50))
    set_col_widths(table, [Inches(3.55), Inches(1.20), Inches(4.35), Inches(1.15), Inches(2.35)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=11, body_size=12)
    footer(s, 8, TOTAL_SLIDES)
    return s


def s09_networking_bom(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking software — detail 1 of 2", "Meraki, Catalyst Networking, and Nexus", "Qty Found = installed base  ·  Qty Desired = EA quantity  ·  68% subscription discount on most PIDs")

    headers = ["PID", "DESCRIPTION", "TIER", "IB", "QTY", "NET TCV"]
    data = [
        ["E3N-MX-XL-E", "Meraki MX X-Large Essentials", "—", "2", "3", "$17,495"],
        ["E3N-MR-E", "Meraki MR Essentials", "—", "77", "193", "$23,345"],
        ["E3N-MS-100-S/M/L-E", "Meraki MS100 Small/Med/Large Essentials", "—", "0", "16", "$1,390"],
        ["E3N-MS-300-M/L-A", "Meraki MS300 Medium/Large Advantage", "Adv", "4", "4", "$3,473"],
        ["E3N-MV-E", "Meraki MV Large Essentials", "—", "0", "2", "$415"],
        ["E3N-CW-A / E3N-CW-E", "Cisco Networking Wireless", "Adv/Ess", "0", "2", "$380"],
        ["E3N-CS-AC1-L-E", "Switching Access T1 Large Essential", "Ess", "0", "49", "$39,126"],
        ["E3N-CS-AC1-M-E", "Switching Access T1 Medium Essential", "Ess", "0", "42", "$18,204"],
        ["E3N-N9300-XF2-A", "Nexus 9300 XF2 or higher (to 6.4T)", "Adv", "2", "2", "$35,440"],
        ["E3N-N9300-XF-A", "Nexus 9300 XF 10G+ (to 3.6T)", "Adv", "22", "30", "$243,119"],
        ["E3N-N9300-XF-E", "Nexus 9300 XF 10G+ (to 3.6T)", "Ess", "0", "4", "$26,112"],
        ["E3N-SPACES-ACT", "Cisco Spaces ACT subscription", "Add-on", "0", "350", "$64,638"],
    ]
    table = add_table(s, 13, 6, Inches(0.30), Inches(1.40), Inches(12.7), Inches(5.55))
    set_col_widths(table, [Inches(2.45), Inches(4.55), Inches(1.15), Inches(0.85), Inches(0.95), Inches(1.75)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)
    footer(s, 9, TOTAL_SLIDES)
    return s


def s10_networking_dna(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking software — detail 2 of 2", "Cisco DNA Switching and Wireless", "Full-commitment DNA suites  ·  C3850 48-port Advantage is the single largest software PID")

    headers = ["PID", "DESCRIPTION", "TIER", "IB", "QTY", "NET TCV"]
    data = [
        ["E3N-C38502-A", "C3850 48-Port DNA EA", "Adv", "0", "218", "$306,464"],
        ["E3N-C95005-A", "C9500 48Y4C DNA EA", "Adv", "43", "16", "$70,791"],
        ["E3N-C93002-A", "C9300/C9300X 48-Port DNA EA", "Adv", "43", "19", "$26,619"],
        ["E3N-C95003-A", "C9500 32C DNA EA", "Adv", "4", "4", "$18,048"],
        ["E3N-C9200CX2-A", "C9200CX 12-Port DNA EA", "Adv", "24", "18", "$6,690"],
        ["E3N-C36502-A", "C3650 48-Port DNA EA", "Adv", "0", "4", "$5,623"],
        ["E3N-C9300X2-A", "C9300X 24-Port DNA EA", "Adv", "6", "6", "$4,497"],
        ["E3N-C95002-A", "C9500 Low Port (12Q/16X) DNA EA", "Adv", "0", "2", "$5,296"],
        ["E3N-C38501-A", "C3850 24-Port DNA EA", "Adv", "0", "3", "$2,249"],
        ["E3N-C95005-E", "C9500 48Y4C DNA EA", "Ess", "3", "3", "$1,863"],
        ["E3N-C93001-A / others", "C9300 24 / C3560CX / C9200L 48", "Mix", "18", "7", "$1,407"],
        ["E3N-AIRWLAN-A", "Cisco DNA Wireless", "Adv", "132", "525", "$107,361"],
        ["TOTAL DNA + WIRELESS", "", "", "", "", "$556,908"],
    ]
    table = add_table(s, 14, 6, Inches(0.30), Inches(1.40), Inches(12.7), Inches(5.55))
    set_col_widths(table, [Inches(2.45), Inches(4.55), Inches(1.15), Inches(0.85), Inches(0.95), Inches(1.75)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)
    footer(s, 10, TOTAL_SLIDES)
    return s


def s11_networking_cx(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking CX", "Support & lifecycle services", "Portfolio services TCV $1,127,847  ·  Cisco Support Standard unless noted Enhanced")

    headers = ["PID", "COVERAGE", "IB ASSETS", "DISC %", "NET TCV"]
    data = [
        ["E3-CX-ST-T1NBD", "Switching  8x5xNBD", "120", "35%", "$385,433"],
        ["E3-CX-ST-T1NCD", "Switching  8x7xNCD", "1,053", "35%", "$246,936"],
        ["E3-CX-ST-T1SWT", "Switching  software support", "—", "35%", "$138,048"],
        ["E3-CX-CS-ENBD", "Cisco Switching  8x5xNBD", "91", "23%", "$81,638"],
        ["E3-CX-DCN-T14HR", "Nexus  24x7x4", "10", "21%", "$83,260"],
        ["E3-CX-ST-T14HR", "Switching  24x7x4", "7", "35%", "$57,020"],
        ["E3-CXST-AIR-T1SWP", "Wireless  software support", "—", "35%", "$33,114"],
        ["E3-CX-DCN-L1SWPERP", "Nexus  perpetual", "24", "22%", "$27,504"],
        ["E3-CX-DCN-T1NBD", "Nexus  8x5xNBD", "26", "30%", "$23,171"],
        ["E3-CX-DNS-T1SWC", "Spaces / DNAS cloud SW (Enhanced)", "—", "35%", "$19,705"],
        ["E3-CX-DCN-T1NCD", "Nexus  8x7xNCD", "16", "23%", "$15,612"],
        ["E3-CX-ST-T1NOS", "Switching  8x5xNBD OS", "7", "35%", "$12,789"],
        ["E3-CX-CW-ENCD", "Cisco Wireless  8x7xNCD", "16", "31%", "$3,615"],
        ["TOTAL NETWORKING CX", "", "", "", "$1,127,847"],
    ]
    table = add_table(s, 15, 5, Inches(0.35), Inches(1.40), Inches(12.6), Inches(5.55))
    set_col_widths(table, [Inches(2.70), Inches(4.40), Inches(1.55), Inches(1.30), Inches(2.65)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)
    footer(s, 11, TOTAL_SLIDES)
    return s


def s12_apps(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Applications Infrastructure", "Splunk Cloud and ThousandEyes", "Software TCV $437,058  ·  No attached CX on this portfolio  ·  Full commitment  ·  5% programmatic")

    # KPI row
    for i, (lab, val, cap, acc) in enumerate([
        ("PORTFOLIO TCV", "$437,058", "100% software", GOLD),
        ("SPLUNK CLOUD", "$90,918", "50 GB/day  ·  42% sub disc.", BLUE),
        ("THOUSANDEYES", "$346,140", "Agents + EUM  ·  30% sub disc.", CYAN),
        ("SUCCESS PLAN", "Included", "Splunk Standard Success", GREEN),
    ]):
        kpi_card(s, Inches(0.40) + Inches(i * 3.20), Inches(1.42), Inches(3.05), Inches(1.22), lab, val, cap, acc)

    headers = ["PID", "DESCRIPTION", "QTY", "UNIT LIST / MO", "SUB DISC.", "NET TCV"]
    data = [
        ["E3A-SK-SE-CLD-S", "Splunk Cloud Subscription — Std success (GB/day)", "50", "$87.09", "42%", "$90,918"],
        ["E3-CX-SK-SUP-ST", "Splunk Standard Success Plan", "1", "$0.00", "—", "$0"],
        ["E3A-TE-UNITS", "ThousandEyes Cloud & Enterprise Agents (per unit)", "1,500", "$0.84", "30%", "$31,860"],
        ["E3A-TE-USERS", "ThousandEyes End User Monitoring Advantage", "1,500", "$8.31", "30%", "$314,280"],
        ["E3A-TE-S", "Cisco Support Basic for EA ThousandEyes", "1", "$0.00", "—", "$0"],
        ["TOTAL", "", "", "", "", "$437,058"],
    ]
    table = add_table(s, 7, 6, Inches(0.35), Inches(2.90), Inches(12.6), Inches(3.05))
    set_col_widths(table, [Inches(2.20), Inches(4.70), Inches(1.10), Inches(1.70), Inches(1.30), Inches(1.60)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=12)

    add_textbox(
        s,
        Inches(0.45),
        Inches(6.10),
        Inches(12.4),
        Inches(0.85),
        "Use case fit: Splunk Cloud gives Hallmark a 50 GB/day security/observability ingest baseline on-contract. ThousandEyes (1,500 cloud/enterprise units + 1,500 EUM seats) extends digital-experience visibility across retail, HQ, and partner paths — aligned to the CX EA motion.",
        13,
        False,
        MUTED,
    )
    footer(s, 12, TOTAL_SLIDES)
    return s


def s13_security(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Security", "Identity Services Engine (Zero Trust)", "SCU count 1,000  ·  Software $76,727  ·  CX $115,233  ·  Combined $191,960  ·  Full commitment")

    for i, item in enumerate([
        ("ISE ADVANTAGE", "3,450", "IB 400  →  desired 3,450", GOLD),
        ("ISE ESSENTIALS", "1,100", "IB 300  →  desired 1,100", BLUE),
        ("SOFTWARE TCV", "$76,727", "42% subscription discount", CYAN),
        ("ISE CX TCV", "$115,233", "Support Standard + credits", GREEN),
    ]):
        kpi_card(s, Inches(0.40) + Inches(i * 3.20), Inches(1.42), Inches(3.05), Inches(1.22), *item)

    headers = ["PID", "DESCRIPTION", "IB", "QTY", "DISC.", "NET TCV"]
    data = [
        ["E3S-ISE-ADV", "ISE Advantage", "400", "3,450", "42%", "$70,470"],
        ["E3S-ISE-ESS", "ISE Essentials", "300", "1,100", "42%", "$6,256"],
        ["SVS-E3S-ISE-B", "Cisco Support Basic for ISE", "0", "1", "23%", "$0"],
        ["E3-CX-ISE-ENBD", "Support Standard  8x5xNBD for ISE", "6", "1", "23%", "$51,486"],
        ["E3-CX-ISE-EPER", "Cisco Support Standard Perpetual for ISE", "20", "1", "24%", "$48,012"],
        ["E3-CX-ISE-ESWP", "Support Standard SW Support OP for ISE", "0", "1", "23%", "$15,736"],
        ["E3-CX-ISE-ENCD", "Support Standard  8x7xNCD for ISE", "1", "1", "23%", "$0"],
        ["TOTAL SECURITY", "", "", "", "", "$191,960"],
    ]
    table = add_table(s, 9, 6, Inches(0.35), Inches(2.88), Inches(12.6), Inches(3.95))
    set_col_widths(table, [Inches(2.30), Inches(4.70), Inches(1.05), Inches(1.20), Inches(1.15), Inches(2.20)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=12)
    footer(s, 13, TOTAL_SLIDES)
    return s


def s14_collab(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Collaboration", "CX only — Employee Experience support", "No new Collaboration software on this proposal  ·  Services TCV $359,682  ·  Full commitment")

    add_round(s, Inches(0.40), Inches(1.45), Inches(12.5), Inches(1.35), WHITE)
    add_textbox(s, Inches(0.65), Inches(1.58), Inches(12.0), Inches(0.35), "What is in scope", 16, True, NAVY)
    add_textbox(
        s,
        Inches(0.65),
        Inches(1.95),
        Inches(12.0),
        Inches(0.65),
        "This EA continues Cisco Support Standard on the existing Collaboration installed base (Employee Experience). There is no Webex or Calling software subscription on Proposal 9325586. Uncovered-asset credits and a program-migration incentive reduce net CX versus list.",
        13,
        False,
        MUTED,
    )

    headers = ["PID", "DESCRIPTION", "IB", "QTY", "LIST", "CREDITS / OTD", "NET TCV"]
    data = [
        ["E3-CX-COL-ENBD", "Support Standard  8x5xNBD  COLLAB", "110", "1", "$414,060", "Migration $44,362", "$273,770"],
        ["E3-CX-COL-ENCD", "Support Standard  8x7xNCD  COLLAB", "92", "1", "$145,359", "Uncovered $21,425", "$85,912"],
        ["TOTAL", "", "202", "", "$559,419", "", "$359,682"],
    ]
    table = add_table(s, 4, 7, Inches(0.35), Inches(3.05), Inches(12.6), Inches(2.15))
    set_col_widths(table, [Inches(2.15), Inches(3.55), Inches(0.85), Inches(0.80), Inches(1.35), Inches(1.85), Inches(2.05)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=12)

    add_round(s, Inches(0.40), Inches(5.45), Inches(12.5), Inches(1.50), WHITE)
    add_textbox(s, Inches(0.65), Inches(5.58), Inches(12.0), Inches(0.30), "Why keep Collaboration CX on the EA", 15, True, NAVY)
    add_textbox(
        s,
        Inches(0.65),
        Inches(5.95),
        Inches(12.0),
        Inches(0.80),
        "Folding 202 Collaboration support assets into the same 36-month EA avoids a separate TAC/hardware contract, applies EA service discounts (23–26%) plus IB credits, and keeps one partner (WWT) and one book date for campus, DC, ISE, and collab hardware support.",
        13,
        False,
        MUTED,
    )
    footer(s, 14, TOTAL_SLIDES)
    return s


def s15_schedule(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Commercial schedule", "Annual billing over 36 months", "Even annual cash view of net TCV — actual invoices follow Cisco EA billing rules and true-forward")

    headers = ["PERIOD", "SOFTWARE", "SERVICES", "ANNUAL NET", "CUMULATIVE"]
    data = [
        ["Year 1  (from 16 Oct 2026)", "$514,609", "$534,254", "$1,048,863", "$1,048,863"],
        ["Year 2", "$514,609", "$534,254", "$1,048,863", "$2,097,727"],
        ["Year 3", "$514,609", "$534,254", "$1,048,863", "$3,146,590"],
        ["3-YEAR TCV", "$1,543,828", "$1,602,762", "$3,146,590", "$3,146,590"],
    ]
    table = add_table(s, 5, 5, Inches(0.45), Inches(1.45), Inches(12.4), Inches(2.55))
    set_col_widths(table, [Inches(3.30), Inches(2.25), Inches(2.25), Inches(2.30), Inches(2.30)])
    aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 4
    fill_table(table, headers, data, col_align=aligns, header_size=12, body_size=14)

    notes = [
        ("True-forward", "EA quantities are a floor. Additional consumption during the year is licensed at EA rates and billed at the annual true-forward — no surprise list-price true-ups."),
        ("IB credits", "One-time discounts ($475K) and uncovered-asset credits are already netted into TCV. They are not an extra year-1 concession on top of these figures."),
        ("Partner path", "WWT is reseller of record (services bill-to 1001363275). Cisco AM Robin Randolph sponsors the deal through expected book date 16 Oct 2026."),
    ]
    for i, (t, b) in enumerate(notes):
        x = Inches(0.45) + Inches(i * 4.20)
        add_round(s, x, Inches(4.25), Inches(4.00), Inches(2.65), WHITE)
        add_rect(s, x, Inches(4.25), Inches(4.00), Inches(0.08), GOLD if i == 0 else BLUE)
        add_textbox(s, x + Inches(0.20), Inches(4.45), Inches(3.60), Inches(0.40), t, 15, True, NAVY)
        add_textbox(s, x + Inches(0.20), Inches(4.90), Inches(3.60), Inches(1.80), b, 12, False, MUTED)
    footer(s, 15, TOTAL_SLIDES)
    return s


def s16_value(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Value", "Discount architecture versus list", "List $6,987,826  ·  Programmatic + subscription discounts $3,366,176  ·  One-time $475,060")

    headers = ["PORTFOLIO", "LIST", "NET", "SAVINGS", "EFFECTIVE %"]
    data = [
        ["Networking software", "$3,252,178", "$1,030,044", "$2,222,134", "68%"],
        ["Applications software", "$650,862", "$437,058", "$213,804", "33%"],
        ["Security software", "$136,244", "$76,727", "$59,517", "44%"],
        ["CX services (all)", "$2,948,543", "$1,602,762", "$1,345,781", "46%"],
        ["TOTAL", "$6,987,826", "$3,146,590", "$3,841,236", "55%"],
    ]
    table = add_table(s, 6, 5, Inches(0.45), Inches(1.42), Inches(12.4), Inches(2.85))
    set_col_widths(table, [Inches(3.40), Inches(2.25), Inches(2.25), Inches(2.25), Inches(2.25)])
    aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 4
    fill_table(table, headers, data, col_align=aligns, header_size=12, body_size=13)

    points = [
        ("68% Networking software", "EA 3.0 suite discount (typically 68% after 10% programmatic) on Catalyst, Nexus, Meraki, and Spaces — the primary savings engine."),
        ("CX at 23–35%", "Support Standard discounts plus program-migration incentives and uncovered-asset credits on switching, Nexus, ISE, and Collaboration."),
        ("Apps at 30–42%", "ThousandEyes 30% and Splunk Cloud 42% subscription discounts on a full-commitment Applications suite."),
    ]
    for i, (t, b) in enumerate(points):
        x = Inches(0.45) + Inches(i * 4.20)
        add_round(s, x, Inches(4.50), Inches(4.00), Inches(2.40), WHITE)
        add_textbox(s, x + Inches(0.20), Inches(4.65), Inches(3.60), Inches(0.45), t, 14, True, NAVY)
        add_textbox(s, x + Inches(0.20), Inches(5.15), Inches(3.60), Inches(1.50), b, 12, False, MUTED)
    footer(s, 16, TOTAL_SLIDES)
    return s


def s17_next(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Path to book", "Recommended next steps", "Target book / RSD: 16 October 2026")

    steps = [
        ("1", "Confirm quantities", "Validate Qty Desired vs. IB for Catalyst DNA (especially C3850 x218 and DNA Wireless x525), ISE 3,450/1,100, and ThousandEyes 1,500/1,500."),
        ("2", "Lock suite mix", "Keep Full commitment on core suites; confirm Partial is still correct for Spaces ACT and Meraki cameras."),
        ("3", "WWT quote path", "Convert this EAMP estimate to a WWT / Cisco approved quote. This file is indicative only."),
        ("4", "Smart Account", "Confirm HALLMARK CARDS, INCORPORATED Smart Account is the entitlement destination before book."),
        ("5", "Legal / EA paper", "Standard Cisco EA 3.0 terms, true-forward, and CX service descriptions via WWT."),
        ("6", "Book 16 Oct 2026", "Align PO, annual billing, and RSD so software and CX start together."),
    ]
    for i, (n, title, body) in enumerate(steps):
        col = i % 3
        row = i // 3
        x = Inches(0.40) + Inches(col * 4.25)
        y = Inches(1.45) + Inches(row * 2.70)
        add_round(s, x, y, Inches(4.05), Inches(2.50), WHITE)
        add_round(s, x + Inches(0.20), y + Inches(0.22), Inches(0.48), Inches(0.48), NAVY)
        add_textbox(s, x + Inches(0.20), y + Inches(0.22), Inches(0.48), Inches(0.48), n, 16, True, WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, x + Inches(0.80), y + Inches(0.28), Inches(3.00), Inches(0.40), title, 16, True, NAVY)
        add_textbox(s, x + Inches(0.20), y + Inches(0.85), Inches(3.65), Inches(1.45), body, 13, False, MUTED)
    footer(s, 17, TOTAL_SLIDES)
    return s


def s18_close(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, GOLD)
    add_textbox(s, Inches(0.7), Inches(0.85), Inches(12), Inches(0.30), "HALLMARK CX EA  ·  PROPOSAL 9325586", 13, True, GOLD)
    add_textbox(s, Inches(0.7), Inches(1.25), Inches(12), Inches(0.70), "Ready when Hallmark is.", 32, True, WHITE)
    add_textbox(
        s,
        Inches(0.7),
        Inches(2.05),
        Inches(12),
        Inches(0.55),
        "A 36-month Enterprise Agreement covering campus, data center, Meraki, Spaces,\nISE, Splunk, ThousandEyes, and Cisco CX — net $3.15M versus $6.99M list.",
        16,
        False,
        CYAN,
    )

    contacts = [
        ("CUSTOMER", "Hallmark Cards, Incorporated\n2501 McGee Traffic Way\nKansas City, MO 64141"),
        ("PARTNER", "World Wide Technology, Inc\nReseller of record\nServices bill-to 1001363275"),
        ("CISCO", "Robin Randolph\nCisco Account Manager\nrobrando@cisco.com"),
    ]
    for i, (h, body) in enumerate(contacts):
        x = Inches(0.7) + Inches(i * 4.05)
        add_rect(s, x, Inches(3.00), Inches(3.70), Inches(0.06), GOLD)
        add_textbox(s, x, Inches(3.20), Inches(3.70), Inches(0.30), h, 12, True, GOLD)
        add_textbox(s, x, Inches(3.52), Inches(3.70), Inches(1.40), body, 14, False, WHITE)

    add_textbox(
        s,
        Inches(0.7),
        Inches(5.40),
        Inches(12),
        Inches(1.40),
        "Disclaimer: This presentation is derived from Cisco EAMP Price Estimate ID 9325586, generated 08-Sep-2026.\n"
        "It contains indicative pricing only as of that date and is NOT an approved Cisco quote. Pricing is subject to change.\n"
        "CISCO CONFIDENTIAL — do not distribute outside Hallmark, World Wide Technology, and Cisco deal teams.",
        12,
        False,
        RGBColor(0xC5, 0xD4, 0xE3),
    )
    return s


def build(path: str):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    s01_title(prs)
    s02_agenda(prs)
    s03_snapshot(prs)
    s04_why_ea(prs)
    s05_investment(prs)
    s06_financials(prs)
    s07_mix(prs)
    s08_networking_overview(prs)
    s09_networking_bom(prs)
    s10_networking_dna(prs)
    s11_networking_cx(prs)
    s12_apps(prs)
    s13_security(prs)
    s14_collab(prs)
    s15_schedule(prs)
    s16_value(prs)
    s17_next(prs)
    s18_close(prs)

    # Verify slide count matches footer
    actual = len(prs.slides)
    if actual != TOTAL_SLIDES:
        raise SystemExit(f"Slide count mismatch: built {actual}, footers assume {TOTAL_SLIDES}")

    prs.save(path)
    print(f"Wrote {path} ({actual} slides)")


if __name__ == "__main__":
    import os

    out = os.path.join(os.path.dirname(__file__), "..", "Hallmark_Cards_Cisco_Enterprise_Agreement.pptx")
    out = os.path.abspath(out)
    build(out)
