#!/usr/bin/env python3
"""Build the Hallmark Cards Cisco Enterprise Agreement proposal PPTX.

Pricing is sourced from EAMP Price Estimate ID 9325586 (indicative only;
not an approved Cisco quote). Customer-facing figures include 22.5 points
of partner gross margin on Cisco EA net (sell = net / 0.775).
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

# 22.5 points of gross margin on Cisco EA net (not a 22.5% markup).
# Hallmark price = Cisco net / (1 - 0.225) = Cisco net / 0.775
MARGIN_PTS = 22.5
MARGIN = MARGIN_PTS / 100.0


def sell(cisco_net: float) -> float:
    """Customer / partner sell price at 22.5 pts GM."""
    return cisco_net / (1.0 - MARGIN)


def gm_dollars(cisco_net: float) -> float:
    return sell(cisco_net) - cisco_net


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
        "CISCO CONFIDENTIAL  |  Indicative  ·  22.5 pts GM on Cisco EA net  |  Not an approved Cisco quote",
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

TOTAL_SLIDES = 19


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
        ("HALLMARK TCV", "$4.06M"),
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
        "CISCO CONFIDENTIAL  ·  Indicative  ·  22.5 pts partner GM on Cisco EA net  ·  Not an approved Cisco quote",
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
        ("02", "Investment with 22.5 pts margin", "Cisco EA net, partner GM, Hallmark price vs. list"),
        ("03", "Portfolio architecture", "Networking, Applications, Security, and Collaboration CX"),
        ("04", "Bill of materials", "Suite commitments, quantities, and Hallmark TCV"),
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
        ("Customer pricing", "Cisco EA net + 22.5 pts partner gross margin  ·  Hallmark TCV $4.06M"),
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


def s05_margin(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(
        s,
        "Commercial wrap",
        "22.5 points of partner margin",
        "Gross margin on Cisco EA net  ·  Hallmark price = Cisco net ÷ 0.775  ·  Applied to every line in this deck",
    )

    cisco = 3146589.96
    hallmark = sell(cisco)
    margin = gm_dollars(cisco)

    cards = [
        ("CISCO EA NET", money_short(cisco), money(cisco), BLUE),
        ("PARTNER GM", "22.5 pts", money(margin) + "  ·  29.0% markup", GOLD),
        ("HALLMARK PRICE", money_short(hallmark), money(hallmark), GOLD),
        ("VS. CISCO LIST", "41.9% off", "List $6.99M  ·  Save $2.93M", GREEN),
    ]
    for i, (lab, val, cap, acc) in enumerate(cards):
        kpi_card(s, Inches(0.40) + Inches(i * 3.20), Inches(1.42), Inches(3.05), Inches(1.22), lab, val, cap, acc)

    headers = ["PORTFOLIO", "CISCO EA NET", "22.5 PTS GM", "HALLMARK PRICE"]
    data = [
        ["Networking software", money(1030043.52), money(gm_dollars(1030043.52)), money(sell(1030043.52))],
        ["Applications software", money(437058.00), money(gm_dollars(437058.00)), money(sell(437058.00))],
        ["Security software", money(76726.68), money(gm_dollars(76726.68)), money(sell(76726.68))],
        ["CX services (all)", money(1602761.76), money(gm_dollars(1602761.76)), money(sell(1602761.76))],
        ["TOTAL", money(cisco), money(margin), money(hallmark)],
    ]
    table = add_table(s, 6, 4, Inches(0.45), Inches(2.88), Inches(12.4), Inches(2.85))
    set_col_widths(table, [Inches(3.40), Inches(3.00), Inches(3.00), Inches(3.00)])
    aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 3
    fill_table(table, headers, data, col_align=aligns, header_size=12, body_size=13)

    add_textbox(
        s,
        Inches(0.50),
        Inches(5.90),
        Inches(12.3),
        Inches(1.05),
        "All software and services TCV figures from this slide forward are Hallmark (sell) prices. Cisco list, subscription-discount %, IB credits, and one-time discounts are unchanged from the EAMP export — only the customer-facing contract value includes the 22.5-point wrap. True-forward of additional quantity will use the same margin method.",
        13,
        False,
        MUTED,
    )
    footer(s, 5, TOTAL_SLIDES)
    return s


def s05_investment(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Investment summary", "Hallmark 3-year contract value", "Includes 22.5 pts partner GM on Cisco EA net  ·  USD")

    kpis = [
        ("3-YEAR HALLMARK TCV", "$4.06M", money(sell(3146589.96))),
        ("SOFTWARE", "$1.99M", "49% of Hallmark TCV"),
        ("CISCO CX SERVICES", "$2.07M", "51% of Hallmark TCV"),
        ("VS. LIST", "42% off", "List $6.99M  ·  Save $2.93M"),
    ]
    for i, (lab, val, cap) in enumerate(kpis):
        kpi_card(s, Inches(0.40) + Inches(i * 3.20), Inches(1.45), Inches(3.05), Inches(1.22), lab, val, cap, GOLD if i == 0 else BLUE)

    add_round(s, Inches(0.40), Inches(2.90), Inches(6.20), Inches(4.05), WHITE)
    add_textbox(s, Inches(0.65), Inches(3.05), Inches(5.8), Inches(0.35), "How the investment is built", 16, True, NAVY)
    bullet_block(
        s,
        Inches(0.65),
        Inches(3.45),
        Inches(5.8),
        Inches(3.30),
        [
            f"Networking software is the largest software block at {money_short(sell(1030043.52))} (Catalyst DNA, Nexus, Meraki, Spaces).",
            f"Applications add Splunk Cloud (50 GB/day) and ThousandEyes (1,500 agents + 1,500 EUM users) for {money_short(sell(437058.00))}.",
            f"Security is ISE Advantage (3,450) and Essentials (1,100) at {money_short(sell(76726.68))} software plus {money_short(sell(115233.48))} CX.",
            f"CX services ({money_short(sell(1602761.76))}) cover switching, wireless, Nexus, ISE, and Collaboration hardware support.",
        ],
        size=13,
    )

    add_round(s, Inches(6.80), Inches(2.90), Inches(6.10), Inches(4.05), WHITE)
    add_textbox(s, Inches(7.05), Inches(3.05), Inches(5.7), Inches(0.35), "Commercial mechanics", 16, True, NAVY)
    rows = [
        ("Annual payment (even)", money(sell(3146589.96) / 3, cents=False)),
        ("Implied monthly", money(sell(3146589.96) / 36, cents=False)),
        ("Partner margin", "22.5 pts GM  ($913,526)"),
        ("Subscription discount", "Up to 68% (Networking, Cisco net)"),
        ("Programmatic discount", "5–10% by suite (Cisco net)"),
        ("One-time discounts", "$475,060  (Cisco net, already in)"),
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
    footer(s, 6, TOTAL_SLIDES)
    return s


def s06_financials(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Financial summary", "Portfolio Hallmark pricing (USD)", "Cisco EA net + 22.5 pts GM  ·  List prices unchanged from EAMP")

    headers = ["PORTFOLIO", "SOFTWARE", "SERVICES", "HALLMARK TCV", "CISCO NET", "LIST PRICE"]
    data = [
        ["Networking Infrastructure", money(sell(1030043.52)), "$0.00*", money(sell(1030043.52)), money(1030043.52), "$3,252,177.72"],
        ["Applications Infrastructure", money(sell(437058.00)), "$0.00", money(sell(437058.00)), money(437058.00), "$650,862.00"],
        ["Security", money(sell(76726.68)), "$0.00*", money(sell(76726.68)), money(76726.68), "$136,243.50"],
        ["Services (CX rollup)", "$0.00", money(sell(1602761.76)), money(sell(1602761.76)), money(1602761.76), "$2,948,542.56"],
        ["TOTAL", money(sell(1543828.20)), money(sell(1602761.76)), money(sell(3146589.96)), money(3146589.96), "$6,987,825.78"],
    ]
    table = add_table(s, 6, 6, Inches(0.28), Inches(1.45), Inches(12.75), Inches(3.15))
    widths = [Inches(2.85), Inches(1.95), Inches(1.95), Inches(2.15), Inches(1.90), Inches(1.95)]
    set_col_widths(table, widths)
    aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 5
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)

    add_textbox(
        s,
        Inches(0.40),
        Inches(4.75),
        Inches(12.5),
        Inches(0.35),
        "* CX is rolled up on the Services row. Hallmark CX: Networking "
        + money(sell(1127846.52), cents=False)
        + "  ·  Security "
        + money(sell(115233.48), cents=False)
        + "  ·  Collaboration "
        + money(sell(359681.76), cents=False)
        + ".",
        11,
        False,
        MUTED,
    )

    chips = [
        ("Effective discount", "41.9%", "List $6.99M → Hallmark $4.06M"),
        ("Software mix", "49 / 51", f"SW {money_short(sell(1543828.20))}  ·  CX {money_short(sell(1602761.76))}"),
        ("Partner wrap", "22.5 pts", f"GM {money_short(gm_dollars(3146589.96))} on Cisco net"),
    ]
    for i, (lab, val, cap) in enumerate(chips):
        x = Inches(0.40) + Inches(i * 4.25)
        add_round(s, x, Inches(5.20), Inches(4.05), Inches(1.75), WHITE)
        add_textbox(s, x + Inches(0.20), Inches(5.35), Inches(3.65), Inches(0.28), lab.upper(), 11, True, MUTED)
        add_textbox(s, x + Inches(0.20), Inches(5.62), Inches(3.65), Inches(0.50), val, 26, True, NAVY)
        add_textbox(s, x + Inches(0.20), Inches(6.18), Inches(3.65), Inches(0.45), cap, 12, False, MUTED)
    footer(s, 7, TOTAL_SLIDES)
    return s


def s07_mix(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Where the investment goes", "Hallmark TCV by portfolio (software + attached CX)", "CX from each portfolio sheet is shown with its parent technology  ·  22.5 pts GM included")

    blocks = [
        ("Networking", sell(1030043.52 + 1127846.52), f"Software {money_short(sell(1030043.52))}\nCX {money_short(sell(1127846.52))}", NAVY),
        ("Collaboration CX", sell(359681.76), "Support only\nno new collab SW", BLUE),
        ("Applications", sell(437058.00), "Splunk + ThousandEyes", CYAN),
        ("Security", sell(76726.68 + 115233.48), "ISE software\n+ ISE CX", GOLD),
    ]
    total = sell(3146589.96)
    max_w = Inches(8.6)
    add_textbox(s, Inches(0.45), Inches(1.40), Inches(12), Inches(0.30), "Relative scale of 3-year Hallmark contract value", 12, True, MUTED)
    for i, (name, val, cap, color) in enumerate(blocks):
        y = Inches(1.80) + Inches(i * 1.15)
        add_textbox(s, Inches(0.45), y, Inches(2.35), Inches(0.85), name, 14, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        bar_w = int(max_w * (val / total))
        add_round(s, Inches(2.90), y + Inches(0.12), bar_w, Inches(0.48), color)
        add_textbox(s, Inches(2.90) + bar_w + Inches(0.12), y, Inches(2.4), Inches(0.48), money_short(val), 16, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(2.90), y + Inches(0.58), Inches(8.5), Inches(0.35), cap.replace("\n", "  ·  "), 11, False, MUTED)

    add_round(s, Inches(10.55), Inches(1.80), Inches(2.40), Inches(4.70), NAVY)
    add_textbox(s, Inches(10.70), Inches(2.10), Inches(2.10), Inches(0.30), "HALLMARK TCV", 11, True, GOLD)
    add_textbox(s, Inches(10.70), Inches(2.45), Inches(2.10), Inches(0.90), "$4.06M", 26, True, WHITE)
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
    footer(s, 8, TOTAL_SLIDES)
    return s


def s08_networking_overview(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking Infrastructure", "Campus, data center, Meraki, and Spaces", f"Software {money(sell(1030043.52), cents=False)}  ·  CX {money(sell(1127846.52), cents=False)}  ·  36 months  ·  22.5 pts GM included")

    headers = ["SUITE", "COMMIT", "KEY ENTITLEMENT", "QTY", "SW TCV"]
    data = [
        ["Meraki – Network Infrastructure", "Full", "MX XL Essentials, MR, MS100/MS300", "216", money(sell(17494.92+23344.56+342.36+529.20+518.40+1458.72+2014.20), cents=False)],
        ["Meraki – Camera Systems", "Partial", "MV Large Essentials", "2", money(sell(414.72), cents=False)],
        ["Cisco Networking Wireless", "Full", "CW Advantage + Essential", "2", money(sell(380.16), cents=False)],
        ["Cisco Networking Switching", "Full", "Access T1 Large/Medium Essential", "91", money(sell(57330.00), cents=False)],
        ["Nexus Switching", "Full", "N9300 XF/XF2 Advantage + Essential", "36", money(sell(304670.16), cents=False)],
        ["Cisco Spaces (wireless add-on)", "Partial", "Spaces ACT subscription", "350", money(sell(64638.00), cents=False)],
        ["Cisco DNA Switching", "Full", "C3850/C9300/C9500/C9200 DNA", "298", money(sell(449546.76), cents=False)],
        ["Cisco DNA Wireless", "Full", "DNA Wireless Advantage", "525", money(sell(107361.36), cents=False)],
        ["TOTAL SOFTWARE", "", "", "", money(sell(1030043.52), cents=False)],
    ]
    table = add_table(s, 10, 5, Inches(0.35), Inches(1.42), Inches(12.6), Inches(5.50))
    set_col_widths(table, [Inches(3.55), Inches(1.20), Inches(4.35), Inches(1.15), Inches(2.35)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=11, body_size=12)
    footer(s, 9, TOTAL_SLIDES)
    return s


def s09_networking_bom(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking software — detail 1 of 2", "Meraki, Catalyst Networking, and Nexus", "Hallmark TCV includes 22.5 pts GM  ·  Qty Found = IB  ·  68% Cisco subscription discount on most PIDs")

    headers = ["PID", "DESCRIPTION", "TIER", "IB", "QTY", "HALLMARK TCV"]
    data = [
        ["E3N-MX-XL-E", "Meraki MX X-Large Essentials", "—", "2", "3", money(sell(17494.92), cents=False)],
        ["E3N-MR-E", "Meraki MR Essentials", "—", "77", "193", money(sell(23344.56), cents=False)],
        ["E3N-MS-100-S/M/L-E", "Meraki MS100 Small/Med/Large Essentials", "—", "0", "16", money(sell(1389.96), cents=False)],
        ["E3N-MS-300-M/L-A", "Meraki MS300 Medium/Large Advantage", "Adv", "4", "4", money(sell(3472.92), cents=False)],
        ["E3N-MV-E", "Meraki MV Large Essentials", "—", "0", "2", money(sell(414.72), cents=False)],
        ["E3N-CW-A / E3N-CW-E", "Cisco Networking Wireless", "Adv/Ess", "0", "2", money(sell(380.16), cents=False)],
        ["E3N-CS-AC1-L-E", "Switching Access T1 Large Essential", "Ess", "0", "49", money(sell(39125.52), cents=False)],
        ["E3N-CS-AC1-M-E", "Switching Access T1 Medium Essential", "Ess", "0", "42", money(sell(18204.48), cents=False)],
        ["E3N-N9300-XF2-A", "Nexus 9300 XF2 or higher (to 6.4T)", "Adv", "2", "2", money(sell(35439.84), cents=False)],
        ["E3N-N9300-XF-A", "Nexus 9300 XF 10G+ (to 3.6T)", "Adv", "22", "30", money(sell(243118.80), cents=False)],
        ["E3N-N9300-XF-E", "Nexus 9300 XF 10G+ (to 3.6T)", "Ess", "0", "4", money(sell(26111.52), cents=False)],
        ["E3N-SPACES-ACT", "Cisco Spaces ACT subscription", "Add-on", "0", "350", money(sell(64638.00), cents=False)],
    ]
    table = add_table(s, 13, 6, Inches(0.30), Inches(1.40), Inches(12.7), Inches(5.55))
    set_col_widths(table, [Inches(2.45), Inches(4.55), Inches(1.15), Inches(0.85), Inches(0.95), Inches(1.75)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)
    footer(s, 10, TOTAL_SLIDES)
    return s


def s10_networking_dna(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking software — detail 2 of 2", "Cisco DNA Switching and Wireless", "Hallmark TCV includes 22.5 pts GM  ·  C3850 48-port Advantage is the single largest software PID")

    headers = ["PID", "DESCRIPTION", "TIER", "IB", "QTY", "HALLMARK TCV"]
    data = [
        ["E3N-C38502-A", "C3850 48-Port DNA EA", "Adv", "0", "218", money(sell(306464.40), cents=False)],
        ["E3N-C95005-A", "C9500 48Y4C DNA EA", "Adv", "43", "16", money(sell(70791.12), cents=False)],
        ["E3N-C93002-A", "C9300/C9300X 48-Port DNA EA", "Adv", "43", "19", money(sell(26619.12), cents=False)],
        ["E3N-C95003-A", "C9500 32C DNA EA", "Adv", "4", "4", money(sell(18047.52), cents=False)],
        ["E3N-C9200CX2-A", "C9200CX 12-Port DNA EA", "Adv", "24", "18", money(sell(6690.24), cents=False)],
        ["E3N-C36502-A", "C3650 48-Port DNA EA", "Adv", "0", "4", money(sell(5623.20), cents=False)],
        ["E3N-C9300X2-A", "C9300X 24-Port DNA EA", "Adv", "6", "6", money(sell(4497.12), cents=False)],
        ["E3N-C95002-A", "C9500 Low Port (12Q/16X) DNA EA", "Adv", "0", "2", money(sell(5295.60), cents=False)],
        ["E3N-C38501-A", "C3850 24-Port DNA EA", "Adv", "0", "3", money(sell(2248.56), cents=False)],
        ["E3N-C95005-E", "C9500 48Y4C DNA EA", "Ess", "3", "3", money(sell(1863.00), cents=False)],
        ["E3N-C93001-A / others", "C9300 24 / C3560CX / C9200L 48", "Mix", "18", "7", money(sell(1406.88), cents=False)],
        ["E3N-AIRWLAN-A", "Cisco DNA Wireless", "Adv", "132", "525", money(sell(107361.36), cents=False)],
        ["TOTAL DNA + WIRELESS", "", "", "", "", money(sell(449546.76 + 107361.36), cents=False)],
    ]
    table = add_table(s, 14, 6, Inches(0.30), Inches(1.40), Inches(12.7), Inches(5.55))
    set_col_widths(table, [Inches(2.45), Inches(4.55), Inches(1.15), Inches(0.85), Inches(0.95), Inches(1.75)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)
    footer(s, 11, TOTAL_SLIDES)
    return s


def s11_networking_cx(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Networking CX", "Support & lifecycle services", f"Hallmark CX {money(sell(1127846.52), cents=False)}  ·  Cisco Support Standard unless noted Enhanced  ·  22.5 pts GM")

    headers = ["PID", "COVERAGE", "IB ASSETS", "DISC %", "HALLMARK TCV"]
    data = [
        ["E3-CX-ST-T1NBD", "Switching  8x5xNBD", "120", "35%", money(sell(385433.28), cents=False)],
        ["E3-CX-ST-T1NCD", "Switching  8x7xNCD", "1,053", "35%", money(sell(246936.24), cents=False)],
        ["E3-CX-ST-T1SWT", "Switching  software support", "—", "35%", money(sell(138047.76), cents=False)],
        ["E3-CX-CS-ENBD", "Cisco Switching  8x5xNBD", "91", "23%", money(sell(81637.56), cents=False)],
        ["E3-CX-DCN-T14HR", "Nexus  24x7x4", "10", "21%", money(sell(83260.08), cents=False)],
        ["E3-CX-ST-T14HR", "Switching  24x7x4", "7", "35%", money(sell(57020.40), cents=False)],
        ["E3-CXST-AIR-T1SWP", "Wireless  software support", "—", "35%", money(sell(33114.24), cents=False)],
        ["E3-CX-DCN-L1SWPERP", "Nexus  perpetual", "24", "22%", money(sell(27504.00), cents=False)],
        ["E3-CX-DCN-T1NBD", "Nexus  8x5xNBD", "26", "30%", money(sell(23171.40), cents=False)],
        ["E3-CX-DNS-T1SWC", "Spaces / DNAS cloud SW (Enhanced)", "—", "35%", money(sell(19705.32), cents=False)],
        ["E3-CX-DCN-T1NCD", "Nexus  8x7xNCD", "16", "23%", money(sell(15612.12), cents=False)],
        ["E3-CX-ST-T1NOS", "Switching  8x5xNBD OS", "7", "35%", money(sell(12789.00), cents=False)],
        ["E3-CX-CW-ENCD", "Cisco Wireless  8x7xNCD", "16", "31%", money(sell(3615.12), cents=False)],
        ["TOTAL NETWORKING CX", "", "", "", money(sell(1127846.52), cents=False)],
    ]
    table = add_table(s, 15, 5, Inches(0.35), Inches(1.40), Inches(12.6), Inches(5.55))
    set_col_widths(table, [Inches(2.70), Inches(4.40), Inches(1.55), Inches(1.30), Inches(2.65)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=11)
    footer(s, 12, TOTAL_SLIDES)
    return s


def s12_apps(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Applications Infrastructure", "Splunk Cloud and ThousandEyes", f"Hallmark TCV {money(sell(437058.00), cents=False)}  ·  No attached CX  ·  Full commitment  ·  22.5 pts GM")

    for i, (lab, val, cap, acc) in enumerate([
        ("PORTFOLIO TCV", money(sell(437058.00), cents=False), "100% software", GOLD),
        ("SPLUNK CLOUD", money(sell(90918.00), cents=False), "50 GB/day  ·  42% Cisco sub disc.", BLUE),
        ("THOUSANDEYES", money(sell(346140.00), cents=False), "Agents + EUM  ·  30% Cisco sub disc.", CYAN),
        ("SUCCESS PLAN", "Included", "Splunk Standard Success", GREEN),
    ]):
        kpi_card(s, Inches(0.40) + Inches(i * 3.20), Inches(1.42), Inches(3.05), Inches(1.22), lab, val, cap, acc)

    headers = ["PID", "DESCRIPTION", "QTY", "UNIT LIST / MO", "SUB DISC.", "HALLMARK TCV"]
    data = [
        ["E3A-SK-SE-CLD-S", "Splunk Cloud Subscription — Std success (GB/day)", "50", "$87.09", "42%", money(sell(90918.00), cents=False)],
        ["E3-CX-SK-SUP-ST", "Splunk Standard Success Plan", "1", "$0.00", "—", "$0"],
        ["E3A-TE-UNITS", "ThousandEyes Cloud & Enterprise Agents (per unit)", "1,500", "$0.84", "30%", money(sell(31860.00), cents=False)],
        ["E3A-TE-USERS", "ThousandEyes End User Monitoring Advantage", "1,500", "$8.31", "30%", money(sell(314280.00), cents=False)],
        ["E3A-TE-S", "Cisco Support Basic for EA ThousandEyes", "1", "$0.00", "—", "$0"],
        ["TOTAL", "", "", "", "", money(sell(437058.00), cents=False)],
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
    footer(s, 13, TOTAL_SLIDES)
    return s


def s13_security(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(
        s,
        "Security",
        "Identity Services Engine (Zero Trust)",
        f"SCU 1,000  ·  Software {money(sell(76726.68), cents=False)}  ·  CX {money(sell(115233.48), cents=False)}  ·  Combined {money(sell(191960.16), cents=False)}  ·  22.5 pts GM",
    )

    for i, item in enumerate([
        ("ISE ADVANTAGE", "3,450", "IB 400  →  desired 3,450", GOLD),
        ("ISE ESSENTIALS", "1,100", "IB 300  →  desired 1,100", BLUE),
        ("SOFTWARE TCV", money(sell(76726.68), cents=False), "42% Cisco subscription discount", CYAN),
        ("ISE CX TCV", money(sell(115233.48), cents=False), "Support Standard + credits", GREEN),
    ]):
        kpi_card(s, Inches(0.40) + Inches(i * 3.20), Inches(1.42), Inches(3.05), Inches(1.22), *item)

    headers = ["PID", "DESCRIPTION", "IB", "QTY", "DISC.", "HALLMARK TCV"]
    data = [
        ["E3S-ISE-ADV", "ISE Advantage", "400", "3,450", "42%", money(sell(70470.36), cents=False)],
        ["E3S-ISE-ESS", "ISE Essentials", "300", "1,100", "42%", money(sell(6256.32), cents=False)],
        ["SVS-E3S-ISE-B", "Cisco Support Basic for ISE", "0", "1", "23%", "$0"],
        ["E3-CX-ISE-ENBD", "Support Standard  8x5xNBD for ISE", "6", "1", "23%", money(sell(51485.76), cents=False)],
        ["E3-CX-ISE-EPER", "Cisco Support Standard Perpetual for ISE", "20", "1", "24%", money(sell(48011.76), cents=False)],
        ["E3-CX-ISE-ESWP", "Support Standard SW Support OP for ISE", "0", "1", "23%", money(sell(15735.96), cents=False)],
        ["E3-CX-ISE-ENCD", "Support Standard  8x7xNCD for ISE", "1", "1", "23%", "$0"],
        ["TOTAL SECURITY", "", "", "", "", money(sell(191960.16), cents=False)],
    ]
    table = add_table(s, 9, 6, Inches(0.35), Inches(2.88), Inches(12.6), Inches(3.95))
    set_col_widths(table, [Inches(2.30), Inches(4.70), Inches(1.05), Inches(1.20), Inches(1.15), Inches(2.20)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT]
    fill_table(table, headers, data, col_align=aligns, header_size=10, body_size=12)
    footer(s, 14, TOTAL_SLIDES)
    return s


def s14_collab(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Collaboration", "CX only — Employee Experience support", f"No new Collaboration software  ·  Hallmark CX {money(sell(359681.76), cents=False)}  ·  22.5 pts GM")

    add_round(s, Inches(0.40), Inches(1.45), Inches(12.5), Inches(1.35), WHITE)
    add_textbox(s, Inches(0.65), Inches(1.58), Inches(12.0), Inches(0.35), "What is in scope", 16, True, NAVY)
    add_textbox(
        s,
        Inches(0.65),
        Inches(1.95),
        Inches(12.0),
        Inches(0.65),
        "This EA continues Cisco Support Standard on the existing Collaboration installed base (Employee Experience). There is no Webex or Calling software subscription on Proposal 9325586. Uncovered-asset credits and a program-migration incentive reduce Cisco net CX versus list; 22.5 pts GM is then applied.",
        13,
        False,
        MUTED,
    )

    headers = ["PID", "DESCRIPTION", "IB", "QTY", "LIST", "CREDITS / OTD", "HALLMARK TCV"]
    data = [
        ["E3-CX-COL-ENBD", "Support Standard  8x5xNBD  COLLAB", "110", "1", "$414,060", "Migration $44,362", money(sell(273769.56), cents=False)],
        ["E3-CX-COL-ENCD", "Support Standard  8x7xNCD  COLLAB", "92", "1", "$145,359", "Uncovered $21,425", money(sell(85912.20), cents=False)],
        ["TOTAL", "", "202", "", "$559,419", "", money(sell(359681.76), cents=False)],
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
    footer(s, 15, TOTAL_SLIDES)
    return s


def s15_schedule(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Commercial schedule", "Annual billing over 36 months", "Even annual cash view of Hallmark TCV (Cisco net + 22.5 pts GM)  —  invoices follow Cisco EA billing + partner wrap")

    y_sw = sell(1543828.20) / 3
    y_svc = sell(1602761.76) / 3
    y_tot = sell(3146589.96) / 3
    headers = ["PERIOD", "SOFTWARE", "SERVICES", "ANNUAL", "CUMULATIVE"]
    data = [
        ["Year 1  (from 16 Oct 2026)", money(y_sw, cents=False), money(y_svc, cents=False), money(y_tot, cents=False), money(y_tot, cents=False)],
        ["Year 2", money(y_sw, cents=False), money(y_svc, cents=False), money(y_tot, cents=False), money(y_tot * 2, cents=False)],
        ["Year 3", money(y_sw, cents=False), money(y_svc, cents=False), money(y_tot, cents=False), money(sell(3146589.96), cents=False)],
        ["3-YEAR TCV", money(sell(1543828.20), cents=False), money(sell(1602761.76), cents=False), money(sell(3146589.96), cents=False), money(sell(3146589.96), cents=False)],
    ]
    table = add_table(s, 5, 5, Inches(0.45), Inches(1.45), Inches(12.4), Inches(2.55))
    set_col_widths(table, [Inches(3.30), Inches(2.25), Inches(2.25), Inches(2.30), Inches(2.30)])
    aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 4
    fill_table(table, headers, data, col_align=aligns, header_size=12, body_size=14)

    notes = [
        ("True-forward", "EA quantities are a floor. Additional consumption is licensed at EA rates, then wrapped with the same 22.5 pts GM — no surprise list-price true-ups."),
        ("IB credits", "Cisco one-time discounts ($475K) are already in Cisco net. The 22.5-pt wrap is applied after those credits, not on top of list."),
        ("Partner path", "WWT is reseller of record (services bill-to 1001363275). 22.5 pts GM is the partner commercial wrap on this EA."),
    ]
    for i, (t, b) in enumerate(notes):
        x = Inches(0.45) + Inches(i * 4.20)
        add_round(s, x, Inches(4.25), Inches(4.00), Inches(2.65), WHITE)
        add_rect(s, x, Inches(4.25), Inches(4.00), Inches(0.08), GOLD if i == 0 else BLUE)
        add_textbox(s, x + Inches(0.20), Inches(4.45), Inches(3.60), Inches(0.40), t, 15, True, NAVY)
        add_textbox(s, x + Inches(0.20), Inches(4.90), Inches(3.60), Inches(1.80), b, 12, False, MUTED)
    footer(s, 16, TOTAL_SLIDES)
    return s


def s16_value(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(
        s,
        "Value",
        "Hallmark price versus Cisco list",
        f"List $6,987,826  ·  Hallmark {money(sell(3146589.96), cents=False)}  ·  Savings {money(6987825.78 - sell(3146589.96), cents=False)}  (41.9%)",
    )

    def row(name, lst, net):
        h = sell(net)
        return [name, money(lst, cents=False), money(h, cents=False), money(lst - h, cents=False), f"{(lst - h) / lst * 100:.1f}%"]

    headers = ["PORTFOLIO", "LIST", "HALLMARK", "SAVINGS", "EFFECTIVE %"]
    data = [
        row("Networking software", 3252177.72, 1030043.52),
        row("Applications software", 650862.00, 437058.00),
        row("Security software", 136243.50, 76726.68),
        row("CX services (all)", 2948542.56, 1602761.76),
        row("TOTAL", 6987825.78, 3146589.96),
    ]
    table = add_table(s, 6, 5, Inches(0.45), Inches(1.42), Inches(12.4), Inches(2.85))
    set_col_widths(table, [Inches(3.40), Inches(2.25), Inches(2.25), Inches(2.25), Inches(2.25)])
    aligns = [PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 4
    fill_table(table, headers, data, col_align=aligns, header_size=12, body_size=13)

    points = [
        ("Cisco discounts still apply", "EA 3.0 suite discounts (up to 68% Networking, 30–42% Apps, 23–35% CX) sit underneath the partner wrap."),
        ("22.5 pts GM wrap", "Hallmark price = Cisco EA net ÷ 0.775. Partner margin is $913,526 over three years — not a reduction of Cisco list discount."),
        ("Still 41.9% vs list", "After margin, Hallmark is $2.93M under Cisco list $6.99M, with predictable annual true-forward at the same method."),
    ]
    for i, (t, b) in enumerate(points):
        x = Inches(0.45) + Inches(i * 4.20)
        add_round(s, x, Inches(4.50), Inches(4.00), Inches(2.40), WHITE)
        add_textbox(s, x + Inches(0.20), Inches(4.65), Inches(3.60), Inches(0.50), t, 14, True, NAVY)
        add_textbox(s, x + Inches(0.20), Inches(5.15), Inches(3.60), Inches(1.50), b, 12, False, MUTED)
    footer(s, 17, TOTAL_SLIDES)
    return s


def s17_next(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Path to book", "Recommended next steps", "Target book / RSD: 16 October 2026  ·  Hallmark TCV $4.06M with 22.5 pts GM")

    steps = [
        ("1", "Confirm quantities", "Validate Qty Desired vs. IB for Catalyst DNA (especially C3850 x218 and DNA Wireless x525), ISE 3,450/1,100, and ThousandEyes 1,500/1,500."),
        ("2", "Lock suite mix", "Keep Full commitment on core suites; confirm Partial is still correct for Spaces ACT and Meraki cameras."),
        ("3", "WWT quote path", "Convert this EAMP estimate to a WWT / Cisco approved quote at 22.5 pts GM. This file is indicative only."),
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
    footer(s, 18, TOTAL_SLIDES)
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
        "A 36-month Enterprise Agreement covering campus, data center, Meraki, Spaces,\nISE, Splunk, ThousandEyes, and Cisco CX — $4.06M Hallmark price versus $6.99M list.",
        16,
        False,
        CYAN,
    )

    contacts = [
        ("CUSTOMER", "Hallmark Cards, Incorporated\n2501 McGee Traffic Way\nKansas City, MO 64141"),
        ("PARTNER", "World Wide Technology, Inc\nReseller of record\n22.5 pts GM wrap"),
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
        "Disclaimer: Derived from Cisco EAMP Price Estimate ID 9325586 (08-Sep-2026), then wrapped with 22.5 points of partner gross margin.\n"
        "Indicative pricing only — NOT an approved Cisco quote. Pricing is subject to change.\n"
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
    s05_margin(prs)
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
