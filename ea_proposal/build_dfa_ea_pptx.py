#!/usr/bin/env python3
"""Build a Cisco Enterprise Agreement Deal Financial Analysis PPTX from an EAMP Excel export."""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "pptx"))

from eamp_parser import EAData, LineItem, parse_eamp_excel  # noqa: E402

NAVY = RGBColor(0x05, 0x2B, 0x4E)
NAVY_HEX = "052B4E"
BLUE = RGBColor(0x04, 0x9F, 0xD9)
CYAN = RGBColor(0x00, 0xBC, 0xEB)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
GREEN = RGBColor(0x2E, 0xA8, 0x4E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
OFF_WHITE = RGBColor(0xF4, 0xF7, 0xFB)
INK = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x5A, 0x64, 0x72)
ROW_ALT = "EEF3F8"
ROW_WHITE = "FFFFFF"
PALE = RGBColor(0xC5, 0xD4, 0xE3)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

MARGIN_PTS = 0.0


def sell(cisco_net: float) -> float:
    if MARGIN_PTS <= 0:
        return cisco_net
    return cisco_net / (1.0 - MARGIN_PTS / 100.0)


def gm_dollars(cisco_net: float) -> float:
    return sell(cisco_net) - cisco_net


def money(n: float, cents: bool = True) -> str:
    if abs(n) < 0.005:
        return "—"
    if cents:
        return f"${n:,.2f}"
    return f"${n:,.0f}"


def money_short(n: float) -> str:
    if abs(n) >= 1_000_000:
        return f"${n / 1_000_000:.2f}M"
    if abs(n) >= 1_000:
        return f"${n / 1_000:.0f}K"
    return money(n)


def pct(n: float) -> str:
    if n >= 1:
        n = n if n > 1.5 else n * 100
    return f"{n:.1f}%" if n else "—"


def display_name(data: EAData) -> str:
    raw = (data.customer_name or data.smart_account or "Enterprise Customer").strip()
    titled = raw.title().replace("'S", "'s")
    for token in ("Inc", "Llc", "Lp"):
        titled = titled.replace(token, token.upper())
    return titled


def years(data: EAData) -> float:
    return (data.term_months or 36) / 12.0


def term_label(data: EAData) -> str:
    months = data.term_months or 36
    if abs(months - round(months)) < 0.05:
        return f"{int(round(months))} months"
    return f"{months:.1f} months"


def set_run(run, text, size=14, bold=False, color=INK, font="Calibri"):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def _ensure_run(p):
    return p.runs[0] if p.runs else p.add_run()


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
    tcPr = cell._tc.get_or_add_tcPr()
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
    table = slide.shapes.add_table(rows, cols, l, t, w, h).table
    table.horz_banding = False
    table.vert_banding = False
    return table


def fill_table(table, headers, data, col_align=None, header_fill=NAVY_HEX, header_size=11, body_size=10):
    n_cols = len(headers)
    col_align = col_align or [PP_ALIGN.LEFT] * n_cols
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, size=header_size, bold=True, color=WHITE, fill=header_fill, align=col_align[j])
    for i, row in enumerate(data, 1):
        fill = ROW_ALT if i % 2 == 0 else ROW_WHITE
        label = str(row[0]).upper()
        is_total = label.startswith("TOTAL")
        for j, val in enumerate(row):
            style_cell(
                table.cell(i, j),
                str(val),
                size=body_size,
                bold=is_total,
                color=WHITE if is_total else INK,
                fill=NAVY_HEX if is_total else fill,
                align=col_align[j],
            )


def set_col_widths(table, widths):
    for i, w in enumerate(widths):
        table.columns[i].width = w


def conf_note() -> str:
    gm = f"  ·  {MARGIN_PTS:g} pts GM on Cisco EA net" if MARGIN_PTS else ""
    return f"CISCO CONFIDENTIAL  |  Deal Financial Analysis{gm}  |  Indicative — not an approved Cisco quote"


def footer(slide, page, total, data: EAData):
    add_rect(slide, Inches(0), Inches(7.28), SLIDE_W, Inches(0.22), NAVY)
    add_textbox(slide, Inches(0.35), Inches(7.28), Inches(9.2), Inches(0.22), conf_note(), size=9, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    add_textbox(
        slide,
        Inches(9.6),
        Inches(7.28),
        Inches(3.4),
        Inches(0.22),
        f"EA {data.proposal_id or ''}  ·  {page} / {total}",
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
    add_rect(slide, l, t, Inches(0.08), h, accent)
    add_textbox(slide, l + Inches(0.18), t + Inches(0.10), w - Inches(0.28), Inches(0.26), label.upper(), size=10, bold=True, color=MUTED)
    add_textbox(slide, l + Inches(0.18), t + Inches(0.34), w - Inches(0.28), Inches(0.42), value, size=20, bold=True, color=NAVY)
    if caption:
        add_textbox(slide, l + Inches(0.18), t + Inches(0.78), w - Inches(0.28), Inches(0.32), caption, size=10, color=MUTED)


def in_scope(data: EAData) -> list[dict]:
    return [p for p in data.portfolios if p.get("total", 0) > 0]


def suite_groups(items: list[LineItem]) -> list[tuple[str, list[LineItem], float]]:
    groups: dict[str, list[LineItem]] = defaultdict(list)
    for item in items:
        groups[item.suite or item.solution or "Other"].append(item)
    out = []
    for name, rows in groups.items():
        out.append((name, rows, sum(i.tcv for i in rows)))
    return sorted(out, key=lambda x: -x[2])


def paying(items: list[LineItem]) -> list[LineItem]:
    return [i for i in items if i.tcv > 0]


# ---------------------------------------------------------------------------
# Slides
# ---------------------------------------------------------------------------


def s_title(prs, data: EAData):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, GOLD)
    add_textbox(
        s, Inches(0.7), Inches(1.25), Inches(12), Inches(0.35),
        f"CISCO ENTERPRISE AGREEMENT  ·  DEAL FINANCIAL ANALYSIS  ·  PROPOSAL {data.proposal_id}",
        13, True, GOLD,
    )
    add_textbox(s, Inches(0.7), Inches(1.75), Inches(12), Inches(1.05), display_name(data), 34, True, WHITE)
    add_textbox(
        s, Inches(0.7), Inches(2.80), Inches(12), Inches(0.45),
        data.proposal_name or data.project_name or "Cisco Enterprise Agreement proposal",
        18, False, CYAN,
    )
    add_rect(s, Inches(0.7), Inches(3.40), Inches(2.2), Inches(0.06), GOLD)

    tcv = sell(data.total_tcv)
    facts = [
        ("TERM", term_label(data)),
        ("BILLING", data.billing or "Annual"),
        ("CISCO NET TCV" if MARGIN_PTS == 0 else "CUSTOMER TCV", money_short(tcv)),
        ("BOOK / RSD", data.rsd or "—"),
    ]
    for i, (k, v) in enumerate(facts):
        x = Inches(0.7) + Inches(i * 3.05)
        add_textbox(s, x, Inches(3.70), Inches(2.8), Inches(0.24), k, 11, True, GOLD)
        add_textbox(s, x, Inches(3.94), Inches(2.8), Inches(0.40), v, 20, True, WHITE)

    partner_line = f"Partner: {data.partner or '—'}    ·    Cisco AM: {data.account_manager or '—'}"
    ids = f"Buying Program {data.buying_program_id or '—'}    ·    Deal {data.deal_id or '—'}    ·    Quote {data.quote_id or '—'}"
    add_textbox(s, Inches(0.7), Inches(5.05), Inches(11.5), Inches(0.70), partner_line + "\n" + ids, 13, False, PALE)
    add_textbox(s, Inches(0.7), Inches(6.55), Inches(12), Inches(0.45), conf_note(), 11, False, GOLD)
    return s


def s_agenda(prs, data: EAData, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "DFA", "Agenda", display_name(data))
    items = [
        ("01", "Deal snapshot & commercial terms", "Customer, partner, buying program, duration, and book date"),
        ("02", "Deal financial analysis", "List → EA discount → one-time credit → Cisco net TCV"),
        ("03", "Portfolio mix & bill of materials", "In-scope suites, quantities, and Cisco net by PID"),
        ("04", "Annual operating model", "Even annual billing over the EA term"),
        ("05", "Path to book", "Quantity checks, partner quote, and Smart Account"),
    ]
    for i, (num, title, desc) in enumerate(items):
        y = Inches(1.50) + Inches(i * 1.05)
        add_round(s, Inches(0.45), y, Inches(12.4), Inches(0.92), WHITE)
        add_rect(s, Inches(0.45), y, Inches(0.12), Inches(0.92), GOLD if i % 2 == 0 else BLUE)
        add_textbox(s, Inches(0.80), y + Inches(0.16), Inches(1.0), Inches(0.60), num, 22, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(1.90), y + Inches(0.14), Inches(10.5), Inches(0.36), title, 18, True, NAVY)
        add_textbox(s, Inches(1.90), y + Inches(0.48), Inches(10.5), Inches(0.32), desc, 13, False, MUTED)
    footer(s, 2, total, data)
    return s


def s_snapshot(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Deal snapshot", f"{display_name(data)} at a glance", data.proposal_name or data.project_name)
    portfolios = ", ".join(p["name"] for p in in_scope(data)) or "—"
    rows = [
        ("End customer", data.customer_name or "—"),
        ("Sold-to address", data.address or "—"),
        ("Account", data.account or "—"),
        ("Smart Account", data.smart_account or "—"),
        ("Partner / reseller", data.partner or data.reseller or "—"),
        ("Cisco Account Manager", data.account_manager or "—"),
        ("Project / motion", f"{data.project_name or '—'}  ·  {data.ea_motion or 'EA'}"),
        ("Buying Program ID", data.buying_program_id or "—"),
        ("Deal / Quote / Proposal", f"{data.deal_id or '—'}  /  {data.quote_id or '—'}  /  {data.proposal_id or '—'}"),
        ("Price list", data.price_list or "—"),
        ("Requested ship / book", data.rsd or "—"),
        ("Duration / billing", f"{term_label(data)}  ·  {data.billing or 'Annual'}"),
        ("In-scope portfolios", portfolios),
        ("Report date", data.report_date or "—"),
    ]
    table = add_table(s, len(rows) + 1, 2, Inches(0.45), Inches(1.40), Inches(12.4), Inches(5.55))
    set_col_widths(table, [Inches(3.4), Inches(9.0)])
    fill_table(table, ["FIELD", "DETAIL"], rows, header_size=11, body_size=12)
    footer(s, page, total, data)
    return s


def s_why_ea(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Why this structure", "Cisco Enterprise Agreement", "One buying program, predictable spend, suite coverage")
    cards = [
        ("Predictable term spend", f"Fixed EA rates for {term_label(data)} with {data.billing.lower() if data.billing else 'annual'} true-forward. Growth is licensed at contracted EA pricing."),
        ("Suite coverage", "Buy the software suites needed now; add users, GB/day, or devices later at EA rates instead of one-off quotes."),
        ("Software + CX aligned", "Subscriptions and Cisco support ride together so coverage, SLAs, and entitlement stay in lockstep."),
        ("Installed-base credit", "In-use assets (IB) generate one-time discounts that reduce net TCV versus buying the same footprint a la carte."),
        ("Partner-delivered", f"{data.partner or 'The reseller of record'} transacts the EA against the customer Smart Account."),
        ("Operational simplicity", f"A single buying program ({data.buying_program_id or 'EA'}) and one expected book date ({data.rsd or 'TBD'}) in place of staggered renewals."),
    ]
    for i, (title, body) in enumerate(cards):
        col, row = i % 3, i // 3
        x = Inches(0.40) + Inches(col * 4.25)
        y = Inches(1.45) + Inches(row * 2.70)
        add_round(s, x, y, Inches(4.05), Inches(2.50), WHITE)
        add_rect(s, x, y, Inches(4.05), Inches(0.08), GOLD if row == 0 else BLUE)
        add_textbox(s, x + Inches(0.20), y + Inches(0.25), Inches(3.65), Inches(0.55), title, 16, True, NAVY)
        add_textbox(s, x + Inches(0.20), y + Inches(0.85), Inches(3.65), Inches(1.45), body, 13, False, MUTED)
    footer(s, page, total, data)
    return s


def s_dfa(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    savings = (data.list_total or 0) - (data.total_tcv or 0)
    save_pct = (100 * savings / data.list_total) if data.list_total else 0
    header_bar(
        s, "Deal financial analysis",
        "List → discount → Cisco net TCV",
        f"Indicative as of {data.report_date or 'export'}  ·  Proposal {data.proposal_id}",
    )
    kpi_card(s, Inches(0.40), Inches(1.42), Inches(3.00), Inches(1.25), "Cisco list", money_short(data.list_total), "Before EA discounts", BLUE)
    kpi_card(s, Inches(3.55), Inches(1.42), Inches(3.00), Inches(1.25), "EA discount", money_short(data.discount_total), f"{save_pct:.1f}% vs list" if data.list_total else "", GOLD)
    kpi_card(s, Inches(6.70), Inches(1.42), Inches(3.00), Inches(1.25), "One-time credit", money_short(data.otd_total), "IB / OTD ramp", BLUE)
    kpi_card(s, Inches(9.85), Inches(1.42), Inches(3.05), Inches(1.25), "Cisco net TCV", money_short(data.total_tcv), conf_caption(), GOLD)

    rows = [["PORTFOLIO", "SOFTWARE", "SERVICES", "NET TCV", "LIST", "DISCOUNT"]]
    table_data = []
    for p in in_scope(data):
        table_data.append([
            p["name"],
            money(sell(p.get("software", 0))),
            money(sell(p.get("services", 0))),
            money(sell(p.get("total", 0))),
            money(p.get("list", 0)),
            money(p.get("discount", 0)),
        ])
    table_data.append([
        "TOTAL",
        money(sell(data.total_software)),
        money(sell(data.total_services)),
        money(sell(data.total_tcv)),
        money(data.list_total),
        money(data.discount_total),
    ])
    table = add_table(s, len(table_data) + 1, 6, Inches(0.40), Inches(2.90), Inches(12.5), Inches(0.38 * (len(table_data) + 1) + 0.2))
    set_col_widths(table, [Inches(3.1)] + [Inches(1.88)] * 5)
    fill_table(table, rows[0], table_data, col_align=[PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 5, body_size=12)

    notes = data.highlights[:3] or ["Cisco net is the EAMP Total Contract Value after suite and subscription discounts."]
    y = Inches(5.55)
    for i, note in enumerate(notes):
        add_textbox(s, Inches(0.50), y + Inches(i * 0.38), Inches(12.3), Inches(0.36), "•  " + note, 13, False, MUTED)
    footer(s, page, total, data)
    return s


def conf_caption() -> str:
    return f"{MARGIN_PTS:g} pts GM wrap" if MARGIN_PTS else "EAMP export net"


def s_investment(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    yr = years(data)
    annual = sell(data.total_tcv) / yr if yr else 0
    header_bar(s, "Investment", "Contract value at a glance", f"{term_label(data)}  ·  {data.billing} billing")
    kpi_card(s, Inches(0.40), Inches(1.45), Inches(4.05), Inches(1.35), "Net TCV", money(sell(data.total_tcv)), f"{term_label(data)} Cisco EA", GOLD)
    kpi_card(s, Inches(4.65), Inches(1.45), Inches(4.05), Inches(1.35), "Software", money(sell(data.total_software)), "Subscriptions", BLUE)
    kpi_card(s, Inches(8.90), Inches(1.45), Inches(4.00), Inches(1.35), "Services / CX", money(sell(data.total_services)), "Support & lifecycle", BLUE)
    kpi_card(s, Inches(0.40), Inches(3.05), Inches(4.05), Inches(1.35), "Est. annual", money(annual), f"TCV ÷ {yr:.2f} years", GOLD)
    kpi_card(s, Inches(4.65), Inches(3.05), Inches(4.05), Inches(1.35), "Vs. list", money((data.list_total or 0) - sell(data.total_tcv)), f"{(100 * ((data.list_total or 0) - data.total_tcv) / data.list_total):.1f}% off list" if data.list_total else "", GREEN)
    kpi_card(s, Inches(8.90), Inches(3.05), Inches(4.00), Inches(1.35), "OTD / IB credit", money(data.otd_total), "Consumed at activation", BLUE)

    if MARGIN_PTS:
        add_round(s, Inches(0.40), Inches(4.65), Inches(12.5), Inches(2.20), WHITE)
        add_textbox(s, Inches(0.65), Inches(4.80), Inches(12.0), Inches(0.35), f"Partner wrap — {MARGIN_PTS:g} points of gross margin", 16, True, NAVY)
        add_textbox(
            s, Inches(0.65), Inches(5.25), Inches(12.0), Inches(1.35),
            f"Customer / partner price = Cisco net ÷ {1 - MARGIN_PTS/100:.3f}.  "
            f"Gross profit on this EA is {money(gm_dollars(data.total_tcv))} "
            f"({MARGIN_PTS:g}% of sell price).  Cisco list discounts are unchanged.",
            14, False, MUTED,
        )
    else:
        add_round(s, Inches(0.40), Inches(4.65), Inches(12.5), Inches(2.20), WHITE)
        add_textbox(s, Inches(0.65), Inches(4.80), Inches(12.0), Inches(0.35), "Pricing basis", 16, True, NAVY)
        add_textbox(
            s, Inches(0.65), Inches(5.25), Inches(12.0), Inches(1.35),
            "Figures are Cisco EA net from the EAMP Price Estimate (Total Contract Value). "
            "This is indicative only and is not an approved Cisco quote. Partner margin has not been applied.",
            14, False, MUTED,
        )
    footer(s, page, total, data)
    return s


def s_mix(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Mix", "Where the investment sits", "Cisco net TCV by in-scope portfolio")
    scoped = in_scope(data)
    max_total = max((p["total"] for p in scoped), default=1) or 1
    for i, p in enumerate(scoped[:6]):
        y = Inches(1.50) + Inches(i * 0.85)
        add_textbox(s, Inches(0.50), y, Inches(4.0), Inches(0.35), p["name"], 14, True, NAVY)
        add_textbox(s, Inches(0.50), y + Inches(0.32), Inches(4.0), Inches(0.28), f"SW {money(p['software'])}   CX {money(p['services'])}", 11, False, MUTED)
        bar_w = Inches(6.8) * (p["total"] / max_total)
        add_round(s, Inches(4.70), y + Inches(0.12), Inches(6.8), Inches(0.42), RGBColor(0xE4, 0xEC, 0xF4))
        add_round(s, Inches(4.70), y + Inches(0.12), bar_w, Inches(0.42), GOLD if i == 0 else BLUE)
        share = 100 * p["total"] / data.total_tcv if data.total_tcv else 0
        add_textbox(s, Inches(11.60), y + Inches(0.12), Inches(1.3), Inches(0.42), f"{share:.0f}%", 14, True, NAVY, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
    footer(s, page, total, data)
    return s


def s_bom(prs, data: EAData, title: str, items: list[LineItem], tcv: float, page: int, total: int, subtitle: str = ""):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Bill of materials", title, subtitle or f"Cisco net TCV {money(sell(tcv))}")
    rows = paying(items) or items
    rows = sorted(rows, key=lambda x: -x.tcv)[:12]
    data_rows = []
    for item in rows:
        disc = item.sub_disc
        disc_s = f"{disc:.2f}%" if disc and disc <= 100 else (f"{disc*100:.1f}%" if disc else "—")
        qty = int(item.desired) if item.desired else int(item.ib)
        data_rows.append([
            (item.suite or item.solution)[:28],
            item.part_number,
            (item.description or "")[:42],
            f"{qty:,}",
            disc_s,
            money(sell(item.tcv)),
        ])
    if not data_rows:
        add_textbox(s, Inches(0.5), Inches(2.0), Inches(12), Inches(1), "No line items with contract value in this portfolio.", 16, False, MUTED)
        footer(s, page, total, data)
        return s
    table = add_table(s, len(data_rows) + 1, 6, Inches(0.35), Inches(1.42), Inches(12.6), Inches(min(5.4, 0.38 * (len(data_rows) + 1) + 0.25)))
    set_col_widths(table, [Inches(2.3), Inches(2.15), Inches(4.15), Inches(1.1), Inches(1.2), Inches(1.7)])
    fill_table(
        table,
        ["SUITE", "PART NUMBER", "DESCRIPTION", "QTY", "DISC %", "NET TCV"],
        data_rows,
        col_align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT],
        body_size=10,
    )
    footer(s, page, total, data)
    return s


def s_services(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Support & lifecycle", "Cisco CX on this EA", f"Services TCV {money(sell(data.total_services))}")
    if data.support_breakdown:
        rows = [[r["category"][:70], money(sell(r["tcv"]))] for r in data.support_breakdown]
        rows.append(["TOTAL", money(sell(data.total_services))])
        table = add_table(s, len(rows) + 1, 2, Inches(0.45), Inches(1.45), Inches(12.4), Inches(0.42 * (len(rows) + 1) + 0.2))
        set_col_widths(table, [Inches(9.4), Inches(3.0)])
        fill_table(table, ["CATEGORY", "NET TCV"], rows, col_align=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT], body_size=13)
    else:
        add_textbox(s, Inches(0.5), Inches(2.0), Inches(12), Inches(1), "No paid Cisco CX / support line items on this estimate.", 16, False, MUTED)
    footer(s, page, total, data)
    return s


def s_schedule(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    yr = years(data)
    annual_sw = sell(data.total_software) / yr if yr else 0
    annual_svc = sell(data.total_services) / yr if yr else 0
    annual = sell(data.total_tcv) / yr if yr else 0
    header_bar(s, "Operating model", f"{data.billing or 'Annual'} cash flow", f"{term_label(data)} term  ·  even {data.billing.lower() if data.billing else 'annual'} invoices")
    n_years = max(1, int(round(yr))) if abs(yr - round(yr)) < 0.05 else yr
    year_phrase = f"{int(n_years)} invoice cycles" if isinstance(n_years, int) else f"{yr:.2f} years"
    headers = ["CATEGORY", f"{term_label(data).upper()} TCV", "EST. ANNUAL"]
    rows = [
        ["Software subscriptions", money(sell(data.total_software)), money(annual_sw)],
        ["Support & lifecycle", money(sell(data.total_services)), money(annual_svc)],
        ["TOTAL", money(sell(data.total_tcv)), money(annual)],
    ]
    table = add_table(s, 4, 3, Inches(0.45), Inches(1.50), Inches(12.4), Inches(1.85))
    set_col_widths(table, [Inches(5.4), Inches(3.5), Inches(3.5)])
    fill_table(table, headers, rows, col_align=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT], body_size=14, header_size=13)

    notes = [
        f"Approximate annual spend: {money(annual)} per year across {year_phrase}.",
        f"Billing model: {data.billing or 'Annual'} over a {term_label(data)} EA term.",
        "Programmatic and suite discounts are already in the Cisco net TCV.",
        "Pre-EA install-base assets (and any OTD) are consumed when the EA activates.",
    ]
    for i, note in enumerate(notes):
        add_textbox(s, Inches(0.55), Inches(3.70) + Inches(i * 0.55), Inches(12.2), Inches(0.50), "•  " + note, 14, False, MUTED)
    footer(s, page, total, data)
    return s


def s_oos(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Scope", "Portfolios not in this estimate", "Can be added later by EA amendment")
    items = data.out_of_scope or ["None — all standard EA portfolios appear in this estimate."]
    for i, item in enumerate(items[:8]):
        y = Inches(1.50) + Inches(i * 0.65)
        add_round(s, Inches(0.45), y, Inches(12.4), Inches(0.55), WHITE)
        add_rect(s, Inches(0.45), y, Inches(0.10), Inches(0.55), GOLD)
        add_textbox(s, Inches(0.80), y, Inches(11.8), Inches(0.55), item, 16, False, NAVY, anchor=MSO_ANCHOR.MIDDLE)
    footer(s, page, total, data)
    return s


def s_next(prs, data: EAData, page: int, total: int):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, OFF_WHITE)
    header_bar(s, "Path to book", "Recommended next steps", f"Target book / RSD: {data.rsd or 'TBD'}  ·  Net TCV {money(sell(data.total_tcv))}")
    qty_hint = "Validate Quantity Desired versus install base on every paying PID."
    top = paying(data.applications + data.security_sw + data.networking_sw)
    if top:
        sample = ", ".join(f"{i.part_number} × {int(i.desired or i.ib)}" for i in top[:3])
        qty_hint = f"Validate quantities against IB, starting with {sample}."
    steps = [
        ("1", "Confirm quantities", qty_hint),
        ("2", "Lock suite mix", "Keep Full commitment on core suites; confirm Partial suites are still correct."),
        ("3", "Partner quote path", f"Convert this EAMP estimate to a {(data.partner or 'partner')} / Cisco approved quote. This file is indicative only."),
        ("4", "Smart Account", f"Confirm {data.smart_account or 'the customer Smart Account'} is the entitlement destination before book."),
        ("5", "Legal / EA paper", "Standard Cisco EA terms, true-forward, and CX service descriptions via the partner."),
        ("6", f"Book {data.rsd or 'on RSD'}", f"Align PO, {data.billing.lower() if data.billing else 'annual'} billing, and RSD so software and CX start together."),
    ]
    for i, (n, title, body) in enumerate(steps):
        col, row = i % 3, i // 3
        x = Inches(0.40) + Inches(col * 4.25)
        y = Inches(1.45) + Inches(row * 2.70)
        add_round(s, x, y, Inches(4.05), Inches(2.50), WHITE)
        add_round(s, x + Inches(0.20), y + Inches(0.22), Inches(0.48), Inches(0.48), NAVY)
        add_textbox(s, x + Inches(0.20), y + Inches(0.22), Inches(0.48), Inches(0.48), n, 16, True, WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, x + Inches(0.80), y + Inches(0.28), Inches(3.00), Inches(0.40), title, 16, True, NAVY)
        add_textbox(s, x + Inches(0.20), y + Inches(0.85), Inches(3.65), Inches(1.45), body, 13, False, MUTED)
    footer(s, page, total, data)
    return s


def s_close(prs, data: EAData):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_rect(s, 0, 0, SLIDE_W, SLIDE_H, NAVY)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, GOLD)
    add_textbox(s, Inches(0.7), Inches(0.85), Inches(12), Inches(0.30), f"DEAL FINANCIAL ANALYSIS  ·  PROPOSAL {data.proposal_id}", 13, True, GOLD)
    add_textbox(s, Inches(0.7), Inches(1.25), Inches(12), Inches(0.70), display_name(data), 32, True, WHITE)
    scoped = ", ".join(p["name"].replace(" Infrastructure", "") for p in in_scope(data)) or "Cisco EA portfolios"
    add_textbox(
        s, Inches(0.7), Inches(2.10), Inches(12), Inches(0.70),
        f"{term_label(data)} Enterprise Agreement covering {scoped} — "
        f"{money_short(sell(data.total_tcv))} Cisco net versus {money_short(data.list_total)} list.",
        16, False, CYAN,
    )
    contacts = [
        ("CUSTOMER", f"{data.customer_name or '—'}\n{data.address or ''}\n{data.account or ''}"),
        ("PARTNER", f"{data.partner or 'Reseller of record'}\n{data.reseller or ''}\nBuying Program {data.buying_program_id or '—'}"),
        ("CISCO", f"{data.account_manager or 'Cisco Account Manager'}\nDeal {data.deal_id or '—'}\nQuote {data.quote_id or '—'}"),
    ]
    for i, (h, body) in enumerate(contacts):
        x = Inches(0.7) + Inches(i * 4.05)
        add_rect(s, x, Inches(3.10), Inches(3.70), Inches(0.06), GOLD)
        add_textbox(s, x, Inches(3.30), Inches(3.70), Inches(0.30), h, 12, True, GOLD)
        add_textbox(s, x, Inches(3.62), Inches(3.70), Inches(1.50), body.strip(), 14, False, WHITE)
    add_textbox(
        s, Inches(0.7), Inches(5.40), Inches(12), Inches(1.50),
        f"Disclaimer: Derived from Cisco EAMP Price Estimate ID {data.proposal_id} ({data.report_date or 'export date'}).\n"
        "Indicative pricing only — NOT an approved Cisco quote. Pricing is subject to change.\n"
        "CISCO CONFIDENTIAL — do not distribute outside the customer, partner, and Cisco deal teams.",
        12, False, PALE,
    )
    return s


def build_presentation(data: EAData) -> Presentation:
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    def _suite_sub(items: list[LineItem], fallback: str) -> str:
        names = []
        for name, _rows, tcv in suite_groups(items):
            if tcv > 0 or name not in names:
                names.append(name)
        return ", ".join(names[:5]) or fallback

    boms = []
    if data.applications or data.applications_sw_tcv:
        boms.append(("Applications Infrastructure", data.applications, data.applications_sw_tcv, _suite_sub(data.applications, "Application suites")))
    if data.security_sw or data.security_sw_tcv:
        boms.append(("Security software", data.security_sw, data.security_sw_tcv, _suite_sub(data.security_sw, "Security suites")))
    if data.networking_sw or data.networking_sw_tcv:
        boms.append(("Networking Infrastructure", data.networking_sw, data.networking_sw_tcv, _suite_sub(data.networking_sw, "Networking suites")))
    if data.collaboration_sw or data.collaboration_sw_tcv:
        boms.append(("Collaboration software", data.collaboration_sw, data.collaboration_sw_tcv, _suite_sub(data.collaboration_sw, "Collaboration suites")))

    has_svc = bool(data.total_services or data.support_breakdown)
    has_oos = bool(data.out_of_scope)
    # title, agenda, snapshot, why, dfa, investment, mix, N boms, services?, schedule, oos?, next, close
    total = 7 + len(boms) + (1 if has_svc else 0) + 1 + (1 if has_oos else 0) + 2

    s_title(prs, data)
    s_agenda(prs, data, total)
    page = 3
    s_snapshot(prs, data, page, total); page += 1
    s_why_ea(prs, data, page, total); page += 1
    s_dfa(prs, data, page, total); page += 1
    s_investment(prs, data, page, total); page += 1
    s_mix(prs, data, page, total); page += 1
    for title, items, tcv, sub in boms:
        s_bom(prs, data, title, items, tcv, page, total, sub); page += 1
    if has_svc:
        s_services(prs, data, page, total); page += 1
    s_schedule(prs, data, page, total); page += 1
    if has_oos:
        s_oos(prs, data, page, total); page += 1
    s_next(prs, data, page, total); page += 1
    s_close(prs, data)

    if len(prs.slides) != total:
        # Footers can tolerate off-by-one if a section is empty; don't fail the build.
        pass
    return prs


def find_excel(explicit: Path | None) -> Path | None:
    if explicit and explicit.exists():
        return explicit
    needles = [
        "*9322988*.xlsx",
        "*EAMP Price Estimate*.xlsx",
        "*HALIFAX*EAMP*.xlsx",
        "*DFA*EAMP*.xlsx",
    ]
    roots = [
        Path.cwd(),
        ROOT,
        Path("/home/ubuntu/.cursor/projects/workspace/uploads"),
        Path("/home/ubuntu/.cursor/projects/agent-repos-cxvx/uploads"),
        Path("/tmp"),
        Path("/opt/cursor/artifacts"),
        Path("/cursor/stores/self"),
    ]
    for root in roots:
        if not root.exists():
            continue
        for needle in needles:
            hits = sorted(root.rglob(needle))
            if hits:
                return hits[0]
    return None


def inspect_workbook(path: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    print(f"Workbook: {path.name}")
    for name in wb.sheetnames:
        ws = wb[name]
        print(f"\n  Sheet: {name!r} ({ws.max_row} rows x {ws.max_column} cols)")
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=min(18, ws.max_row or 18), values_only=True), 1):
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cells:
                print(f"    {i:3d} | " + " | ".join(cells[:8]))


def main(argv: list[str] | None = None) -> int:
    global MARGIN_PTS
    parser = argparse.ArgumentParser(description="Generate a DFA Enterprise Agreement PPTX from an EAMP Excel file")
    parser.add_argument("excel", nargs="?", type=Path, help="EAMP Price Estimate .xlsx")
    parser.add_argument("-o", "--output", type=Path, default=None)
    parser.add_argument("--margin-pts", type=float, default=0.0, help="Partner gross-margin points on Cisco net (sell = net / (1 - pts/100))")
    parser.add_argument("--inspect", action="store_true")
    args = parser.parse_args(argv)

    MARGIN_PTS = args.margin_pts
    excel = find_excel(args.excel)
    if not excel:
        print("Error: EAMP Excel not found. Pass the .xlsx path or place *9322988*.xlsx in the workspace.", file=sys.stderr)
        return 1

    if args.inspect:
        inspect_workbook(excel)
        return 0

    data = parse_eamp_excel(excel)
    slug = (data.customer_name or "Customer").replace(" ", "_").replace("/", "_")
    out = args.output or (ROOT / "output" / f"{slug}_DFA_EA_Proposal_{data.proposal_id or 'estimate'}.pptx")
    out.parent.mkdir(parents=True, exist_ok=True)
    prs = build_presentation(data)
    prs.save(str(out))
    print(f"Generated: {out} ({len(prs.slides)} slides)")
    print(f"  Customer : {data.customer_name}")
    print(f"  Proposal : {data.proposal_id}  |  Cisco net TCV: {money(data.total_tcv)}")
    print(f"  List     : {money(data.list_total)}  |  Discount: {money(data.discount_total)}  |  OTD: {money(data.otd_total)}")
    print(f"  Term     : {term_label(data)}  |  Billing: {data.billing}  |  Partner: {data.partner or '—'}")
    if MARGIN_PTS:
        print(f"  Sell TCV : {money(sell(data.total_tcv))}  ({MARGIN_PTS:g} pts GM, GP {money(gm_dollars(data.total_tcv))})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
