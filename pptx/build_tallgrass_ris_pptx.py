#!/usr/bin/env python3
"""Build the Tallgrass IQ Onboarded RIS Assessment opportunity PPTX.

Source: Tallgrass Region IQ Onboarded Accounts & RIS Propensity Guide
(Cisco Services & Software Buying Programs — internal sales enablement).
"""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# --- Dark executive palette ---
BG = RGBColor(0x0E, 0x1A, 0x2B)
BG_HEX = "0E1A2B"
HEADER = RGBColor(0x0B, 0x16, 0x24)
CARD = RGBColor(0x17, 0x28, 0x3C)
CARD_HEX = "17283C"
CARD_ALT_HEX = "1C3048"
NAVY_HEX = "052B4E"
BLUE = RGBColor(0x04, 0x9F, 0xD9)
BLUE_HEX = "049FD9"
CYAN = RGBColor(0x00, 0xBC, 0xEB)
GOLD = RGBColor(0xC4, 0xA3, 0x5A)
GOLD_HEX = "C4A35A"
TEAL = RGBColor(0x2B, 0xBB, 0xAD)
TEAL_HEX = "2BBBAD"
GREEN = RGBColor(0x3D, 0xDC, 0x97)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
WHITE_HEX = "FFFFFF"
BODY = RGBColor(0xD5, 0xE0, 0xEA)
MUTED = RGBColor(0x9B, 0xB0, 0xC3)
FOOTER_BG = RGBColor(0x08, 0x12, 0x1E)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

TOTAL_SLIDES = 14
OUT_PATH = Path(__file__).resolve().parents[1] / "output" / (
    "Tallgrass_IQ_Onboarded_RIS_Assessment_Opportunity_Guide.pptx"
)


def set_run(run, text, size=14, bold=False, color=BODY, font="Calibri"):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def _ensure_run(p):
    if p.runs:
        return p.runs[0]
    return p.add_run()


def add_textbox(
    slide,
    l,
    t,
    w,
    h,
    text,
    size=14,
    bold=False,
    color=BODY,
    align=PP_ALIGN.LEFT,
    font="Calibri",
    anchor=MSO_ANCHOR.TOP,
):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    try:
        tf._txBody.bodyPr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor],
        )
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
    shp.shadow.inherit = False
    try:
        shp.adjustments[0] = 0.08
    except Exception:
        pass
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


def set_cell_margins(cell, l=70000, t=50000, r=70000, b=50000):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.set("marL", str(l))
    tcPr.set("marR", str(r))
    tcPr.set("marT", str(t))
    tcPr.set("marB", str(b))


def style_cell(
    cell,
    text,
    size=11,
    bold=False,
    color=BODY,
    fill=None,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.MIDDLE,
):
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


def set_col_widths(table, widths):
    for i, w in enumerate(widths):
        table.columns[i].width = w


def paint_bg(slide):
    add_rect(slide, 0, 0, SLIDE_W, SLIDE_H, BG)


def footer(slide, page):
    add_rect(slide, Inches(0), Inches(7.26), SLIDE_W, Inches(0.24), FOOTER_BG)
    add_rect(slide, Inches(0), Inches(7.26), SLIDE_W, Inches(0.03), GOLD)
    add_textbox(
        slide,
        Inches(0.35),
        Inches(7.27),
        Inches(9.4),
        Inches(0.22),
        "CISCO CONFIDENTIAL  ·  Internal Sales Enablement  ·  Tallgrass Region  ·  Services & Software Buying Programs",
        size=9,
        color=MUTED,
        anchor=MSO_ANCHOR.MIDDLE,
    )
    add_textbox(
        slide,
        Inches(10.4),
        Inches(7.27),
        Inches(2.6),
        Inches(0.22),
        f"RIS Pipeline  ·  {page} / {TOTAL_SLIDES}",
        size=9,
        color=GOLD,
        align=PP_ALIGN.RIGHT,
        anchor=MSO_ANCHOR.MIDDLE,
    )


def header_bar(slide, eyebrow, title, subtitle=None):
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(1.16), HEADER)
    add_rect(slide, Inches(0), Inches(0), Inches(0.12), Inches(1.16), GOLD)
    add_textbox(
        slide,
        Inches(0.40),
        Inches(0.10),
        Inches(12.4),
        Inches(0.26),
        eyebrow.upper(),
        size=11,
        bold=True,
        color=GOLD,
    )
    add_textbox(
        slide,
        Inches(0.40),
        Inches(0.34),
        Inches(12.4),
        Inches(0.42),
        title,
        size=24,
        bold=True,
        color=WHITE,
    )
    if subtitle:
        add_textbox(
            slide,
            Inches(0.40),
            Inches(0.78),
            Inches(12.4),
            Inches(0.28),
            subtitle,
            size=12,
            color=CYAN,
        )


def kpi_card(slide, l, t, w, h, label, value, caption=None, accent=GOLD):
    add_round(slide, l, t, w, h, CARD)
    add_rect(slide, l, t, Inches(0.08), h, accent)
    add_textbox(
        slide,
        l + Inches(0.20),
        t + Inches(0.10),
        w - Inches(0.30),
        Inches(0.24),
        label.upper(),
        size=10,
        bold=True,
        color=MUTED,
    )
    add_textbox(
        slide,
        l + Inches(0.20),
        t + Inches(0.34),
        w - Inches(0.30),
        Inches(0.42),
        value,
        size=22,
        bold=True,
        color=WHITE,
    )
    if caption:
        add_textbox(
            slide,
            l + Inches(0.20),
            t + Inches(0.78),
            w - Inches(0.30),
            Inches(0.34),
            caption,
            size=11,
            color=CYAN,
        )


def bullet_block(slide, l, t, w, h, items, size=13, color=BODY):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.level = 0
        p.space_after = Pt(8)
        set_run(p.add_run(), "•  " + item, size, False, color)
    return box


def blank_slide(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    paint_bg(s)
    return s


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------


def s01_title(prs):
    s = blank_slide(prs)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, GOLD)
    add_textbox(
        s,
        Inches(0.70),
        Inches(1.15),
        Inches(12.0),
        Inches(0.32),
        "CISCO SERVICES & SOFTWARE BUYING PROGRAMS  ·  PIPELINE GENERATION PLAN",
        13,
        True,
        GOLD,
    )
    add_textbox(s, Inches(0.70), Inches(1.55), Inches(12.0), Inches(0.70), "Tallgrass Region", 40, True, WHITE)
    add_textbox(
        s,
        Inches(0.70),
        Inches(2.25),
        Inches(12.0),
        Inches(0.50),
        "IQ Onboarded Accounts & RIS Assessment Propensity Guide",
        22,
        False,
        CYAN,
    )
    add_rect(s, Inches(0.70), Inches(2.90), Inches(2.40), Inches(0.06), GOLD)

    add_textbox(
        s,
        Inches(0.70),
        Inches(3.15),
        Inches(11.8),
        Inches(0.70),
        "Convert live Cisco IQ telemetry into Routing & Infrastructure Services (RIS) Professional Assessments — then into multi-million-dollar routing, switching, and cloud security modernization programs.",
        16,
        False,
        BODY,
    )

    facts = [
        ("ADOPTING ACCOUNTS", "233"),
        ("CORE TARGETS", "10"),
        ("INSTALL FOOTPRINT", "$319M+"),
        ("RIS ASSESSMENT TAM", "$1.2–1.8M"),
        ("PULL-THROUGH", "$25M+"),
    ]
    for i, (k, v) in enumerate(facts):
        x = Inches(0.70) + Inches(i * 2.45)
        add_round(s, x, Inches(4.15), Inches(2.30), Inches(1.35), CARD)
        add_textbox(s, x + Inches(0.12), Inches(4.28), Inches(2.06), Inches(0.28), k, 10, True, MUTED)
        add_textbox(s, x + Inches(0.12), Inches(4.58), Inches(2.06), Inches(0.70), v, 22, True, WHITE)

    add_textbox(
        s,
        Inches(0.70),
        Inches(5.80),
        Inches(12.0),
        Inches(0.50),
        "Account alignment: Robin Randolph  ·  Gerry Garwood  ·  Jeff Pavalone  ·  Paige Ball",
        14,
        False,
        GOLD,
    )
    add_textbox(
        s,
        Inches(0.70),
        Inches(6.55),
        Inches(12.0),
        Inches(0.35),
        "CISCO CONFIDENTIAL  ·  Internal Sales Enablement  ·  Not for customer distribution",
        12,
        False,
        MUTED,
    )
    return s


def s02_agenda(prs):
    s = blank_slide(prs)
    header_bar(s, "How to use this deck", "Agenda", "Five moves from IQ telemetry to staged RIS SOWs")
    items = [
        ("01", "Why this motion, why now", "IQ-onboarded accounts are the highest-propensity RIS Assessment buyers in Tallgrass."),
        ("02", "Propensity evaluation framework", "Four pillars: live telemetry, scale, LDOS exposure, and transformation in flight."),
        ("03", "Tier 1 and Tier 2 target accounts", "Ten core names, recommended plays, and architectural triggers by SAVM."),
        ("04", "Sales plays and discovery scripts", "Telemetry Unlock vs. Pre-Transformation Derisking — with executive language."),
        ("05", "AE action plan and SOW staging", "Six-week path to $75k–$150k RIS assessments attached to FY27 buying programs."),
    ]
    for i, (num, title, desc) in enumerate(items):
        y = Inches(1.40) + Inches(i * 1.10)
        add_round(s, Inches(0.40), y, Inches(12.50), Inches(0.98), CARD)
        add_rect(s, Inches(0.40), y, Inches(0.10), Inches(0.98), GOLD if i % 2 == 0 else BLUE)
        add_textbox(s, Inches(0.70), y + Inches(0.18), Inches(1.05), Inches(0.62), num, 22, True, GOLD, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(1.85), y + Inches(0.16), Inches(10.7), Inches(0.36), title, 18, True, WHITE)
        add_textbox(s, Inches(1.85), y + Inches(0.52), Inches(10.7), Inches(0.32), desc, 13, False, MUTED)
    footer(s, 2)
    return s


def s03_snapshot(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Pipeline snapshot",
        "Tallgrass RIS Assessment opportunity at a glance",
        "IQ-onboarded install base is already instrumented — the assessment is the consulting unlock",
    )
    kpis = [
        ("TOTAL ADOPTING", "233", "IQ / telemetry live", BLUE),
        ("CORE TARGETS", "10", "Named Tier 1 & 2", GOLD),
        ("INSTALL FOOTPRINT", "$319M+", "Cisco estate, core 10", TEAL),
        ("DIRECT RIS TAM", "$1.2–1.8M", "Assessment SOW TAM", GOLD),
        ("DOWNSTREAM", "$25M+", "Routing & security", GREEN),
    ]
    for i, (lab, val, cap, acc) in enumerate(kpis):
        kpi_card(s, Inches(0.32) + Inches(i * 2.58), Inches(1.42), Inches(2.46), Inches(1.22), lab, val, cap, acc)

    add_round(s, Inches(0.40), Inches(2.88), Inches(6.20), Inches(4.05), CARD)
    add_textbox(s, Inches(0.65), Inches(3.05), Inches(5.75), Inches(0.36), "The unlock", 16, True, GOLD)
    bullet_block(
        s,
        Inches(0.65),
        Inches(3.48),
        Inches(5.75),
        Inches(3.25),
        [
            "Customers already onboarded to Catalyst Center, Intersight, CX Cloud, Splunk FSO, or ISE stream configuration, health, and software data automatically.",
            "Cisco CX can run a high-impact architectural resilience and security audit with zero customer onboarding friction.",
            "The RIS Professional Assessment turns telemetry into a board-ready 3-year modernization roadmap.",
            "That roadmap is the on-ramp to multi-million-dollar routing, switching, and cloud security programs.",
        ],
        size=13,
    )

    add_round(s, Inches(6.80), Inches(2.88), Inches(6.10), Inches(4.05), CARD)
    add_textbox(s, Inches(7.05), Inches(3.05), Inches(5.65), Inches(0.36), "What to stage", 16, True, GOLD)
    rows = [
        ("Assessment SOW range", "$75k – $150k"),
        ("Buying program attach", "FY27 services programs"),
        ("Primary motion", "RIS Professional Assessment"),
        ("Telemetry sources", "Cat Center · Intersight · CX Cloud"),
        ("Also in play", "Splunk FSO · ISE"),
        ("Pull-through thesis", "Routing · switching · cloud security"),
        ("First-wave AMs", "Randolph · Garwood · Pavalone · Ball"),
        ("First-wave names", "UMB · Garmin · DFA · Lockton"),
    ]
    table = add_table(s, len(rows) + 1, 2, Inches(7.05), Inches(3.48), Inches(5.60), Inches(3.20))
    set_col_widths(table, [Inches(2.55), Inches(3.05)])
    style_cell(table.cell(0, 0), "FIELD", 10, True, WHITE, BLUE_HEX)
    style_cell(table.cell(0, 1), "DETAIL", 10, True, WHITE, BLUE_HEX)
    for i, (k, v) in enumerate(rows, 1):
        fill = CARD_ALT_HEX if i % 2 == 0 else CARD_HEX
        style_cell(table.cell(i, 0), k, 11, False, MUTED, fill)
        style_cell(table.cell(i, 1), v, 11, True, WHITE, fill)
    footer(s, 3)
    return s


def s04_rationale(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Strategic rationale",
        "IQ telemetry is the RIS Assessment unlock",
        "Highest propensity sits with accounts already streaming automated configuration, health, and software data",
    )

    sources = [
        ("Catalyst Center", "Campus fabric, switching, and wireless telemetry"),
        ("Intersight", "Compute, DC fabric, and operations health"),
        ("CX Cloud", "Installed-base, LDOS, and contract telemetry"),
        ("Splunk FSO", "Full-stack observability and Zero Trust signals"),
        ("ISE", "Identity, segmentation, and policy posture"),
    ]
    for i, (title, body) in enumerate(sources):
        x = Inches(0.35) + Inches(i * 2.58)
        add_round(s, x, Inches(1.42), Inches(2.46), Inches(1.70), CARD)
        add_rect(s, x, Inches(1.42), Inches(2.46), Inches(0.07), GOLD if i % 2 == 0 else BLUE)
        add_textbox(s, x + Inches(0.14), Inches(1.60), Inches(2.18), Inches(0.55), title, 14, True, WHITE)
        add_textbox(s, x + Inches(0.14), Inches(2.18), Inches(2.18), Inches(0.75), body, 12, False, MUTED)

    cards = [
        (
            "Zero onboarding friction",
            "Because the customer is already sending IQ data, CX engineers can benchmark routing resilience, failover latency, and configuration drift without a lengthy manual discovery cycle.",
        ),
        (
            "Board-ready findings",
            "The deliverable is an executive findings deck: routing convergence, latent failover risk, LDOS exposure, and a prioritized 3-year modernization roadmap.",
        ),
        (
            "Assessment is the wedge",
            "A $75k–$150k RIS Professional Assessment is the consulting SOW. Downstream pull-through is the $25M+ routing, switching, and cloud security program it qualifies.",
        ),
    ]
    for i, (title, body) in enumerate(cards):
        x = Inches(0.35) + Inches(i * 4.30)
        add_round(s, x, Inches(3.35), Inches(4.12), Inches(3.55), CARD)
        add_textbox(s, x + Inches(0.22), Inches(3.52), Inches(3.70), Inches(0.70), title, 16, True, GOLD)
        add_textbox(s, x + Inches(0.22), Inches(4.28), Inches(3.70), Inches(2.35), body, 14, False, BODY)
    footer(s, 4)
    return s


def s05_framework(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Qualification",
        "Propensity evaluation framework",
        "Use these four pillars to prioritize outreach — every Tier 1 name scores on at least three",
    )
    pillars = [
        (
            "01",
            "Active IQ Telemetry",
            "Catalyst Center, Intersight, Cloud Monitoring, or Splunk connected.",
            "Pre-existing telemetry removes data-collection friction so CX can benchmark network resilience immediately.",
            BLUE,
        ),
        (
            "02",
            "Scale & Complexity",
            "Cisco footprint > $15M; distributed campus / multi-site WAN.",
            "Large distributed environments carry architectural debt, route-table fragmentation, and configuration drift.",
            GOLD,
        ),
        (
            "03",
            "Imminent LDOS Exposure",
            "Hardware / software reaching Last Date of Support in FY26–FY28.",
            "Creates a C-level compelling event to fund an architectural risk review before hardware failure.",
            TEAL,
        ),
        (
            "04",
            "Strategic Transformation",
            "Active SD-WAN, data center builds, or Zero Trust initiatives.",
            "RIS is the independent architectural validation SOW that de-risks multi-cloud transit and campus segmentation.",
            GREEN,
        ),
    ]
    for i, (num, title, criteria, why, accent) in enumerate(pillars):
        col, row = i % 2, i // 2
        x = Inches(0.40) + Inches(col * 6.45)
        y = Inches(1.42) + Inches(row * 2.75)
        add_round(s, x, y, Inches(6.25), Inches(2.55), CARD)
        add_rect(s, x, y, Inches(0.10), Inches(2.55), accent)
        add_textbox(s, x + Inches(0.30), y + Inches(0.18), Inches(1.00), Inches(0.40), num, 20, True, accent)
        add_textbox(s, x + Inches(1.30), y + Inches(0.22), Inches(4.65), Inches(0.40), title, 18, True, WHITE)
        add_textbox(s, x + Inches(0.30), y + Inches(0.78), Inches(5.70), Inches(0.70), "Qualify:  " + criteria, 13, False, CYAN)
        add_textbox(s, x + Inches(0.30), y + Inches(1.50), Inches(5.70), Inches(0.85), "Why it converts:  " + why, 13, False, BODY)
    footer(s, 5)
    return s


def s06_tier1_table(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Tier 1  ·  High propensity",
        "Ranked target accounts",
        "Six named accounts  ·  $216M+ combined Cisco footprint  ·  IQ already live",
    )
    headers = ["CUSTOMER / SAVM", "VERTICAL", "FOOTPRINT", "IQ MATURITY", "RECOMMENDED PLAY"]
    data = [
        ["UMB Financial Corp\nSAVM 203733540", "Banking / Finance", "$30.6M", "Adopt / Optimize\nSplunk FSO & Zero Trust", "Financial Hybrid Routing Audit"],
        ["Garmin International\nSAVM 203707065", "Global Tech / Mfg", "$66.4M", "Onboarded\nCat Center, Intersight, ACI", "DC Fabric & Core Routing Assessment"],
        ["Dairy Farmers of America\nSAVM 203860067", "Mfg & Distribution", "$41.7M", "Active Onboard\nCampus & Auto-WAN", "Industrial Edge & WAN Survivability"],
        ["Children's Mercy Hospital\nSAVM 203771956", "Pediatric Healthcare", "$33.4M", "Connected\nCat Center & Intersight", "Clinical Resilience & Risk Assessment"],
        ["Lockton Companies\nSAVM 203707061", "Insurance / Global", "$25.7M", "Cloud Monitoring\n+$14.8M WxC renewal", "Global Collaboration Routing Assessment"],
        ["Truman Medical (UHKC)\nSAVM 203771980", "Academic Healthcare", "$18.5M", "Use / Adopt\nCampus Analytics", "Healthcare Campus Network Audit"],
    ]
    table = add_table(s, 7, 5, Inches(0.28), Inches(1.40), Inches(12.75), Inches(5.55))
    set_col_widths(table, [Inches(2.85), Inches(2.15), Inches(1.55), Inches(2.85), Inches(3.35)])
    aligns = [PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT]
    for j, h in enumerate(headers):
        style_cell(table.cell(0, j), h, 11, True, WHITE, BLUE_HEX, align=aligns[j])
    for i, row in enumerate(data, 1):
        fill = CARD_ALT_HEX if i % 2 == 0 else CARD_HEX
        for j, val in enumerate(row):
            color = GOLD if j == 4 else WHITE
            bold = j in (0, 2, 4)
            style_cell(table.cell(i, j), val, 12, bold, color, fill, align=aligns[j])
    footer(s, 6)
    return s


def _account_card(slide, x, y, w, h, name, meta, triggers, play, accent):
    add_round(slide, x, y, w, h, CARD)
    add_rect(slide, x, y, w, Inches(0.08), accent)
    add_textbox(slide, x + Inches(0.20), y + Inches(0.20), w - Inches(0.40), Inches(0.38), name, 16, True, WHITE)
    add_textbox(slide, x + Inches(0.20), y + Inches(0.56), w - Inches(0.40), Inches(0.55), meta, 12, False, CYAN)
    add_textbox(slide, x + Inches(0.20), y + Inches(1.18), w - Inches(0.40), Inches(0.24), "ARCHITECTURAL TRIGGERS", 10, True, GOLD)
    add_textbox(slide, x + Inches(0.20), y + Inches(1.42), w - Inches(0.40), Inches(1.35), triggers, 13, False, BODY)
    add_textbox(slide, x + Inches(0.20), y + Inches(2.85), w - Inches(0.40), Inches(0.24), "RECOMMENDED SALES PLAY", 10, True, GOLD)
    add_textbox(slide, x + Inches(0.20), y + Inches(3.10), w - Inches(0.40), Inches(1.15), play, 13, False, WHITE)


def s07_tier1_wave_a(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Tier 1  ·  First-wave conversations",
        "UMB  ·  Garmin  ·  Dairy Farmers of America",
        "Primary Sales Play A — Telemetry Unlock Architectural Review",
    )
    cards = [
        (
            "UMB Financial Corp",
            "SAVM 203733540  ·  Banking  ·  $30.6M  ·  Tier 1 Adopt/Optimize",
            "Financial regulatory compliance, multi-site routing resilience, and a $1.5M Catalyst Center opportunity already in pipeline.",
            "Financial Hybrid Routing Audit — core banking transit resilience, SD-WAN failover latency, and Zero Trust segmentation integrity.",
            GOLD,
        ),
        (
            "Garmin International",
            "SAVM 203707065  ·  Global Tech / Mfg  ·  $66.4M  ·  Cat Center, Intersight, ACI",
            "$2.5M new data center build in pipeline. Global manufacturing supply-chain dependency on campus and DC fabric routing.",
            "DC Fabric & Core Routing Assessment — anchor RIS to de-risk the new DC build and optimize high-throughput campus fabric routing.",
            BLUE,
        ),
        (
            "Dairy Farmers of America",
            "SAVM 203860067  ·  Mfg & Distribution  ·  $41.7M  ·  Campus & Auto-WAN",
            "Distributed plant-network routing, $2M Zscaler takeout in flight, and multi-cloud interconnect dependencies.",
            "Industrial Edge & WAN Survivability — plant-to-cloud routing redundancy, SD-WAN path optimization, and IoT edge resilience.",
            TEAL,
        ),
    ]
    for i, args in enumerate(cards):
        _account_card(s, Inches(0.32) + Inches(i * 4.32), Inches(1.40), Inches(4.16), Inches(5.50), *args)
    footer(s, 7)
    return s


def s08_tier1_wave_b(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Tier 1  ·  Clinical, insurance, academic",
        "Children's Mercy  ·  Lockton  ·  Truman Medical",
        "Mix of Play A (Children's Mercy) and Play B (Lockton, Truman) — Pre-Transformation Derisking",
    )
    cards = [
        (
            "Children's Mercy Hospital",
            "SAVM 203771956  ·  Pediatric Healthcare  ·  $33.4M  ·  Cat Center & Intersight",
            "Zero-downtime clinical mandate, medical-device IoT segmentation, and a $1.09M upcoming LDOS hardware runway.",
            "Clinical Resilience & Risk Assessment — unpatched vulnerability mapping, medical-device isolation routing, and hospital core resilience.",
            GOLD,
        ),
        (
            "Lockton Companies",
            "SAVM 203707061  ·  Insurance / Global  ·  $25.7M  ·  Cloud Monitoring + $14.8M WxC",
            "Rapid global office expansion, multi-carrier SIP/WAN routing consolidation, and C1 modernization ahead of Webex Calling renewal.",
            "Global Collaboration Routing Assessment — WAN reliability, QoS posture for Webex Calling, and branch survivability.",
            BLUE,
        ),
        (
            "Truman Medical (UHKC)",
            "SAVM 203771980  ·  Academic Healthcare  ·  $18.5M  ·  Campus Analytics",
            "$1.35M Webex Calling DI 5-year renewal, high-density hospital campus switching refresh, and EHR telemetry dependencies.",
            "Healthcare Campus Network Audit — core campus switching resilience, wireless telemetry health, and EHR traffic routing redundancy.",
            TEAL,
        ),
    ]
    for i, args in enumerate(cards):
        _account_card(s, Inches(0.32) + Inches(i * 4.32), Inches(1.40), Inches(4.16), Inches(5.50), *args)
    footer(s, 8)
    return s


def s09_tier2(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Tier 2  ·  Strategic growth",
        "Next-wave IQ-connected opportunities",
        "Live telemetry plus a transformation trigger — stage after first-wave SOWs are in motion",
    )
    cards = [
        (
            "HNTB Holdings Inc",
            "Play B  ·  Derisking",
            "SAVM 203707063  ·  Engineering  ·  $21.6M",
            "Catalyst Center telemetry live. $1.2M Webex refresh in Commit.",
            "Hybrid Studio WAN Audit — multi-gigabit design-studio bandwidth, CAD/BIM cloud sync, and hybrid routing latency.",
        ),
        (
            "Black & Veatch Inc",
            "Play B  ·  SASE",
            "SAVM 203771954  ·  Critical Infrastructure  ·  $19.4M",
            "Intersight compute live. $1.14M Secure Access opportunity.",
            "Critical Infrastructure SASE Readiness — secure remote engineering transit and critical-infrastructure routing mesh.",
        ),
        (
            "WellSky Corporation",
            "Play A  ·  Telemetry",
            "SAVM 203839175  ·  Healthcare SaaS  ·  $16.2M",
            "21 adoption use cases active (Meraki, Compute, Zero Trust).",
            "Multi-Region SaaS Edge Resilience — cloud edge ingress/egress routing throughput, DC firewall resilience, and SD-WAN links.",
        ),
        (
            "Salina Regional Health",
            "Play A  ·  LDOS",
            "SAVM 203793443  ·  Regional Health  ·  $20.6M",
            "Campus switching and compute telemetry connected.",
            "Regional Clinic WAN Survivability — remote clinic router failover, telemedicine QoS, and LDOS replacement roadmaps.",
        ),
    ]
    for i, (name, tag, meta, iq, play) in enumerate(cards):
        col, row = i % 2, i // 2
        x = Inches(0.35) + Inches(col * 6.48)
        y = Inches(1.40) + Inches(row * 2.78)
        add_round(s, x, y, Inches(6.28), Inches(2.58), CARD)
        add_rect(s, x, y, Inches(0.10), Inches(2.58), GOLD if i % 2 == 0 else BLUE)
        add_textbox(s, x + Inches(0.28), y + Inches(0.14), Inches(3.70), Inches(0.36), name, 16, True, WHITE)
        add_textbox(s, x + Inches(4.00), y + Inches(0.16), Inches(2.05), Inches(0.32), tag, 11, True, GOLD, align=PP_ALIGN.RIGHT)
        add_textbox(s, x + Inches(0.28), y + Inches(0.52), Inches(5.75), Inches(0.32), meta, 12, False, CYAN)
        add_textbox(s, x + Inches(0.28), y + Inches(0.90), Inches(5.75), Inches(0.48), iq, 13, False, MUTED)
        add_textbox(s, x + Inches(0.28), y + Inches(1.42), Inches(5.75), Inches(0.95), play, 13, False, BODY)
    footer(s, 9)
    return s


def s10_play_a(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Sales Play A",
        'The "Telemetry Unlock" architectural review',
        "Primary motion for Garmin, Dairy Farmers of America, and UMB Financial",
    )
    add_round(s, Inches(0.40), Inches(1.40), Inches(6.20), Inches(5.50), CARD)
    add_textbox(s, Inches(0.65), Inches(1.58), Inches(5.75), Inches(0.36), "Positioning", 16, True, GOLD)
    add_textbox(
        s,
        Inches(0.65),
        Inches(2.05),
        Inches(5.75),
        Inches(1.70),
        "Position the RIS Professional Assessment as the natural consulting deliverable that turns raw IQ telemetry into a prioritized C-level executive roadmap. Do not sell a data-collection exercise — the data is already flowing.",
        15,
        False,
        BODY,
    )
    add_textbox(s, Inches(0.65), Inches(3.85), Inches(5.75), Inches(0.30), "Talk tracks", 14, True, CYAN)
    bullet_block(
        s,
        Inches(0.65),
        Inches(4.20),
        Inches(5.75),
        Inches(2.40),
        [
            "Your IQ integration is already complete — we start from live telemetry, not questionnaires.",
            "Deliverable is a board-ready findings deck, not a raw dump of configs.",
            "Output: routing convergence, latent failover risk, and a 3-year modernization sequence.",
        ],
        size=13,
    )

    add_round(s, Inches(6.80), Inches(1.40), Inches(6.10), Inches(5.50), CARD)
    add_rect(s, Inches(6.80), Inches(1.40), Inches(0.10), Inches(5.50), GOLD)
    add_textbox(s, Inches(7.15), Inches(1.58), Inches(5.50), Inches(0.30), "Executive discovery script", 14, True, GOLD)
    add_textbox(
        s,
        Inches(7.15),
        Inches(2.05),
        Inches(5.50),
        Inches(3.55),
        '"Because your environment is already integrated with Cisco IQ telemetry, our Advanced Services architecture team can conduct a comprehensive Routing & Infrastructure Resilience Assessment without lengthy manual network audits. We will deliver a board-ready findings deck analyzing routing convergence, latent failover risks, and a prioritized 3-year modernization roadmap."',
        15,
        False,
        WHITE,
    )
    add_textbox(
        s,
        Inches(7.15),
        Inches(5.80),
        Inches(5.50),
        Inches(0.80),
        "Ask next: Which of routing convergence, DC fabric, or WAN failover is the board most exposed to in the next 18 months?",
        13,
        False,
        CYAN,
    )
    footer(s, 10)
    return s


def s11_play_b(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Sales Play B",
        "The Pre-Transformation Derisking Audit",
        "Primary motion for Lockton, Truman Medical, and HNTB — attach RIS before the migration, not after the outage",
    )
    add_round(s, Inches(0.40), Inches(1.40), Inches(6.20), Inches(5.50), CARD)
    add_textbox(s, Inches(0.65), Inches(1.58), Inches(5.75), Inches(0.36), "Positioning", 16, True, GOLD)
    add_textbox(
        s,
        Inches(0.65),
        Inches(2.05),
        Inches(5.75),
        Inches(1.70),
        "Position RIS as mandatory risk mitigation before executing major cloud calling, SD-WAN, or data center migrations. The assessment is how the CIO buys day-one SLA confidence.",
        15,
        False,
        BODY,
    )
    add_textbox(s, Inches(0.65), Inches(3.85), Inches(5.75), Inches(0.30), "Attach it in front of", 14, True, CYAN)
    bullet_block(
        s,
        Inches(0.65),
        Inches(4.20),
        Inches(5.75),
        Inches(2.40),
        [
            "Lockton — $14.8M Webex Calling renewal and global WAN consolidation.",
            "Truman — $1.35M WxC DI 5-year renewal and campus switching refresh.",
            "HNTB — $1.2M Webex refresh in Commit plus hybrid studio WAN load.",
        ],
        size=13,
    )

    add_round(s, Inches(6.80), Inches(1.40), Inches(6.10), Inches(5.50), CARD)
    add_rect(s, Inches(6.80), Inches(1.40), Inches(0.10), Inches(5.50), BLUE)
    add_textbox(s, Inches(7.15), Inches(1.58), Inches(5.50), Inches(0.30), "Executive discovery script", 14, True, GOLD)
    add_textbox(
        s,
        Inches(7.15),
        Inches(2.05),
        Inches(5.50),
        Inches(3.20),
        '"Prior to rolling out our planned cloud and collaboration migration, our CX Advisory Practice conducts a high-fidelity Infrastructure Assessment to validate path diversity, QoS queueing integrity, and core route table stability, guaranteeing day-one performance SLAs."',
        15,
        False,
        WHITE,
    )
    add_textbox(
        s,
        Inches(7.15),
        Inches(5.50),
        Inches(5.50),
        Inches(1.10),
        "Ask next: If calling, CAD sync, or EHR traffic failed over incorrectly on day one, which executive owns that outage — and have we proven path diversity yet?",
        13,
        False,
        CYAN,
    )
    footer(s, 11)
    return s


def s12_action_plan(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Account executive action plan",
        "Six weeks from alignment to staged SOWs",
        "Target: $75k–$150k RIS Assessment SOWs in CCW, attached to FY27 service buying programs",
    )
    phases = [
        (
            "PHASE 1",
            "Weeks 1–2",
            "AM alignment",
            "Align with Robin Randolph, Gerry Garwood, Jeff Pavalone, and Paige Ball. Review target IQ profiles and confirm architectural priorities for UMB, Garmin, DFA, and Lockton.",
            "Deliverable: confirmed first-wave account list, named CIO / VP Infrastructure, and agreed play (A or B) per account.",
            GOLD,
        ),
        (
            "PHASE 2",
            "Weeks 3–4",
            "Executive discovery",
            "Outreach to customer CIO / VP of Infrastructure. Run 30-minute CX Advisory briefings using pre-populated IQ telemetry snapshots — no blank-page discovery.",
            "Deliverable: 30-minute CX Advisory Briefing complete; verbal sponsor for a RIS Assessment SOW.",
            BLUE,
        ),
        (
            "PHASE 3",
            "Weeks 5–6",
            "Scope and stage",
            "Scope Professional Services SOWs and stage them in CCW against FY27 service buying programs. Keep the assessment as the wedge; attach pull-through separately.",
            "Deliverable: $75k–$150k RIS Assessment SOW staged per first-wave account.",
            TEAL,
        ),
    ]
    for i, (phase, when, title, body, deliverable, accent) in enumerate(phases):
        x = Inches(0.35) + Inches(i * 4.30)
        add_round(s, x, Inches(1.40), Inches(4.12), Inches(5.50), CARD)
        add_rect(s, x, Inches(1.40), Inches(4.12), Inches(0.08), accent)
        add_textbox(s, x + Inches(0.22), Inches(1.60), Inches(3.70), Inches(0.26), f"{phase}  ·  {when}", 12, True, accent)
        add_textbox(s, x + Inches(0.22), Inches(1.95), Inches(3.70), Inches(0.45), title, 20, True, WHITE)
        add_textbox(s, x + Inches(0.22), Inches(2.50), Inches(3.70), Inches(2.35), body, 14, False, BODY)
        add_textbox(s, x + Inches(0.22), Inches(4.95), Inches(3.70), Inches(1.60), deliverable, 13, False, CYAN)
    footer(s, 12)
    return s


def s13_economics(prs):
    s = blank_slide(prs)
    header_bar(
        s,
        "Commercial thesis",
        "From assessment SOW to pull-through",
        "The RIS Professional Assessment is the wedge. Do not stop at the $75k–$150k SOW.",
    )
    steps = [
        ("1", "Stage the assessment", "$75k–$150k", "RIS Professional Assessment SOW in CCW, attached to FY27 service buying programs."),
        ("2", "Cover the core 10", "$1.2M–$1.8M", "Direct RIS Assessment TAM if the ten Tier 1 and Tier 2 names convert."),
        ("3", "Fund the roadmap", "$25M+", "Downstream routing, switching, and cloud security modernization pull-through."),
    ]
    for i, (num, title, value, body) in enumerate(steps):
        y = Inches(1.42) + Inches(i * 1.55)
        add_round(s, Inches(0.40), y, Inches(12.50), Inches(1.42), CARD)
        add_round(s, Inches(0.62), y + Inches(0.32), Inches(0.78), Inches(0.78), RGBColor(0x0B, 0x16, 0x24))
        add_textbox(s, Inches(0.62), y + Inches(0.32), Inches(0.78), Inches(0.78), num, 22, True, GOLD, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(1.65), y + Inches(0.18), Inches(6.20), Inches(0.40), title, 18, True, WHITE)
        add_textbox(s, Inches(1.65), y + Inches(0.62), Inches(6.20), Inches(0.60), body, 14, False, MUTED)
        add_textbox(s, Inches(8.20), y + Inches(0.28), Inches(4.40), Inches(0.85), value, 28, True, GOLD, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    add_textbox(
        s,
        Inches(0.50),
        Inches(6.20),
        Inches(12.3),
        Inches(0.75),
        "Anchor every briefing to a named architectural trigger already in motion — Garmin DC build, DFA Zscaler takeout, Lockton WxC renewal, Children's Mercy LDOS — so the assessment is risk insurance on a program the customer has already funded.",
        14,
        False,
        BODY,
    )
    footer(s, 13)
    return s


def s14_close(prs):
    s = blank_slide(prs)
    add_rect(s, 0, 0, Inches(0.18), SLIDE_H, GOLD)
    add_textbox(s, Inches(0.70), Inches(0.55), Inches(12.0), Inches(0.28), "NEXT 10 WORKING DAYS", 12, True, GOLD)
    add_textbox(s, Inches(0.70), Inches(0.90), Inches(12.0), Inches(0.55), "Open the first four conversations", 28, True, WHITE)
    add_textbox(
        s,
        Inches(0.70),
        Inches(1.50),
        Inches(12.0),
        Inches(0.40),
        "UMB  ·  Garmin  ·  Dairy Farmers of America  ·  Lockton  — then expand to Children's Mercy and Truman.",
        16,
        False,
        CYAN,
    )

    actions = [
        ("This week", "Sit with Randolph, Garwood, Pavalone, and Ball. Confirm IQ snapshots exist and pick Play A or Play B per account."),
        ("Week 2", "Book 30-minute CX Advisory briefings with CIO / VP Infrastructure. Lead with the script; leave with a SOW sponsor."),
        ("Weeks 3–6", "Scope $75k–$150k RIS Assessment SOWs and stage them in CCW against FY27 service buying programs."),
    ]
    for i, (when, body) in enumerate(actions):
        y = Inches(2.10) + Inches(i * 1.15)
        add_round(s, Inches(0.70), y, Inches(12.00), Inches(1.02), CARD)
        add_rect(s, Inches(0.70), y, Inches(0.10), Inches(1.02), GOLD if i == 0 else BLUE)
        add_textbox(s, Inches(1.05), y + Inches(0.12), Inches(2.20), Inches(0.78), when, 16, True, GOLD, anchor=MSO_ANCHOR.MIDDLE)
        add_textbox(s, Inches(3.40), y + Inches(0.16), Inches(9.05), Inches(0.72), body, 15, False, BODY, anchor=MSO_ANCHOR.MIDDLE)

    add_textbox(
        s,
        Inches(0.70),
        Inches(5.75),
        Inches(12.0),
        Inches(0.70),
        "The telemetry is already on. The assessment is how Tallgrass turns 233 IQ-onboarded accounts into a $25M+ routing and infrastructure modernization pipeline.",
        16,
        False,
        WHITE,
    )
    add_textbox(
        s,
        Inches(0.70),
        Inches(6.55),
        Inches(12.0),
        Inches(0.35),
        "CISCO CONFIDENTIAL  ·  Internal Sales Enablement  ·  Services & Software Buying Programs  ·  Circuit AI Super Agent Platform",
        12,
        False,
        MUTED,
    )
    return s


def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    s01_title(prs)
    s02_agenda(prs)
    s03_snapshot(prs)
    s04_rationale(prs)
    s05_framework(prs)
    s06_tier1_table(prs)
    s07_tier1_wave_a(prs)
    s08_tier1_wave_b(prs)
    s09_tier2(prs)
    s10_play_a(prs)
    s11_play_b(prs)
    s12_action_plan(prs)
    s13_economics(prs)
    s14_close(prs)

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUT_PATH))
    print(f"Wrote {OUT_PATH} ({TOTAL_SLIDES} slides)")


if __name__ == "__main__":
    build()
