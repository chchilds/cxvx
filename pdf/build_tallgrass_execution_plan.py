#!/usr/bin/env python3
"""Cisco Tallgrass Region — RIS Assessment Execution Plan (PDF).

Branding follows Cisco's core enablement palette used in the source
Tallgrass IQ / RIS propensity guide:

  Indigo  #005073   (PMS 2210C)
  Cyan    #00BCEB   (PMS 2995C)  Cisco Blue
  Grass   #6EBE4A   (PMS 360C)   Tallgrass regional accent
  Wheat   #FBAB18   (PMS 130C)   phase / milestone accent

Source: Tallgrass Region IQ Onboarded Accounts & RIS Propensity Guide
(Services & Software Buying Programs — internal sales enablement).
"""

from __future__ import annotations

import math
from pathlib import Path

from reportlab.lib.colors import Color, HexColor, white
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# --- Fonts ---
FONT_DIR = Path("/usr/share/fonts/truetype/liberation")
pdfmetrics.registerFont(TTFont("Sans", str(FONT_DIR / "LiberationSans-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Sans-Bold", str(FONT_DIR / "LiberationSans-Bold.ttf")))
pdfmetrics.registerFont(TTFont("Sans-Italic", str(FONT_DIR / "LiberationSans-Italic.ttf")))
pdfmetrics.registerFont(TTFont("Sans-BoldItalic", str(FONT_DIR / "LiberationSans-BoldItalic.ttf")))

# --- Cisco / Tallgrass palette ---
INDIGO = HexColor("#005073")
CYAN = HexColor("#00BCEB")
GRASS = HexColor("#6EBE4A")
WHEAT = HexColor("#FBAB18")
INK = HexColor("#1B2733")
MUTED = HexColor("#586574")
RULE = HexColor("#D9DCE1")
BAND = HexColor("#F8FAFC")
PILL_BG = HexColor("#E8F6DC")
CYAN_BG = HexColor("#E6F7FC")
INDIGO_SOFT = HexColor("#E8F1F5")
FOOTER_LINE = HexColor("#005073")

PAGE_W, PAGE_H = LETTER
LEFT = 0.55 * inch
RIGHT = PAGE_W - 0.55 * inch
TOP = PAGE_H - 0.42 * inch
BOTTOM = 0.48 * inch
CONTENT_W = RIGHT - LEFT

OUT_PATH = Path(__file__).resolve().parents[1] / "output" / (
    "Tallgrass_Region_RIS_Assessment_Execution_Plan.pdf"
)

TOTAL_PAGES = 12


def _wrap(text: str, font: str, size: float, width: float) -> list[str]:
    words = text.replace("\n", " \n ").split(" ")
    lines, cur = [], ""
    for w in words:
        if w == "\n":
            lines.append(cur)
            cur = ""
            continue
        trial = w if not cur else cur + " " + w
        if pdfmetrics.stringWidth(trial, font, size) <= width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


def draw_text(c, text, x, y, font="Sans", size=10, color=INK, align="left"):
    c.setFont(font, size)
    c.setFillColor(color)
    if align == "right":
        c.drawRightString(x, y, text)
    elif align == "center":
        c.drawCentredString(x, y, text)
    else:
        c.drawString(x, y, text)


def draw_para(c, text, x, y, w, font="Sans", size=9.5, color=INK, leading=12.5, align="left"):
    lines = _wrap(text, font, size, w)
    for i, line in enumerate(lines):
        draw_text(c, line, x if align != "center" else x + w / 2, y - i * leading, font, size, color, align)
    return y - len(lines) * leading


def round_rect(c, x, y, w, h, r=6, fill=None, stroke=None, sw=0.8):
    c.saveState()
    if fill:
        c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(sw)
    p = c.beginPath()
    p.moveTo(x + r, y)
    p.lineTo(x + w - r, y)
    p.arcTo(x + w - 2 * r, y, x + w, y + 2 * r, -90, 90)
    p.lineTo(x + w, y + h - r)
    p.arcTo(x + w - 2 * r, y + h - 2 * r, x + w, y + h, 0, 90)
    p.lineTo(x + r, y + h)
    p.arcTo(x, y + h - 2 * r, x + 2 * r, y + h, 90, 90)
    p.lineTo(x, y + r)
    p.arcTo(x, y, x + 2 * r, y + 2 * r, 180, 90)
    p.close()
    c.drawPath(p, fill=1 if fill else 0, stroke=1 if stroke else 0)
    c.restoreState()


def pill(c, text, x, y, font="Sans-Bold", size=7.5, fg=INDIGO, bg=CYAN_BG, pad_x=7, pad_y=3.5):
    tw = pdfmetrics.stringWidth(text, font, size)
    w, h = tw + pad_x * 2, size + pad_y * 2
    round_rect(c, x, y, w, h, r=h / 2, fill=bg)
    draw_text(c, text, x + pad_x, y + pad_y + 0.4, font, size, fg)
    return w


def grass_blades(c, x, y, scale=1.0, color=GRASS, alpha=1.0):
    """Simple prairie-grass cluster — Tallgrass regional mark."""
    c.saveState()
    col = Color(color.red, color.green, color.blue, alpha=alpha)
    c.setStrokeColor(col)
    c.setFillColor(col)
    c.setLineWidth(1.1 * scale)
    c.setLineCap(1)
    blades = [
        (0, 18, -4),
        (3, 22, -1),
        (6, 16, 2),
        (9, 20, 4),
        (12, 14, 6),
    ]
    for dx, h, lean in blades:
        p = c.beginPath()
        p.moveTo(x + dx * scale, y)
        p.curveTo(
            x + (dx + lean * 0.3) * scale,
            y + h * 0.45 * scale,
            x + (dx + lean) * scale,
            y + h * 0.75 * scale,
            x + (dx + lean * 1.2) * scale,
            y + h * scale,
        )
        c.drawPath(p, stroke=1, fill=0)
    c.restoreState()


def weave(c, x, y, w, h):
    """Cisco-style intersecting arcs used as cover atmosphere."""
    c.saveState()
    c.setLineCap(1)
    curves = [
        (CYAN, 1.6, 0.35),
        (INDIGO, 1.3, 0.22),
        (GRASS, 1.4, 0.28),
    ]
    for i, (col, lw, a) in enumerate(curves):
        c.setStrokeColor(Color(col.red, col.green, col.blue, alpha=a))
        c.setLineWidth(lw)
        p = c.beginPath()
        y0 = y + h * (0.15 + i * 0.22)
        p.moveTo(x, y0)
        p.curveTo(x + w * 0.25, y0 + h * 0.55, x + w * 0.55, y0 - h * 0.15, x + w, y0 + h * 0.35)
        c.drawPath(p, stroke=1, fill=0)
    c.restoreState()


def header(c, section: str):
    # Top indigo rule
    c.setFillColor(INDIGO)
    c.rect(0, PAGE_H - 8, PAGE_W, 8, fill=1, stroke=0)
    c.setFillColor(GRASS)
    c.rect(0, PAGE_H - 11, PAGE_W, 3, fill=1, stroke=0)

    grass_blades(c, LEFT, PAGE_H - 36, scale=0.72, color=GRASS)
    draw_text(c, "CISCO", LEFT + 22, PAGE_H - 32, "Sans-Bold", 9, INDIGO)
    c.setFillColor(CYAN)
    c.rect(LEFT + 58, PAGE_H - 31, 1.2, 10, fill=1, stroke=0)
    draw_text(c, "TALLGRASS REGION", LEFT + 64, PAGE_H - 32, "Sans-Bold", 9, GRASS)

    pill(c, section.upper(), RIGHT - pdfmetrics.stringWidth(section.upper(), "Sans-Bold", 7.5) - 16, PAGE_H - 36)

    c.setStrokeColor(RULE)
    c.setLineWidth(0.6)
    c.line(LEFT, PAGE_H - 44, RIGHT, PAGE_H - 44)


def footer(c, page: int):
    c.setStrokeColor(INDIGO)
    c.setLineWidth(1.1)
    c.line(LEFT, 32, RIGHT, 32)
    draw_text(
        c,
        "CISCO CONFIDENTIAL  ·  Internal Sales Enablement  ·  Services & Software Buying Programs",
        LEFT,
        18,
        "Sans",
        7,
        MUTED,
    )
    draw_text(c, f"{page}  /  {TOTAL_PAGES}", RIGHT, 18, "Sans-Bold", 8, INDIGO, align="right")


def h1(c, title, y, subtitle=None):
    draw_text(c, title, LEFT, y, "Sans-Bold", 16, INDIGO)
    y -= 6
    c.setStrokeColor(CYAN)
    c.setLineWidth(2)
    c.line(LEFT, y, LEFT + 42, y)
    y -= 16
    if subtitle:
        y = draw_para(c, subtitle, LEFT, y, CONTENT_W, "Sans", 9.5, MUTED, 12.5) - 6
    return y


def kpi_row(c, items, y, h=72):
    n = len(items)
    gap = 8
    w = (CONTENT_W - gap * (n - 1)) / n
    for i, (val, label) in enumerate(items):
        x = LEFT + i * (w + gap)
        round_rect(c, x, y - h, w, h, r=7, fill=white, stroke=INDIGO, sw=1.3)
        draw_text(c, val, x + w / 2, y - 26, "Sans-Bold", 14, INDIGO, align="center")
        draw_para(c, label.upper(), x + 6, y - 44, w - 12, "Sans", 6.5, MUTED, 8.5, align="center")
    return y - h - 12


def table(c, headers, rows, y, col_w, row_h=28, header_h=20, fonts=None, sizes=None, aligns=None, wrap_cols=None):
    """Draw a branded table. Returns y under the table."""
    fonts = fonts or ["Sans"] * len(headers)
    sizes = sizes or [8] * len(headers)
    aligns = aligns or ["left"] * len(headers)
    wrap_cols = set(wrap_cols or range(len(headers)))
    x0 = LEFT
    c.setFillColor(INDIGO)
    c.rect(x0, y - header_h, sum(col_w), header_h, fill=1, stroke=0)
    cx = x0
    for i, h in enumerate(headers):
        draw_text(c, h.upper(), cx + 6, y - header_h + 6.5, "Sans-Bold", 7, white)
        cx += col_w[i]
    y -= header_h
    for r_i, row in enumerate(rows):
        wrapped = []
        rh = row_h
        for i, cell in enumerate(row):
            font = "Sans-Bold" if i == 0 else fonts[i]
            lines = _wrap(str(cell), font, sizes[i], col_w[i] - 12)
            wrapped.append((lines, font))
            rh = max(rh, 14 + len(lines) * 11)
        bg = BAND if r_i % 2 == 0 else white
        c.setFillColor(bg)
        c.rect(x0, y - rh, sum(col_w), rh, fill=1, stroke=0)
        c.setStrokeColor(RULE)
        c.setLineWidth(0.4)
        c.line(x0, y - rh, x0 + sum(col_w), y - rh)
        cx = x0
        for i, (lines, font) in enumerate(wrapped):
            yy = y - 13
            for line in lines:
                if aligns[i] == "center":
                    draw_text(c, line, cx + col_w[i] / 2, yy, font, sizes[i], INK, "center")
                else:
                    draw_text(c, line, cx + 6, yy, font, sizes[i], INK)
                yy -= 11
            cx += col_w[i]
        y -= rh
    return y - 10


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------


def page_cover(c):
    # Full-bleed indigo
    c.setFillColor(INDIGO)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)
    c.setFillColor(GRASS)
    c.rect(0, 0, 14, PAGE_H, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.rect(14, 0, 4, PAGE_H, fill=1, stroke=0)

    weave(c, 40, 420, 540, 220)
    grass_blades(c, 48, 700, scale=1.6, color=GRASS, alpha=0.95)
    grass_blades(c, 78, 690, scale=1.2, color=CYAN, alpha=0.55)

    draw_text(c, "CISCO", 48, 660, "Sans-Bold", 11, CYAN)
    draw_text(c, "TALLGRASS REGION", 48, 628, "Sans-Bold", 22, GRASS)
    c.setFillColor(CYAN)
    c.rect(48, 612, 72, 3, fill=1, stroke=0)

    draw_text(c, "RIS Assessment", 48, 555, "Sans-Bold", 28, white)
    draw_text(c, "Execution Plan", 48, 518, "Sans-Bold", 28, white)
    draw_text(
        c,
        "IQ Onboarded Accounts  ·  Routing & Infrastructure Services",
        48,
        488,
        "Sans",
        12,
        CYAN,
    )

    y = 430
    facts = [
        ("233", "Adopting accounts"),
        ("10", "Core targets"),
        ("$319M+", "Install footprint"),
        ("$1.2–1.8M", "RIS Assessment TAM"),
        ("$25M+", "Pull-through"),
    ]
    fw = 96
    for i, (v, lab) in enumerate(facts):
        x = 48 + i * (fw + 8)
        round_rect(c, x, y - 58, fw, 58, r=6, fill=HexColor("#083D54"), stroke=None)
        draw_text(c, v, x + fw / 2, y - 26, "Sans-Bold", 12, white, "center")
        draw_text(c, lab.upper(), x + fw / 2, y - 44, "Sans", 6.2, CYAN, "center")

    draw_text(c, "PIPELINE GENERATION PLAN", 48, 330, "Sans-Bold", 9, WHEAT)
    y = draw_para(
        c,
        "Six-week path to staged $75k–$150k RIS Professional Assessment SOWs on ten IQ-onboarded named accounts, attached to FY27 service buying programs — and a 90-day horizon to convert findings into routing, switching, and cloud security pull-through.",
        48,
        310,
        500,
        "Sans",
        10,
        HexColor("#D5E6EE"),
        14,
    )

    draw_text(c, "ACCOUNT ALIGNMENT", 48, 210, "Sans-Bold", 8, CYAN)
    draw_text(
        c,
        "Robin Randolph   ·   Gerry Garwood   ·   Jeff Pavalone   ·   Paige Ball",
        48,
        192,
        "Sans",
        11,
        white,
    )
    draw_text(c, "Services & Software Buying Programs  ·  CX Advisory Practice", 48, 172, "Sans", 9, HexColor("#9BB8C4"))

    draw_text(
        c,
        "CISCO CONFIDENTIAL  ·  Internal Sales Enablement  ·  Not for customer distribution",
        48,
        48,
        "Sans",
        8,
        HexColor("#7FA3B3"),
    )
    draw_text(c, "Circuit AI Super Agent Platform", 48, 32, "Sans", 8, HexColor("#7FA3B3"))


def page_objective(c):
    header(c, "Why this motion")
    y = h1(
        c,
        "Turn live IQ telemetry into staged RIS SOWs",
        PAGE_H - 68,
        "Customers in Tallgrass who already stream Catalyst Center, Intersight, CX Cloud, Splunk FSO, or ISE data are the highest-propensity buyers of the Routing & Infrastructure Services Professional Assessment.",
    )
    y -= 4
    round_rect(c, LEFT, y - 78, CONTENT_W, 78, r=8, fill=BAND, stroke=None)
    draw_text(c, "MISSION  ·  NEXT 6 WEEKS", LEFT + 14, y - 18, "Sans-Bold", 8, GRASS)
    draw_para(
        c,
        "Align the four named AMs, run 30-minute CX Advisory briefings from pre-populated IQ snapshots, and stage $75k–$150k RIS Assessment SOWs in CCW on the first-wave accounts (UMB, Garmin, DFA, Lockton) — then Children's Mercy and Truman — attached to FY27 service buying programs.",
        LEFT + 14,
        y - 36,
        CONTENT_W - 28,
        "Sans",
        9.5,
        INK,
        13,
    )
    y -= 96

    draw_text(c, "WHY TALLGRASS, WHY NOW", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 14
    cards = [
        (
            "Zero onboarding friction",
            "IQ is already flowing. CX can benchmark routing resilience, failover latency, and configuration drift without a manual discovery cycle.",
        ),
        (
            "Compelling events are funded",
            "Garmin DC build, DFA Zscaler takeout, Lockton WxC renewal, Children's Mercy LDOS — the customer has already budgeted the program RIS de-risks.",
        ),
        (
            "Assessment is the wedge",
            "The $75k–$150k SOW is not the prize. It qualifies the $25M+ routing, switching, and cloud security modernization that follows.",
        ),
    ]
    cw = (CONTENT_W - 16) / 3
    for i, (t, b) in enumerate(cards):
        x = LEFT + i * (cw + 8)
        round_rect(c, x, y - 118, cw, 118, r=8, fill=white, stroke=RULE, sw=0.8)
        c.setFillColor(GRASS if i == 0 else CYAN if i == 1 else WHEAT)
        c.rect(x, y - 4, cw, 4, fill=1, stroke=0)
        draw_text(c, t, x + 10, y - 24, "Sans-Bold", 9.5, INDIGO)
        draw_para(c, b, x + 10, y - 42, cw - 20, "Sans", 8.5, MUTED, 11.5)
    y -= 140

    draw_text(c, "SUCCESS LOOKS LIKE", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 8
    kpis = [
        ("4", "First-wave briefings by week 4"),
        ("6", "Tier 1 SOWs staged by week 6"),
        ("$75–150k", "Assessment envelope per account"),
        ("$25M+", "Pull-through held as next action"),
    ]
    y = kpi_row(c, kpis, y, h=74)

    draw_text(c, "TELEMETRY SOURCES ALREADY ON THE WIRE", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 16
    sources = [
        ("Catalyst Center", "Campus fabric, switching, wireless"),
        ("Intersight", "Compute, DC fabric, operations"),
        ("CX Cloud", "Installed base, LDOS, contracts"),
        ("Splunk FSO", "Observability, Zero Trust signals"),
        ("ISE", "Identity, segmentation, policy"),
    ]
    sw = (CONTENT_W - 24) / 5
    for i, (t, b) in enumerate(sources):
        x = LEFT + i * (sw + 6)
        round_rect(c, x, y - 62, sw, 62, r=6, fill=INDIGO_SOFT, stroke=None)
        draw_text(c, t, x + 8, y - 20, "Sans-Bold", 8, INDIGO)
        draw_para(c, b, x + 8, y - 36, sw - 14, "Sans", 7.5, MUTED, 10)
    footer(c, 2)


def page_framework(c):
    header(c, "Qualification")
    y = h1(
        c,
        "Propensity evaluation framework",
        PAGE_H - 68,
        "Every first-wave name scores on at least three of four pillars. Use this as the go / no-go before a CIO briefing is booked.",
    )
    pillars = [
        (
            "01",
            "Active IQ Telemetry",
            "Qualify",
            "Catalyst Center, Intersight, Cloud Monitoring, or Splunk connected.",
            "Why it converts",
            "Pre-existing telemetry removes data-collection friction so CX can benchmark network resilience immediately.",
            GRASS,
        ),
        (
            "02",
            "Scale & Complexity",
            "Qualify",
            "Cisco footprint greater than $15M; distributed campus / multi-site WAN.",
            "Why it converts",
            "Large distributed environments carry architectural debt, route-table fragmentation, and configuration drift.",
            CYAN,
        ),
        (
            "03",
            "Imminent LDOS Exposure",
            "Qualify",
            "Hardware / software reaching Last Date of Support in FY26–FY28.",
            "Why it converts",
            "Creates a C-level compelling event to fund an architectural risk review before hardware failure.",
            WHEAT,
        ),
        (
            "04",
            "Strategic Transformation",
            "Qualify",
            "Active SD-WAN migrations, data center builds, or Zero Trust initiatives.",
            "Why it converts",
            "RIS is the independent architectural validation SOW that de-risks multi-cloud transit and campus segmentation.",
            INDIGO,
        ),
    ]
    for i, (num, title, qlab, q, wlab, w, accent) in enumerate(pillars):
        yy = y - i * 108
        round_rect(c, LEFT, yy - 98, CONTENT_W, 98, r=8, fill=white, stroke=RULE, sw=0.7)
        c.setFillColor(accent)
        c.rect(LEFT, yy - 98, 7, 98, fill=1, stroke=0)
        draw_text(c, num, LEFT + 20, yy - 28, "Sans-Bold", 16, accent)
        draw_text(c, title, LEFT + 58, yy - 26, "Sans-Bold", 13, INDIGO)
        draw_text(c, qlab.upper(), LEFT + 58, yy - 48, "Sans-Bold", 7.5, GRASS)
        draw_para(c, q, LEFT + 130, yy - 48, 380, "Sans", 9, INK, 12)
        draw_text(c, wlab.upper(), LEFT + 58, yy - 72, "Sans-Bold", 7.5, CYAN)
        draw_para(c, w, LEFT + 130, yy - 72, 380, "Sans", 9, MUTED, 12)
    footer(c, 3)


def page_tier1_roster(c):
    header(c, "Tier 1 roster")
    y = h1(
        c,
        "Ranked target accounts — high propensity",
        PAGE_H - 68,
        "Six named accounts. $216M combined Cisco footprint. IQ already live. First-wave conversations: UMB, Garmin, DFA, Lockton.",
    )
    y -= 4
    headers = ["Customer / SAVM", "Vertical", "Footprint", "IQ maturity", "Recommended play"]
    col_w = [128, 88, 62, 118, 144]
    rows = [
        ["UMB Financial Corp\nSAVM 203733540", "Banking / Finance", "$30.6M", "Adopt / Optimize\nSplunk FSO & Zero Trust", "Financial Hybrid Routing Audit"],
        ["Garmin International\nSAVM 203707065", "Global Tech / Mfg", "$66.4M", "Onboarded\nCat Center, Intersight, ACI", "DC Fabric & Core Routing Assessment"],
        ["Dairy Farmers of America\nSAVM 203860067", "Mfg & Distribution", "$41.7M", "Active Onboard\nCampus & Auto-WAN", "Industrial Edge & WAN Survivability"],
        ["Children's Mercy Hospital\nSAVM 203771956", "Pediatric Healthcare", "$33.4M", "Connected\nCat Center & Intersight", "Clinical Resilience & Risk Assessment"],
        ["Lockton Companies\nSAVM 203707061", "Insurance / Global", "$25.7M", "Cloud Monitoring\n+$14.8M WxC renewal", "Global Collaboration Routing Assessment"],
        ["Truman Medical (UHKC)\nSAVM 203771980", "Academic Healthcare", "$18.5M", "Use / Adopt\nCampus Analytics", "Healthcare Campus Network Audit"],
    ]
    y = table(c, headers, rows, y, col_w, row_h=36, wrap_cols=[0, 1, 3, 4], sizes=[8, 8, 8.5, 7.5, 8])

    round_rect(c, LEFT, y - 72, CONTENT_W, 72, r=8, fill=CYAN_BG, stroke=None)
    draw_text(c, "FIRST-WAVE RULE", LEFT + 14, y - 18, "Sans-Bold", 8, INDIGO)
    draw_para(
        c,
        "Do not open all six in week 1. Phase 1 of the source plan is AM alignment on UMB, Garmin, DFA, and Lockton. Children's Mercy and Truman run in parallel only if IQ snapshots and a named VP Infrastructure are already in hand. Tier 2 (HNTB, Black & Veatch, WellSky, Salina Regional) is Horizon 2 — after first-wave SOWs are staged.",
        LEFT + 14,
        y - 36,
        CONTENT_W - 28,
        "Sans",
        9,
        INK,
        12.5,
    )
    footer(c, 4)


def _account_block(c, x, y, w, h, name, meta, play, trigger, next_action, accent):
    round_rect(c, x, y - h, w, h, r=8, fill=white, stroke=RULE, sw=0.7)
    c.setFillColor(accent)
    c.rect(x, y - 5, w, 5, fill=1, stroke=0)
    inner = w - 20
    cy = y - 22
    draw_text(c, name, x + 10, cy, "Sans-Bold", 10, INDIGO)
    cy = draw_para(c, meta, x + 10, cy - 14, inner, "Sans", 7.5, MUTED, 10) - 8
    draw_text(c, "SALES PLAY", x + 10, cy, "Sans-Bold", 6.5, GRASS)
    cy = draw_para(c, play, x + 10, cy - 12, inner, "Sans-Bold", 8, INK, 10.5) - 8
    draw_text(c, "ARCHITECTURAL TRIGGER", x + 10, cy, "Sans-Bold", 6.5, CYAN)
    cy = draw_para(c, trigger, x + 10, cy - 12, inner, "Sans", 8, MUTED, 10.5) - 8
    draw_text(c, "THIS WEEK'S MOVE", x + 10, cy, "Sans-Bold", 6.5, WHEAT)
    draw_para(c, next_action, x + 10, cy - 12, inner, "Sans", 8, INK, 10.5)


def page_wave_a(c):
    header(c, "First-wave execution")
    y = h1(
        c,
        "UMB  ·  Garmin  ·  Dairy Farmers of America",
        PAGE_H - 68,
        "Primary motion: Sales Play A — Telemetry Unlock Architectural Review. These three already have IQ density and a funded architectural trigger.",
    )
    w = (CONTENT_W - 16) / 3
    h = 292
    cards = [
        (
            "UMB Financial Corp",
            "SAVM 203733540  ·  Banking  ·  $30.6M  ·  Splunk FSO & Zero Trust",
            "Financial Hybrid Routing Audit — core banking transit, SD-WAN failover latency, Zero Trust segmentation.",
            "Regulatory compliance plus a $1.5M Catalyst Center opportunity already in pipeline. Multi-site routing resilience is the board risk.",
            "AM confirms IQ snapshot (Splunk FSO + campus). Book 30-min CIO / VP Infra briefing. Lead with Play A script.",
            GRASS,
        ),
        (
            "Garmin International",
            "SAVM 203707065  ·  Global Tech / Mfg  ·  $66.4M  ·  Cat Center, Intersight, ACI",
            "DC Fabric & Core Routing Assessment — de-risk the new data center build; optimize campus fabric throughput.",
            "$2.5M new DC build in pipeline. Global manufacturing supply chain depends on campus and DC fabric routing.",
            "Pull Cat Center + Intersight + ACI health into one snapshot. Position RIS as day-one insurance on the DC program.",
            CYAN,
        ),
        (
            "Dairy Farmers of America",
            "SAVM 203860067  ·  Mfg & Dist.  ·  $41.7M  ·  Campus & Auto-WAN",
            "Industrial Edge & WAN Survivability — plant-to-cloud redundancy, SD-WAN pathing, IoT edge resilience.",
            "Distributed plant network. $2M Zscaler takeout in flight. Multi-cloud interconnect is live risk.",
            "Map plant-to-cloud paths from Auto-WAN telemetry. Attach RIS in front of the Zscaler takeout, not after cutover.",
            WHEAT,
        ),
    ]
    for i, args in enumerate(cards):
        _account_block(c, LEFT + i * (w + 8), y, w, h, *args)

    y = y - h - 18
    round_rect(c, LEFT, y - 88, CONTENT_W, 88, r=8, fill=BAND, stroke=None)
    draw_text(c, "PLAY A  ·  OPENING LINE", LEFT + 14, y - 18, "Sans-Bold", 8, INDIGO)
    draw_para(
        c,
        '"Because your environment is already integrated with Cisco IQ telemetry, our Advanced Services architecture team can conduct a comprehensive Routing & Infrastructure Resilience Assessment without lengthy manual network audits. We will deliver a board-ready findings deck analyzing routing convergence, latent failover risks, and a prioritized 3-year modernization roadmap."',
        LEFT + 14,
        y - 36,
        CONTENT_W - 28,
        "Sans-Italic",
        9,
        INK,
        12.5,
    )
    footer(c, 5)


def page_wave_b(c):
    header(c, "Tier 1  ·  Wave 1b")
    y = h1(
        c,
        "Children's Mercy  ·  Lockton  ·  Truman Medical",
        PAGE_H - 68,
        "Lockton and Truman are Play B (pre-transformation derisking). Children's Mercy is Play A with an LDOS compelling event. Sequence Lockton with the first wave; CMH and Truman as snapshots land.",
    )
    w = (CONTENT_W - 16) / 3
    h = 292
    cards = [
        (
            "Children's Mercy Hospital",
            "SAVM 203771956  ·  Pediatric Healthcare  ·  $33.4M  ·  Cat Center & Intersight",
            "Clinical Resilience & Risk Assessment — unpatched exposure, device isolation, hospital core.",
            "Zero-downtime clinical mandate. Medical-device IoT segmentation. $1.09M LDOS hardware runway.",
            "Build LDOS + unpatched inventory from Cat Center/Intersight. Brief clinical infrastructure leadership on isolation routing.",
            GRASS,
        ),
        (
            "Lockton Companies",
            "SAVM 203707061  ·  Insurance / Global  ·  $25.7M  ·  Cloud Monitoring + $14.8M WxC",
            "Global Collaboration Routing Assessment — WAN reliability, Webex Calling QoS, branch survivability.",
            "Rapid global office expansion. Multi-carrier SIP/WAN consolidation. C1 modernization ahead of WxC renewal.",
            "First-wave account. Attach RIS in front of the $14.8M Webex Calling renewal. Play B script — day-one SLA, not a survey.",
            CYAN,
        ),
        (
            "Truman Medical (UHKC)",
            "SAVM 203771980  ·  Academic Healthcare  ·  $18.5M  ·  Campus Analytics",
            "Healthcare Campus Network Audit — core switching, wireless telemetry, EHR path redundancy.",
            "$1.35M Webex Calling DI 5-year renewal. High-density campus switching refresh. EHR telemetry dependencies.",
            "Confirm campus analytics snapshot covers EHR paths. Position RIS as the condition of the WxC DI renewal, not a side project.",
            WHEAT,
        ),
    ]
    for i, args in enumerate(cards):
        _account_block(c, LEFT + i * (w + 8), y, w, h, *args)

    y = y - h - 18
    round_rect(c, LEFT, y - 88, CONTENT_W, 88, r=8, fill=CYAN_BG, stroke=None)
    draw_text(c, "PLAY B  ·  OPENING LINE", LEFT + 14, y - 18, "Sans-Bold", 8, INDIGO)
    draw_para(
        c,
        '"Prior to rolling out our planned cloud and collaboration migration, our CX Advisory Practice conducts a high-fidelity Infrastructure Assessment to validate path diversity, QoS queueing integrity, and core route table stability, guaranteeing day-one performance SLAs."',
        LEFT + 14,
        y - 36,
        CONTENT_W - 28,
        "Sans-Italic",
        9,
        INK,
        12.5,
    )
    footer(c, 6)


def page_tier2(c):
    header(c, "Horizon 2")
    y = h1(
        c,
        "Tier 2 — stage after first-wave SOWs are in motion",
        PAGE_H - 68,
        "Same propensity math, slightly smaller footprint. Do not let these four displace UMB, Garmin, DFA, or Lockton on the week-1 calendar.",
    )
    headers = ["Customer / SAVM", "Footprint", "IQ now", "Targeted RIS positioning", "Play"]
    col_w = [130, 58, 130, 162, 60]
    rows = [
        [
            "HNTB Holdings Inc\nSAVM 203707063  ·  Engineering",
            "$21.6M",
            "Catalyst Center live.\n$1.2M Webex refresh in Commit.",
            "Hybrid Studio WAN Audit — multi-gigabit design-studio bandwidth, CAD/BIM cloud sync, hybrid routing latency.",
            "Play B",
        ],
        [
            "Black & Veatch Inc\nSAVM 203771954  ·  Critical Infra",
            "$19.4M",
            "Intersight compute live.\n$1.14M Secure Access opportunity.",
            "Critical Infrastructure SASE Readiness — secure remote engineering transit and routing mesh.",
            "Play B",
        ],
        [
            "WellSky Corporation\nSAVM 203839175  ·  Healthcare SaaS",
            "$16.2M",
            "21 adoption use cases active\n(Meraki, Compute, Zero Trust).",
            "Multi-Region SaaS Edge Resilience — cloud edge throughput, DC firewall resilience, SD-WAN links.",
            "Play A",
        ],
        [
            "Salina Regional Health\nSAVM 203793443  ·  Regional Health",
            "$20.6M",
            "Campus switching and compute\ntelemetry connected.",
            "Regional Clinic WAN Survivability — clinic router failover, telemedicine QoS, LDOS replacement.",
            "Play A",
        ],
    ]
    y = table(c, headers, rows, y, col_w, row_h=40, wrap_cols=[0, 2, 3], sizes=[8, 8.5, 7.5, 7.5, 8])

    draw_text(c, "HORIZON 2 ENTRY CRITERIA", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 14
    bullets = [
        "At least three first-wave briefings complete and two SOWs staged in CCW.",
        "Named AM owner and VP Infrastructure identified for the Tier 2 account.",
        "IQ snapshot actually populated — do not brief from a blank Catalyst Center tenant.",
        "A funded trigger exists (Webex refresh, Secure Access, LDOS, or Zero Trust) so RIS is insurance on a live program.",
    ]
    for b in bullets:
        c.setFillColor(GRASS)
        c.circle(LEFT + 6, y + 3, 2.4, fill=1, stroke=0)
        y = draw_para(c, b, LEFT + 16, y, CONTENT_W - 16, "Sans", 9.5, INK, 13) - 6
    footer(c, 7)


def page_plays(c):
    header(c, "Sales plays")
    y = h1(
        c,
        "Two plays. Pick one per account. Do not mix in the room.",
        PAGE_H - 68,
        "Play A sells the fact that telemetry is already on. Play B sells the cost of being wrong on day one of a migration already funded.",
    )

    # two columns
    col_w = (CONTENT_W - 12) / 2
    card_h = 278
    # Play A
    round_rect(c, LEFT, y - card_h, col_w, card_h, r=8, fill=white, stroke=GRASS, sw=1.4)
    draw_text(c, "SALES PLAY A", LEFT + 14, y - 20, "Sans-Bold", 8, GRASS)
    draw_text(c, "Telemetry Unlock", LEFT + 14, y - 38, "Sans-Bold", 13, INDIGO)
    draw_para(c, "Garmin  ·  DFA  ·  UMB  ·  Children's Mercy  ·  WellSky  ·  Salina", LEFT + 14, y - 56, col_w - 28, "Sans", 8, MUTED, 11)
    draw_para(
        c,
        "Position the RIS Professional Assessment as the consulting deliverable that turns raw IQ telemetry into a prioritized C-level roadmap. Do not sell a data-collection exercise — the data is already flowing.",
        LEFT + 14,
        y - 84,
        col_w - 28,
        "Sans",
        9,
        INK,
        12,
    )
    draw_text(c, "TALK TRACKS", LEFT + 14, y - 150, "Sans-Bold", 7.5, CYAN)
    tracks_a = [
        "IQ integration is complete — we start from live telemetry, not questionnaires.",
        "Deliverable is a board-ready findings deck, not a config dump.",
        "Output: routing convergence, latent failover risk, 3-year sequence.",
    ]
    ty = y - 166
    for t in tracks_a:
        ty = draw_para(c, "•  " + t, LEFT + 14, ty, col_w - 28, "Sans", 8.5, MUTED, 11.5) - 4
    draw_para(
        c,
        "Ask next: Which of routing convergence, DC fabric, or WAN failover is the board most exposed to in the next 18 months?",
        LEFT + 14,
        ty - 4,
        col_w - 28,
        "Sans-Bold",
        8,
        INDIGO,
        11,
    )

    # Play B
    x2 = LEFT + col_w + 12
    round_rect(c, x2, y - card_h, col_w, card_h, r=8, fill=white, stroke=CYAN, sw=1.4)
    draw_text(c, "SALES PLAY B", x2 + 14, y - 20, "Sans-Bold", 8, CYAN)
    draw_text(c, "Pre-Transformation Derisking", x2 + 14, y - 38, "Sans-Bold", 13, INDIGO)
    draw_para(c, "Lockton  ·  Truman  ·  HNTB  ·  Black & Veatch", x2 + 14, y - 56, col_w - 28, "Sans", 8, MUTED, 11)
    draw_para(
        c,
        "Position RIS as mandatory risk mitigation before cloud calling, SD-WAN, or data center migrations. The assessment is how the CIO buys day-one SLA confidence.",
        x2 + 14,
        y - 84,
        col_w - 28,
        "Sans",
        9,
        INK,
        12,
    )
    draw_text(c, "ATTACH IT IN FRONT OF", x2 + 14, y - 150, "Sans-Bold", 7.5, WHEAT)
    tracks_b = [
        "Lockton — $14.8M Webex Calling renewal and global WAN consolidation.",
        "Truman — $1.35M WxC DI 5-year renewal and campus refresh.",
        "HNTB — $1.2M Webex refresh in Commit plus studio WAN load.",
    ]
    ty = y - 166
    for t in tracks_b:
        ty = draw_para(c, "•  " + t, x2 + 14, ty, col_w - 28, "Sans", 8.5, MUTED, 11.5) - 4
    draw_para(
        c,
        "Ask next: If calling, CAD sync, or EHR traffic failed over incorrectly on day one, which executive owns that outage?",
        x2 + 14,
        ty - 4,
        col_w - 28,
        "Sans-Bold",
        8,
        INDIGO,
        11,
    )

    y = y - card_h - 18
    round_rect(c, LEFT, y - 70, CONTENT_W, 70, r=8, fill=BAND, stroke=None)
    draw_text(c, "RULE FOR THE ROOM", LEFT + 14, y - 16, "Sans-Bold", 8, INDIGO)
    draw_para(
        c,
        "Open with the play, not the SKU. Do not walk into a CIO briefing with a CCW quote. Leave with a named sponsor for a $75k–$150k RIS Assessment SOW and a date for CX to return with the statement of work. Pull-through ($25M+) is the finding of the assessment, not the ask of the first meeting.",
        LEFT + 14,
        y - 34,
        CONTENT_W - 28,
        "Sans",
        9,
        INK,
        12,
    )
    footer(c, 8)


def page_operating_model(c):
    header(c, "Operating model")
    y = h1(
        c,
        "Who does what — and the weekly drumbeat",
        PAGE_H - 68,
        "Four AMs own the relationships. CX Advisory owns the briefing and the SOW. Buying Programs owns staging against FY27. One 30-minute standup keeps it honest.",
    )

    draw_text(c, "RACI", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 8
    headers = ["Workstream", "AM (pool of 4)", "CX Advisory", "Buying Programs", "AE / Overlay"]
    col_w = [120, 105, 105, 105, 105]
    rows = [
        ["IQ snapshot quality", "A / C", "R", "I", "C"],
        ["CIO / VP Infra access", "R / A", "C", "I", "C"],
        ["Play A vs Play B call", "A", "R", "I", "C"],
        ["30-min CX briefing", "C", "R / A", "I", "C"],
        ["SOW scope $75–150k", "C", "R / A", "C", "C"],
        ["CCW stage + FY27 attach", "C", "C", "R / A", "R"],
        ["Pull-through hypothesis", "A", "R", "C", "C"],
        ["Weekly standup", "C", "R", "A", "C"],
    ]
    y = table(c, headers, rows, y, col_w, row_h=18, sizes=[8, 8, 8, 8, 8], aligns=["left", "center", "center", "center", "center"])
    draw_text(c, "R = Responsible   A = Accountable   C = Consulted   I = Informed     AM pool: Randolph · Garwood · Pavalone · Ball", LEFT, y, "Sans", 7.5, MUTED)
    y -= 18

    draw_text(c, "WEEKLY TALLGRASS RIS STANDUP  ·  30 MINUTES  ·  MONDAYS", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 12
    agenda = [
        ("0–5 min", "Scoreboard", "Briefings booked / held, SOWs staged, stalled accounts, IQ snapshot gaps."),
        ("5–20 min", "First-wave accounts", "One account at a time: trigger, play, next meeting, blocker, owner."),
        ("20–25 min", "SOW / CCW", "Which statements of work are ready to stage against FY27 buying programs."),
        ("25–30 min", "Asks", "CX capacity, AM introductions, snapshot production, executive time."),
    ]
    for when, title, body in agenda:
        round_rect(c, LEFT, y - 36, CONTENT_W, 34, r=5, fill=BAND, stroke=None)
        draw_text(c, when, LEFT + 10, y - 22, "Sans-Bold", 8, GRASS)
        draw_text(c, title, LEFT + 78, y - 22, "Sans-Bold", 9, INDIGO)
        draw_para(c, body, LEFT + 200, y - 16, 330, "Sans", 8, MUTED, 10)
        y -= 40

    y -= 14
    draw_text(c, "PHASE GATES  ·  DO NOT SKIP", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 12
    gates = [
        ("Gate 1", "Weeks 1–2", "Named AM, named VP Infra, Play A/B locked, IQ snapshot in hand."),
        ("Gate 2", "Weeks 3–4", "Briefing held. Verbal sponsor. SOW envelope agreed ($75k–$150k)."),
        ("Gate 3", "Weeks 5–6", "SOW staged in CCW and attached to an FY27 service buying program."),
    ]
    gw = (CONTENT_W - 16) / 3
    for i, (g, when, body) in enumerate(gates):
        x = LEFT + i * (gw + 8)
        round_rect(c, x, y - 78, gw, 78, r=7, fill=white, stroke=INDIGO, sw=1.1)
        draw_text(c, g.upper(), x + 10, y - 18, "Sans-Bold", 8, GRASS)
        draw_text(c, when, x + 10, y - 34, "Sans-Bold", 10, INDIGO)
        draw_para(c, body, x + 10, y - 50, gw - 20, "Sans", 8, MUTED, 11)
    footer(c, 9)


def page_calendar(c):
    header(c, "Six-week calendar")
    y = h1(
        c,
        "Execution calendar — weeks 1 through 6",
        PAGE_H - 68,
        "This is the source-plan calendar, unpacked into owners, exit criteria, and what 'done' means. Horizon 2 (Tier 2) does not start until Gate 3 is green on at least two first-wave accounts.",
    )

    weeks = [
        (
            "WEEK 1",
            "Stand up the motion",
            [
                "Kickoff with Randolph, Garwood, Pavalone, Ball — 45 minutes, this document as the agenda.",
                "Assign an AM owner to each of the six Tier 1 names. First-wave four cannot be unowned.",
                "Confirm IQ telemetry is actually live (Cat Center / Intersight / Splunk / ISE) per account.",
                "Lock Play A or Play B. No account enters week 2 without a play.",
            ],
            "Exit: owner + play + IQ source listed on the tracker (page 11).",
        ),
        (
            "WEEK 2",
            "Snapshots and calendar",
            [
                "CX Advisory produces a one-page IQ snapshot per first-wave account (no blank-page discovery).",
                "AM identifies CIO / VP of Infrastructure and sends the briefing hold.",
                "Internal dry-run of Play A and Play B scripts — 20 minutes, critique the open.",
                "Confirm FY27 buying-program attach path with Buying Programs so CCW staging is not a surprise in week 5.",
            ],
            "Exit: four briefing holds on the calendar; four snapshots in the working folder.",
        ),
        (
            "WEEKS 3–4",
            "Executive discovery",
            [
                "Run 30-minute CX Advisory briefings. AM opens, CX delivers, AM closes on sponsor + date.",
                "Sequence: UMB, Garmin, DFA, Lockton. Slip Children's Mercy / Truman only if Gate 1 is green.",
                "Capture: verbal sponsor name, preferred SOW envelope, any security / change-window constraints.",
                "Same week: start SOW outline so week 5 is editing, not inventing.",
            ],
            "Exit: Gate 2 green — briefing held, sponsor named, envelope agreed.",
        ),
        (
            "WEEKS 5–6",
            "Scope and stage",
            [
                "CX finalizes the RIS Professional Assessment SOW ($75k–$150k) per first-wave account.",
                "AE / Buying Programs stages the SOW in CCW and attaches it to the FY27 service buying program.",
                "Pull-through ($25M+ routing / switching / cloud security) is written as the next-best-action, not line-itemed.",
                "Standup reviews: which SOWs are staged vs. stuck on legal, Smart Account, or executive signature.",
            ],
            "Exit: Gate 3 green — SOW in CCW on first-wave names.",
        ),
    ]
    for label, title, bullets, exit_c in weeks:
        round_rect(c, LEFT, y - 118, CONTENT_W, 114, r=7, fill=white, stroke=RULE, sw=0.7)
        c.setFillColor(INDIGO)
        c.roundRect(LEFT + 8, y - 36, 62, 18, 3, fill=1, stroke=0)
        draw_text(c, label, LEFT + 39, y - 31, "Sans-Bold", 7, white, "center")
        draw_text(c, title, LEFT + 78, y - 30, "Sans-Bold", 11, INDIGO)
        by = y - 48
        for b in bullets:
            draw_text(c, "•  " + b, LEFT + 12, by, "Sans", 8, INK)
            by -= 11.5
        draw_text(c, exit_c, LEFT + 12, y - 106, "Sans-Bold", 8, GRASS)
        y -= 124
    footer(c, 10)


def page_tracker(c):
    header(c, "Execution tracker")
    y = h1(
        c,
        "Account execution tracker — fill this in week 1",
        PAGE_H - 68,
        "AM owner is assigned at the kickoff. Status stays red until Gate 1 is green. This page is the standup artifact — print it or keep it live, but review it every Monday.",
    )
    headers = ["Account", "AM owner", "Play", "IQ snapshot", "VP Infra", "Briefing", "SOW $", "CCW", "Next action"]
    col_w = [78, 58, 40, 58, 58, 52, 48, 42, 106]
    rows = [
        ["UMB Financial", "Assign W1", "A", "Due W2", "Name W1", "W3–4", "75–150k", "W5–6", "Confirm Splunk FSO snapshot"],
        ["Garmin", "Assign W1", "A", "Due W2", "Name W1", "W3–4", "75–150k", "W5–6", "Anchor to $2.5M DC build"],
        ["DFA", "Assign W1", "A", "Due W2", "Name W1", "W3–4", "75–150k", "W5–6", "Attach ahead of Zscaler takeout"],
        ["Lockton", "Assign W1", "B", "Due W2", "Name W1", "W3–4", "75–150k", "W5–6", "Attach ahead of $14.8M WxC"],
        ["Children's Mercy", "Assign W1", "A", "If live", "Name W2", "W4–5", "75–150k", "W6", "LDOS + device isolation story"],
        ["Truman / UHKC", "Assign W1", "B", "If live", "Name W2", "W4–5", "75–150k", "W6", "WxC DI + EHR path diversity"],
        ["HNTB (T2)", "Hold", "B", "Hold", "Hold", "Horizon 2", "75–150k", "H2", "After two first-wave SOWs staged"],
        ["Black & Veatch (T2)", "Hold", "B", "Hold", "Hold", "Horizon 2", "75–150k", "H2", "Pair with Secure Access $1.14M"],
        ["WellSky (T2)", "Hold", "A", "Hold", "Hold", "Horizon 2", "75–150k", "H2", "21 use cases — Play A density"],
        ["Salina Regional (T2)", "Hold", "A", "Hold", "Hold", "Horizon 2", "75–150k", "H2", "Clinic WAN + LDOS roadmap"],
    ]
    y = table(
        c,
        headers,
        rows,
        y,
        col_w,
        row_h=22,
        sizes=[7, 6.5, 7.5, 6.5, 6.5, 6.5, 6.5, 6.5, 6.5],
        wrap_cols=[8],
    )

    y -= 4
    draw_text(c, "STATUS KEY FOR STANDUP", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 16
    keys = [
        (GRASS, "Green", "Gate met this week — snapshot in, briefing held, or SOW staged."),
        (WHEAT, "Amber", "Owned, but a dependency is slipping (calendar, snapshot, sponsor)."),
        (HexColor("#C41230"), "Red", "Unowned, no IQ, or first-wave account still missing Gate 1 after week 2."),
    ]
    kw = (CONTENT_W - 16) / 3
    for i, (col, lab, body) in enumerate(keys):
        x = LEFT + i * (kw + 8)
        round_rect(c, x, y - 52, kw, 52, r=6, fill=white, stroke=RULE, sw=0.6)
        c.setFillColor(col)
        c.circle(x + 12, y - 18, 5, fill=1, stroke=0)
        draw_text(c, lab, x + 24, y - 22, "Sans-Bold", 9, INDIGO)
        draw_para(c, body, x + 10, y - 36, kw - 20, "Sans", 7.5, MUTED, 10)
    footer(c, 11)


def page_close(c):
    header(c, "Briefing kit  ·  staging  ·  next 10 days")
    y = h1(
        c,
        "What goes in the room — and what happens after",
        PAGE_H - 68,
        "A CIO briefing without an IQ snapshot is a discovery call. We are not selling discovery. We are selling a reading of data they already send us.",
    )

    # two columns top
    col_w = (CONTENT_W - 12) / 2
    round_rect(c, LEFT, y - 168, col_w, 168, r=8, fill=white, stroke=RULE, sw=0.7)
    draw_text(c, "BRIEFING KIT  ·  30 MINUTES", LEFT + 12, y - 18, "Sans-Bold", 9, INDIGO)
    kit = [
        "One-page IQ snapshot (health, LDOS, fabric, WAN) — CX produces, AM reviews.",
        "Named architectural trigger (the program already funded).",
        "Play A or Play B script — one paragraph, memorized, not read.",
        "Assessment envelope: $75k–$150k, 4–6 week CX delivery, board-ready findings deck.",
        "Close: sponsor name, SOW review date, FY27 buying-program path.",
        "Leave-behind: this plan's Play A/B quote and the three-year roadmap promise — not a quote.",
    ]
    ky = y - 36
    for item in kit:
        ky = draw_para(c, "•  " + item, LEFT + 12, ky, col_w - 24, "Sans", 8, INK, 11) - 3

    x2 = LEFT + col_w + 12
    round_rect(c, x2, y - 168, col_w, 168, r=8, fill=white, stroke=RULE, sw=0.7)
    draw_text(c, "CCW STAGING CHECKLIST", x2 + 12, y - 18, "Sans-Bold", 9, INDIGO)
    sow = [
        "SOW title: RIS Professional Assessment — <account>.",
        "Value: $75k–$150k. Do not bundle pull-through SKUs.",
        "Attach to FY27 service buying program in CCW.",
        "Smart Account and SAVM (see roster) confirmed before submit.",
        "CX delivery window aligned to customer's change freeze.",
        "Next-best-action note: routing / switching / cloud security hypothesis for post-findings.",
    ]
    ky = y - 36
    for item in sow:
        ky = draw_para(c, "•  " + item, x2 + 12, ky, col_w - 24, "Sans", 8, INK, 11) - 3

    y -= 186
    draw_text(c, "NEXT 10 WORKING DAYS", LEFT, y, "Sans-Bold", 9, INDIGO)
    y -= 12
    days = [
        ("This week", "Sit with Randolph, Garwood, Pavalone, and Ball. Assign AM owners. Lock Play A/B on UMB, Garmin, DFA, Lockton. Confirm IQ is live."),
        ("Week 2", "CX produces four snapshots. AM books four 30-minute CIO / VP Infra briefings. Dry-run both scripts."),
        ("Weeks 3–6", "Brief, sponsor, scope, stage. Do not open Tier 2 until two first-wave SOWs are in CCW."),
    ]
    for when, body in days:
        round_rect(c, LEFT, y - 36, CONTENT_W, 34, r=5, fill=BAND, stroke=None)
        draw_text(c, when, LEFT + 12, y - 21, "Sans-Bold", 8.5, GRASS)
        draw_para(c, body, LEFT + 88, y - 16, CONTENT_W - 100, "Sans", 8, INK, 10.5)
        y -= 40

    y -= 4
    # commercial thesis strip
    c.setFillColor(INDIGO)
    c.roundRect(LEFT, y - 92, CONTENT_W, 92, 8, fill=1, stroke=0)
    c.setFillColor(GRASS)
    c.rect(LEFT, y - 92, 7, 92, fill=1, stroke=0)
    draw_text(c, "COMMERCIAL THESIS  ·  DO NOT STOP AT THE ASSESSMENT", LEFT + 20, y - 20, "Sans-Bold", 8, GRASS)
    draw_text(c, "$75k–$150k   →   $1.2M–$1.8M   →   $25M+", LEFT + 20, y - 42, "Sans-Bold", 13, white)
    draw_para(
        c,
        "Stage the assessment. Cover the core ten. Fund the roadmap. Anchor every briefing to a trigger already in motion — Garmin DC, DFA Zscaler, Lockton WxC, Children's Mercy LDOS — so RIS is risk insurance on a program the customer has already funded.",
        LEFT + 20,
        y - 62,
        CONTENT_W - 40,
        "Sans",
        8.5,
        HexColor("#D5E6EE"),
        12,
    )
    footer(c, 12)


def build():
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    c = canvas.Canvas(str(OUT_PATH), pagesize=LETTER)
    c.setTitle("Cisco Tallgrass Region — RIS Assessment Execution Plan")
    c.setAuthor("Cisco Services & Software Buying Programs")
    c.setSubject("Internal sales enablement — IQ onboarded RIS Assessment pipeline")
    c.setKeywords("Cisco,Tallgrass,RIS,IQ,telemetry,execution plan")

    pages = [
        page_cover,
        page_objective,
        page_framework,
        page_tier1_roster,
        page_wave_a,
        page_wave_b,
        page_tier2,
        page_plays,
        page_operating_model,
        page_calendar,
        page_tracker,
        page_close,
    ]
    assert len(pages) == TOTAL_PAGES
    for fn in pages:
        fn(c)
        c.showPage()
    c.save()
    print(f"Wrote {OUT_PATH} ({TOTAL_PAGES} pages)")


if __name__ == "__main__":
    build()
