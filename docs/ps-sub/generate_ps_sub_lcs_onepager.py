#!/usr/bin/env python3
"""Generate a one-page landscape PDF: Cisco PS-Sub / LCS abilities.

Aligned to the FY26 Cisco Services Offering (Support vs PS-Sub vs PS-T),
the Professional Services lifecycle, and the PS / LCS service descriptions.
"""

from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Cisco_PS-Sub_LCS_Abilities_OnePager.pdf"

FONT_DIR = Path("/usr/share/fonts/truetype/macos")
pdfmetrics.registerFont(TTFont("Inter", str(FONT_DIR / "Inter-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Inter-Med", str(FONT_DIR / "Inter-Medium.ttf")))
pdfmetrics.registerFont(TTFont("Inter-Semi", str(FONT_DIR / "Inter-SemiBold.ttf")))
pdfmetrics.registerFont(TTFont("Inter-Bold", str(FONT_DIR / "Inter-Bold.ttf")))

NAVY = HexColor("#0D274D")
DEEP = HexColor("#081A33")
CYAN = HexColor("#00BCEB")
INK = HexColor("#1B2A3A")
MUTED = HexColor("#5C6B7A")
OFF = HexColor("#F4F7F9")
LINE = HexColor("#D5DEE7")
GOLD = HexColor("#C6A552")
GROW = HexColor("#0E6B5C")
STRONG = HexColor("#8A5A12")
WARN = HexColor("#8B2E2E")


def rrect(c, x, y, w, h, r=6, fill=None, stroke=None, sw=0.7):
    c.saveState()
    if fill:
        c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke)
        c.setLineWidth(sw)
    p = c.beginPath()
    p.roundRect(x, y, w, h, r)
    if fill and stroke:
        c.drawPath(p, fill=1, stroke=1)
    elif fill:
        c.drawPath(p, fill=1, stroke=0)
    else:
        c.drawPath(p, fill=0, stroke=1)
    c.restoreState()


def wrap(c, text, font, size, max_w):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def draw_wrapped(c, text, x, y, font, size, max_w, color=INK, leading=None, max_lines=None):
    leading = leading or size + 2.2
    lines = wrap(c, text, font, size, max_w)
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
        last = lines[-1]
        while c.stringWidth(last + "…", font, size) > max_w and last:
            last = last[:-1]
        lines[-1] = last.rstrip(" ,.;:") + "…"
    c.setFillColor(color)
    c.setFont(font, size)
    for i, line in enumerate(lines):
        c.drawString(x, y - i * leading, line)
    return len(lines) * leading


def bullet_block(c, items, x, y, max_w, font="Inter", size=6.5, color=INK, leading=8.4, bullet=CYAN):
    top = y
    for item in items:
        c.setFillColor(bullet)
        c.circle(x + 2.0, y + 1.5, 1.25, fill=1, stroke=0)
        h = draw_wrapped(c, item, x + 8.2, y, font, size, max_w - 8.2, color=color, leading=leading)
        y -= h
    return top - y


def kicker(c, text, x, y, color=CYAN):
    c.setFillColor(color)
    c.setFont("Inter-Semi", 6.2)
    c.drawString(x, y, text.upper())


def chip(c, text, x, y, h=13, fill=DEEP, fg=white, font="Inter-Med", size=6.0, pad=6.5):
    w = c.stringWidth(text, font, size) + pad * 2
    rrect(c, x, y, w, h, r=3.2, fill=fill)
    c.setFillColor(fg)
    c.setFont(font, size)
    c.drawString(x + pad, y + 3.5, text)
    return w


def draw():
    W, H = landscape(letter)
    c = canvas.Canvas(str(OUT), pagesize=(W, H))
    c.setTitle("Cisco PS-Sub / LCS Abilities at a Glance")
    c.setAuthor("Cisco Customer Experience")
    c.setSubject("Professional Services Subscription abilities from the FY26 Services Offering")

    header_h = 44
    footer_h = 70
    m = 13
    gap = 7

    # Header
    c.setFillColor(NAVY)
    c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.rect(0, H - header_h, 6, header_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.setFont("Inter-Bold", 10.5)
    c.drawString(16, H - 18, "CISCO")
    c.setFillColor(white)
    c.setFont("Inter-Bold", 14)
    c.drawString(58, H - 19, "Professional Services Subscription")
    c.setFillColor(CYAN)
    c.setFont("Inter-Med", 8.2)
    c.drawString(58, H - 34, "PS-Sub  ·  formerly Lifecycle Services (LCS)  ·  FY26 Services Offering")
    c.setFillColor(white)
    c.setFont("Inter-Semi", 10.5)
    c.drawRightString(W - 16, H - 18, "Abilities at a Glance")
    c.setFillColor(HexColor("#9BB4C9"))
    c.setFont("Inter", 6.8)
    c.drawRightString(W - 16, H - 34, "Seller conversation sheet  ·  October 2026")

    # Where it sits in the offering
    sit_h = 72
    sit_y = H - header_h - 7 - sit_h
    offers = [
        (
            False,
            "Cisco Support",
            "Keep it running",
            "Basic / Standard / Signature, powered by Cisco IQ. Lead Standard+. TAC, RMA, lifecycle insights, RIS Playbook. Not Professional Services.",
        ),
        (
            True,
            "PS-Subscription  ·  this sheet",
            "Govern, advise, measure",
            "Ongoing outcomes, Digital Insights, named experts, 90-day plans, QBRs. Flexible consumption — not T&M.",
        ),
        (
            False,
            "PS-Transactional  ·  SOW",
            "Do the hands-on work",
            "Defined-scope PDI / migrate / harden / LDoS. Preferred vehicle for high-frequency remediation. Do not fold into PS-Sub.",
        ),
    ]
    ow = (W - 2 * m - 2 * gap) / 3
    for i, (hot, title, sub, body) in enumerate(offers):
        x = m + i * (ow + gap)
        if hot:
            rrect(c, x, sit_y, ow, sit_h, r=6, fill=NAVY)
            c.setFillColor(CYAN)
            c.setFont("Inter-Semi", 6.0)
            c.drawString(x + 10, sit_y + sit_h - 12, "SERVICES OFFERING")
            c.setFillColor(white)
            c.setFont("Inter-Bold", 8.6)
            c.drawString(x + 10, sit_y + sit_h - 24, title)
            c.setFillColor(CYAN)
            c.setFont("Inter-Med", 6.4)
            c.drawString(x + 10, sit_y + sit_h - 35, sub)
            draw_wrapped(c, body, x + 10, sit_y + sit_h - 47, "Inter", 6.2, ow - 20, color=white, leading=7.8)
        else:
            rrect(c, x, sit_y, ow, sit_h, r=6, fill=OFF, stroke=LINE)
            c.setFillColor(MUTED)
            c.setFont("Inter-Semi", 6.0)
            c.drawString(x + 10, sit_y + sit_h - 12, "SERVICES OFFERING")
            c.setFillColor(NAVY)
            c.setFont("Inter-Bold", 8.6)
            c.drawString(x + 10, sit_y + sit_h - 24, title)
            c.setFillColor(MUTED)
            c.setFont("Inter-Med", 6.4)
            c.drawString(x + 10, sit_y + sit_h - 35, sub)
            draw_wrapped(c, body, x + 10, sit_y + sit_h - 47, "Inter", 6.2, ow - 20, color=INK, leading=7.8)

    # Lifecycle abilities — the offering spine
    life_h = 88
    life_y = sit_y - gap - life_h
    phases = [
        ("Plan", ["Requirements & readiness", "Architecture strategy", "KPI / value baseline"]),
        ("Design", ["Design review vs requirements", "Validated-design guidance", "Standards alignment (SA)"]),
        ("Implement", ["Implementation planning", "Test plan / execute support", "Services as Code / IaC"]),
        ("Migrate", ["Migration planning", "Cutover / NRFU guidance", "Post-migration support"]),
        ("Operate", ["Day-2 assurance", "Change MOP review", "AMO / program governance"]),
        ("Optimize", ["Health / resiliency assessments", "PSIRT & software strategy", "KPI snapshots / QBR"]),
    ]
    pw = (W - 2 * m - 5 * gap) / 6
    rrect(c, m, life_y, W - 2 * m, life_h, r=7, fill=OFF)
    c.setFillColor(CYAN)
    c.setFont("Inter-Semi", 6.0)
    c.drawString(m + 10, life_y + life_h - 12, "LIFECYCLE ABILITIES  ·  PLAN → OPTIMIZE  ·  SUBSCRIPTION CAN ADVISE, ASSIST, OR LEAD IN EACH PHASE")
    inner_h = life_h - 20
    inner_y = life_y + 8
    inner_w = (W - 2 * m - 20 - 5 * 6) / 6
    for i, (title, items) in enumerate(phases):
        x = m + 10 + i * (inner_w + 6)
        rrect(c, x, inner_y, inner_w, inner_h, r=5, fill=white)
        c.setFillColor(CYAN)
        c.circle(x + 9, inner_y + inner_h - 11, 6.2, fill=1, stroke=0)
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 6.4)
        c.drawCentredString(x + 9, inner_y + inner_h - 13.2, str(i + 1))
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 7.4)
        c.drawString(x + 18, inner_y + inner_h - 14, title)
        bullet_block(c, items, x + 7, inner_y + inner_h - 28, inner_w - 14, size=6.05, leading=9.0)

    # Main 3 columns
    col_y = footer_h + 7
    col_h = life_y - gap - col_y
    cw = (W - 2 * m - 2 * gap) / 3

    # Col 1 — how the subscription is built
    x = m
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "How PS-Sub is packaged", x + 10, col_y + col_h - 13)
    packs = [
        (
            "01  Outcome Services",
            "LCS engine. Outcome Consulting + Digital Insights + PM. Priorities, baseline, analyze, recommend, action plan, implementation assist, KPI snapshots. 90-day plan and quarterly activity / artifact report.",
        ),
        (
            "02  Expert-as-a-Service",
            "Named expert on a quoted technology. Consulting Engineer: design, implement, operate, KT. Solution Architect: vision, strategy, standards. Includes Cisco tools, analytics, and knowledge systems.",
        ),
        (
            "03  Attached Services",
            "Product / family-specific consulting for the Services Term: onboarding, implementation, adoption, optimization. Not a one-shot PDI SOW.",
        ),
    ]
    yy = col_y + col_h - 28
    for title, body in packs:
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 7.5)
        c.drawString(x + 10, yy, title)
        yy -= 11
        yy -= draw_wrapped(c, body, x + 10, yy, "Inter", 6.35, cw - 22, color=INK, leading=8.1) + 7

    kicker(c, "Five-step cadence  ·  continuous", x + 10, yy)
    yy -= 14
    steps = ["1 Baseline", "2 Analyze", "3 Recommend", "4 Execute", "5 Measure"]
    sx = x + 10
    for s in steps:
        sx += chip(c, s, sx, yy - 2, h=13, fill=NAVY, fg=white, size=5.7, pad=5) + 4
    yy -= 22
    rrect(c, x + 8, yy - 40, cw - 16, 48, r=5, fill=OFF)
    c.setFillColor(NAVY)
    c.setFont("Inter-Semi", 6.3)
    c.drawString(x + 16, yy - 8, "Engage the way the customer works")
    modes = [("Advise Me", "Guided"), ("Do It With Me", "Assisted"), ("Do It For Me", "Assured")]
    for i, (t, s) in enumerate(modes):
        mx = x + 16 + i * 78
        rrect(c, mx, yy - 36, 74, 24, r=4, fill=NAVY)
        c.setFillColor(white)
        c.setFont("Inter-Semi", 6.1)
        c.drawCentredString(mx + 37, yy - 22, t)
        c.setFillColor(CYAN)
        c.setFont("Inter", 5.5)
        c.drawCentredString(mx + 37, yy - 32, s)

    # Col 2 — Digital Insights + ability catalogue
    x = m + cw + gap
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "Digital Insights  ·  telemetry + expert judgment", x + 10, col_y + col_h - 13)
    insights = [
        "Config complexity, misconfig, inconsistency",
        "Customer config-standard audit and reporting",
        "Policy / design mapping to simplify support",
        "Field Notices and hardware lifecycle (EoS, LDoS, EoL)",
        "PSIRT Critical / High and crash-risk guidance",
        "Scalability vs verified limits; software version strategy",
        "Syslog service-impact risk; air-gapped collector option",
    ]
    used = bullet_block(
        c, insights, x + 10, col_y + col_h - 27, cw - 22, size=6.35, leading=8.55
    )
    cat_top = col_y + col_h - 27 - used - 6
    kicker(c, "Offering catalogue examples  ·  resize to available capacity", x + 10, cat_top)
    cats = [
        "Assessments — resiliency, architecture, health (EN / Sec / DC / Collab)",
        "Software management — PSIRT impact, bug scrub, version recommendations",
        "Architecture Management Office — weekly / biweekly governance",
        "Change & test — MOP review, test plan, validate-with-test",
        "Implementation / migration support — assist, not a full PDI SOW",
        "Knowledge transfer — workshops, runbooks, enablement",
        "Automation — Services as Code, digital twins, DevOps pipelines",
    ]
    bullet_block(c, cats, x + 10, cat_top - 13, cw - 22, size=6.25, leading=8.4, bullet=HexColor("#049FD9"))

    # Col 3 — outcomes, cadence, guardrail
    x = m + 2 * (cw + gap)
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "Business outcomes on the quote", x + 10, col_y + col_h - 13)
    groups = [
        (GROW, "Grow faster", "Digital transformation · new capabilities · CX · end-user experience · revenue"),
        (STRONG, "Be stronger", "Reduce risk · compliance · enhance security · reputation / brand"),
        (HexColor("#005073"), "Deliver better", "Optimize productivity · reduce cost"),
    ]
    yy = col_y + col_h - 28
    for color, title, body in groups:
        c.setFillColor(color)
        c.roundRect(x + 10, yy - 1.4, 6, 9, 1.4, fill=1, stroke=0)
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 7.4)
        c.drawString(x + 20, yy, title)
        yy -= 11
        yy -= draw_wrapped(c, body, x + 10, yy, "Inter", 6.25, cw - 22, color=INK, leading=8.0) + 6

    kicker(c, "90-day operating cadence", x + 10, yy)
    yy -= 12
    yy -= draw_wrapped(
        c,
        "Kickoff agrees activities and request format. Then 90-day plan → deliver → quarterly report / QBR → re-plan. New work is resized to remaining resources. Baseline → Analyze → Recommend → Execute → Measure, continuously.",
        x + 10,
        yy,
        "Inter",
        6.25,
        cw - 22,
        color=INK,
        leading=8.1,
    )
    box_h = 92
    by = max(col_y + 8, yy - box_h - 6)
    rrect(c, x + 8, by, cw - 16, box_h, r=5, fill=HexColor("#F7F1E4"), stroke=GOLD, sw=0.85)
    kicker(c, "Keep the motion clean", x + 16, by + box_h - 12, color=STRONG)
    c.setFillColor(GROW)
    c.setFont("Inter-Semi", 6.3)
    c.drawString(x + 16, by + box_h - 24, "Use PS-Sub for")
    h1 = draw_wrapped(
        c,
        "Advisory, governance, architecture, KPI baselining, PSIRT / lifecycle insight, QBRs, named CE / SA, catalogue items that fit remaining capacity.",
        x + 16,
        by + box_h - 35,
        "Inter",
        6.15,
        cw - 34,
        color=INK,
        leading=7.8,
    )
    c.setFillColor(WARN)
    c.setFont("Inter-Semi", 6.3)
    c.drawString(x + 16, by + box_h - 38 - h1, "Put on a PS-T SOW")
    draw_wrapped(
        c,
        "High-frequency hands-on, OS upgrades, hardening, LDoS refresh, transformation, incident response, third-party gear. Do not absorb into an existing subscription.",
        x + 16,
        by + box_h - 49 - h1,
        "Inter",
        6.15,
        cw - 34,
        color=INK,
        leading=7.8,
    )

    # Footer
    c.setFillColor(NAVY)
    c.rect(0, 0, W, footer_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.rect(0, 0, 6, footer_h, fill=1, stroke=0)
    kicker(c, "Boundaries  ·  quote and service description control scope", 16, footer_h - 12)
    c.setFillColor(HexColor("#8AA0B5"))
    c.setFont("Inter", 5.3)
    c.drawRightString(W - 16, footer_h - 12, "Cisco Confidential — Cisco and partner use")
    draw_wrapped(
        c,
        "Supplements Cisco Support or Partner Support on covered products. Not T&M and not an hours bank. Cisco technologies only. “Outcome” does not guarantee a result. Delivery stops the last business day of the term. Digital Insights need data collection (or a quoted air-gapped collector). Scrum / Outcome Scale / EaaS are subscription add-ons. Using PS-Sub for RIS-style remediation requires CX Delivery + SVP path — default is a new PS-T engagement.",
        16,
        footer_h - 25,
        "Inter",
        6.05,
        W - 32,
        color=white,
        leading=7.7,
        max_lines=3,
    )
    arch = ["Networking", "Data Center", "Security", "Collaboration", "SP Mobility", "IoT"]
    ax = 16
    ay = 7
    c.setFillColor(CYAN)
    c.setFont("Inter-Semi", 5.6)
    c.drawString(ax, ay + 3.3, "ARCHITECTURES")
    ax += c.stringWidth("ARCHITECTURES", "Inter-Semi", 5.6) + 7
    for a in arch:
        ax += chip(c, a, ax, ay, h=12.5, fill=DEEP, fg=white, size=5.8) + 4
    c.setFillColor(HexColor("#8AA0B5"))
    c.setFont("Inter", 5.15)
    c.drawRightString(
        W - 16,
        ay + 3.4,
        "PS SD EDCS-25902390 · LCS SD EDCS-24374819 · PS AAG Dec 2025 · FY26 Services Offering",
    )

    c.showPage()
    c.save()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    draw()
