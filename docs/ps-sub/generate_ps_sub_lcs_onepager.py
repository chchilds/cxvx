#!/usr/bin/env python3
"""Generate a one-page landscape PDF: Cisco PS-Sub / LCS abilities."""

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
BETTER = HexColor("#005073")
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


def card_with_bar(c, x, y, w, h, r=7, fill=white, stroke=LINE, bar=CYAN):
    c.saveState()
    p = c.beginPath()
    p.roundRect(x, y, w, h, r)
    c.clipPath(p, stroke=0)
    c.setFillColor(fill)
    c.rect(x, y, w, h, fill=1, stroke=0)
    c.setFillColor(bar)
    c.rect(x, y, 4.5, h, fill=1, stroke=0)
    c.restoreState()
    c.saveState()
    c.setStrokeColor(stroke)
    c.setLineWidth(0.8)
    p = c.beginPath()
    p.roundRect(x, y, w, h, r)
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


def bullet_block(c, items, x, y, max_w, font="Inter", size=7.0, color=INK, leading=9.2, bullet=CYAN):
    top = y
    for item in items:
        c.setFillColor(bullet)
        c.circle(x + 2.1, y + 1.5, 1.3, fill=1, stroke=0)
        h = draw_wrapped(c, item, x + 8.5, y, font, size, max_w - 8.5, color=color, leading=leading)
        y -= h
    return top - y


def kicker(c, text, x, y, color=CYAN):
    c.setFillColor(color)
    c.setFont("Inter-Semi", 6.3)
    c.drawString(x, y, text.upper())


def chip(c, text, x, y, h=13, fill=DEEP, fg=white, font="Inter-Med", size=6.0, pad=7):
    w = c.stringWidth(text, font, size) + pad * 2
    rrect(c, x, y, w, h, r=3.2, fill=fill)
    c.setFillColor(fg)
    c.setFont(font, size)
    c.drawString(x + pad, y + 3.6, text)
    return w


def draw():
    W, H = landscape(letter)
    c = canvas.Canvas(str(OUT), pagesize=(W, H))
    c.setTitle("Cisco PS-Sub / LCS Abilities at a Glance")
    c.setAuthor("Cisco Customer Experience")
    c.setSubject("Professional Services Subscription and Lifecycle Services abilities one-pager")

    header_h = 46
    footer_h = 76
    m = 14
    gap = 8

    # Header
    c.setFillColor(NAVY)
    c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.rect(0, H - header_h, 6, header_h, fill=1, stroke=0)

    c.setFillColor(CYAN)
    c.setFont("Inter-Bold", 11)
    c.drawString(18, H - 20, "CISCO")
    c.setFillColor(white)
    c.setFont("Inter-Bold", 15)
    c.drawString(62, H - 21, "Professional Services Subscription")
    c.setFillColor(CYAN)
    c.setFont("Inter-Med", 8.6)
    c.drawString(62, H - 36, "PS-Sub  ·  formerly Lifecycle Services (LCS)")

    c.setFillColor(white)
    c.setFont("Inter-Semi", 11)
    c.drawRightString(W - 18, H - 20, "Abilities at a Glance")
    c.setFillColor(HexColor("#9BB4C9"))
    c.setFont("Inter", 7.0)
    c.drawRightString(W - 18, H - 36, "Seller & customer conversation sheet  ·  October 2026")

    # Intro
    intro_h = 78
    intro_y = H - header_h - 8 - intro_h
    left_w = 478
    rrect(c, m, intro_y, left_w, intro_h, r=7, fill=OFF)
    kicker(c, "What it is", m + 12, intro_y + intro_h - 14)
    draw_wrapped(
        c,
        "PS-Sub is Cisco’s flexible-consumption professional services subscription. It translates IT initiatives into measurable business outcomes using Cisco experts, telemetry-based Digital Insights, and a co-defined KPI cadence. LCS is the outcomes engine inside Outcome Services — same engagement model, current EA / Services packaging. Transactional SME and SOW work is purchased separately.",
        m + 12,
        intro_y + intro_h - 29,
        "Inter",
        7.5,
        left_w - 24,
        color=INK,
        leading=10.0,
    )

    right_x = m + left_w + gap
    right_w = W - m - right_x
    rrect(c, right_x, intro_y, right_w, intro_h, r=7, fill=HexColor("#F7F1E4"), stroke=GOLD, sw=0.9)
    pad = 10
    yy = intro_y + intro_h - 14
    kicker(c, "Keep the motion clean", right_x + pad, yy, color=STRONG)
    yy -= 14
    c.setFillColor(GROW)
    c.setFont("Inter-Semi", 6.8)
    c.drawString(right_x + pad, yy, "Use PS-Sub / LCS for")
    yy -= 11
    yy -= draw_wrapped(
        c,
        "Advisory, governance, architecture, KPI baselining, PSIRT / lifecycle insight, 90-day plans, QBRs, named CE / SA.",
        right_x + pad,
        yy,
        "Inter",
        6.5,
        right_w - 20,
        color=INK,
        leading=8.3,
    )
    yy -= 4
    c.setFillColor(WARN)
    c.setFont("Inter-Semi", 6.8)
    c.drawString(right_x + pad, yy, "Do not use it as")
    yy -= 11
    draw_wrapped(
        c,
        "A T&M hours bank, high-frequency hands-on, incident response, third-party gear, or a place to fold a PS SOW.",
        right_x + pad,
        yy,
        "Inter",
        6.5,
        right_w - 20,
        color=INK,
        leading=8.3,
    )

    # Three pillars
    pill_h = 102
    pill_y = intro_y - gap - pill_h
    pw = (W - 2 * m - 2 * gap) / 3
    pillars = [
        (
            "01",
            "Outcome Services",
            "LCS core  ·  business-outcome aligned",
            [
                "Outcome Consulting: set priorities, baseline KPIs, analyze, recommend, plan, assist implementation, snapshot vs baseline.",
                "Digital Insights: regular baseline comparisons and recommendations from Cisco telemetry plus agreed new insights.",
                "Project Management: 90-day delivery plan; quarterly activity and artifact report.",
            ],
        ),
        (
            "02",
            "Expert-as-a-Service",
            "Named expert  ·  specified tech / architecture",
            [
                "Consulting Engineer: design, implementation, operations, and knowledge transfer for the quoted technology.",
                "Solution Architect: architecture vision, strategy development, and standards alignment.",
                "Includes Cisco tools, analytics, software, knowledge resources, and systems for the Services Term.",
            ],
        ),
        (
            "03",
            "Attached Services",
            "Product / family-specific consulting",
            [
                "Targeted to a Cisco product or product family quoted on the subscription.",
                "Typical categories: onboarding, implementation, adoption, and optimization.",
                "Consumed across the Services Term — not a one-shot PDI SOW and not T&M.",
            ],
        ),
    ]
    for i, (num, title, sub, bullets) in enumerate(pillars):
        x = m + i * (pw + gap)
        card_with_bar(c, x, pill_y, pw, pill_h)
        c.setFillColor(CYAN)
        c.setFont("Inter-Bold", 8)
        c.drawString(x + 14, pill_y + pill_h - 15, num)
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 10)
        c.drawString(x + 34, pill_y + pill_h - 15, title)
        c.setFillColor(MUTED)
        c.setFont("Inter", 6.5)
        c.drawString(x + 14, pill_y + pill_h - 27, sub)
        bullet_block(c, bullets, x + 14, pill_y + pill_h - 41, pw - 28, size=6.6, leading=8.6)

    # Process bar
    proc_h = 50
    proc_y = pill_y - gap - proc_h
    rrect(c, m, proc_y, W - 2 * m, proc_h, r=7, fill=NAVY)
    kicker(c, "Five-step engagement  ·  continuous", m + 12, proc_y + proc_h - 12, color=CYAN)
    steps = [
        ("1", "Baseline", "KPIs & measurement"),
        ("2", "Analyze", "Risks & opportunities"),
        ("3", "Recommend", "Actions & practices"),
        ("4", "Execute", "Chosen workstyle"),
        ("5", "Measure", "KPI / QBR cadence"),
    ]
    for i, (n, t, s) in enumerate(steps):
        x = m + 10 + i * 104
        c.setFillColor(CYAN)
        c.circle(x + 8, proc_y + 16, 7.4, fill=1, stroke=0)
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 7.6)
        c.drawCentredString(x + 8, proc_y + 13.4, n)
        c.setFillColor(white)
        c.setFont("Inter-Semi", 8.0)
        c.drawString(x + 20, proc_y + 18.4, t)
        c.setFillColor(HexColor("#9BB4C9"))
        c.setFont("Inter", 6.2)
        c.drawString(x + 20, proc_y + 8.2, s)
        if i < 4:
            c.setStrokeColor(CYAN)
            c.setLineWidth(1.15)
            c.line(x + 90, proc_y + 16, x + 100, proc_y + 16)

    mx = 548
    c.setFillColor(CYAN)
    c.setFont("Inter-Semi", 6.2)
    c.drawString(mx, proc_y + proc_h - 12, "ENGAGE THE WAY YOU WORK")
    modes = [("Advise Me", "Cisco advises"), ("Do It With Me", "Side-by-side"), ("Do It For Me", "Cisco / partner execute")]
    for i, (t, s) in enumerate(modes):
        x = mx + i * 76
        rrect(c, x, proc_y + 7, 72, 24, r=4, fill=DEEP)
        c.setFillColor(white)
        c.setFont("Inter-Semi", 6.2)
        c.drawCentredString(x + 36, proc_y + 20, t)
        c.setFillColor(CYAN)
        c.setFont("Inter", 5.5)
        c.drawCentredString(x + 36, proc_y + 10.5, s)

    # Main columns
    col_y = footer_h + 8
    col_h = proc_y - gap - col_y
    cw = (W - 2 * m - 2 * gap) / 3

    # Outcomes
    x = m
    rrect(c, x, col_y, cw, col_h, r=7, fill=OFF)
    kicker(c, "Business outcomes  ·  specified on the quote", x + 12, col_y + col_h - 13)
    groups = [
        (
            GROW,
            "Grow faster",
            [
                "Digital transformation — new models, automation, data, tools",
                "Enable capabilities through new technology — adoption to value",
                "Improve customer experience — fewer client-impacting outages",
                "Improve end-user experience — hybrid / collab productivity",
                "Grow revenue — stable, evolvable infrastructure",
            ],
        ),
        (
            STRONG,
            "Be stronger",
            [
                "Reduce risk — find issues, prioritize remediation",
                "Demonstrate compliance — gap ID and remediation steps",
                "Enhance security — posture, config, policy, controls",
                "Manage reputation / brand — secure digital experience",
            ],
        ),
        (
            BETTER,
            "Deliver better",
            [
                "Optimize productivity — infrastructure and operating practices",
                "Reduce cost — innovation, recommended practices, process",
            ],
        ),
    ]
    gy = col_y + col_h - 28
    for color, title, items in groups:
        c.setFillColor(color)
        c.roundRect(x + 12, gy - 1.5, 6, 9.5, 1.4, fill=1, stroke=0)
        c.setFillColor(NAVY)
        c.setFont("Inter-Bold", 7.8)
        c.drawString(x + 22, gy, title)
        gy -= 11.5
        h = bullet_block(c, items, x + 12, gy, cw - 26, size=6.35, leading=8.2, bullet=color)
        gy -= h + 5

    rrect(c, x + 10, col_y + 8, cw - 20, 40, r=5, fill=white)
    c.setFillColor(NAVY)
    c.setFont("Inter-Semi", 6.3)
    c.drawString(x + 18, col_y + 36, "90-day operating cadence")
    draw_wrapped(
        c,
        "Kickoff agrees activities and request format. Then 90-day plan → deliver → quarterly report / QBR → re-plan. Intended work is reviewed and resized to available resources each quarter.",
        x + 18,
        col_y + 25,
        "Inter",
        6.0,
        cw - 38,
        color=MUTED,
        leading=7.7,
    )

    # Digital Insights
    x = m + cw + gap
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "Digital Insights abilities", x + 12, col_y + col_h - 13)
    draw_wrapped(
        c,
        "Telemetry + AI/ML from 3M+ Cisco devices, combined with expert judgment. Baseline, compare, recommend.",
        x + 12,
        col_y + col_h - 26,
        "Inter",
        6.4,
        cw - 24,
        color=MUTED,
        leading=8.2,
    )
    insights = [
        "Config complexity, misconfig, and inconsistency on like features",
        "Customer config-standard definition, audit, and reporting",
        "End-to-end policy / design mapping to simplify support",
        "Field Notice exposure tracking",
        "Hardware lifecycle: End-of-Sale, Last Date of Support, End-of-Life",
        "Policy variation that adds design complexity",
        "PSIRT advisories rated Critical or High",
        "Device crash-risk identification and guidance",
        "Scalability versus verified configuration limits",
        "Software version fitness, upgrade triggers, adoption strategy",
        "Syslog operational / performance service-impact risk",
        "Air-gapped collector option for highly secure networks",
    ]
    used = bullet_block(c, insights, x + 12, col_y + col_h - 46, cw - 24, size=6.45, leading=9.05)
    note_y = col_y + 10
    rrect(c, x + 10, note_y, cw - 20, 28, r=4, fill=OFF)
    draw_wrapped(
        c,
        "Data required. Use Cisco collectors, a customer data drop, or a quoted air-gapped appliance for classified / government networks.",
        x + 18,
        note_y + 16,
        "Inter",
        6.0,
        cw - 36,
        color=MUTED,
        leading=7.6,
    )

    # Scrum + EaaS
    x = m + 2 * (cw + gap)
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "Scrum Services & Outcome Scale", x + 12, col_y + col_h - 13)
    draw_wrapped(
        c,
        "Add-on delivery capacity. Outcome Scale is the same idea, smaller increment. Not T&M.",
        x + 12,
        col_y + col_h - 26,
        "Inter",
        6.4,
        cw - 24,
        color=MUTED,
        leading=8.2,
    )
    scrum = [
        "Planning & architecture across the deployment lifecycle",
        "Design engineering for new and incremental Cisco designs",
        "Implementation / migration planning and implementation support",
        "Assessments: capacity, reliability, performance, security",
        "Test strategy, planning, execution, and retest support",
        "Security assessment and resilience (not incident response)",
        "Cloud transformation strategy ready for execution",
        "Observability / assurance on existing Cisco products",
        "Services as Code — automation and DevOps for network state",
        "Day-2 assurance: fault and change management",
        "Matrix analytics for mobility performance / OSS-BSS",
    ]
    bullet_block(c, scrum, x + 12, col_y + col_h - 46, cw - 24, size=6.35, leading=8.55)

    ey = col_y + 10
    eh = 40
    half = (cw - 28) / 2
    c.setFillColor(CYAN)
    c.setFont("Inter-Semi", 6.1)
    c.drawString(x + 12, ey + eh + 6, "EXPERT-AS-A-SERVICE ROLES")
    rrect(c, x + 10, ey, half, eh, r=5, fill=OFF)
    c.setFillColor(NAVY)
    c.setFont("Inter-Bold", 6.8)
    c.drawString(x + 18, ey + eh - 12, "Consulting Engineer")
    draw_wrapped(
        c,
        "Design, implement, operate, knowledge transfer.",
        x + 18,
        ey + eh - 24,
        "Inter",
        6.0,
        half - 16,
        color=MUTED,
        leading=7.6,
    )
    rrect(c, x + 14 + half, ey, half, eh, r=5, fill=OFF)
    c.setFillColor(NAVY)
    c.setFont("Inter-Bold", 6.8)
    c.drawString(x + 22 + half, ey + eh - 12, "Solution Architect")
    draw_wrapped(
        c,
        "Vision, strategy, architecture standards alignment.",
        x + 22 + half,
        ey + eh - 24,
        "Inter",
        6.0,
        half - 16,
        color=MUTED,
        leading=7.6,
    )

    # Footer
    c.setFillColor(NAVY)
    c.rect(0, 0, W, footer_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.rect(0, 0, 6, footer_h, fill=1, stroke=0)

    kicker(c, "Boundaries  ·  quote and service description control scope", 18, footer_h - 13)
    c.setFillColor(HexColor("#8AA0B5"))
    c.setFont("Inter", 5.4)
    c.drawRightString(
        W - 16,
        footer_h - 13,
        "Cisco Confidential — Cisco and partner use",
    )
    draw_wrapped(
        c,
        "Supplements a current Cisco Support or Partner Support agreement on covered products. Not T&M. Cisco technologies only — no work on third-party equipment. “Outcome” does not guarantee a result. Delivery stops the last business day of the term. Digital Insights require data collection (or a purchased air-gapped collector). Scrum / Outcome Scale / EaaS are subscription add-ons, not an hours bank. Hands-on upgrade, hardening, and Last-Date-of-Support remediation belong on a PS SOW (PS-Transactional), not folded into PS-Sub.  Sources: PS SD EDCS-25902390 (Dec 2025) · LCS SD EDCS-24374819 (Feb 2024) · LCS At-a-Glance (May 2024).",
        18,
        footer_h - 26,
        "Inter",
        6.15,
        W - 36,
        color=white,
        leading=7.9,
        max_lines=3,
    )

    arch = ["Networking", "Data Center", "Security", "Collaboration", "SP Mobility", "Internet of Things"]
    ax = 18
    ay = 8
    c.setFillColor(CYAN)
    c.setFont("Inter-Semi", 5.8)
    c.drawString(ax, ay + 3.4, "ARCHITECTURES")
    ax += c.stringWidth("ARCHITECTURES", "Inter-Semi", 5.8) + 8
    for a in arch:
        ax += chip(c, a, ax, ay, h=13, fill=DEEP, fg=white) + 5

    c.showPage()
    c.save()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    draw()
