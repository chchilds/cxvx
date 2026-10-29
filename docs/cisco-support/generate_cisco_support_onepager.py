#!/usr/bin/env python3
"""Generate a one-page landscape PDF: Cisco Support abilities.

Source: Cisco Support Customer Value Deck v8.0 (October 2026).
"""

from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Cisco_Support_Abilities_OnePager.pdf"

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


def bullet_block(c, items, x, y, max_w, font="Inter", size=6.4, color=INK, leading=8.35, bullet=CYAN):
    top = y
    for item in items:
        c.setFillColor(bullet)
        c.circle(x + 2.0, y + 1.5, 1.25, fill=1, stroke=0)
        h = draw_wrapped(c, item, x + 8.2, y, font, size, max_w - 8.2, color=color, leading=leading)
        y -= h
    return top - y


def kicker(c, text, x, y, color=CYAN):
    c.setFillColor(color)
    c.setFont("Inter-Semi", 6.15)
    c.drawString(x, y, text.upper())


def chip(c, text, x, y, h=13, fill=DEEP, fg=white, font="Inter-Med", size=5.9, pad=6.2):
    w = c.stringWidth(text, font, size) + pad * 2
    rrect(c, x, y, w, h, r=3.2, fill=fill)
    c.setFillColor(fg)
    c.setFont(font, size)
    c.drawString(x + pad, y + 3.4, text)
    return w


def draw():
    W, H = landscape(letter)
    c = canvas.Canvas(str(OUT), pagesize=(W, H))
    c.setTitle("Cisco Support Abilities at a Glance")
    c.setAuthor("Cisco Customer Experience")
    c.setSubject("Cisco Support abilities from Customer Value Deck v8.0 October 2026")

    header_h = 44
    footer_h = 62
    m = 13
    gap = 7

    c.setFillColor(NAVY)
    c.rect(0, H - header_h, W, header_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.rect(0, H - header_h, 6, header_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.setFont("Inter-Bold", 10.5)
    c.drawString(16, H - 18, "CISCO")
    c.setFillColor(white)
    c.setFont("Inter-Bold", 14)
    c.drawString(58, H - 19, "Support")
    c.setFillColor(CYAN)
    c.setFont("Inter-Med", 8.2)
    c.drawString(58, H - 34, "Powered by Cisco IQ  ·  Customer Value Deck v8.0  ·  October 2026")
    c.setFillColor(white)
    c.setFont("Inter-Semi", 10.5)
    c.drawRightString(W - 16, H - 18, "Abilities at a Glance")
    c.setFillColor(HexColor("#9BB4C9"))
    c.setFont("Inter", 6.8)
    c.drawRightString(W - 16, H - 34, "Lead with Standard or above  ·  Seller conversation sheet")

    # Three outcomes
    out_h = 52
    out_y = H - header_h - 7 - out_h
    outcomes = [
        ("Complete landscape clarity", "See the estate without gaps — inventory, lifecycle, coverage, criticality."),
        ("Proactive resilience", "Catch exposure before the business feels it — RIS, PSIRT, LDoS, quantum."),
        ("Rapid resolution", "Collapse hours to minutes — contextual TAC, AI troubleshooting, AI RMA."),
    ]
    ow = (W - 2 * m - 2 * gap) / 3
    for i, (t, b) in enumerate(outcomes):
        x = m + i * (ow + gap)
        rrect(c, x, out_y, ow, out_h, r=6, fill=NAVY)
        c.setFillColor(CYAN)
        c.setFont("Inter-Bold", 8.0)
        c.drawString(x + 10, out_y + out_h - 15, f"0{i+1}   {t}")
        draw_wrapped(c, b, x + 10, out_y + out_h - 29, "Inter", 6.4, ow - 20, color=white, leading=8.0)

    # Three tiers
    tier_h = 90
    tier_y = out_y - gap - tier_h
    tiers = [
        (
            False,
            "Basic",
            "Reactive foundational control",
            "TAC 24x7 · S1/2 60 min · S3/4 NBD · software updates · hardware RMA · communities. IQ: inventory, tagging, coverage, LDoS dashboard, PSIRT/FN reports, case + RMA creation, self-serve troubleshooting. Cisco U. free library.",
        ),
        (
            True,
            "Standard  ·  lead here",
            "Proactive risk mitigation  ·  includes Basic",
            "Solution support / centralized triage · S1/2 30 min · S3/4 60 min. AI insights: criticality, LDoS/EOL, peer benchmarking, hardening, config, quantum, certificates, priority bugs, EoVS. RIS Playbook + automated OS upgrades. Contextual troubleshooting. Personalized Cisco U. paths. Standard QuickStart bridges Basic customers mid-term.",
        ),
        (
            False,
            "Signature",
            "Operational excellence  ·  includes Standard",
            "Expert team that knows the estate. S1 15 min · restoration SLA. Systematic RCA, continuous optimization, asset management, periodic reviews, predictive disruption, config/hardening recommendations. Cert paths, 121+ labs, 20 named licenses. Full set requires USD $200k ACV.",
        ),
    ]
    tw = (W - 2 * m - 2 * gap) / 3
    for i, (hot, title, sub, body) in enumerate(tiers):
        x = m + i * (tw + gap)
        if hot:
            rrect(c, x, tier_y, tw, tier_h, r=6, fill=NAVY)
            c.setFillColor(CYAN)
            c.setFont("Inter-Semi", 6.0)
            c.drawString(x + 10, tier_y + tier_h - 11, "CISCO SUPPORT TIER")
            c.setFillColor(white)
            c.setFont("Inter-Bold", 9.0)
            c.drawString(x + 10, tier_y + tier_h - 24, title)
            c.setFillColor(CYAN)
            c.setFont("Inter-Med", 6.3)
            c.drawString(x + 10, tier_y + tier_h - 35, sub)
            draw_wrapped(c, body, x + 10, tier_y + tier_h - 48, "Inter", 6.05, tw - 20, color=white, leading=7.6)
        else:
            rrect(c, x, tier_y, tw, tier_h, r=6, fill=OFF, stroke=LINE)
            c.setFillColor(MUTED)
            c.setFont("Inter-Semi", 6.0)
            c.drawString(x + 10, tier_y + tier_h - 11, "CISCO SUPPORT TIER")
            c.setFillColor(NAVY)
            c.setFont("Inter-Bold", 9.0)
            c.drawString(x + 10, tier_y + tier_h - 24, title)
            c.setFillColor(MUTED)
            c.setFont("Inter-Med", 6.3)
            c.drawString(x + 10, tier_y + tier_h - 35, sub)
            draw_wrapped(c, body, x + 10, tier_y + tier_h - 48, "Inter", 6.05, tw - 20, color=INK, leading=7.6)

    # Main columns
    col_y = footer_h + 7
    col_h = tier_y - gap - col_y
    cw = (W - 2 * m - 2 * gap) / 3

    # Col 1 — Standard IQ abilities
    x = m
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "Standard abilities  ·  Cisco IQ  ·  * = AI-powered", x + 10, col_y + col_h - 13)
    kicker(c, "Complete landscape clarity", x + 10, col_y + col_h - 26, color=NAVY)
    used = 28
    used += bullet_block(
        c,
        [
            "Controller and non-controller install base",
            "Asset inventory, tagging, coverage, last-signal activity",
            "EOL / LDoS dashboard + insights *",
            "Asset criticality scoring and risk rating *",
            "Peer benchmarking; Services EA growth report",
        ],
        x + 10,
        col_y + col_h - 38,
        cw - 22,
        size=6.25,
        leading=8.15,
    )
    y2 = col_y + col_h - 38 - used - 4
    kicker(c, "Proactive resilience", x + 10, y2, color=NAVY)
    bullet_block(
        c,
        [
            "Security advisories, Field Notices, config, hardening *",
            "Software conformance, IOS XE upgrades, priority bugs *",
            "Quantum readiness and PKI certificate expiration insights *",
            "Software end-of-vulnerability support (EoVS) *",
            "RIS Playbook: map attack surface → remediate (automated OS upgrade) → modernize (LDoS, quantum, EoVS)",
            "Standard+ is the digital layer RIS depends on — Basic was not built for this threat velocity",
        ],
        x + 10,
        y2 - 12,
        cw - 22,
        size=6.25,
        leading=8.15,
    )
    kicker(c, "Contextual learning  ·  Cisco U.", x + 10, col_y + 36, color=NAVY)
    bullet_block(
        c,
        [
            "Standard: 155+ courses, unlimited named users, auto-provisioned",
            "Signature: 95+ cert paths, practice exams, 121+ labs, 20 named licenses via SAMT, CE credits",
        ],
        x + 10,
        col_y + 24,
        cw - 22,
        size=6.2,
        leading=8.1,
    )

    # Col 2 — Rapid resolution + Signature
    x = m + cw + gap
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "Rapid resolution", x + 10, col_y + col_h - 13)
    used = bullet_block(
        c,
        [
            "Contextual self-serve troubleshooting: live diagnostics + TAC knowledge + past cases *",
            "Streamlined case creation for serialized assets *",
            "AI-assisted RMA (Basic and Standard) *",
            "Automated RMA on connected DC hardware (Standard / Signature)",
            "Solution support: one case owner, Cisco + Solution Alliance triage (Standard+)",
        ],
        x + 10,
        col_y + col_h - 26,
        cw - 22,
        size=6.25,
        leading=8.2,
    )
    y2 = col_y + col_h - 26 - used - 6
    kicker(c, "Signature extras  ·  $200k ACV for the full set", x + 10, y2)
    bullet_block(
        c,
        [
            "Restoration SLA with service credit if missed",
            "Named expert team familiar with the environment",
            "Systematic root-cause investigation beyond the device",
            "Continuous optimization and maintenance-window support",
            "Proactive asset management (location, contract, entitlement)",
            "Periodic service reviews, KPIs, post-incident improvements",
            "Predictive disruption analysis; config + hardening recommendations *",
            "Software technical onboarding / expert software deployment (Security & Collaboration; not Duo)",
            "Cisco U. cert paths, practice exams, 121+ remote labs, 20 named licenses, CE credits",
        ],
        x + 10,
        y2 - 12,
        cw - 22,
        size=6.15,
        leading=8.05,
    )

    # Col 3 — IQ Link, On-Prem, vs PS
    x = m + 2 * (cw + gap)
    rrect(c, x, col_y, cw, col_h, r=7, fill=white, stroke=LINE)
    kicker(c, "Cisco IQ Link  ·  the value multiplier", x + 10, col_y + col_h - 13)
    draw_wrapped(
        c,
        "One secure on-prem connector. Replaces CSPC / CX Agent. Unifies telemetry from Catalyst Center and non-controller devices. Bidirectional: insights in, commands and OS/security updates out.",
        x + 10,
        col_y + col_h - 26,
        "Inter",
        6.25,
        cw - 22,
        color=INK,
        leading=8.0,
    )
    bullet_block(
        c,
        [
            "Unlocks deeper Standard+ assessments (hardening, config, quantum BOM, certificates, OS upgrade, RIS)",
            "Air-gapped IQ On-Prem VA available now; tethered mode coming soon",
            "Ask AI / local inference needs GPU; core lifecycle intelligence does not",
            "IQ APIs + MCP server: insights into existing tools and AI agents, at scale",
        ],
        x + 10,
        col_y + col_h - 62,
        cw - 22,
        size=6.2,
        leading=8.1,
    )

    box_h = 102
    by = col_y + 8
    rrect(c, x + 8, by, cw - 16, box_h, r=5, fill=HexColor("#F7F1E4"), stroke=GOLD, sw=0.85)
    kicker(c, "Do not confuse Support with PS", x + 16, by + box_h - 12, color=HexColor("#8A5A12"))
    c.setFillColor(GROW)
    c.setFont("Inter-Semi", 6.3)
    c.drawString(x + 16, by + box_h - 24, "Cisco Support / Partner Support")
    h1 = draw_wrapped(
        c,
        "Keep it running. Digital insights, TAC, RMA, RIS Playbook. Same customer-facing capabilities on Cisco Support and Cisco Partner Support.",
        x + 16,
        by + box_h - 35,
        "Inter",
        6.1,
        cw - 34,
        color=INK,
        leading=7.7,
    )
    c.setFillColor(WARN)
    c.setFont("Inter-Semi", 6.3)
    c.drawString(x + 16, by + box_h - 38 - h1, "Professional Services")
    draw_wrapped(
        c,
        "Optional. Use PS-Transactional (or partner PS) for white-glove RIS remediation / modernization. PS-Sub / LCS is advisory + outcomes, not this TAC/IQ layer. IQ will later connect to Professional Services; LCS/BCS customers get IQ + IQ Link next.",
        x + 16,
        by + box_h - 49 - h1,
        "Inter",
        6.1,
        cw - 34,
        color=INK,
        leading=7.7,
    )

    # Footer
    c.setFillColor(NAVY)
    c.rect(0, 0, W, footer_h, fill=1, stroke=0)
    c.setFillColor(CYAN)
    c.rect(0, 0, 6, footer_h, fill=1, stroke=0)
    kicker(c, "Boundaries from the Customer Value Deck", 16, footer_h - 12)
    c.setFillColor(HexColor("#8AA0B5"))
    c.setFont("Inter", 5.3)
    c.drawRightString(W - 16, footer_h - 12, "Cisco Confidential — Cisco and partner use")
    draw_wrapped(
        c,
        "Feature availability varies by platform. AppDynamics, Splunk, ThousandEyes, and CVM are not on Signature. Meraki is SLO only and not on Asset Management. SP Hardware and SIA subscriptions do not include software updates. Signature full deliverables require USD $200k ACV (GUID-aggregated). Partner Support Signature expected GA end of Q2 FY27. RIS is Standard+ digital layer; hands-on modernization is optional PS-T.",
        16,
        footer_h - 25,
        "Inter",
        6.0,
        W - 32,
        color=white,
        leading=7.6,
        max_lines=3,
    )
    ax = 16
    ay = 7
    for a in ["Basic", "Standard", "Signature", "Cisco IQ", "IQ Link", "On-Prem VA", "RIS Playbook"]:
        ax += chip(c, a, ax, ay, h=12.5, fill=DEEP, fg=white, size=5.7) + 4
    c.setFillColor(HexColor("#8AA0B5"))
    c.setFont("Inter", 5.15)
    c.drawRightString(W - 16, ay + 3.4, "Source: Cisco Support Customer Value Deck v8.0 · 25 Oct 2026")

    c.showPage()
    c.save()
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    draw()
