#!/usr/bin/env python3
"""Generate a Cisco Enterprise Agreement summary PowerPoint from an EAMP Price Estimate Excel file."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_FILL
from pptx.util import Inches, Pt

from apply_dark_theme import apply_dark_theme
from eamp_parser import EAData, LineItem, parse_eamp_excel
from theme import CISCO_BLUE, WHITE


def _fmt_currency(value: float) -> str:
    return f"${value:,.0f}" if value else "—"


def _add_title_banner(slide, title: str) -> None:
    banner = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.85))
    banner.fill.solid()
    banner.fill.fore_color.rgb = CISCO_BLUE
    banner.line.fill.background()
    tf = banner.text_frame
    tf.text = title
    p = tf.paragraphs[0]
    p.font.size = Pt(26)
    p.font.bold = True
    p.font.color.rgb = WHITE


def _add_text(slide, left, top, width, height, text: str, size: int = 13, bold: bool = False) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)


def _style_table_header(table, row: int = 0) -> None:
    for col in range(len(table.columns)):
        cell = table.cell(row, col)
        cell.fill.solid()
        cell.fill.fore_color.rgb = CISCO_BLUE
        for p in cell.text_frame.paragraphs:
            for run in p.runs:
                run.font.color.rgb = WHITE
                run.font.bold = True
                run.font.size = Pt(10)


def _fill_table_row(table, row: int, values: list[str], alt: bool = False) -> None:
    for col, val in enumerate(values):
        table.cell(row, col).text = val
        if alt:
            table.cell(row, col).fill.solid()
            table.cell(row, col).fill.fore_color.rgb = RGBColor(0xF5, 0xF5, 0xF5)
        for p in table.cell(row, col).text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(9)


def build_presentation(data: EAData) -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    display_name = data.smart_account or data.customer_name

    # Slide 1 — Title
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, display_name)
    _add_text(s, Inches(0.5), Inches(1.1), Inches(9), Inches(0.9), "Cisco Enterprise Agreement\nPrice Estimate Summary", size=30, bold=True)
    meta = (
        f"Proposal ID: {data.proposal_id}\n"
        f"Project: {data.project_name}\n"
        f"Expected Book Date: {data.rsd}\n"
        f"{data.term_months}-Month Term | {data.billing} Billing\n"
        f"Report Date: {data.report_date}"
    )
    _add_text(s, Inches(0.5), Inches(2.4), Inches(9), Inches(2), meta, size=15)
    _add_text(s, Inches(0.5), Inches(6.6), Inches(9), Inches(0.4), "CISCO CONFIDENTIAL — Indicative pricing only. Not an approved Cisco quote.", size=10)

    # Slide 2 — Deal Overview
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Deal Overview")
    overview = (
        "Customer & Partners\n"
        f"End Customer: {data.customer_name}\n"
        f"Smart Account: {data.smart_account}\n"
        f"Address: {data.address}\n"
        f"Partner: {data.partner}\n"
        f"Reseller Services: {data.reseller}\n"
        f"Cisco Account Manager: {data.account_manager}\n\n"
        "Agreement Details\n"
        f"Enterprise Agreement Motion: {data.ea_motion}\n"
        f"Buying Program ID: {data.buying_program_id}\n"
        f"Deal ID: {data.deal_id}  |  Quote ID: {data.quote_id or '—'}\n"
        f"Price List: {data.price_list}\n"
        f"Duration: {data.term_months} months  |  Billing: {data.billing}  |  RSD: {data.rsd}"
    )
    _add_text(s, Inches(0.5), Inches(1.1), Inches(9), Inches(5.8), overview, size=13)

    # Slide 3 — Executive Summary
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Executive Summary")
    _add_text(s, Inches(0.5), Inches(1.0), Inches(9), Inches(0.5), f"3-Year Total Contract Value: {_fmt_currency(data.total_tcv)}", size=20, bold=True)

    chart_portfolios = [p for p in data.portfolios if p["total"] > 0]
    rows = len(chart_portfolios) + 2
    table = s.shapes.add_table(rows, 4, Inches(0.4), Inches(1.6), Inches(5.6), Inches(0.32 * rows)).table
    _fill_table_row(table, 0, ["Portfolio", "Software (USD)", "Services (USD)", "Total (USD)"])
    _style_table_header(table)
    for idx, p in enumerate(chart_portfolios, start=1):
        _fill_table_row(
            table, idx,
            [p["name"], _fmt_currency(p["software"]), _fmt_currency(p["services"]) if p["services"] else "—", _fmt_currency(p["total"])],
            alt=idx % 2 == 0,
        )
    _fill_table_row(table, rows - 1, ["TOTAL", _fmt_currency(data.total_software), _fmt_currency(data.total_services), _fmt_currency(data.total_tcv)])
    _style_table_header(table, rows - 1)

    if chart_portfolios:
        chart_data = CategoryChartData()
        chart_data.categories = [p["name"].replace(" Infrastructure", "") for p in chart_portfolios]
        chart_data.add_series("TCV", [p["total"] for p in chart_portfolios])
        s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(6.1), Inches(1.6), Inches(3.3), Inches(3.0), chart_data)

    highlights = "Key Highlights\n" + "\n".join(f"• {h}" for h in data.highlights)
    _add_text(s, Inches(0.4), Inches(4.6), Inches(9.2), Inches(2.6), highlights, size=12)

    # Slide 4 — Networking
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Networking Infrastructure")
    _add_text(
        s, Inches(0.5), Inches(1.0), Inches(9), Inches(0.4),
        f"Software TCV: {_fmt_currency(data.networking_sw_tcv)}  |  Services TCV: {_fmt_currency(data.networking_svc_tcv)}",
        size=14, bold=True,
    )
    net_lines = []
    if data.meraki:
        net_lines.append(
            f"Meraki Cloud-Managed\n"
            f"  {data.meraki.get('AP', 0)} × MR Essentials (wireless APs)\n"
            f"  {data.meraki.get('Switch', 0)} × MS switches  |  {data.meraki.get('MX', 0)} × MX firewalls"
            + (f"  |  {data.meraki['Camera']} × MV cameras" if data.meraki.get("Camera") else "")
        )
    top_sw = sorted(data.networking_sw, key=lambda x: -x.tcv)[:6]
    campus = [i for i in top_sw if i.part_number.startswith(("E3N-C", "E3N-N", "E3N-AIR", "E3N-SP", "E3N-CS"))]
    if campus:
        net_lines.append("Data Center & Campus")
        for item in campus[:5]:
            net_lines.append(f"  {item.description[:55]} — {int(item.desired)} desired (IB: {int(item.ib)})")
    top_svc = sorted(data.networking_svc, key=lambda x: -x.tcv)[:4]
    if top_svc:
        net_lines.append("Top Support Services")
        for item in top_svc:
            net_lines.append(f"  {item.description[:50]}: {_fmt_currency(item.tcv)}")
    _add_text(s, Inches(0.5), Inches(1.5), Inches(9), Inches(5.5), "\n".join(net_lines), size=12)

    # Slide 5 — Security
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Security Portfolio")
    _add_text(
        s, Inches(0.5), Inches(1.0), Inches(9), Inches(0.4),
        f"Software TCV: {_fmt_currency(data.security_sw_tcv)}  |  Services TCV: {_fmt_currency(data.security_svc_tcv)}",
        size=14, bold=True,
    )
    if data.security_sw:
        sec_rows = len(data.security_sw) + 1
        table = s.shapes.add_table(sec_rows, 4, Inches(0.4), Inches(1.5), Inches(9.2), Inches(0.3 * sec_rows)).table
        _fill_table_row(table, 0, ["Solution", "Product", "Qty (Desired)", "TCV (USD)"])
        _style_table_header(table)
        for idx, item in enumerate(data.security_sw, start=1):
            qty = int(item.desired) if item.desired else int(item.ib)
            _fill_table_row(table, idx, [item.solution, item.description[:40], str(qty), _fmt_currency(item.tcv)], alt=idx % 2 == 0)
    footer = f"Security Content Users (SCU): {data.scu_count:,}\n"
    if data.security_svc:
        footer += "Top security services: " + ", ".join(f"{i.description[:35]} ({_fmt_currency(i.tcv)})" for i in sorted(data.security_svc, key=lambda x: -x.tcv)[:3])
    _add_text(s, Inches(0.5), Inches(4.5), Inches(9), Inches(2.5), footer, size=12)

    # Slide 6 — Applications
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Applications Infrastructure")
    _add_text(s, Inches(0.5), Inches(1.0), Inches(9), Inches(0.4), f"Software TCV: {_fmt_currency(data.applications_sw_tcv)}", size=14, bold=True)
    app_lines = []
    for item in sorted(data.applications, key=lambda x: -x.tcv):
        app_lines.append(
            f"{item.solution} — {item.description}\n"
            f"  Part Number: {item.part_number}  |  Quantity: {int(item.desired)}  |  TCV: {_fmt_currency(item.tcv)}"
        )
    if not app_lines:
        app_lines.append("No applications portfolio line items in this proposal.")
    _add_text(s, Inches(0.5), Inches(1.5), Inches(9), Inches(5.5), "\n\n".join(app_lines), size=12)

    # Slide 7 — Support & Lifecycle
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Support & Lifecycle Services")
    _add_text(s, Inches(0.5), Inches(1.0), Inches(9), Inches(0.4), f"Total Services TCV: {_fmt_currency(data.total_services)}", size=14, bold=True)
    if data.support_breakdown:
        sb_rows = len(data.support_breakdown) + 1
        table = s.shapes.add_table(sb_rows, 2, Inches(0.4), Inches(1.5), Inches(9.2), Inches(0.32 * sb_rows)).table
        _fill_table_row(table, 0, ["Category", "TCV (USD)"])
        _style_table_header(table)
        for idx, row in enumerate(data.support_breakdown, start=1):
            _fill_table_row(table, idx, [row["category"], _fmt_currency(row["tcv"])], alt=idx % 2 == 0)

    # Slide 8 — Annual Cost
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Annual Cost Breakdown")
    _add_text(s, Inches(0.5), Inches(1.0), Inches(9), Inches(0.35), f"Based on {data.term_months}-month term with {data.billing.lower()} billing", size=12)
    years = data.term_months / 12
    annual_sw = data.total_software / years if years else 0
    annual_svc = data.total_services / years if years else 0
    annual_total = data.total_tcv / years if years else 0
    table = s.shapes.add_table(4, 3, Inches(0.4), Inches(1.5), Inches(9.2), Inches(1.4)).table
    _fill_table_row(table, 0, ["Category", f"{int(years)}-Year TCV", "Estimated Annual"])
    _style_table_header(table)
    _fill_table_row(table, 1, ["Software Subscriptions", _fmt_currency(data.total_software), _fmt_currency(annual_sw)], alt=True)
    _fill_table_row(table, 2, ["Support & Lifecycle Services", _fmt_currency(data.total_services), _fmt_currency(annual_svc)])
    _fill_table_row(table, 3, ["TOTAL", _fmt_currency(data.total_tcv), _fmt_currency(annual_total)], alt=True)
    notes = (
        f"Approximate annual spend: {_fmt_currency(annual_total)}/year\n"
        f"Billing model: {data.billing} (invoiced yearly over {int(years)}-year term)\n"
        "Programmatic discounts applied on EA 3.0 subscriptions\n"
        "Pre-EA install base assets consumed upon EA activation"
    )
    _add_text(s, Inches(0.5), Inches(3.2), Inches(9), Inches(3.5), notes, size=12)

    # Slide 9 — Out of Scope
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Portfolios Not in Scope")
    oos_text = "The following EA portfolios have no line items in this proposal:\n\n"
    oos_text += "\n".join(f"• {item}" for item in data.out_of_scope)
    oos_text += "\n\nThese can be added in future EA amendments if needed."
    _add_text(s, Inches(0.5), Inches(1.1), Inches(9), Inches(5.5), oos_text, size=14)

    # Slide 10 — Professional Services
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Recommended Professional Services")
    ps = (
        "Context\n"
        f"• Professional Services are not included in Proposal {data.proposal_id} "
        f"({_fmt_currency(data.total_services)} support is in scope).\n"
        "• Recommendations align EA portfolios with Cisco Support / PS lifecycle: Activate → Integrate → Optimize.\n\n"
        "Tier 1 — EA Activation (Must-Have)\n"
        "• Enterprise Agreement Management Support (included with Services EA purchase)\n"
        "• Install Base Assessment & Entitlement Reconciliation\n"
        "• EA Activation & License Consumption Planning\n\n"
        "Tier 2 — Highest EA Impact\n"
    )
    if data.meraki.get("AP"):
        ps += f"• Meraki At-Scale Deployment — {data.meraki['AP']} APs, {data.meraki.get('Switch', 0)} switches\n"
    if any("ISE" in i.part_number for i in data.security_sw):
        ps += "• ISE Deployment & Policy Migration — large ISE Advantage footprint\n"
    if any("TE-" in i.part_number for i in data.applications):
        ps += "• ThousandEyes Deployment & Baseline Digital Experience Monitoring\n"
    if any("SK-" in i.part_number for i in data.applications):
        ps += "• Splunk Cloud Onboarding & Data Ingestion Optimization\n"
    ps += (
        "• Catalyst / Nexus DNA EA Onboarding & Automation\n\n"
        "Tier 3 — Support & Enablement\n"
        "• Services EA wrapper, Signature/Standard support tiers\n"
        "• Cisco Learning Credits for admin and engineering teams"
    )
    _add_text(s, Inches(0.5), Inches(1.0), Inches(9), Inches(6.0), ps, size=11)
    _add_text(s, Inches(0.5), Inches(6.8), Inches(9), Inches(0.4), "Cisco Confidential | Indicative PS Recommendations — Not a formal quote", size=9)

    # Slide 11 — Next Steps
    s = prs.slides.add_slide(blank)
    _add_title_banner(s, "Next Steps & Important Notes")
    next_steps = (
        "Important Disclaimers\n"
        "• This is indicative pricing only — NOT an approved Cisco quote\n"
        "• Pricing is subject to change until formal quote approval\n"
        "• Pre-EA install base licenses will be consumed at EA activation\n\n"
        "Recommended Actions\n"
        "• Review portfolio quantities against current install base\n"
    )
    if data.meraki.get("AP"):
        next_steps += f"• Validate Meraki counts ({data.meraki['AP']} APs) against deployment plan\n"
    ise = next((i for i in data.security_sw if "ISE-ADV" in i.part_number), None)
    if ise:
        next_steps += f"• Confirm ISE Advantage license count ({int(ise.desired)} desired, IB: {int(ise.ib)})\n"
    next_steps += (
        "• Confirm support coverage levels (8x5xNBD vs 24x7x4) meet SLA requirements\n"
        f"• Work with {data.partner or 'your partner'} and Cisco AM {data.account_manager or ''} to finalize quote\n"
    )
    if data.proposal_id:
        next_steps += f"• EA Proposal Link: apps.cisco.com/ea/project/proposal/{data.proposal_id}"
    _add_text(s, Inches(0.5), Inches(1.1), Inches(9), Inches(5.8), next_steps, size=12)
    _add_text(s, Inches(0.5), Inches(6.8), Inches(9), Inches(0.4), "Cisco Confidential | Indicative Pricing Only", size=9)

    return prs


def inspect_workbook(path: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=True, read_only=True)
    print(f"Workbook: {path.name}")
    for name in wb.sheetnames:
        ws = wb[name]
        print(f"\n  Sheet: {name!r} ({ws.max_row} rows x {ws.max_column} cols)")
        for row in ws.iter_rows(min_row=1, max_row=min(15, ws.max_row or 15), values_only=True):
            cells = [str(c).strip() for c in row if c is not None and str(c).strip()]
            if cells:
                print("    " + " | ".join(cells[:8]))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate EA summary PowerPoint from EAMP Excel")
    parser.add_argument("excel", type=Path, nargs="?", help="EAMP Price Estimate .xlsx file")
    parser.add_argument("-o", "--output", type=Path, default=Path("Enterprise_Agreement_Summary.pptx"))
    parser.add_argument("--dark", action="store_true", help="Apply Cisco dark theme")
    parser.add_argument("--inspect", action="store_true", help="List workbook sheets and sample rows")
    args = parser.parse_args(argv)

    if not args.excel:
        parser.error("excel file is required")

    if not args.excel.exists():
        print(f"Error: file not found: {args.excel}", file=sys.stderr)
        return 1

    if args.inspect:
        inspect_workbook(args.excel)
        return 0

    data = parse_eamp_excel(args.excel)
    prs = build_presentation(data)
    prs.save(str(args.output))
    print(f"Generated: {args.output} ({len(prs.slides)} slides)")
    print(f"  Customer: {data.customer_name}")
    print(f"  Proposal: {data.proposal_id}  |  TCV: {_fmt_currency(data.total_tcv)}")

    if args.dark:
        apply_dark_theme(args.output)
        print(f"Dark theme applied: {args.output}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
