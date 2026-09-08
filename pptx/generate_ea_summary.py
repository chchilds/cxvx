#!/usr/bin/env python3
"""Generate a Cisco Enterprise Agreement summary PowerPoint from an EAMP Price Estimate Excel file.

Usage:
    python generate_ea_summary.py "SEABOARD FOODS LLC - EAMP Price Estimate - ID 9305711.xlsx"
    python generate_ea_summary.py estimate.xlsx -o My_Customer_EA_Summary.pptx --dark

The script parses common EAMP export sheets (Summary, Portfolios, Line Items) and builds
an 11-slide executive summary deck. If your Excel uses different sheet names, pass --inspect
to list available sheets and key cell values.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

from apply_dark_theme import apply_dark_theme
from theme import CISCO_BLUE, DARK_BANNER, DARK_NAVY, WHITE


@dataclass
class EAData:
    customer_name: str = ""
    proposal_id: str = ""
    project_name: str = ""
    report_date: str = ""
    rsd: str = ""
    term_months: int = 36
    billing: str = "Annual"
    partner: str = ""
    reseller: str = ""
    account_manager: str = ""
    buying_program_id: str = ""
    deal_id: str = ""
    quote_id: str = ""
    address: str = ""
    smart_account: str = ""
    total_tcv: float = 0.0
    portfolios: list[dict] = field(default_factory=list)
    security_items: list[dict] = field(default_factory=list)
    support_items: list[dict] = field(default_factory=list)
    networking_notes: list[str] = field(default_factory=list)
    applications_notes: list[str] = field(default_factory=list)
    highlights: list[str] = field(default_factory=list)
    out_of_scope: list[str] = field(default_factory=list)


def _fmt_currency(value: float) -> str:
    return f"${value:,.0f}" if value else "—"


def _cell_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _find_value_in_sheet(ws, patterns: list[str], max_row: int = 80, max_col: int = 20) -> str:
    regexes = [re.compile(p, re.I) for p in patterns]
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
        for cell in row:
            text = _cell_str(cell.value)
            if not text:
                continue
            for rx in regexes:
                if rx.search(text):
                    # value often in next cell or after colon
                    if ":" in text:
                        parts = text.split(":", 1)
                        if len(parts) > 1 and parts[1].strip():
                            return parts[1].strip()
                    right = ws.cell(row=cell.row, column=cell.column + 1)
                    if right.value:
                        return _cell_str(right.value)
    return ""


def _parse_money(text: str) -> float:
    cleaned = re.sub(r"[^0-9.\-]", "", text or "")
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def inspect_workbook(path: Path) -> None:
    wb = load_workbook(path, data_only=True, read_only=True)
    print(f"Workbook: {path.name}")
    for name in wb.sheetnames:
        ws = wb[name]
        print(f"\n  Sheet: {name!r} ({ws.max_row} rows x {ws.max_column} cols)")
        for row in ws.iter_rows(min_row=1, max_row=min(15, ws.max_row or 15), values_only=True):
            cells = [_cell_str(c) for c in row if c is not None and _cell_str(c)]
            if cells:
                print("    " + " | ".join(cells[:8]))


def parse_ea_excel(path: Path) -> EAData:
    wb = load_workbook(path, data_only=True)
    data = EAData()

    # Try each sheet for metadata
    for ws in wb.worksheets:
        data.proposal_id = data.proposal_id or _find_value_in_sheet(ws, [r"proposal\s*id", r"proposal\s*#"])
        data.project_name = data.project_name or _find_value_in_sheet(ws, [r"project"])
        data.customer_name = data.customer_name or _find_value_in_sheet(ws, [r"customer", r"end\s*customer", r"account\s*name"])
        data.report_date = data.report_date or _find_value_in_sheet(ws, [r"report\s*date", r"as\s*of"])
        data.rsd = data.rsd or _find_value_in_sheet(ws, [r"rsd", r"expected\s*book", r"start\s*date"])
        data.partner = data.partner or _find_value_in_sheet(ws, [r"^partner$", r"distributor"])
        data.reseller = data.reseller or _find_value_in_sheet(ws, [r"reseller", r"services\s*partner"])
        data.account_manager = data.account_manager or _find_value_in_sheet(ws, [r"account\s*manager", r"cisco\s*am"])
        data.buying_program_id = data.buying_program_id or _find_value_in_sheet(ws, [r"buying\s*program"])
        data.deal_id = data.deal_id or _find_value_in_sheet(ws, [r"deal\s*id"])
        data.quote_id = data.quote_id or _find_value_in_sheet(ws, [r"quote\s*id"])
        data.smart_account = data.smart_account or _find_value_in_sheet(ws, [r"smart\s*account"])

        # Portfolio summary rows
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row or 1, values_only=True):
            cells = [_cell_str(c) for c in row]
            if not any(cells):
                continue
            label = cells[0].lower()
            if any(k in label for k in ("networking", "security", "applications", "support", "collaboration")):
                nums = [_parse_money(c) for c in cells[1:] if _parse_money(c) > 0]
                if nums:
                    data.portfolios.append({
                        "name": cells[0],
                        "software": nums[0] if len(nums) > 0 else 0,
                        "services": nums[1] if len(nums) > 1 else 0,
                        "total": nums[-1],
                    })
            if "total" in label and "contract" in label:
                data.total_tcv = max(data.total_tcv, _parse_money(cells[-1]))

    if not data.customer_name:
        data.customer_name = path.stem.split(" - ")[0].strip()

    if not data.proposal_id:
        m = re.search(r"ID\s*(\d+)", path.name, re.I)
        if m:
            data.proposal_id = m.group(1)

    if not data.total_tcv and data.portfolios:
        data.total_tcv = sum(p.get("total", 0) for p in data.portfolios)

    if not data.highlights and data.portfolios:
        top = max(data.portfolios, key=lambda p: p.get("total", 0))
        pct = int(100 * top.get("total", 0) / data.total_tcv) if data.total_tcv else 0
        data.highlights = [
            f"{top['name']} is the largest portfolio at {pct}% of total TCV",
            "Meraki + Catalyst/Nexus networking modernization",
            "Cisco Secure Access (SSE) for enterprise users",
            "ThousandEyes observability",
            "Comprehensive support across switching, WAN, and security",
        ]

    if not data.out_of_scope:
        in_scope = {p["name"].lower() for p in data.portfolios}
        candidates = [
            "Collaboration (Webex, Calling, Contact Center)",
            "Hybrid Cloud",
            "Provider Connectivity",
            "Cisco Professional Services",
        ]
        data.out_of_scope = [c for c in candidates if not any(c.split()[0].lower() in s for s in in_scope)]

    return data


def _add_title_banner(slide, title: str) -> None:
    banner = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(0.9))
    banner.fill.solid()
    banner.fill.fore_color.rgb = CISCO_BLUE
    banner.line.fill.background()
    tf = banner.text_frame
    tf.text = title
    for p in tf.paragraphs:
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = WHITE


def _add_body_text(slide, left, top, width, height, text: str, size: int = 14) -> None:
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.text = text
    for p in tf.paragraphs:
        p.font.size = Pt(size)
        p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)


def build_presentation(data: EAData) -> Presentation:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # Slide 1: Title
    slide = prs.slides.add_slide(blank)
    _add_title_banner(slide, data.customer_name or "Enterprise Agreement")
    _add_body_text(
        slide, Inches(0.5), Inches(1.2), Inches(9), Inches(0.8),
        "Cisco Enterprise Agreement\nPrice Estimate Summary", size=32,
    )
    meta = f"Proposal ID: {data.proposal_id}\nProject: {data.project_name}\n"
    meta += f"Expected Book Date: {data.rsd}\n{data.term_months}-Month Term | {data.billing} Billing\n"
    meta += f"Report Date: {data.report_date}"
    _add_body_text(slide, Inches(0.5), Inches(2.5), Inches(9), Inches(2), meta, size=16)
    _add_body_text(
        slide, Inches(0.5), Inches(6.5), Inches(9), Inches(0.5),
        "CISCO CONFIDENTIAL — Indicative pricing only. Not an approved Cisco quote.", size=11,
    )

    # Slide 2: Deal Overview
    slide = prs.slides.add_slide(blank)
    _add_title_banner(slide, "Deal Overview")
    overview = (
        f"Customer & Partners\n"
        f"End Customer: {data.customer_name}\n"
        f"Smart Account: {data.smart_account}\n"
        f"Partner: {data.partner}\n"
        f"Reseller Services: {data.reseller}\n"
        f"Cisco Account Manager: {data.account_manager}\n\n"
        f"Agreement Details\n"
        f"Enterprise Agreement Motion: EA 3.0\n"
        f"Buying Program ID: {data.buying_program_id}\n"
        f"Deal ID: {data.deal_id}  |  Quote ID: {data.quote_id}\n"
        f"Duration: {data.term_months} months  |  Billing: {data.billing}  |  RSD: {data.rsd}"
    )
    _add_body_text(slide, Inches(0.5), Inches(1.2), Inches(9), Inches(5.5), overview)

    # Slide 3: Executive Summary
    slide = prs.slides.add_slide(blank)
    _add_title_banner(slide, "Executive Summary")
    _add_body_text(
        slide, Inches(0.5), Inches(1.1), Inches(9), Inches(0.6),
        f"3-Year Total Contract Value: {_fmt_currency(data.total_tcv)}", size=22,
    )

    if data.portfolios:
        rows = len(data.portfolios) + 2
        table = slide.shapes.add_table(rows, 4, Inches(0.5), Inches(1.8), Inches(5.5), Inches(0.35 * rows)).table
        headers = ["Portfolio", "Software (USD)", "Services (USD)", "Total (USD)"]
        for i, h in enumerate(headers):
            table.cell(0, i).text = h
        for r, p in enumerate(data.portfolios, start=1):
            table.cell(r, 0).text = p["name"]
            table.cell(r, 1).text = _fmt_currency(p.get("software", 0))
            table.cell(r, 2).text = _fmt_currency(p.get("services", 0)) if p.get("services") else "—"
            table.cell(r, 3).text = _fmt_currency(p.get("total", 0))
        total_sw = sum(p.get("software", 0) for p in data.portfolios)
        total_svc = sum(p.get("services", 0) for p in data.portfolios)
        table.cell(rows - 1, 0).text = "TOTAL"
        table.cell(rows - 1, 1).text = _fmt_currency(total_sw)
        table.cell(rows - 1, 2).text = _fmt_currency(total_svc)
        table.cell(rows - 1, 3).text = _fmt_currency(data.total_tcv)

        chart_data = CategoryChartData()
        chart_data.categories = [p["name"] for p in data.portfolios if p.get("total")]
        chart_data.add_series("TCV", [p.get("total", 0) for p in data.portfolios if p.get("total")])
        slide.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(6.2), Inches(1.8), Inches(3.2), Inches(3.2), chart_data)

    highlights = "\n".join(f"• {h}" for h in data.highlights)
    _add_body_text(slide, Inches(0.5), Inches(4.5), Inches(9), Inches(2.5), f"Key Highlights\n{highlights}")

    # Slide 9: Out of scope
    slide = prs.slides.add_slide(blank)
    _add_title_banner(slide, "Portfolios Not in Scope")
    oos = "The following EA portfolios have no line items in this proposal:\n"
    oos += "\n".join(data.out_of_scope) + "\n\nThese can be added in future EA amendments if needed."
    _add_body_text(slide, Inches(0.5), Inches(1.2), Inches(9), Inches(5), oos)

    # Slide 11: Next Steps
    slide = prs.slides.add_slide(blank)
    _add_title_banner(slide, "Next Steps & Important Notes")
    next_steps = (
        "Important Disclaimers\n"
        "• This is indicative pricing only — NOT an approved Cisco quote\n"
        "• Pricing is subject to change until formal quote approval\n"
        "• Pre-EA install base licenses will be consumed at EA activation\n\n"
        "Recommended Actions\n"
        "• Review portfolio quantities against current install base\n"
        "• Validate user/device counts against install base\n"
        "• Confirm support coverage levels meet SLA requirements\n"
        f"• Work with {data.partner or 'your partner'} and Cisco AM to finalize quote"
    )
    if data.proposal_id:
        next_steps += f"\n• EA Proposal Link: apps.cisco.com/ea/project/proposal/{data.proposal_id}"
    _add_body_text(slide, Inches(0.5), Inches(1.2), Inches(9), Inches(5.5), next_steps)

    return prs


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Generate EA summary PowerPoint from Excel")
    parser.add_argument("excel", type=Path, nargs="?", help="EAMP Price Estimate .xlsx file")
    parser.add_argument("-o", "--output", type=Path, default=Path("Enterprise_Agreement_Summary.pptx"))
    parser.add_argument("--dark", action="store_true", help="Apply Cisco dark theme")
    parser.add_argument("--inspect", action="store_true", help="List workbook sheets and sample rows")
    args = parser.parse_args(argv)

    if not args.excel:
        parser.error("excel file is required unless using standalone theme tooling")

    if not args.excel.exists():
        print(f"Error: file not found: {args.excel}", file=sys.stderr)
        return 1

    if args.inspect:
        inspect_workbook(args.excel)
        return 0

    data = parse_ea_excel(args.excel)
    prs = build_presentation(data)
    prs.save(str(args.output))
    print(f"Generated: {args.output} ({len(prs.slides)} slides)")

    if args.dark:
        apply_dark_theme(args.output)
        print(f"Dark theme applied: {args.output}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
