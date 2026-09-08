"""Parse Cisco EAMP Proposal Export Summary Excel files."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook


@dataclass
class LineItem:
    solution: str
    suite: str
    part_number: str
    description: str
    ib: float
    desired: float
    tcv: float


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
    price_list: str = ""
    ea_motion: str = "EA 3.0"
    proposal_link: str = ""
    total_tcv: float = 0.0
    total_software: float = 0.0
    total_services: float = 0.0
    portfolios: list[dict] = field(default_factory=list)
    networking_sw: list[LineItem] = field(default_factory=list)
    networking_svc: list[LineItem] = field(default_factory=list)
    applications: list[LineItem] = field(default_factory=list)
    security_sw: list[LineItem] = field(default_factory=list)
    security_svc: list[LineItem] = field(default_factory=list)
    collaboration_svc: list[LineItem] = field(default_factory=list)
    support_breakdown: list[dict] = field(default_factory=list)
    highlights: list[str] = field(default_factory=list)
    out_of_scope: list[str] = field(default_factory=list)
    networking_sw_tcv: float = 0.0
    networking_svc_tcv: float = 0.0
    applications_sw_tcv: float = 0.0
    security_sw_tcv: float = 0.0
    security_svc_tcv: float = 0.0
    scu_count: int = 0
    meraki: dict = field(default_factory=dict)


def _cell_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _float(value) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _find_tcv_col(ws, header_row: int) -> int:
    for c in range(1, 70):
        h = ws.cell(header_row, c).value
        if h and "Total Contract Value" in str(h):
            return c
    return 21


def _find_header_row(ws, start: int = 15, end: int = 70) -> int | None:
    for r in range(start, end):
        for c in range(1, 10):
            if ws.cell(r, c).value == "Part Number":
                return r
    return None


def _ib_desired_cols(ws, header_row: int) -> tuple[int, int]:
    ib_col, desired_col = 7, 8
    for c in range(1, 15):
        h = _cell_str(ws.cell(header_row, c).value).lower()
        if "quantity found" in h or h == "quantity found (ib)":
            ib_col = c
        if "quantity desired" in h:
            desired_col = c
    return ib_col, desired_col


def _parse_line_items(ws, start_row: int, end_row: int, header_row: int) -> list[LineItem]:
    tcv_col = _find_tcv_col(ws, header_row)
    ib_col, desired_col = _ib_desired_cols(ws, header_row)
    items: list[LineItem] = []
    for r in range(start_row, end_row + 1):
        pn = ws.cell(r, 4).value
        if not pn or str(pn) in ("Part Number", ""):
            continue
        sol = _cell_str(ws.cell(r, 1).value)
        if sol in ("Solution", "CX PID", "Additional Details", "SERVICES", "SOFTWARE:") or sol.startswith("This is NOT"):
            continue
        desc = _cell_str(ws.cell(r, 5).value)
        ib = _float(ws.cell(r, ib_col).value)
        desired = _float(ws.cell(r, desired_col).value)
        tcv = _float(ws.cell(r, tcv_col).value)
        items.append(
            LineItem(
                solution=sol,
                suite=_cell_str(ws.cell(r, 2).value),
                part_number=_cell_str(pn),
                description=desc,
                ib=ib,
                desired=desired,
                tcv=tcv,
            )
        )
    return items


def _parse_proposal_summary(wb, data: EAData) -> None:
    if "Proposal Summary" not in wb.sheetnames:
        return
    ws = wb["Proposal Summary"]
    field_map = {
        "Report created on": "report_date",
        "Deal ID": "deal_id",
        "Quote ID": "quote_id",
        "Project Name": "project_name",
        "Buying Program ID": "buying_program_id",
        "Smart Account": "smart_account",
        "Proposal ID": "proposal_id",
        "Expected Book Date": "rsd",
        "Price List": "price_list",
        "Enterprise Agreement Motion": "ea_motion",
        "Proposal Link": "proposal_link",
    }
    for r in range(1, 40):
        key = ws.cell(r, 1).value
        if key in field_map:
            val = ws.cell(r, 2).value
            setattr(data, field_map[key], _cell_str(val) if key != "Proposal ID" else str(int(_float(val))))

    # End customer / partner from row 9
    customer_block = _cell_str(ws.cell(9, 5).value)
    if customer_block:
        lines = [ln.strip().lstrip("-").strip() for ln in customer_block.split("\n") if ln.strip()]
        for ln in lines:
            m = re.search(r"Account\s*:\s*(.+)", ln, re.I)
            if m:
                data.customer_name = m.group(1).strip()
            elif re.search(r"\d+.*(?:ST|AVE|WAY|BLVD|DR|RD)\b", ln, re.I) or re.match(r"\d", ln):
                data.address = ln
        if not data.customer_name and lines:
            data.customer_name = lines[-1].replace("Account :", "").strip()

    partner = _cell_str(ws.cell(9, 9).value)
    if partner and partner != "-":
        data.partner = partner

    am = _cell_str(ws.cell(17, 7).value)
    if am:
        data.account_manager = am

    # Financial summary
    for r in range(25, 35):
        label = _cell_str(ws.cell(r, 1).value)
        if label == "PORTFOLIO":
            continue
        if label == "TOTAL":
            data.total_software = _float(ws.cell(r, 2).value)
            data.total_services = _float(ws.cell(r, 3).value)
            data.total_tcv = _float(ws.cell(r, 4).value)
            continue
        if label in ("Networking Infrastructure", "Applications Infrastructure", "Security", "Services"):
            data.portfolios.append(
                {
                    "name": label,
                    "software": _float(ws.cell(r, 2).value),
                    "services": _float(ws.cell(r, 3).value),
                    "total": _float(ws.cell(r, 4).value),
                }
            )


def _parse_portfolio_sheet(wb, sheet_name: str, data: EAData) -> None:
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]

    for r in range(1, 20):
        label = _cell_str(ws.cell(r, 1).value)
        if label == "Duration (months)":
            data.term_months = int(_float(ws.cell(r, 2).value) or 36)
        elif label == "Billing Model":
            data.billing = _cell_str(ws.cell(r, 2).value) or data.billing
        elif label == "Requested Ship Date (RSD)" and not data.rsd:
            data.rsd = _cell_str(ws.cell(r, 2).value)
        elif "Reseller Services Bill to Name" in label:
            data.reseller = _cell_str(ws.cell(r, 2).value)
        elif label == "Portfolio Total Contract Value for Software (USD)":
            sw_tcv = _float(ws.cell(r, 2).value)
            if sheet_name == "Networking Infrastructure":
                data.networking_sw_tcv = sw_tcv
            elif sheet_name == "Applications Infrastructure":
                data.applications_sw_tcv = sw_tcv
            elif sheet_name == "Security":
                data.security_sw_tcv = sw_tcv
        elif label == "Portfolio Total Contract Value for Services (USD)":
            svc_tcv = _float(ws.cell(r, 2).value)
            if sheet_name == "Networking Infrastructure":
                data.networking_svc_tcv = svc_tcv
            elif sheet_name == "Security":
                data.security_svc_tcv = svc_tcv
        elif label == "Security Content User (SCU) Count":
            data.scu_count = int(_float(ws.cell(r, 2).value))

    headers: list[int] = []
    for r in range(15, 80):
        row_vals = [ws.cell(r, c).value for c in range(1, 6)]
        if row_vals[0] in ("SOFTWARE:", "SERVICES"):
            continue
        if row_vals[0] == "Part Number" or row_vals[3] == "Part Number":
            headers.append(r)

    sw_header = headers[0] if headers else None
    svc_header = headers[1] if len(headers) > 1 else None

    if sw_header:
        svc_start = (svc_header - 1) if svc_header else min(sw_header + 45, ws.max_row)
        sw_items = _parse_line_items(ws, sw_header + 1, svc_start, sw_header)
        if sheet_name == "Networking Infrastructure":
            data.networking_sw = [i for i in sw_items if i.tcv > 0 or i.part_number.startswith("E3")]
        elif sheet_name == "Applications Infrastructure":
            data.applications = [i for i in sw_items if i.tcv > 0]
        elif sheet_name == "Security":
            data.security_sw = [i for i in sw_items if i.tcv > 0]

    if svc_header:
        svc_items = _parse_line_items(ws, svc_header + 1, ws.max_row, svc_header)
        svc_items = [i for i in svc_items if i.tcv > 0]
        if sheet_name == "Networking Infrastructure":
            data.networking_svc = svc_items
        elif sheet_name == "Security":
            data.security_svc = svc_items
        elif sheet_name == "Collaboration":
            data.collaboration_svc = svc_items


def _build_support_breakdown(data: EAData) -> None:
    groups: dict[str, float] = {}
    for item in data.networking_svc:
        key = item.suite.replace("Services: ", "") if item.suite else item.description
        groups[f"Networking — {key}"] = groups.get(f"Networking — {key}", 0) + item.tcv
    for item in data.security_svc:
        groups["Security — ISE Support"] = groups.get("Security — ISE Support", 0) + item.tcv
    for item in data.collaboration_svc:
        groups["Collaboration — Employee Experience"] = groups.get("Collaboration — Employee Experience", 0) + item.tcv

    data.support_breakdown = sorted(
        [{"category": k, "tcv": v} for k, v in groups.items()],
        key=lambda x: -x["tcv"],
    )[:8]


def _build_meraki_summary(data: EAData) -> None:
    counts = {"AP": 0.0, "Switch": 0.0, "MX": 0.0, "Camera": 0.0}
    for item in data.networking_sw:
        pn = item.part_number
        d = item.desired or 0
        if "MR-E" in pn or "MR-A" in pn:
            counts["AP"] += d
        elif pn.startswith("E3N-MS"):
            counts["Switch"] += d
        elif "MX" in pn:
            counts["MX"] += d
        elif "MV" in pn:
            counts["Camera"] += d
    data.meraki = {k: int(v) for k, v in counts.items() if v > 0}


def _build_highlights(data: EAData) -> None:
    if not data.portfolios or not data.total_tcv:
        return
    chart_portfolios = [p for p in data.portfolios if p["total"] > 0]
    top = max(chart_portfolios, key=lambda p: p["total"])
    pct = int(100 * top["total"] / data.total_tcv)
    highlights = [f"{top['name']} is the largest portfolio at {pct}% of total TCV"]

    if data.meraki.get("AP"):
        highlights.append(
            f"Meraki cloud-managed networking — {data.meraki.get('AP', 0)} APs, "
            f"{data.meraki.get('Switch', 0)} switches, {data.meraki.get('MX', 0)} MX appliances"
        )
    else:
        highlights.append("Meraki + Catalyst/Nexus campus and data center modernization")

    dna = next((i for i in data.networking_sw if "AIRWLAN" in i.part_number), None)
    if dna:
        highlights.append(f"Cisco DNA Wireless — {int(dna.desired)} AP licenses (IB: {int(dna.ib)})")

    ise_adv = next((i for i in data.security_sw if "ISE-ADV" in i.part_number), None)
    if ise_adv:
        highlights.append(
            f"ISE Advantage — {int(ise_adv.desired)} licenses (IB: {int(ise_adv.ib)})"
        )

    te_users = next((i for i in data.applications if "TE-USERS" in i.part_number), None)
    if te_users:
        highlights.append(f"ThousandEyes End User Monitoring — {int(te_users.desired)} users")

    splunk = next((i for i in data.applications if "SK-SE" in i.part_number), None)
    if splunk:
        highlights.append(f"Splunk Cloud Platform — {int(splunk.desired)} GB/day ingestion")

    highlights.append(f"Comprehensive support portfolio — {_fmt_currency(data.total_services)} services TCV")
    data.highlights = highlights[:6]


def _build_out_of_scope(wb, data: EAData) -> None:
    oos = []
    checks = [
        ("Collaboration", "Collaboration (Webex, Calling, Contact Center) — software only; support services in scope"),
        ("Hybrid", "Hybrid Cloud"),
        ("Provider Connectivity", "Provider Connectivity"),
        ("Cisco Professional Services", "Cisco Professional Services"),
    ]
    for sheet, label in checks:
        if sheet not in wb.sheetnames:
            oos.append(label)
            continue
        ws = wb[sheet]
        sw_tcv = 0.0
        for r in range(1, 20):
            if "Software (USD)" in _cell_str(ws.cell(r, 1).value):
                sw_tcv = _float(ws.cell(r, 2).value)
                break
        if sw_tcv == 0 and sheet == "Collaboration" and data.collaboration_svc:
            continue
        if sw_tcv == 0:
            oos.append(label)
    data.out_of_scope = oos


def _fmt_currency(value: float) -> str:
    return f"${value:,.0f}" if value else "—"


def parse_eamp_excel(path: Path) -> EAData:
    wb = load_workbook(path, data_only=True)
    data = EAData()

    if not data.customer_name:
        data.customer_name = path.stem.split(" - EAMP")[0].replace("_", " ").strip()

    _parse_proposal_summary(wb, data)
    for sheet in ("Networking Infrastructure", "Applications Infrastructure", "Security", "Collaboration"):
        _parse_portfolio_sheet(wb, sheet, data)

    _build_meraki_summary(data)
    _build_support_breakdown(data)
    _build_highlights(data)
    _build_out_of_scope(wb, data)

    if not data.proposal_id:
        m = re.search(r"ID[_\s]*(\d+)", path.name, re.I)
        if m:
            data.proposal_id = m.group(1)

    return data
