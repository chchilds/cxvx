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
    commitment: str = ""
    list_price: float = 0.0
    sub_disc: float = 0.0
    prog_disc: float = 0.0
    unit_list: float = 0.0


@dataclass
class EAData:
    customer_name: str = ""
    account: str = ""
    proposal_id: str = ""
    proposal_name: str = ""
    project_name: str = ""
    report_date: str = ""
    rsd: str = ""
    term_months: float = 36.0
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
    ea_motion: str = "EA"
    proposal_link: str = ""
    total_tcv: float = 0.0
    total_software: float = 0.0
    total_services: float = 0.0
    list_total: float = 0.0
    discount_total: float = 0.0
    otd_total: float = 0.0
    portfolios: list[dict] = field(default_factory=list)
    networking_sw: list[LineItem] = field(default_factory=list)
    networking_svc: list[LineItem] = field(default_factory=list)
    applications: list[LineItem] = field(default_factory=list)
    applications_svc: list[LineItem] = field(default_factory=list)
    security_sw: list[LineItem] = field(default_factory=list)
    security_svc: list[LineItem] = field(default_factory=list)
    collaboration_sw: list[LineItem] = field(default_factory=list)
    collaboration_svc: list[LineItem] = field(default_factory=list)
    support_breakdown: list[dict] = field(default_factory=list)
    highlights: list[str] = field(default_factory=list)
    out_of_scope: list[str] = field(default_factory=list)
    networking_sw_tcv: float = 0.0
    networking_svc_tcv: float = 0.0
    applications_sw_tcv: float = 0.0
    applications_svc_tcv: float = 0.0
    security_sw_tcv: float = 0.0
    security_svc_tcv: float = 0.0
    collaboration_sw_tcv: float = 0.0
    collaboration_svc_tcv: float = 0.0
    scu_count: int = 0
    meraki: dict = field(default_factory=dict)
    source_file: str = ""


def _cell_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _float(value) -> float:
    try:
        if isinstance(value, str):
            value = value.replace(",", "").replace("$", "").replace("%", "").strip()
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


def _header_map(ws, header_row: int) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for c in range(1, 70):
        h = _cell_str(ws.cell(header_row, c).value).lower()
        if h:
            mapping[h] = c
    return mapping


def _col(mapping: dict[str, int], *needles: str, default: int = 0) -> int:
    for needle in needles:
        needle = needle.lower()
        for key, col in mapping.items():
            if needle in key:
                return col
    return default


def _parse_line_items(ws, start_row: int, end_row: int, header_row: int) -> list[LineItem]:
    mapping = _header_map(ws, header_row)
    pn_col = _col(mapping, "part number", default=4)
    sol_col = _col(mapping, "solution", "cx pid", default=1)
    suite_col = _col(mapping, "suite", default=2)
    commit_col = _col(mapping, "commitment", default=3)
    desc_col = _col(mapping, "description", default=5)
    ib_col = _col(mapping, "quantity found", default=7)
    desired_col = _col(mapping, "quantity desired", default=8)
    tcv_col = _col(mapping, "total contract value", default=21)
    list_col = _col(mapping, "total list", default=11)
    sub_col = _col(mapping, "subscription disc", default=12)
    prog_col = _col(mapping, "programmatic disc", default=9)
    unit_col = _col(mapping, "unit list", default=10)

    skip_sol = {
        "solution",
        "cx pid",
        "additional details",
        "services",
        "software:",
        "software",
        "part number",
    }
    items: list[LineItem] = []
    for r in range(start_row, end_row + 1):
        pn = _cell_str(ws.cell(r, pn_col).value)
        if not pn or pn.lower() == "part number":
            continue
        sol = _cell_str(ws.cell(r, sol_col).value)
        if sol.lower() in skip_sol or sol.startswith("This is NOT"):
            continue
        items.append(
            LineItem(
                solution=sol,
                suite=_cell_str(ws.cell(r, suite_col).value) if suite_col else "",
                part_number=pn,
                description=_cell_str(ws.cell(r, desc_col).value) if desc_col else "",
                ib=_float(ws.cell(r, ib_col).value) if ib_col else 0.0,
                desired=_float(ws.cell(r, desired_col).value) if desired_col else 0.0,
                tcv=_float(ws.cell(r, tcv_col).value) if tcv_col else 0.0,
                commitment=_cell_str(ws.cell(r, commit_col).value) if commit_col else "",
                list_price=_float(ws.cell(r, list_col).value) if list_col else 0.0,
                sub_disc=_float(ws.cell(r, sub_col).value) if sub_col else 0.0,
                prog_disc=_float(ws.cell(r, prog_col).value) if prog_col else 0.0,
                unit_list=_float(ws.cell(r, unit_col).value) if unit_col else 0.0,
            )
        )
    return items


def _parse_proposal_summary(wb, data: EAData) -> None:
    if "Proposal Summary" not in wb.sheetnames:
        return
    ws = wb["Proposal Summary"]
    field_map = {
        "report created on": "report_date",
        "deal id": "deal_id",
        "quote id": "quote_id",
        "project name": "project_name",
        "buying program id": "buying_program_id",
        "smart account": "smart_account",
        "proposal id": "proposal_id",
        "proposal name": "proposal_name",
        "expected book date": "rsd",
        "price list": "price_list",
        "enterprise agreement motion": "ea_motion",
        "proposal link": "proposal_link",
        "cisco account manager": "account_manager",
        "account manager": "account_manager",
    }
    for r in range(1, 45):
        for c in range(1, 12):
            key = _cell_str(ws.cell(r, c).value)
            key_l = key.lower().rstrip(":")
            if key_l in field_map:
                val = ws.cell(r, c + 1).value
                attr = field_map[key_l]
                if attr == "proposal_id":
                    data.proposal_id = str(int(_float(val))) if _float(val) else _cell_str(val)
                else:
                    text = _cell_str(val)
                    if text and text != "-":
                        setattr(data, attr, text)
            if key_l in ("partner", "partner name") and not data.partner:
                partner = _cell_str(ws.cell(r, c + 1).value)
                if partner and partner != "-":
                    data.partner = partner

    # End-customer block is often a wrapped cell rather than a label/value pair.
    for r in range(1, 25):
        for c in range(1, 12):
            block = _cell_str(ws.cell(r, c).value)
            if "Account :" in block or "Account:" in block or (
                "HALIFAX" in block.upper() and "\n" in block
            ):
                lines = [ln.strip().lstrip("-").strip() for ln in block.split("\n") if ln.strip()]
                for ln in lines:
                    m = re.search(r"Account\s*:\s*(.+)", ln, re.I)
                    if m:
                        data.account = m.group(1).strip()
                    elif re.search(r"\d+.*(?:ST|AVE|WAY|BLVD|DR|RD|TRAFFIC)\b", ln, re.I) or re.match(
                        r"\d", ln
                    ):
                        data.address = ln
                    elif "INC" in ln.upper() or "LLC" in ln.upper() or "HEALTH" in ln.upper():
                        if not data.customer_name or len(ln) > len(data.customer_name):
                            if "account" not in ln.lower():
                                data.customer_name = ln.replace("Account :", "").strip()
            if not data.partner:
                val = _cell_str(ws.cell(r, c).value)
                if val and val not in ("-", "Partner") and c >= 8 and r <= 12:
                    if val.upper() in ("CDW", "WWT") or "TECHNOLOGY" in val.upper() or "LOGISTICS" in val.upper():
                        data.partner = val

    # Financial summary: find PORTFOLIO header, then rows until TOTAL.
    header_row = None
    for r in range(1, 50):
        if _cell_str(ws.cell(r, 1).value).upper() == "PORTFOLIO":
            header_row = r
            break
    if header_row:
        mapping = _header_map(ws, header_row)
        sw_col = _col(mapping, "software", default=2)
        svc_col = _col(mapping, "services", default=3)
        tot_col = _col(mapping, "total", default=4)
        list_col = _col(mapping, "list", default=5)
        disc_col = _col(mapping, "discount", default=6)
        otd_col = _col(mapping, "one time", "one-time", "otd", default=7)
        known = {
            "networking infrastructure",
            "applications infrastructure",
            "security",
            "services",
            "collaboration",
            "hybrid",
            "provider connectivity",
            "cisco professional services",
        }
        data.portfolios = []
        for r in range(header_row + 1, header_row + 20):
            label = _cell_str(ws.cell(r, 1).value)
            if not label:
                continue
            software = _float(ws.cell(r, sw_col).value)
            services = _float(ws.cell(r, svc_col).value)
            total = _float(ws.cell(r, tot_col).value)
            list_price = _float(ws.cell(r, list_col).value) if list_col else 0.0
            discount = _float(ws.cell(r, disc_col).value) if disc_col else 0.0
            otd = _float(ws.cell(r, otd_col).value) if otd_col else 0.0
            if label.upper() == "TOTAL":
                data.total_software = software
                data.total_services = services
                data.total_tcv = total
                data.list_total = list_price
                data.discount_total = discount
                data.otd_total = otd
                break
            if label.lower() in known or total or software or services:
                data.portfolios.append(
                    {
                        "name": label,
                        "software": software,
                        "services": services,
                        "total": total,
                        "list": list_price,
                        "discount": discount,
                        "otd": otd,
                    }
                )


def _parse_portfolio_sheet(wb, sheet_name: str, data: EAData) -> None:
    if sheet_name not in wb.sheetnames:
        return
    ws = wb[sheet_name]

    sw_tcv = 0.0
    svc_tcv = 0.0
    for r in range(1, 25):
        label = _cell_str(ws.cell(r, 1).value)
        if label == "Duration (months)":
            months = _float(ws.cell(r, 2).value)
            if months:
                data.term_months = months
        elif label == "Billing Model":
            billing = _cell_str(ws.cell(r, 2).value)
            if billing:
                data.billing = billing
        elif label == "Requested Ship Date (RSD)":
            rsd = _cell_str(ws.cell(r, 2).value)
            if rsd:
                data.rsd = rsd
        elif "Reseller Services Bill to Name" in label:
            reseller = _cell_str(ws.cell(r, 2).value)
            if reseller and reseller not in ("-", "N/A"):
                data.reseller = reseller
        elif "Software (USD)" in label and "Services" not in label:
            sw_tcv = _float(ws.cell(r, 2).value)
        elif "Services (USD)" in label:
            svc_tcv = _float(ws.cell(r, 2).value)
        elif label == "Security Content User (SCU) Count":
            data.scu_count = int(_float(ws.cell(r, 2).value))

    headers: list[int] = []
    max_r = min(ws.max_row or 80, 200)
    for r in range(12, max_r + 1):
        row_vals = [_cell_str(ws.cell(r, c).value) for c in range(1, 6)]
        if row_vals[0] in ("SOFTWARE:", "SERVICES"):
            continue
        if "Part Number" in row_vals:
            headers.append(r)

    sw_header = headers[0] if headers else None
    svc_header = headers[1] if len(headers) > 1 else None
    sw_items: list[LineItem] = []
    svc_items: list[LineItem] = []

    if sw_header:
        end = (svc_header - 1) if svc_header else max_r
        sw_items = _parse_line_items(ws, sw_header + 1, end, sw_header)
    if svc_header:
        svc_items = [i for i in _parse_line_items(ws, svc_header + 1, max_r, svc_header) if i.tcv > 0 or i.desired]

    def _keep_sw(items: list[LineItem]) -> list[LineItem]:
        return [i for i in items if i.tcv > 0 or i.desired or i.part_number.startswith("E3")]

    if sheet_name == "Networking Infrastructure":
        data.networking_sw_tcv = sw_tcv
        data.networking_svc_tcv = svc_tcv
        data.networking_sw = _keep_sw(sw_items)
        data.networking_svc = svc_items
    elif sheet_name == "Applications Infrastructure":
        data.applications_sw_tcv = sw_tcv
        data.applications_svc_tcv = svc_tcv
        data.applications = _keep_sw(sw_items)
        data.applications_svc = svc_items
    elif sheet_name == "Security":
        data.security_sw_tcv = sw_tcv
        data.security_svc_tcv = svc_tcv
        data.security_sw = _keep_sw(sw_items)
        data.security_svc = svc_items
    elif sheet_name == "Collaboration":
        data.collaboration_sw_tcv = sw_tcv
        data.collaboration_svc_tcv = svc_tcv
        data.collaboration_sw = _keep_sw(sw_items)
        data.collaboration_svc = svc_items


def _build_support_breakdown(data: EAData) -> None:
    groups: dict[str, float] = {}
    buckets = [
        ("Networking", data.networking_svc),
        ("Applications", data.applications_svc),
        ("Security", data.security_svc),
        ("Collaboration", data.collaboration_svc),
    ]
    for portfolio, items in buckets:
        for item in items:
            key = item.suite.replace("Services: ", "") if item.suite else item.description
            label = f"{portfolio} — {key or item.part_number}"
            groups[label] = groups.get(label, 0) + item.tcv
    data.support_breakdown = sorted(
        [{"category": k, "tcv": v} for k, v in groups.items() if v],
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
    highlights: list[str] = []
    chart_portfolios = [p for p in data.portfolios if p.get("total", 0) > 0]
    if chart_portfolios and data.total_tcv:
        top = max(chart_portfolios, key=lambda p: p["total"])
        pct = int(100 * top["total"] / data.total_tcv)
        highlights.append(f"{top['name']} is the largest portfolio at {pct}% of Cisco net TCV")

    if data.list_total and data.total_tcv:
        savings = data.list_total - data.total_tcv
        pct = 100 * savings / data.list_total
        highlights.append(f"EA discounts take {pct:.1f}% off list ({_fmt_currency(savings)} vs. list)")

    if data.meraki.get("AP") or data.meraki.get("Switch"):
        highlights.append(
            f"Meraki cloud-managed networking — {data.meraki.get('AP', 0)} APs, "
            f"{data.meraki.get('Switch', 0)} switches, {data.meraki.get('MX', 0)} MX appliances"
        )

    dna = next((i for i in data.networking_sw if "AIRWLAN" in i.part_number or "DNA" in i.part_number), None)
    if dna and dna.desired:
        highlights.append(f"{dna.description[:48]} — {int(dna.desired)} desired (IB: {int(dna.ib)})")

    ise_bits = [
        f"{item.description.split(',')[0]} {int(item.desired):,}"
        for item in data.security_sw
        if "ISE" in item.part_number and item.desired and item.tcv > 0
    ]
    if ise_bits:
        highlights.append("ISE — " + " / ".join(ise_bits[:3]))

    splunk = next((i for i in data.applications if "SK-SE" in i.part_number or "SK-ES" in i.part_number), None)
    if splunk and splunk.desired:
        highlights.append(f"{splunk.description[:48]} — {int(splunk.desired)} GB/day")

    te_users = next((i for i in data.applications if "TE-USERS" in i.part_number), None)
    if te_users:
        highlights.append(f"ThousandEyes End User Monitoring — {int(te_users.desired)} users")

    if data.total_services:
        highlights.append(f"Cisco CX / support TCV {_fmt_currency(data.total_services)}")

    data.highlights = highlights[:6]


def _build_out_of_scope(wb, data: EAData) -> None:
    oos = []
    checks = [
        ("Collaboration", "Collaboration (Webex, Calling, Contact Center)"),
        ("Hybrid", "Hybrid Cloud"),
        ("Provider Connectivity", "Provider Connectivity"),
        ("Cisco Professional Services", "Cisco Professional Services"),
        ("Networking Infrastructure", "Networking Infrastructure"),
        ("Applications Infrastructure", "Applications Infrastructure"),
        ("Security", "Security"),
    ]
    populated = {p["name"] for p in data.portfolios if p.get("total", 0) > 0}
    for sheet, label in checks:
        if label.split(" (")[0] in populated or sheet in populated:
            continue
        if sheet not in wb.sheetnames:
            oos.append(label)
            continue
        ws = wb[sheet]
        tcv = 0.0
        for r in range(1, 25):
            cell = _cell_str(ws.cell(r, 1).value)
            if "Total Contract Value" in cell and "USD" in cell:
                tcv += _float(ws.cell(r, 2).value)
        if sheet == "Collaboration" and data.collaboration_svc:
            continue
        if tcv == 0:
            oos.append(label)
    data.out_of_scope = oos


def _fmt_currency(value: float) -> str:
    return f"${value:,.0f}" if value else "—"


def parse_eamp_excel(path: Path) -> EAData:
    wb = load_workbook(path, data_only=True)
    data = EAData(source_file=path.name)

    stem = path.stem.split(" - EAMP")[0].replace("_", " ").strip()
    data.customer_name = re.sub(r"\s+", " ", stem)

    _parse_proposal_summary(wb, data)
    for sheet in ("Networking Infrastructure", "Applications Infrastructure", "Security", "Collaboration"):
        _parse_portfolio_sheet(wb, sheet, data)

    _build_meraki_summary(data)
    _build_support_breakdown(data)
    _build_highlights(data)
    _build_out_of_scope(wb, data)

    if not data.proposal_id:
        m = re.search(r"ID[_\s-]*(\d+)", path.name, re.I)
        if m:
            data.proposal_id = m.group(1)

    if not data.smart_account:
        data.smart_account = data.customer_name

    return data
