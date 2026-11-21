"""Parse a Cisco EAMP Price Estimate workbook into structured data."""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
R_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

SKIP_VALUES = {"", None, "{"}


def _col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _cell_ref(ref: str) -> tuple[str, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        raise ValueError(f"Bad cell ref: {ref}")
    return match.group(1), int(match.group(2))


def _as_number(value: Any) -> Any:
    if value is None or isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if text in ("", "-", "N/A", "NA"):
        return text if text in ("-", "N/A", "NA") else None
    try:
        if "." in text:
            return float(text)
        return int(text)
    except ValueError:
        return value


def _is_placeholder(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    return value.startswith("{") and ("string:" in value or "object:" in value or "array[" in value)


def _clean(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
        if _is_placeholder(text) or text == "":
            return None
        return text
    return value


def _load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall("m:si", NS):
        strings.append("".join(t.text or "" for t in si.findall(".//m:t", NS)))
    return strings


def _sheet_cells(zf: zipfile.ZipFile, target: str, strings: list[str]) -> dict[tuple[int, str], Any]:
    root = ET.fromstring(zf.read(target))
    cells: dict[tuple[int, str], Any] = {}
    for node in root.findall(".//m:c", NS):
        ref = node.attrib.get("r")
        if not ref:
            continue
        cell_type = node.attrib.get("t")
        v_node = node.find("m:v", NS)
        is_node = node.find("m:is", NS)
        if v_node is not None and v_node.text is not None:
            raw = v_node.text
            if cell_type == "s":
                value: Any = strings[int(raw)]
            elif cell_type == "b":
                value = raw == "1"
            else:
                value = _as_number(raw)
        elif is_node is not None:
            value = "".join(t.text or "" for t in is_node.findall(".//m:t", NS))
        else:
            continue
        col, row = _cell_ref(ref)
        cells[(row, col)] = value
    return cells


def _get(cells: dict[tuple[int, str], Any], row: int, col: str, default: Any = None) -> Any:
    return _clean(cells.get((row, col), default))


def _row_map(cells: dict[tuple[int, str], Any], row: int) -> dict[str, Any]:
    mapped: dict[str, Any] = {}
    for (r, col), value in cells.items():
        if r == row:
            mapped[col] = _clean(value)
    return mapped


def _max_row(cells: dict[tuple[int, str], Any]) -> int:
    return max((r for r, _ in cells), default=0)


def _find_row(cells: dict[tuple[int, str], Any], col: str, text: str, start: int = 1) -> int | None:
    needle = text.strip().lower()
    for row in range(start, _max_row(cells) + 1):
        value = _get(cells, row, col)
        if isinstance(value, str) and value.strip().lower() == needle:
            return row
    return None


def _find_row_startswith(cells: dict[tuple[int, str], Any], col: str, prefix: str, start: int = 1) -> int | None:
    needle = prefix.strip().lower()
    for row in range(start, _max_row(cells) + 1):
        value = _get(cells, row, col)
        if isinstance(value, str) and value.strip().lower().startswith(needle):
            return row
    return None


def _header_lookup(header_row: dict[str, Any]) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for col, value in header_row.items():
        if isinstance(value, str) and value.strip():
            lookup[value.strip().lower()] = col
    return lookup


def _normalize_header(name: str) -> str:
    return re.sub(r"\s+", " ", name.strip().lower())


HEADER_ALIASES = {
    "solution": "solution",
    "suite": "suite",
    "commitment": "commitment",
    "part number": "part_number",
    "description": "description",
    "license tiers": "license_tier",
    "quantity found (ib)": "qty_ib",
    "quantity desired": "qty_desired",
    "programmatic discounts (%)": "programmatic_discount_pct",
    "unit list price per month": "unit_list_month",
    "total list price (usd)": "list_price",
    "total list price  (usd)": "list_price",
    "subscription discount (%)": "subscription_discount_pct",
    "surcharge (%)": "surcharge_pct",
    "surcharge amount": "surcharge_amount",
    "otd (ib adjustment)": "otd_ib",
    "ota (strategic offer)": "ota_strategic",
    "otd (ramp)": "otd_ramp",
    "otd (competitive)": "otd_competitive",
    "otd (one time adjustment)": "otd_other",
    "otd (migration adjustment)": "otd_migration",
    "total contract value (usd)": "tcv",
    "service tier": "service_tier",
    "otd (program migration incentive)": "otd_program_migration",
    "otd (uncovered asset credit)": "otd_uncovered",
    "otd (other)": "otd_other",
    "pre-ea assets": "pre_ea_assets",
    "pcm provider": "pcm_provider",
    "pcm program": "pcm_program",
}


def _map_headers(header_row: dict[str, Any]) -> dict[str, str]:
    mapped: dict[str, str] = {}
    for col, value in header_row.items():
        if not isinstance(value, str):
            continue
        key = HEADER_ALIASES.get(_normalize_header(value))
        if key:
            mapped[key] = col
    return mapped


def _row_to_item(cells: dict[tuple[int, str], Any], row: int, field_cols: dict[str, str]) -> dict[str, Any] | None:
    item: dict[str, Any] = {}
    for field, col in field_cols.items():
        item[field] = _get(cells, row, col)
    if not item.get("part_number") and not item.get("description") and not item.get("suite"):
        return None
    if all(v in (None, "", 0, 0.0) for v in item.values()):
        return None
    return item


def _section_items(
    cells: dict[tuple[int, str], Any],
    header_row_num: int,
    stop_prefixes: tuple[str, ...],
) -> list[dict[str, Any]]:
    field_cols = _map_headers(_row_map(cells, header_row_num))
    items: list[dict[str, Any]] = []
    for row in range(header_row_num + 1, _max_row(cells) + 1):
        a_val = _get(cells, row, "A")
        if isinstance(a_val, str):
            lowered = a_val.strip().lower()
            if any(lowered.startswith(prefix) for prefix in stop_prefixes):
                break
            if lowered.startswith("this is not an approved"):
                break
        item = _row_to_item(cells, row, field_cols)
        if item:
            items.append(item)
    return items


def _num(value: Any, default: float = 0.0) -> float:
    if isinstance(value, (int, float)):
        return float(value)
    return default


def parse_proposal_summary(cells: dict[tuple[int, str], Any]) -> dict[str, Any]:
    meta_map = {
        "deal_id": (9, "B"),
        "end_customer": (9, "E"),
        "partner": (9, "I"),
        "disti": (9, "M"),
        "quote_id": (10, "B"),
        "project_name": (11, "B"),
        "buying_program_id": (12, "B"),
        "smart_account": (13, "B"),
        "proposal_name": (14, "B"),
        "proposal_id": (15, "B"),
        "expected_book_date": (16, "B"),
        "price_list": (17, "B"),
        "account_manager": (17, "G"),
        "ea_motion": (18, "B"),
        "entitlement": (19, "B"),
        "proposal_link": (23, "B"),
        "report_created_on": (8, "B"),
    }
    meta = {key: _get(cells, row, col) for key, (row, col) in meta_map.items()}
    if isinstance(meta.get("proposal_id"), float) and meta["proposal_id"].is_integer():
        meta["proposal_id"] = str(int(meta["proposal_id"]))
    elif meta.get("proposal_id") is not None:
        meta["proposal_id"] = str(meta["proposal_id"]).replace(".0", "")

    customer = meta.get("end_customer") or ""
    customer_lines = [line.strip() for line in str(customer).split("\n") if line.strip() and line.strip() != "-"]
    meta["end_customer_raw"] = customer
    meta["end_customer_address_lines"] = customer_lines
    if customer_lines:
        last = customer_lines[-1]
        if last.lower().startswith("account"):
            meta["end_customer_account"] = last.split(":", 1)[-1].strip()
            meta["end_customer_address"] = "\n".join(customer_lines[:-1])
        else:
            meta["end_customer_account"] = None
            meta["end_customer_address"] = "\n".join(customer_lines)
    else:
        meta["end_customer_account"] = None
        meta["end_customer_address"] = None

    header_row = 26
    financial: list[dict[str, Any]] = []
    for row in range(27, 35):
        portfolio = _get(cells, row, "A")
        if not portfolio:
            break
        financial.append(
            {
                "portfolio": portfolio,
                "software": _num(_get(cells, row, "B")),
                "services": _num(_get(cells, row, "C")),
                "total": _num(_get(cells, row, "D")),
                "list_price": _num(_get(cells, row, "E")),
                "discount_amount": _num(_get(cells, row, "F")),
                "one_time_discount": _num(_get(cells, row, "G")),
            }
        )
    return {"meta": meta, "financial_summary": financial, "header_row": header_row}


def parse_portfolio_sheet(name: str, cells: dict[tuple[int, str], Any]) -> dict[str, Any]:
    header_fields = {
        "requested_ship_date": "Requested Ship Date",
        "billing_model": "Billing Model",
        "capital_financing_frequency": "Capital Financing Frequency",
        "duration_months": "Duration (months)",
        "distributor_services_bill_to": "Distributor Services Bill to Name",
        "reseller_services_bill_to": "Reseller Services Bill to Name",
        "price_list": "Price List",
        "software_tcv": "Portfolio Total Contract Value for Software",
        "services_tcv": "Portfolio Total Contract Value for Services",
        "scu_count": "Security Content User",
    }
    header: dict[str, Any] = {"sheet_name": name, "title": _get(cells, 7, "A") or name}
    for row in range(8, 20):
        label = _get(cells, row, "A")
        if not isinstance(label, str):
            continue
        for key, prefix in header_fields.items():
            if label.lower().startswith(prefix.lower()):
                header[key] = _get(cells, row, "B")
    software_row = _find_row(cells, "A", "SOFTWARE:")
    services_row = _find_row(cells, "A", "SERVICES")
    software: list[dict[str, Any]] = []
    services: list[dict[str, Any]] = []
    if software_row:
        software = _section_items(
            cells,
            software_row + 1,
            stop_prefixes=("services", "additional details", "*license"),
        )
    if services_row:
        services = _section_items(
            cells,
            services_row + 1,
            stop_prefixes=("additional details", "*license", "this is not"),
        )
    header["software_line_tcv"] = round(sum(_num(i.get("tcv")) for i in software), 2)
    header["services_line_tcv"] = round(sum(_num(i.get("tcv")) for i in services), 2)
    populated = bool(software or services or header.get("software_tcv") or header.get("services_tcv"))
    return {
        "name": name,
        "header": header,
        "software": software,
        "services": services,
        "populated": populated,
    }


def parse_eamp_workbook(path: str | Path) -> dict[str, Any]:
    path = Path(path)
    with zipfile.ZipFile(path) as zf:
        strings = _load_shared_strings(zf)
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets: list[tuple[str, str]] = []
        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            rid = sheet.attrib[R_ATTR]
            target = rid_to_target[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target
            sheets.append((sheet.attrib["name"], target))

        summary_cells = None
        portfolios: list[dict[str, Any]] = []
        for name, target in sheets:
            cells = _sheet_cells(zf, target, strings)
            if name == "Proposal Summary":
                summary_cells = cells
            else:
                portfolios.append(parse_portfolio_sheet(name, cells))

    if summary_cells is None:
        raise ValueError("Proposal Summary sheet not found")
    summary = parse_proposal_summary(summary_cells)
    return {
        "source_file": path.name,
        "meta": summary["meta"],
        "financial_summary": summary["financial_summary"],
        "portfolios": portfolios,
    }


def suite_rollups(items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[tuple[str, str], dict[str, Any]] = {}
    order: list[tuple[str, str]] = []
    for item in items:
        key = (str(item.get("solution") or ""), str(item.get("suite") or ""))
        if key not in groups:
            groups[key] = {
                "solution": key[0],
                "suite": key[1],
                "commitment": item.get("commitment"),
                "lines": 0,
                "qty_ib": 0.0,
                "qty_desired": 0.0,
                "list_price": 0.0,
                "tcv": 0.0,
                "otd": 0.0,
            }
            order.append(key)
        bucket = groups[key]
        bucket["lines"] += 1
        # Skip $0 embed/support PIDs so suite quantities reflect licensed assets.
        if _num(item.get("tcv")) or _num(item.get("list_price")) or _num(item.get("unit_list_month")):
            bucket["qty_ib"] += _num(item.get("qty_ib"))
            bucket["qty_desired"] += _num(item.get("qty_desired"))
        bucket["list_price"] += _num(item.get("list_price"))
        bucket["tcv"] += _num(item.get("tcv"))
        bucket["otd"] += sum(
            _num(item.get(field))
            for field in (
                "otd_ib",
                "ota_strategic",
                "otd_ramp",
                "otd_competitive",
                "otd_other",
                "otd_migration",
                "otd_program_migration",
                "otd_uncovered",
            )
        )
        if item.get("commitment") and item.get("commitment") != bucket.get("commitment"):
            bucket["commitment"] = "Mixed"
    rolled = []
    for key in order:
        bucket = groups[key]
        for numeric in ("qty_ib", "qty_desired", "list_price", "tcv", "otd"):
            bucket[numeric] = round(bucket[numeric], 2)
        rolled.append(bucket)
    return rolled
