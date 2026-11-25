#!/usr/bin/env python3
"""Build the Tallgrass H1 Services Remediation Action Plan workbook."""

from __future__ import annotations

import argparse
import re
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel
import xlsxwriter

AS_OF = datetime(2026, 11, 25)
AS_OF_LABEL = "November 25, 2026 7:38 AM EST"
REGION = "Tallgrass"
TERRITORY = "TALLGRASS REGION"
ACTIVE_FC = {"Commit", "Most Likely", "Upside"}
H1_QUARTERS = {1, 2}

# Theme
NAVY = "#0B2545"
NAVY2 = "#163A5F"
TEAL = "#0E7C7B"
GOLD = "#C4A35A"
WHITE = "#FFFFFF"
OFFWHITE = "#F7F9FC"
ALT = "#EEF3F8"
LINE = "#D5DEE8"
MUTED = "#5B6B7C"
P1 = "#B42318"
P2 = "#B54708"
P3 = "#875A00"
COMMIT = "#0E7C7B"
ML = "#1D4E89"
UPSIDE = "#7A5C00"
GAP_BG = "#FDECEC"
OK_BG = "#E7F6EE"
KPI_NAVY = "#0B2545"
KPI_TEAL = "#0E5F5E"
KPI_GOLD = "#8A6A2F"
KPI_RED = "#8F2D24"

MASTER_HEADERS = [
    "Priority",
    "Rank",
    "Opportunity Owner",
    "AE Tab",
    "Account Name",
    "Opportunity Name",
    "Deal ID",
    "Forecast Category",
    "Stage",
    "Close Date",
    "Days to Close",
    "Fiscal Quarter",
    "Fiscal Year",
    "Quarter Label",
    "Product TCV ($)",
    "Service TCV ($)",
    "Attach Rate %",
    "5% Basic Target ($)",
    "5% Gap ($)",
    "12% Standard Target ($)",
    "12% Gap ($)",
    "Expected Service Amount CCW ($)",
    "CCW vs Service TCV ($)",
    "Age (Days)",
    "Forecasting Position",
    "Recommended Action",
]


def parse_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if hasattr(value, "year"):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, (int, float)) and value > 20000:
        try:
            dt = from_excel(value)
            return datetime(dt.year, dt.month, dt.day)
        except Exception:
            pass
    text = str(value).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def to_float(value) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def money(n: float) -> str:
    n = float(n)
    sign = "-" if n < 0 else ""
    n = abs(n)
    if n >= 1_000_000:
        return f"{sign}${n / 1_000_000:.2f}M"
    if n >= 1_000:
        text = f"{n / 1_000:.1f}"
        if text.endswith(".0"):
            text = text[:-2]
        return f"{sign}${text}K"
    return f"{sign}${n:,.0f}"


def pct(n: float) -> str:
    return f"{n * 100:.2f}%"


def ae_tab_name(owner: str) -> str:
    parts = [p for p in str(owner).split() if p]
    if not parts:
        return "Unknown"
    first = parts[0]
    last_initial = parts[-1][0].upper()
    return f"{first} {last_initial}"


def safe_sheet_name(name: str) -> str:
    cleaned = re.sub(r'[:\\/?*\[\]]', " ", name).strip()
    return cleaned[:31] or "AE"


def load_source(path: Path) -> tuple[list[dict], dict]:
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    header_row = 14
    headers = [c.value for c in next(ws.iter_rows(min_row=header_row, max_row=header_row))]

    def col(name: str) -> int:
        return headers.index(name)

    raw_rows = []
    for row in ws.iter_rows(min_row=header_row + 1, max_row=ws.max_row, values_only=True):
        if row[col("Deal ID")] is None:
            continue
        if row[col("Opportunity Name")] in (None, "Total"):
            continue
        raw_rows.append(row)

    funnel = Counter()
    deals = []
    for row in raw_rows:
        funnel["source_open_opps"] += 1
        name = str(row[col("Opportunity Name")] or "").strip()
        typ = str(row[col("Type")] or "").strip()
        fc = row[col("Forecast Status")]
        fq_raw = row[col("Fiscal Quarter")]
        fy = str(row[col("Fiscal Year")] or "").strip()
        try:
            fq = int(float(fq_raw))
        except (TypeError, ValueError):
            funnel["invalid_quarter"] += 1
            continue

        product = to_float(row[col("Product Amount (TCV) ($000s)")]) * 1000.0
        service = to_float(row[col("Service Amount (TCV) ($000s)")]) * 1000.0
        svc_acv = to_float(row[col("Total ACV Service TS-S Annual ($000s)")]) * 1000.0
        ccw = to_float(row[col("Expected Service Amount CCW")])

        is_renewal = name.upper().startswith("RNW") or typ.lower() == "renewal"
        if is_renewal:
            funnel["excluded_renewal"] += 1
            continue
        funnel["non_renewal"] += 1

        if fq not in H1_QUARTERS:
            funnel["excluded_not_h1"] += 1
            continue
        funnel["h1"] += 1

        if fc not in ACTIVE_FC:
            funnel["excluded_inactive_fc"] += 1
            continue
        funnel["active_fc"] += 1

        if product < 10_000 or service >= 10_000:
            funnel["excluded_not_low_services"] += 1
            continue
        funnel["cohort"] += 1

        denom = product + service
        attach = (service / denom) if denom else 0.0
        t5 = product * 0.05
        t12 = product * 0.12
        g5 = max(0.0, t5 - service)
        g12 = max(0.0, t12 - service)
        close_dt = parse_date(row[col("Close Date")])
        created_dt = parse_date(row[col("Created Date")])
        days_to_close = (close_dt - AS_OF).days if close_dt else None
        if service <= 0:
            priority = "P1 - Zero Services"
            action = (
                "Insert a services SKU on this order before the next forecast lock. "
                "Quote the 5% Basic floor immediately and package 12% Standard "
                "(CX / TAS / Solution Support / adoption) for the same close date."
            )
        elif g5 > 0:
            priority = "P2 - Below 5% Floor"
            action = (
                "Uplift services to the 5% Basic floor this week, then expand the "
                "same quote to the 12% Standard package before Commit upgrade."
            )
        else:
            priority = "P3 - Below 12% Standard"
            action = (
                "Hold the 5% floor already on the deal and expand to 12% Standard "
                "(adoption, TAS, migration, CX Preferred) on the same close date."
            )
        if ccw > service + 1:
            action += (
                f" CCW expected services ({money(ccw)}) exceed Salesforce Service TCV "
                f"({money(service)}) — sync Service TCV to CCW."
            )

        owner = str(row[col("Opportunity Owner")] or "").strip()
        deals.append(
            {
                "priority": priority,
                "owner": owner,
                "ae_tab": ae_tab_name(owner),
                "account": str(row[col("Account Name")] or "").strip(),
                "opp_name": name,
                "deal_id": str(row[col("Deal ID")]),
                "forecast": str(fc),
                "stage": str(row[col("Stage")] or ""),
                "close_dt": close_dt,
                "days_to_close": days_to_close,
                "fq": fq,
                "fy": fy,
                "q_label": f"Q{fq} FY{fy}",
                "product": product,
                "service": service,
                "attach": attach,
                "t5": t5,
                "g5": g5,
                "t12": t12,
                "g12": g12,
                "ccw": ccw,
                "ccw_delta": ccw - service,
                "age": to_float(row[col("Age")]),
                "position": str(row[col("Forecasting Position")] or ""),
                "lead_source": str(row[col("Lead Source")] or ""),
                "type": typ,
                "svc_acv": svc_acv,
                "created": created_dt,
                "action": action,
            }
        )

    deals.sort(key=lambda d: (-d["product"], d["owner"], d["opp_name"]))
    for i, deal in enumerate(deals, start=1):
        deal["rank"] = i

    meta = {
        "funnel": dict(funnel),
        "source_rows": funnel["source_open_opps"],
        "title": str(ws["B2"].value or ""),
        "generated": str(ws["B3"].value or ""),
        "filters": [
            str(ws[f"B{r}"].value)
            for r in range(7, 13)
            if ws[f"B{r}"].value
        ],
    }
    return deals, meta


def ae_rollup(deals: list[dict]) -> list[dict]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for deal in deals:
        grouped[deal["owner"]].append(deal)

    rows = []
    for owner, items in grouped.items():
        items_sorted = sorted(items, key=lambda d: -d["product"])
        product = sum(d["product"] for d in items)
        service = sum(d["service"] for d in items)
        g5 = sum(d["g5"] for d in items)
        g12 = sum(d["g12"] for d in items)
        denom = product + service
        acct_gap = Counter()
        for d in items:
            acct_gap[d["account"]] += d["g12"]
        top3 = acct_gap.most_common(3)
        rows.append(
            {
                "owner": owner,
                "ae_tab": ae_tab_name(owner),
                "n": len(items),
                "accounts": len({d["account"] for d in items}),
                "product": product,
                "service": service,
                "attach": (service / denom) if denom else 0.0,
                "g5": g5,
                "g12": g12,
                "zero": sum(1 for d in items if d["service"] <= 0),
                "commit_n": sum(1 for d in items if d["forecast"] == "Commit"),
                "commit_g12": sum(d["g12"] for d in items if d["forecast"] == "Commit"),
                "ml_n": sum(1 for d in items if d["forecast"] == "Most Likely"),
                "up_n": sum(1 for d in items if d["forecast"] == "Upside"),
                "p1": sum(1 for d in items if d["priority"].startswith("P1")),
                "p2": sum(1 for d in items if d["priority"].startswith("P2")),
                "p3": sum(1 for d in items if d["priority"].startswith("P3")),
                "top3": top3,
                "top3_label": " | ".join(
                    f"{acct} ({money(gap)} 12% gap)" for acct, gap in top3
                ),
                "deals": items_sorted,
            }
        )
    rows.sort(key=lambda r: (-r["g12"], -r["product"], r["owner"]))
    for i, row in enumerate(rows, start=1):
        row["rank"] = i
    return rows


def quarter_rollup(deals: list[dict]) -> list[dict]:
    grouped: dict[str, list[dict]] = defaultdict(list)
    for deal in deals:
        grouped[deal["q_label"]].append(deal)
    rows = []
    for label, items in grouped.items():
        product = sum(d["product"] for d in items)
        service = sum(d["service"] for d in items)
        denom = product + service
        closes = [d["close_dt"] for d in items if d["close_dt"]]
        rows.append(
            {
                "label": label,
                "fy": items[0]["fy"],
                "fq": items[0]["fq"],
                "n": len(items),
                "product": product,
                "service": service,
                "attach": (service / denom) if denom else 0.0,
                "g5": sum(d["g5"] for d in items),
                "g12": sum(d["g12"] for d in items),
                "commit": sum(1 for d in items if d["forecast"] == "Commit"),
                "ml": sum(1 for d in items if d["forecast"] == "Most Likely"),
                "upside": sum(1 for d in items if d["forecast"] == "Upside"),
                "zero": sum(1 for d in items if d["service"] <= 0),
                "min_close": min(closes) if closes else None,
                "max_close": max(closes) if closes else None,
            }
        )
    rows.sort(key=lambda r: (r["fy"], r["fq"]))
    return rows


class Book:
    def __init__(self, path: Path):
        self.wb = xlsxwriter.Workbook(str(path))
        self.fmt = self._formats()

    def _formats(self):
        wb = self.wb
        f = {}
        f["title"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 20,
                "bold": True,
                "font_color": WHITE,
                "bg_color": NAVY,
                "valign": "vcenter",
                "align": "left",
            }
        )
        f["subtitle"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "font_color": WHITE,
                "bg_color": NAVY2,
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        f["section"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 13,
                "bold": True,
                "font_color": NAVY,
                "bottom": 2,
                "bottom_color": TEAL,
                "valign": "vcenter",
            }
        )
        f["label"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": MUTED,
                "bg_color": OFFWHITE,
                "valign": "vcenter",
            }
        )
        f["body"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "font_color": NAVY,
                "text_wrap": True,
                "valign": "top",
            }
        )
        f["body_bold"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "bold": True,
                "font_color": NAVY,
                "text_wrap": True,
                "valign": "top",
            }
        )
        f["muted"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "italic": True,
                "font_color": MUTED,
                "text_wrap": True,
                "valign": "top",
            }
        )
        f["kpi_value"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 16,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_NAVY,
                "align": "center",
                "valign": "vcenter",
            }
        )
        f["kpi_label"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 8,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_NAVY,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        f["kpi_value_teal"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 16,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_TEAL,
                "align": "center",
                "valign": "vcenter",
            }
        )
        f["kpi_label_teal"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 8,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_TEAL,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        f["kpi_value_gold"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 16,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_GOLD,
                "align": "center",
                "valign": "vcenter",
            }
        )
        f["kpi_label_gold"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 8,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_GOLD,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        f["kpi_value_red"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 16,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_RED,
                "align": "center",
                "valign": "vcenter",
            }
        )
        f["kpi_label_red"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 8,
                "bold": True,
                "font_color": WHITE,
                "bg_color": KPI_RED,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        f["th"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": NAVY,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "border": 1,
                "border_color": NAVY,
            }
        )
        f["td"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": WHITE,
                "valign": "vcenter",
                "border": 1,
                "border_color": LINE,
                "text_wrap": True,
            }
        )
        f["td_alt"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": ALT,
                "valign": "vcenter",
                "border": 1,
                "border_color": LINE,
                "text_wrap": True,
            }
        )
        f["td_num"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": WHITE,
                "valign": "vcenter",
                "align": "right",
                "border": 1,
                "border_color": LINE,
                "num_format": "$#,##0",
            }
        )
        f["td_num_alt"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": ALT,
                "valign": "vcenter",
                "align": "right",
                "border": 1,
                "border_color": LINE,
                "num_format": "$#,##0",
            }
        )
        f["td_gap"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": P1,
                "bg_color": GAP_BG,
                "bold": True,
                "valign": "vcenter",
                "align": "right",
                "border": 1,
                "border_color": LINE,
                "num_format": "$#,##0",
            }
        )
        f["td_gap_alt"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": P1,
                "bg_color": "#F8E1E0",
                "bold": True,
                "valign": "vcenter",
                "align": "right",
                "border": 1,
                "border_color": LINE,
                "num_format": "$#,##0",
            }
        )
        f["td_ok"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": "#17663A",
                "bg_color": OK_BG,
                "valign": "vcenter",
                "align": "right",
                "border": 1,
                "border_color": LINE,
                "num_format": "$#,##0",
            }
        )
        f["td_pct"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": WHITE,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": LINE,
                "num_format": "0.00%",
            }
        )
        f["td_pct_alt"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": ALT,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": LINE,
                "num_format": "0.00%",
            }
        )
        f["td_date"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": WHITE,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": LINE,
                "num_format": "yyyy-mm-dd",
            }
        )
        f["td_date_alt"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": ALT,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": LINE,
                "num_format": "yyyy-mm-dd",
            }
        )
        f["td_center"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": WHITE,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": LINE,
            }
        )
        f["td_center_alt"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "bg_color": ALT,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": LINE,
            }
        )
        f["total"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": NAVY2,
                "valign": "vcenter",
                "border": 1,
                "border_color": NAVY2,
            }
        )
        f["total_num"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": NAVY2,
                "valign": "vcenter",
                "align": "right",
                "border": 1,
                "border_color": NAVY2,
                "num_format": "$#,##0",
            }
        )
        f["total_pct"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": NAVY2,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": NAVY2,
                "num_format": "0.00%",
            }
        )
        f["total_center"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": NAVY2,
                "valign": "vcenter",
                "align": "center",
                "border": 1,
                "border_color": NAVY2,
            }
        )
        f["p1"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": P1,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": P1,
            }
        )
        f["p2"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": P2,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": P2,
            }
        )
        f["p3"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": P3,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": P3,
            }
        )
        f["fc_commit"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": COMMIT,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": COMMIT,
            }
        )
        f["fc_ml"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": ML,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": ML,
            }
        )
        f["fc_up"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "bold": True,
                "font_color": WHITE,
                "bg_color": UPSIDE,
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "border_color": UPSIDE,
            }
        )
        f["rule_num"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "font_color": WHITE,
                "bg_color": TEAL,
                "align": "center",
                "valign": "vcenter",
            }
        )
        f["rule_title"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10,
                "bold": True,
                "font_color": NAVY,
                "valign": "vcenter",
            }
        )
        f["rule_body"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 9,
                "font_color": NAVY,
                "text_wrap": True,
                "valign": "top",
            }
        )
        f["footer"] = wb.add_format(
            {
                "font_name": "Calibri",
                "font_size": 8,
                "italic": True,
                "font_color": MUTED,
            }
        )
        return f

    def close(self):
        self.wb.close()


def apply_sheet_chrome(ws, freeze_row=None, freeze_col=None):
    ws.hide_gridlines(2)
    ws.set_default_row(16)
    ws.set_zoom(90)
    if freeze_row is not None:
        ws.freeze_panes(freeze_row, freeze_col or 0)
    ws.fit_to_pages(1, 0)
    ws.set_landscape()
    ws.set_paper(1)
    ws.set_margins(0.4, 0.4, 0.5, 0.5)
    ws.repeat_rows(0, 1)
    ws.set_header("&LTallgrass Region — H1 Services Remediation&RConfidential")
    ws.set_footer("&LAs of November 25, 2026 EST&CPage &P of &N&RNot for distribution")


def write_title_banner(ws, fmt, title, subtitle, last_col):
    ws.merge_range(0, 0, 0, last_col, title, fmt["title"])
    ws.merge_range(1, 0, 1, last_col, subtitle, fmt["subtitle"])
    ws.set_row(0, 28)
    ws.set_row(1, 34)


def kpi_pair(ws, fmt, row, col, value, label, variant="navy"):
    vkey = {
        "navy": "kpi_value",
        "teal": "kpi_value_teal",
        "gold": "kpi_value_gold",
        "red": "kpi_value_red",
    }[variant]
    lkey = {
        "navy": "kpi_label",
        "teal": "kpi_label_teal",
        "gold": "kpi_label_gold",
        "red": "kpi_label_red",
    }[variant]
    ws.merge_range(row, col, row, col + 1, value, fmt[vkey])
    ws.merge_range(row + 1, col, row + 1, col + 1, label, fmt[lkey])
    ws.set_row(row, 26)
    ws.set_row(row + 1, 28)


def row_formats(fmt, alt: bool):
    return {
        "td": fmt["td_alt"] if alt else fmt["td"],
        "num": fmt["td_num_alt"] if alt else fmt["td_num"],
        "pct": fmt["td_pct_alt"] if alt else fmt["td_pct"],
        "date": fmt["td_date_alt"] if alt else fmt["td_date"],
        "center": fmt["td_center_alt"] if alt else fmt["td_center"],
        "gap": fmt["td_gap_alt"] if alt else fmt["td_gap"],
    }


def priority_fmt(fmt, priority: str):
    if priority.startswith("P1"):
        return fmt["p1"]
    if priority.startswith("P2"):
        return fmt["p2"]
    return fmt["p3"]


def fc_fmt(fmt, forecast: str):
    if forecast == "Commit":
        return fmt["fc_commit"]
    if forecast == "Most Likely":
        return fmt["fc_ml"]
    return fmt["fc_up"]


def gap_fmt(fmt, alt: bool, value: float):
    if value <= 0:
        return fmt["td_ok"]
    return fmt["td_gap_alt"] if alt else fmt["td_gap"]


def write_deal_header(ws, fmt, start_row, headers):
    for i, header in enumerate(headers):
        ws.write(start_row, i, header, fmt["th"])
    ws.set_row(start_row, 30)


def write_deal_row(ws, fmt, row_idx, deal, alt: bool):
    rf = row_formats(fmt, alt)
    values = [
        (deal["priority"], priority_fmt(fmt, deal["priority"])),
        (deal["rank"], rf["center"]),
        (deal["owner"], rf["td"]),
        (deal["ae_tab"], rf["center"]),
        (deal["account"], rf["td"]),
        (deal["opp_name"], rf["td"]),
        (deal["deal_id"], rf["center"]),
        (deal["forecast"], fc_fmt(fmt, deal["forecast"])),
        (deal["stage"], rf["td"]),
        (deal["close_dt"], rf["date"] if deal["close_dt"] else rf["center"]),
        (deal["days_to_close"] if deal["days_to_close"] is not None else "—", rf["center"]),
        (deal["fq"], rf["center"]),
        (int(deal["fy"]) if str(deal["fy"]).isdigit() else deal["fy"], rf["center"]),
        (deal["q_label"], rf["center"]),
        (deal["product"], rf["num"]),
        (deal["service"], rf["num"]),
        (deal["attach"], rf["pct"]),
        (deal["t5"], rf["num"]),
        (deal["g5"], gap_fmt(fmt, alt, deal["g5"])),
        (deal["t12"], rf["num"]),
        (deal["g12"], gap_fmt(fmt, alt, deal["g12"])),
        (deal["ccw"], rf["num"]),
        (deal["ccw_delta"], rf["num"]),
        (deal["age"], rf["center"]),
        (deal["position"], rf["td"]),
        (deal["action"], rf["td"]),
    ]
    for col, (value, cell_fmt) in enumerate(values):
        if value is None:
            ws.write(row_idx, col, "", cell_fmt)
        else:
            ws.write(row_idx, col, value, cell_fmt)
    ws.set_row(row_idx, 36)


def write_totals_row(ws, fmt, row_idx, deals, first_label="TOTAL"):
    product = sum(d["product"] for d in deals)
    service = sum(d["service"] for d in deals)
    attach = service / (product + service) if (product + service) else 0.0
    ws.write(row_idx, 0, first_label, fmt["total"])
    ws.write(row_idx, 1, len(deals), fmt["total_center"])
    for col in range(2, 14):
        ws.write(row_idx, col, "", fmt["total"])
    ws.write(row_idx, 14, product, fmt["total_num"])
    ws.write(row_idx, 15, service, fmt["total_num"])
    ws.write(row_idx, 16, attach, fmt["total_pct"])
    ws.write(row_idx, 17, sum(d["t5"] for d in deals), fmt["total_num"])
    ws.write(row_idx, 18, sum(d["g5"] for d in deals), fmt["total_num"])
    ws.write(row_idx, 19, sum(d["t12"] for d in deals), fmt["total_num"])
    ws.write(row_idx, 20, sum(d["g12"] for d in deals), fmt["total_num"])
    ws.write(row_idx, 21, sum(d["ccw"] for d in deals), fmt["total_num"])
    ws.write(row_idx, 22, sum(d["ccw_delta"] for d in deals), fmt["total_num"])
    for col in range(23, 26):
        ws.write(row_idx, col, "", fmt["total"])
    ws.set_row(row_idx, 20)


def set_deal_column_widths(ws):
    widths = {
        0: 20,
        1: 7,
        2: 18,
        3: 10,
        4: 32,
        5: 42,
        6: 12,
        7: 14,
        8: 24,
        9: 12,
        10: 12,
        11: 10,
        12: 10,
        13: 14,
        14: 16,
        15: 16,
        16: 12,
        17: 16,
        18: 13,
        19: 18,
        20: 13,
        21: 22,
        22: 18,
        23: 11,
        24: 26,
        25: 55,
    }
    for col, width in widths.items():
        ws.set_column(col, col, width)


def write_executive_summary(book: Book, deals, aes, quarters, meta):
    ws = book.wb.add_worksheet("Executive Summary")
    fmt = book.fmt
    last_col = 13
    apply_sheet_chrome(ws, freeze_row=3)
    ws.repeat_rows(0, 2)
    write_title_banner(
        ws,
        fmt,
        "Tallgrass Region  —  H1 Services Remediation Action Plan",
        "Non-renewal growth pipeline  |  Fiscal Q1 & Q2  |  Commit / Most Likely / Upside  |  "
        "Product TCV ≥ $10,000 and Service TCV < $10,000  |  As of November 25, 2026 7:38 AM EST  |  "
        "Source: Kristina TALLGRAS Forecast View  |  Confidential",
        last_col,
    )
    for col in range(last_col + 1):
        ws.set_column(col, col, 12)
    ws.set_column(0, 0, 14)
    ws.set_column(4, 5, 16)
    ws.set_column(10, 13, 18)

    product = sum(d["product"] for d in deals)
    service = sum(d["service"] for d in deals)
    attach = service / (product + service) if (product + service) else 0.0
    g5 = sum(d["g5"] for d in deals)
    g12 = sum(d["g12"] for d in deals)
    zero = sum(1 for d in deals if d["service"] <= 0)
    commit_n = sum(1 for d in deals if d["forecast"] == "Commit")
    commit_g12 = sum(d["g12"] for d in deals if d["forecast"] == "Commit")
    commit_zero = sum(1 for d in deals if d["forecast"] == "Commit" and d["service"] <= 0)
    accounts = len({d["account"] for d in deals})
    p1 = sum(1 for d in deals if d["priority"].startswith("P1"))
    p2 = sum(1 for d in deals if d["priority"].startswith("P2"))
    p3 = sum(1 for d in deals if d["priority"].startswith("P3"))

    # KPI strip
    kpi_row = 3
    kpis = [
        (str(len(deals)), "Flagged H1 Deals", "navy"),
        (str(accounts), "Unique Accounts", "navy"),
        (money(product), "Product TCV", "teal"),
        (money(service), "Service TCV", "teal"),
        (pct(attach), "Attach Rate", "red"),
        (money(g5), "5% Basic Gap", "gold"),
        (money(g12), "12% Standard Gap", "red"),
    ]
    for i, (value, label, variant) in enumerate(kpis):
        kpi_pair(ws, fmt, kpi_row, i * 2, value, label, variant)

    # Narrative
    r = 6
    ws.merge_range(r, 0, r, last_col, "Leadership Readout", fmt["section"])
    ws.set_row(r, 22)
    r = 7
    readout = (
        f"Tallgrass has {len(deals)} non-renewal H1 deals in Commit, Most Likely, or Upside "
        f"with Product TCV of {money(product)} and Service TCV of only {money(service)} "
        f"({pct(attach)} attach versus a 5% Basic floor and 12% Standard target). "
        f"{zero} of {len(deals)} flagged deals ({zero / len(deals):.0%}) have $0 services. "
        f"Closing the 5% floor is a {money(g5)} quoting action; reaching Standard is {money(g12)}. "
        f"Commit already contains {commit_n} of these deals ({commit_zero} at $0 services) representing "
        f"{money(commit_g12)} of 12% gap — that is in-quarter forecast risk, not pipeline theory. "
        f"Q2 FY2027 (current H1, close dates 7 Nov 2026–23 Jan 2027) holds 94 deals and "
        f"{money(sum(d['g12'] for d in deals if d['q_label']=='Q2 FY2027'))} of the 12% gap. "
        f"Q1 FY2028 (next-year H1 openers) holds 8 deals and "
        f"{money(sum(d['g12'] for d in deals if d['q_label']=='Q1 FY2028'))}. "
        f"Robin Randolph's Lockton WxC Recast ($7.0M product, $0 services, Upside) is 33% of the "
        f"regional 12% gap by itself. Garmin New Data Center ($2.0M, Commit, $0 services) is the "
        f"largest in-Commit miss."
    )
    ws.merge_range(r, 0, r + 2, last_col, readout, fmt["body"])
    ws.set_row(r, 18)
    ws.set_row(r + 1, 18)
    ws.set_row(r + 2, 18)

    r = 11
    ws.merge_range(r, 0, r, last_col, "Scope, Governance, and Filter Funnel", fmt["section"])
    ws.set_row(r, 22)
    r = 12
    funnel = meta["funnel"]
    scope = (
        "Rules applied in this order on the Kristina TALLGRAS Forecast View (Territory Level 05 = "
        "TALLGRASS REGION, Open opportunities). Amounts in the export are in $000s and are converted "
        "to USD for all metrics. Attach Rate % = Service TCV / (Product TCV + Service TCV). "
        "5% Basic Target = Product TCV × 5%; 5% Gap = MAX(0, Target − Service TCV). "
        "12% Standard Target = Product TCV × 12%; 12% Gap = MAX(0, Target − Service TCV)."
    )
    ws.merge_range(r, 0, r + 1, last_col, scope, fmt["body"])
    ws.set_row(r, 18)
    ws.set_row(r + 1, 18)

    r = 15
    funnel_rows = [
        ("0", "Source open opportunities with a Deal ID", funnel.get("source_open_opps", 0),
         "Salesforce export, Close Date = Current and Next 3 FQ"),
        ("1", "Exclude Opportunity Name starting with RNW, or Type = Renewal",
         funnel.get("non_renewal", 0),
         f"{funnel.get('excluded_renewal', 0)} renewals removed (Type was blank on all rows; RNW-prefix did the work)"),
        ("2", "Keep Fiscal Quarter Q1 and Q2 only (any FY in the extract)",
         funnel.get("h1", 0),
         "Q2 FY2027 current H1 + Q1 FY2028 next H1. FY2027 Q1 had 2 leftover RNW deals, already excluded."),
        ("3", "Keep Forecast Category in Commit, Most Likely, Upside",
         funnel.get("active_fc", 0),
         f"{funnel.get('excluded_inactive_fc', 0)} Not Forecastable / blank dropped"),
        ("4", "Product TCV ≥ $10,000 AND Service TCV < $10,000",
         funnel.get("cohort", 0),
         f"{funnel.get('excluded_not_low_services', 0)} deals had either small product or services already ≥ $10K"),
    ]
    ws.write(r, 0, "Step", fmt["th"])
    ws.merge_range(r, 1, r, 2, "Governance Rule", fmt["th"])
    ws.write(r, 3, "Remaining", fmt["th"])
    ws.merge_range(r, 4, r, last_col, "What dropped / note", fmt["th"])
    for i, (step, rule, remaining, note) in enumerate(funnel_rows):
        rr = r + 1 + i
        alt = i % 2 == 1
        rf = row_formats(fmt, alt)
        ws.write(rr, 0, step, rf["center"])
        ws.merge_range(rr, 1, rr, 2, rule, rf["td"])
        ws.write(rr, 3, remaining, rf["center"])
        ws.merge_range(rr, 4, rr, last_col, note, rf["td"])
        ws.set_row(rr, 28)

    r = 21
    ws.merge_range(r, 0, r, last_col, "Q1 vs Q2 Breakdown", fmt["section"])
    ws.set_row(r, 22)
    r = 22
    q_headers = [
        "Quarter",
        "Close Window",
        "Deals",
        "Zero-Svc",
        "Commit",
        "Most Likely",
        "Upside",
        "Product TCV",
        "Service TCV",
        "Attach Rate",
        "5% Target",
        "5% Gap",
        "12% Target",
        "12% Gap",
    ]
    for i, h in enumerate(q_headers):
        ws.write(r, i, h, fmt["th"])
    ws.set_row(r, 28)
    for i, q in enumerate(quarters):
        rr = r + 1 + i
        alt = i % 2 == 1
        rf = row_formats(fmt, alt)
        close_window = "—"
        if q["min_close"] and q["max_close"]:
            close_window = f"{q['min_close']:%d %b %Y} – {q['max_close']:%d %b %Y}"
        ws.write(rr, 0, q["label"], rf["td"])
        ws.write(rr, 1, close_window, rf["td"])
        ws.write(rr, 2, q["n"], rf["center"])
        ws.write(rr, 3, q["zero"], rf["center"])
        ws.write(rr, 4, q["commit"], rf["center"])
        ws.write(rr, 5, q["ml"], rf["center"])
        ws.write(rr, 6, q["upside"], rf["center"])
        ws.write(rr, 7, q["product"], rf["num"])
        ws.write(rr, 8, q["service"], rf["num"])
        ws.write(rr, 9, q["attach"], rf["pct"])
        ws.write(rr, 10, q["product"] * 0.05, rf["num"])
        ws.write(rr, 11, q["g5"], gap_fmt(fmt, alt, q["g5"]))
        ws.write(rr, 12, q["product"] * 0.12, rf["num"])
        ws.write(rr, 13, q["g12"], gap_fmt(fmt, alt, q["g12"]))
        ws.set_row(rr, 22)
    tot_row = r + 1 + len(quarters)
    ws.write(tot_row, 0, "H1 Total", fmt["total"])
    ws.write(tot_row, 1, "Q1 + Q2 in extract", fmt["total"])
    ws.write(tot_row, 2, len(deals), fmt["total_center"])
    ws.write(tot_row, 3, zero, fmt["total_center"])
    ws.write(tot_row, 4, commit_n, fmt["total_center"])
    ws.write(tot_row, 5, sum(1 for d in deals if d["forecast"] == "Most Likely"), fmt["total_center"])
    ws.write(tot_row, 6, sum(1 for d in deals if d["forecast"] == "Upside"), fmt["total_center"])
    ws.write(tot_row, 7, product, fmt["total_num"])
    ws.write(tot_row, 8, service, fmt["total_num"])
    ws.write(tot_row, 9, attach, fmt["total_pct"])
    ws.write(tot_row, 10, product * 0.05, fmt["total_num"])
    ws.write(tot_row, 11, g5, fmt["total_num"])
    ws.write(tot_row, 12, product * 0.12, fmt["total_num"])
    ws.write(tot_row, 13, g12, fmt["total_num"])

    r = tot_row + 2
    ws.merge_range(r, 0, r, last_col, "AE Gap Ranking (sorted by 12% Standard Gap) with Top 3 Accounts", fmt["section"])
    ws.set_row(r, 22)
    r += 1
    ae_headers = [
        "Rank",
        "Account Executive",
        "Tab",
        "Deals",
        "P1 / P2 / P3",
        "Commit Deals",
        "Product TCV",
        "Service TCV",
        "Attach Rate",
        "5% Gap",
        "12% Gap",
        "Commit 12% Gap",
        "Top 3 Accounts by 12% Gap",
        "",
    ]
    ws.write(r, 0, "Rank", fmt["th"])
    ws.write(r, 1, "Account Executive", fmt["th"])
    ws.write(r, 2, "Tab", fmt["th"])
    ws.write(r, 3, "Deals", fmt["th"])
    ws.write(r, 4, "P1 / P2 / P3", fmt["th"])
    ws.write(r, 5, "Commit", fmt["th"])
    ws.write(r, 6, "Product TCV", fmt["th"])
    ws.write(r, 7, "Service TCV", fmt["th"])
    ws.write(r, 8, "Attach Rate", fmt["th"])
    ws.write(r, 9, "5% Gap", fmt["th"])
    ws.write(r, 10, "12% Gap", fmt["th"])
    ws.write(r, 11, "Commit 12% Gap", fmt["th"])
    ws.merge_range(r, 12, r, last_col, "Top 3 Accounts by 12% Gap", fmt["th"])
    ws.set_row(r, 28)
    header_row_ae = r
    first_ae_data = r + 1
    for i, ae in enumerate(aes):
        rr = first_ae_data + i
        alt = i % 2 == 1
        rf = row_formats(fmt, alt)
        ws.write(rr, 0, ae["rank"], rf["center"])
        ws.write(rr, 1, ae["owner"], rf["td"])
        ws.write_url(rr, 2, f"internal:'{ae['ae_tab']}'!A1", None, ae["ae_tab"])
        ws.write(rr, 3, ae["n"], rf["center"])
        ws.write(rr, 4, f"{ae['p1']} / {ae['p2']} / {ae['p3']}", rf["center"])
        ws.write(rr, 5, ae["commit_n"], rf["center"])
        ws.write(rr, 6, ae["product"], rf["num"])
        ws.write(rr, 7, ae["service"], rf["num"])
        ws.write(rr, 8, ae["attach"], rf["pct"])
        ws.write(rr, 9, ae["g5"], gap_fmt(fmt, alt, ae["g5"]))
        ws.write(rr, 10, ae["g12"], gap_fmt(fmt, alt, ae["g12"]))
        ws.write(rr, 11, ae["commit_g12"], gap_fmt(fmt, alt, ae["commit_g12"]))
        ws.merge_range(rr, 12, rr, last_col, ae["top3_label"], rf["td"])
        ws.set_row(rr, 32)
    last_ae_data = first_ae_data + len(aes) - 1
    tot = first_ae_data + len(aes)
    ws.write(tot, 0, "", fmt["total"])
    ws.write(tot, 1, "Region Total", fmt["total"])
    ws.write(tot, 2, "", fmt["total"])
    ws.write(tot, 3, len(deals), fmt["total_center"])
    ws.write(tot, 4, f"{p1} / {p2} / {p3}", fmt["total_center"])
    ws.write(tot, 5, commit_n, fmt["total_center"])
    ws.write(tot, 6, product, fmt["total_num"])
    ws.write(tot, 7, service, fmt["total_num"])
    ws.write(tot, 8, attach, fmt["total_pct"])
    ws.write(tot, 9, g5, fmt["total_num"])
    ws.write(tot, 10, g12, fmt["total_num"])
    ws.write(tot, 11, commit_g12, fmt["total_num"])
    ws.merge_range(tot, 12, tot, last_col, f"{accounts} unique accounts across 7 AE action tabs", fmt["total"])

    # Hidden helper data for the chart (to the right of the print area)
    helper_col = 16
    ws.write(header_row_ae, helper_col, "AE")
    ws.write(header_row_ae, helper_col + 1, "12% Gap")
    for i, ae in enumerate(aes):
        ws.write(first_ae_data + i, helper_col, ae["ae_tab"])
        ws.write(first_ae_data + i, helper_col + 1, ae["g12"])
    ws.set_column(helper_col, helper_col + 3, 12, None, {"hidden": True})
    chart = book.wb.add_chart({"type": "bar"})
    chart.add_series(
        {
            "name": "12% Standard Gap ($)",
            "categories": ["Executive Summary", first_ae_data, helper_col, last_ae_data, helper_col],
            "values": ["Executive Summary", first_ae_data, helper_col + 1, last_ae_data, helper_col + 1],
            "fill": {"color": TEAL},
            "border": {"color": TEAL},
            "data_labels": {"num_format": "$#,##0", "font": {"size": 8, "color": NAVY}},
        }
    )
    chart.set_title({"name": "12% Standard Services Gap by AE", "name_font": {"size": 11, "color": NAVY, "bold": True}})
    chart.set_x_axis({"name": "12% Gap (USD)", "num_format": "$#,##0", "major_gridlines": {"visible": False}})
    chart.set_y_axis({"reverse": True, "major_gridlines": {"visible": False}})
    chart.set_legend({"none": True})
    chart.set_style(10)
    chart.set_size({"width": 720, "height": 280})
    ws.insert_chart(tot + 2, 0, chart, {"x_offset": 0, "y_offset": 8})

    r = tot + 18
    ws.merge_range(r, 0, r, last_col, "Priority Mix and This-Week Must-Do (largest 12% gaps)", fmt["section"])
    ws.set_row(r, 22)
    r += 1
    mix = (
        f"Priority mix: {p1} P1 Zero Services  |  {p2} P2 Below 5% Floor  |  {p3} P3 Below 12% Standard. "
        f"Work P1 Commit first (already in the forecast), then P1 Most Likely, then large P1 Upside. "
        f"Do not move Upside → Most Likely/Commit until a services SKU is on the quote."
    )
    ws.merge_range(r, 0, r, last_col, mix, fmt["body"])
    r += 2
    must_headers = ["#", "Priority", "AE", "Account", "Opportunity", "FC", "Product TCV", "Svc TCV", "5% Gap", "12% Gap", "Close", "Why this week", "", ""]
    for i, h in enumerate(["#", "Priority", "AE", "Account", "Opportunity", "FC", "Product TCV", "Service TCV", "5% Gap", "12% Gap", "Close Date"]):
        ws.write(r, i, h, fmt["th"])
    ws.merge_range(r, 11, r, last_col, "Why this week", fmt["th"])
    ws.set_row(r, 26)
    top10 = sorted(deals, key=lambda d: -d["g12"])[:10]
    for i, d in enumerate(top10):
        rr = r + 1 + i
        alt = i % 2 == 1
        rf = row_formats(fmt, alt)
        why = (
            "Largest regional miss — 5% floor alone is "
            + money(d["g5"])
            + ". "
            + ("Already in Commit with $0 services." if d["forecast"] == "Commit" and d["service"] <= 0 else d["priority"])
        )
        ws.write(rr, 0, i + 1, rf["center"])
        ws.write(rr, 1, d["priority"], priority_fmt(fmt, d["priority"]))
        ws.write(rr, 2, d["ae_tab"], rf["center"])
        ws.write(rr, 3, d["account"], rf["td"])
        ws.write(rr, 4, d["opp_name"], rf["td"])
        ws.write(rr, 5, d["forecast"], fc_fmt(fmt, d["forecast"]))
        ws.write(rr, 6, d["product"], rf["num"])
        ws.write(rr, 7, d["service"], rf["num"])
        ws.write(rr, 8, d["g5"], gap_fmt(fmt, alt, d["g5"]))
        ws.write(rr, 9, d["g12"], gap_fmt(fmt, alt, d["g12"]))
        ws.write(rr, 10, d["close_dt"], rf["date"] if d["close_dt"] else rf["center"])
        ws.merge_range(rr, 11, rr, last_col, why, rf["td"])
        ws.set_row(rr, 30)

    r = r + 12
    ws.merge_range(r, 0, r, last_col, "Quoting Rules (mandatory on every flagged deal)", fmt["section"])
    ws.set_row(r, 22)
    rules = [
        (
            "1",
            "Eligibility gate",
            "Every non-renewal H1 opportunity with Product TCV ≥ $10,000 and Service TCV < $10,000 "
            "in Commit / Most Likely / Upside is in this plan. No exceptions by stage, partner, or 'services to follow'.",
        ),
        (
            "2",
            "5% Basic floor (must-quote)",
            "Quote Basic services at Product TCV × 5% (CX Essentials, Solution Support, or equivalent). "
            "The 5% Gap column is the dollar amount to add. $0-service Commit deals cannot remain $0.",
        ),
        (
            "3",
            "12% Standard package (target quote)",
            "Quote Standard services at Product TCV × 12% covering CX Preferred / Unified, TAS, Solution Support, "
            "deployment, and adoption. The 12% Gap column is the stretch ask on the same close date as product.",
        ),
        (
            "4",
            "Same-order, same close date",
            "Services book on the same opportunity and close date as product. A delayed services follow-on is not a close plan. "
            "If the partner will deliver, still capture Cisco-paper or Expected Service Amount so attach is visible.",
        ),
        (
            "5",
            "CCW is source of truth for the quote",
            "If Expected Service Amount CCW > Salesforce Service TCV, sync Service TCV to CCW before the next forecast lock. "
            "Do not forecast product-only when CCW already has a services line.",
        ),
        (
            "6",
            "Forecast hygiene",
            "Do not promote Upside → Most Likely or Most Likely → Commit until a services SKU is on the quote. "
            "Commit deals in this workbook with $0 services are a forecast-quality defect and a P1 this week.",
        ),
        (
            "7",
            "Exception path",
            "Customer refusal, overlay-owned services, or partner-delivered off Cisco paper requires a written exception "
            "on the opportunity (who, why, residual Cisco attach). Unwritten $0 is not an exception.",
        ),
        (
            "8",
            "Priority order",
            "P1 Zero Services, then P2 Below 5% Floor, then P3 Below 12% Standard. Inside each band: Commit first, "
            "then Most Likely, then largest Product TCV Upside. AE tabs are pre-sorted by Product TCV.",
        ),
    ]
    r += 1
    for i, (num, title, body) in enumerate(rules):
        rr = r + i
        ws.write(rr, 0, num, fmt["rule_num"])
        ws.write(rr, 1, title, fmt["rule_title"])
        ws.merge_range(rr, 2, rr, last_col, body, fmt["rule_body"])
        ws.set_row(rr, 36)

    r = r + len(rules) + 1
    ws.merge_range(r, 0, r, last_col, "Notes the filters did not drop (by design)", fmt["section"])
    ws.set_row(r, 22)
    r += 1
    notes = (
        "Strict RNW / Type=Renewal exclusion retained three name-contains-'Renewal' deals that do not start with RNW "
        "and have blank Type: NBH-Logicalis-TE Renewal only (Amy K, $198K), Stinson-Webex Calling & Collab Renewal-On Premise "
        "(Amy K, $90K), and Renewal Sub645037 (Jeff P, $80K). Review with the AE; if they are true renewals, exclude them "
        "from the next snapshot. Ashton Dunn (191 opps) and Jeff Vavak (180 opps) are almost entirely RNW and correctly "
        "have no action tab. Paige Ball and Christopher Chatham have H1 non-renewal deals that are Not Forecastable, so they "
        "are out of scope until forecast category is raised."
    )
    ws.merge_range(r, 0, r + 2, last_col, notes, fmt["muted"])
    r += 4
    ws.merge_range(
        r,
        0,
        r,
        last_col,
        "How to use this workbook: Tab 2 is the regional queue (sorted by Product TCV). Tabs 3–9 are AE worklists named "
        "FirstName LastInitial. Filter Priority = P1 and Forecast Category = Commit for Monday's inspection. Confidential — do not distribute.",
        fmt["footer"],
    )
    ws.print_area(0, 0, r, last_col)
    ws.set_landscape()
    return ws


def write_master(book: Book, deals):
    ws = book.wb.add_worksheet("H1 Master Action Plan")
    fmt = book.fmt
    last_col = len(MASTER_HEADERS) - 1
    apply_sheet_chrome(ws, freeze_row=3, freeze_col=6)
    write_title_banner(
        ws,
        fmt,
        "H1 Master Action Plan  —  All Flagged Deals (Product TCV descending)",
        f"{len(deals)} non-renewal H1 Commit / Most Likely / Upside deals  |  "
        f"Product TCV ≥ $10,000 and Service TCV < $10,000  |  "
        f"Product {money(sum(d['product'] for d in deals))}  |  "
        f"5% Gap {money(sum(d['g5'] for d in deals))}  |  "
        f"12% Gap {money(sum(d['g12'] for d in deals))}",
        last_col,
    )
    set_deal_column_widths(ws)
    write_deal_header(ws, fmt, 2, MASTER_HEADERS)
    for i, deal in enumerate(deals):
        write_deal_row(ws, fmt, 3 + i, deal, alt=i % 2 == 1)
    total_row = 3 + len(deals)
    write_totals_row(ws, fmt, total_row, deals, "H1 TOTAL")
    ws.autofilter(2, 0, 2 + len(deals), last_col)
    ws.freeze_panes(3, 6)
    ws.print_area(0, 0, total_row, last_col)
    return ws


def write_ae_tab(book: Book, ae: dict):
    name = safe_sheet_name(ae["ae_tab"])
    ws = book.wb.add_worksheet(name)
    fmt = book.fmt
    deals = ae["deals"]
    last_col = len(MASTER_HEADERS) - 1
    apply_sheet_chrome(ws, freeze_row=8, freeze_col=6)
    write_title_banner(
        ws,
        fmt,
        f"{ae['owner']}  —  H1 Services Remediation Worklist",
        f"Tab {ae['ae_tab']}  |  Rank {ae['rank']} of 7 by 12% gap  |  "
        f"{ae['n']} flagged deals  |  {ae['accounts']} accounts  |  "
        f"Product {money(ae['product'])}  |  Attach {pct(ae['attach'])}  |  "
        f"5% Gap {money(ae['g5'])}  |  12% Gap {money(ae['g12'])}",
        last_col,
    )
    set_deal_column_widths(ws)

    kpi_row = 3
    kpis = [
        (str(ae["n"]), "Flagged Deals", "navy"),
        (str(ae["zero"]), "Zero-Service Deals", "red"),
        (money(ae["product"]), "Product TCV", "teal"),
        (money(ae["service"]), "Service TCV", "teal"),
        (pct(ae["attach"]), "Attach Rate", "red"),
        (money(ae["g5"]), "5% Basic Gap", "gold"),
        (money(ae["g12"]), "12% Standard Gap", "red"),
        (money(ae["commit_g12"]), "Commit 12% Gap", "gold"),
    ]
    # KPIs sit in first 16 columns; master table is wider so this is fine
    for i, (value, label, variant) in enumerate(kpis):
        col = i * 2
        if col + 1 <= last_col:
            kpi_pair(ws, fmt, kpi_row, col, value, label, variant)

    r = 6
    ws.merge_range(r, 0, r, last_col, "Top 3 Accounts by 12% Gap — inspect these first", fmt["section"])
    ws.set_row(r, 20)
    r = 7
    ws.write(r, 0, "#", fmt["th"])
    ws.merge_range(r, 1, r, 4, "Account", fmt["th"])
    ws.write(r, 5, "Deals", fmt["th"])
    ws.write(r, 6, "Product TCV", fmt["th"])
    ws.write(r, 7, "Service TCV", fmt["th"])
    ws.write(r, 8, "5% Gap", fmt["th"])
    ws.write(r, 9, "12% Gap", fmt["th"])
    ws.merge_range(r, 10, r, last_col, "Largest deal on this account", fmt["th"])
    for i, (acct, gap) in enumerate(ae["top3"]):
        rr = r + 1 + i
        subset = [d for d in deals if d["account"] == acct]
        biggest = max(subset, key=lambda d: d["g12"])
        alt = i % 2 == 1
        rf = row_formats(fmt, alt)
        ws.write(rr, 0, i + 1, rf["center"])
        ws.merge_range(rr, 1, rr, 4, acct, rf["td"])
        ws.write(rr, 5, len(subset), rf["center"])
        ws.write(rr, 6, sum(d["product"] for d in subset), rf["num"])
        ws.write(rr, 7, sum(d["service"] for d in subset), rf["num"])
        ws.write(rr, 8, sum(d["g5"] for d in subset), gap_fmt(fmt, alt, sum(d["g5"] for d in subset)))
        ws.write(rr, 9, gap, gap_fmt(fmt, alt, gap))
        ws.merge_range(
            rr,
            10,
            rr,
            last_col,
            f"{biggest['opp_name']}  |  {biggest['forecast']}  |  Product {money(biggest['product'])}  |  "
            f"Ask 5% {money(biggest['g5'])} / 12% {money(biggest['g12'])}",
            rf["td"],
        )
        ws.set_row(rr, 24)

    header_row = 12
    write_deal_header(ws, fmt, header_row, MASTER_HEADERS)
    for i, deal in enumerate(deals):
        write_deal_row(ws, fmt, header_row + 1 + i, deal, alt=i % 2 == 1)
    total_row = header_row + 1 + len(deals)
    write_totals_row(ws, fmt, total_row, deals, f"{ae['ae_tab']} TOTAL")
    ws.autofilter(header_row, 0, header_row + len(deals), last_col)
    ws.freeze_panes(header_row + 1, 6)
    ws.print_area(0, 0, total_row, last_col)
    return ws


def build(source: Path, output: Path):
    deals, meta = load_source(source)
    if not deals:
        raise SystemExit("No deals remained after governance filters.")
    aes = ae_rollup(deals)
    quarters = quarter_rollup(deals)
    output.parent.mkdir(parents=True, exist_ok=True)
    book = Book(output)
    write_executive_summary(book, deals, aes, quarters, meta)
    write_master(book, deals)
    for ae in aes:
        write_ae_tab(book, ae)
    book.close()
    return {
        "deals": len(deals),
        "aes": len(aes),
        "product": sum(d["product"] for d in deals),
        "service": sum(d["service"] for d in deals),
        "g5": sum(d["g5"] for d in deals),
        "g12": sum(d["g12"] for d in deals),
        "output": str(output),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    summary = build(Path(args.source), Path(args.output))
    print(summary)


if __name__ == "__main__":
    main()
