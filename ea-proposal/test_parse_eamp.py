#!/usr/bin/env python3
"""Sanity checks against the Hallmark EAMP workbook."""

from pathlib import Path

from parse_eamp import apply_partner_margin, parse_eamp_workbook, suite_rollups

SOURCE = Path(__file__).resolve().parent / "source" / "HALLMARK_CARDS_INCORPORATED_EAMP_Price_Estimate_ID_9343374.xlsx"


def test_hallmark_totals():
    data = parse_eamp_workbook(SOURCE)
    assert data["meta"]["proposal_id"] == "9343374"
    assert data["meta"]["smart_account"] == "HALLMARK CARDS, INCORPORATED"
    total = next(r for r in data["financial_summary"] if r["portfolio"] == "TOTAL")
    assert round(total["total"], 2) == 2191236.13
    assert round(total["software"], 2) == 1187059.90
    assert round(total["services"], 2) == 1004176.23
    net = next(p for p in data["portfolios"] if p["name"] == "Networking Infrastructure")
    assert round(net["header"]["software_line_tcv"], 2) == 1112861.40
    assert round(net["header"]["services_line_tcv"], 2) == 559662.55
    sec = next(p for p in data["portfolios"] if p["name"] == "Security")
    assert round(sec["header"]["software_line_tcv"], 2) == 74198.50
    col = next(p for p in data["portfolios"] if p["name"] == "Collaboration")
    assert round(col["header"]["services_line_tcv"], 2) == 444513.68
    rollups = suite_rollups(net["software"])
    dna_wl = next(r for r in rollups if r["suite"] == "Cisco DNA Wireless")
    assert dna_wl["qty_desired"] == 525
    print("ok", data["meta"]["proposal_id"], total["total"])


def test_partner_margin_20_points():
    data = apply_partner_margin(parse_eamp_workbook(SOURCE), points=20)
    total = next(r for r in data["financial_summary"] if r["portfolio"] == "TOTAL")
    assert round(total["eamp_total"], 2) == 2191236.13
    assert round(total["total"], 2) == 2739045.16
    assert round(total["partner_margin"], 2) == 547809.03
    assert abs(total["total"] - total["eamp_total"] / 0.8) < 0.02
    net = next(p for p in data["portfolios"] if p["name"] == "Networking Infrastructure")
    assert round(net["header"]["software_line_tcv"], 2) == 1391076.75
    print("ok margin", total["eamp_total"], "->", total["total"], "margin", total["partner_margin"])


if __name__ == "__main__":
    test_hallmark_totals()
    test_partner_margin_20_points()
