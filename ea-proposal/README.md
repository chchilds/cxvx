# Hallmark Cards — Cisco Enterprise Agreement proposal

Generates a customer-facing Enterprise Agreement PPTX from the Cisco EAMP Price Estimate workbook.

Customer TCV replaces EAMP net with **20 points of partner margin** (20% gross margin; sell = EAMP net ÷ 0.80). Cisco list, quantities, subscription %, and OTDs are unchanged.

```bash
pip install -r ea-proposal/requirements.txt
python ea-proposal/build_ea_pptx.py
python ea-proposal/build_ea_pptx.py --margin 20
```

Source workbook: `source/HALLMARK_CARDS_INCORPORATED_EAMP_Price_Estimate_ID_9343374.xlsx` (Proposal ID 9343374).

The output file is `HALLMARK_CARDS_INCORPORATED_Enterprise_Agreement_Proposal.pptx`.

Pricing in the deck is indicative, not an approved Cisco quote.
