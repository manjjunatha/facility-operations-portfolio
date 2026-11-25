# Vendor Evaluation and Cost Analysis Dashboard
**MS Excel · Power BI**

A weighted vendor scorecard that rates suppliers on **cost, delivery time, quality and complaints**, plus a spend analysis by category, vendor and month, to support data-driven vendor selection.

| File | |
|---|---|
| `Vendor_Evaluation_Cost_Analysis.xlsx` | Excel model and dashboard |
| `PowerBI/*.csv` | Same data for Power BI |
| `PowerBI/PowerBI_Guide.md` | Data model, DAX measures, report layout, and expected values to check against |

## Excel sheets
| Sheet | What it does |
|---|---|
| **Scorecard** | Per vendor: POs, spend, average price index, on-time %, average quality, complaints. Four 0–100 criterion scores, weighted score, rank and rating (Preferred / Approved / Under Review). Ranked list, recommended vendor per category, score chart |
| **Spend_Analysis** | KPI tiles, spend by category (share %, average PO), spend by vendor, top-5 vendors, monthly spend by category, 3 charts |
| **PO_Data** | 186 purchase orders (inputs). Formulas add category, amount, benchmark price, price index, delay, on-time flag, month |
| **Complaints** | 18 vendor complaints linked to POs |
| **Settings** | Criterion weights (checked to total 100%), complaint penalty, rating thresholds, vendor list, item master with benchmark prices |

## Scoring method
| Criterion | Weight | Score (0–100) |
|---|---|---|
| Cost | 35% | lowest avg. price index among vendors ÷ vendor's avg. price index × 100 (price index = unit price ÷ benchmark) |
| Delivery | 25% | on-time % × 100 |
| Quality | 25% | avg. rating (1–5) ÷ 5 × 100 |
| Complaints | 15% | 100 − complaints per 10 POs × 15, floored at 0 |

Weighted score = Σ score × weight. Preferred ≥ 88, Approved ≥ 80, else Under Review. All weights and thresholds are editable in Settings.

Functions used: `SUMIF(S)`, `COUNTIF(S)`, `AVERAGEIF`, `INDEX/MATCH`, `MAXIFS`, `LARGE`, `RANK`, `EDATE`, data validation, conditional formatting (data bars, rating colours).

## Sample results
Data is fictional (12 vendors, 6 categories, Jun–Nov 2026, ₹24.8 lakh spend). Top-ranked: PaperPlus Stationers (92.5) and TechNet Systems (91.5). Hygiene Hub Distributors (77.6) is *Under Review*: the cheapest housekeeping supplier, but with the lowest quality rating (3.3/5) and the most complaints (7). Furniture is 44% of spend, and its two vendors score only *Approved*, so it is the main category to renegotiate.

The Excel scores were cross-checked against an independent Python calculation from the CSVs. All matched.
