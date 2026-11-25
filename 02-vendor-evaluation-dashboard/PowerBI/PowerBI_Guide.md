# Vendor Evaluation & Cost Analysis — Power BI Build Guide

This folder holds the same data as the Excel workbook, as CSV files, so the dashboard can be rebuilt in **Power BI Desktop** (free, Windows). It takes about 30–40 minutes.

| File | Contents |
|---|---|
| `Vendors.csv` | Vendor name and category |
| `Items.csv` | Item, category, benchmark unit price (Rs) |
| `PurchaseOrders.csv` | 186 POs: date, vendor, item, qty, unit price, promised/delivered dates, quality rating |
| `Complaints.csv` | Vendor complaints linked to PO numbers |
| `Weights.csv` | Scoring weights (Cost 35%, Delivery 25%, Quality 25%, Complaints 15%) |

---

## 1. Load the data
1. Power BI Desktop → **Get data → Text/CSV** → load all five CSV files.
2. In **Power Query** (Transform data), check the data types:
   - `PODate`, `PromisedDate`, `DeliveredDate`, `ComplaintDate` → **Date**
   - `Qty`, `UnitPrice`, `BenchmarkPrice`, `QualityRating` → **Whole/Decimal number**
   - `Weight` → **Decimal number**
3. **Close & Apply**.

## 2. Build the model (Model view)
Create these relationships (all many-to-one, single direction):

| From (many) | To (one) |
|---|---|
| `PurchaseOrders[Vendor]` | `Vendors[Vendor]` |
| `PurchaseOrders[Item]` | `Items[Item]` |
| `Complaints[Vendor]` | `Vendors[Vendor]` |

Optional date table (for month slicers): **Modeling → New table**
```DAX
Calendar = ADDCOLUMNS ( CALENDAR ( DATE(2026,6,1), DATE(2026,12,31) ),
    "Month", FORMAT ( [Date], "MMM yyyy" ),
    "MonthStart", DATE ( YEAR([Date]), MONTH([Date]), 1 ) )
```
Relate `PurchaseOrders[PODate]` → `Calendar[Date]` and `Complaints[ComplaintDate]` → `Calendar[Date]`. Sort `Calendar[Month]` by `Calendar[MonthStart]`.

## 3. Calculated columns (table: PurchaseOrders)
```DAX
Amount         = PurchaseOrders[Qty] * PurchaseOrders[UnitPrice]
Category       = RELATED ( Items[Category] )
BenchmarkPrice = RELATED ( Items[BenchmarkPrice] )
PriceIndex     = DIVIDE ( PurchaseOrders[UnitPrice], PurchaseOrders[BenchmarkPrice] )
OnTime         = IF ( PurchaseOrders[DeliveredDate] <= PurchaseOrders[PromisedDate], 1, 0 )
DelayDays      = MAX ( 0, DATEDIFF ( PurchaseOrders[PromisedDate], PurchaseOrders[DeliveredDate], DAY ) )
```

## 4. Measures
Create a blank table called `_Measures` (Enter data → one empty column) and add these measures to it.

```DAX
Total Spend        = SUM ( PurchaseOrders[Amount] )
PO Count           = COUNTROWS ( PurchaseOrders )
Spend Share %      = DIVIDE ( [Total Spend], CALCULATE ( [Total Spend], ALLSELECTED () ) )
On-Time %          = AVERAGE ( PurchaseOrders[OnTime] )
Avg Quality        = AVERAGE ( PurchaseOrders[QualityRating] )
Avg Price Index    = AVERAGE ( PurchaseOrders[PriceIndex] )
Complaint Count    = COUNTROWS ( Complaints ) + 0
Complaints per 10 POs = DIVIDE ( [Complaint Count], [PO Count] ) * 10

-- criterion scores (0-100), same logic as the Excel Scorecard
Best Price Index   = MINX ( ALL ( Vendors[Vendor] ), [Avg Price Index] )
Cost Score         = DIVIDE ( [Best Price Index], [Avg Price Index] ) * 100
Delivery Score     = [On-Time %] * 100
Quality Score      = DIVIDE ( [Avg Quality], 5 ) * 100
Complaint Score    = MAX ( 0, 100 - [Complaints per 10 POs] * 15 )

Weighted Score =
VAR wC = LOOKUPVALUE ( Weights[Weight], Weights[Criterion], "Cost" )
VAR wD = LOOKUPVALUE ( Weights[Weight], Weights[Criterion], "Delivery" )
VAR wQ = LOOKUPVALUE ( Weights[Weight], Weights[Criterion], "Quality" )
VAR wK = LOOKUPVALUE ( Weights[Weight], Weights[Criterion], "Complaints" )
RETURN
    IF ( [PO Count] > 0,
         [Cost Score] * wC + [Delivery Score] * wD + [Quality Score] * wQ + [Complaint Score] * wK )

Vendor Rank =
IF ( HASONEVALUE ( Vendors[Vendor] ),
     RANKX ( ALL ( Vendors[Vendor] ), [Weighted Score], , DESC, Dense ) )

Rating =
SWITCH ( TRUE (),
    ISBLANK ( [Weighted Score] ), "No data",
    [Weighted Score] >= 88, "Preferred",
    [Weighted Score] >= 80, "Approved",
    "Under Review" )
```
Formats: `Total Spend` → Whole number with thousands separator, prefix "₹"; `On-Time %`, `Spend Share %` → Percentage (1 decimal); scores → Decimal (1 decimal); `Avg Price Index` → 2 decimals.

## 5. Report pages

**Page 1 — Vendor Scorecard**
- Slicers: `Vendors[Category]`, `Calendar[Month]`
- Cards: Total Spend, PO Count, On-Time %, Avg Price Index, Complaint Count
- Table: `Vendors[Vendor]`, `Vendors[Category]`, PO Count, Total Spend, Cost Score, Delivery Score, Quality Score, Complaint Score, Weighted Score, Vendor Rank, Rating
  - Conditional formatting → Data bars on Weighted Score; background colour rules on Rating (Preferred green, Approved amber, Under Review red)
  - Sort by Weighted Score, descending
- Clustered bar chart: Vendor (axis) × Weighted Score (value), sorted descending

**Page 2 — Spend Analysis**
- Donut: `PurchaseOrders[Category]` × Total Spend
- Bar chart: Vendor × Total Spend (top-N filter = 5 for a "Top 5 vendors" visual)
- Stacked column: `Calendar[Month]` × Total Spend, legend = `PurchaseOrders[Category]`
- Scatter: X = Total Spend, Y = Weighted Score, details = Vendor — shows high-spend vendors that score low (renegotiation or replacement candidates)

## 6. Check it matches Excel
With no slicers applied, the Power BI table should show these values (from the Excel Scorecard sheet):

| Vendor | POs | Total Spend (Rs) | Weighted Score | Rating |
|---|---|---|---|---|
| Sri Durga Electricals | 22 | 140,211 | 90.97 | Preferred |
| Coastal Power Traders | 12 | 104,108 | 83.54 | Approved |
| Mangala Plumbing Mart | 10 | 88,239 | 89.33 | Preferred |
| Karavali Sanitary Stores | 15 | 65,854 | 90.76 | Preferred |
| CleanWay Supplies | 18 | 138,672 | 80.58 | Approved |
| Hygiene Hub Distributors | 18 | 140,744 | 77.56 | Under Review |
| TechNet Systems | 14 | 309,318 | 91.51 | Preferred |
| Digital Point Mangalore | 14 | 341,828 | 84.40 | Approved |
| PaperPlus Stationers | 13 | 46,869 | 92.49 | Preferred |
| Office Mart Wholesale | 10 | 21,861 | 86.83 | Approved |
| Comfort Seating Co. | 23 | 614,379 | 86.22 | Approved |
| Woodline Furnishers | 17 | 469,957 | 80.30 | Approved |

Total spend: Rs 2,482,040 across 186 POs, on-time delivery 83.3%, 18 complaints.

> When a Category or Month slicer is applied, `Best Price Index` is recalculated within the filtered period, so scores shift slightly. That is intended: vendors are compared on the same period.

Save as `Vendor_Evaluation_Dashboard.pbix`.
