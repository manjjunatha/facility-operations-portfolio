"""Project 2: Vendor Evaluation and Cost Analysis Dashboard (MS Excel + Power BI kit)."""
import os
import random, sys, os, csv, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

random.seed(7)
OUT_XLSX, CSV_DIR = sys.argv[1], sys.argv[2]
os.makedirs(CSV_DIR, exist_ok=True)

CATS = ["Electrical Supplies", "Plumbing Supplies", "Housekeeping Supplies", "IT Hardware", "Stationery", "Furniture"]
ITEMS = {  # item: benchmark unit price (Rs)
    "Electrical Supplies": [("LED Tube Light 20W", 320), ("MCB 32A", 450), ("Copper Wire 2.5 sqmm (90 m)", 2600), ("Ceiling Fan", 2100)],
    "Plumbing Supplies": [("PVC Pipe 1 inch (3 m)", 380), ("Bib Tap Brass", 520), ("Flush Valve", 1450), ("Ball Valve 1 inch", 340)],
    "Housekeeping Supplies": [("Floor Cleaner 5 L", 420), ("Garbage Bags (100)", 260), ("Hand Wash 5 L", 480), ("Mop Set", 650)],
    "IT Hardware": [("LAN Cable Cat6 (305 m)", 7800), ("Wireless Mouse", 550), ("Network Switch 24-port", 14500), ("Printer Toner", 2900)],
    "Stationery": [("A4 Paper Ream", 290), ("Whiteboard Marker (10)", 230), ("Box File", 140), ("Register 200 pg", 120)],
    "Furniture": [("Office Chair", 5200), ("Student Desk", 3800), ("Steel Cupboard", 11500), ("Visitor Chair", 2400)],
}
# vendor, category, price factor, on-time prob, quality mean, complaint prob per PO
VENDORS = [
    ("Sri Durga Electricals", "Electrical Supplies", 0.97, 0.90, 4.3, 0.05),
    ("Coastal Power Traders", "Electrical Supplies", 0.91, 0.70, 3.6, 0.20),
    ("Mangala Plumbing Mart", "Plumbing Supplies", 1.03, 0.92, 4.4, 0.04),
    ("Karavali Sanitary Stores", "Plumbing Supplies", 0.95, 0.80, 3.9, 0.10),
    ("CleanWay Supplies", "Housekeeping Supplies", 0.99, 0.88, 4.1, 0.06),
    ("Hygiene Hub Distributors", "Housekeeping Supplies", 0.93, 0.66, 3.4, 0.22),
    ("TechNet Systems", "IT Hardware", 1.05, 0.93, 4.6, 0.03),
    ("Digital Point Mangalore", "IT Hardware", 0.96, 0.78, 3.8, 0.12),
    ("PaperPlus Stationers", "Stationery", 0.98, 0.95, 4.2, 0.03),
    ("Office Mart Wholesale", "Stationery", 0.92, 0.84, 3.7, 0.09),
    ("Comfort Seating Co.", "Furniture", 1.06, 0.86, 4.5, 0.05),
    ("Woodline Furnishers", "Furniture", 0.94, 0.72, 3.5, 0.18),
]
CTYPES = ["Late delivery", "Damaged goods", "Wrong item", "Short supply", "Billing error", "Poor service"]
WEIGHTS = [("Cost", 0.35), ("Delivery", 0.25), ("Quality", 0.25), ("Complaints", 0.15)]

# ---------- sample data ----------
pos, comps = [], []
start = dt.date(2026, 6, 1)
for i in range(186):
    v, cat, pf, ot, qm, cp = random.choice(VENDORS)
    item, bench = random.choice(ITEMS[cat])
    d = start + dt.timedelta(days=random.randint(0, 172))
    qty = random.choice([2, 4, 5, 6, 10, 12, 15, 20, 25, 30, 50]) if bench < 1000 else random.choice([1, 2, 3, 4, 5, 6, 8, 10])
    price = round(bench * pf * random.uniform(0.95, 1.05))
    lead = random.choice([3, 5, 7, 7, 10])
    promised = d + dt.timedelta(days=lead)
    delay = 0 if random.random() < ot else random.randint(1, 6)
    delivered = promised + dt.timedelta(days=delay - (random.randint(0, 1) if delay == 0 else 0))
    q = max(1, min(5, round(random.gauss(qm, 0.6))))
    po = f"PO-{26001 + i}"
    pos.append([po, d, v, item, qty, price, promised, delivered, q])
    if random.random() < cp or (delay > 3 and random.random() < 0.5):
        ctype = "Late delivery" if delay > 0 and random.random() < 0.6 else random.choice(CTYPES[1:])
        comps.append([None, delivered + dt.timedelta(days=random.randint(0, 3)), v, po, ctype,
                      random.choice(["Low", "Medium", "Medium", "High"]), random.choice(["Resolved", "Resolved", "Open"])])
pos.sort(key=lambda r: r[1])
for i, p in enumerate(pos):  # renumber chronologically
    old = p[0]; p[0] = f"PO-{26001 + i}"
    for c in comps:
        if c[3] == old: c[3] = "TMP" + p[0]
for c in comps: c[3] = c[3].replace("TMP", "")
comps.sort(key=lambda r: r[1])
for i, c in enumerate(comps): c[0] = f"CMP-{301 + i}"

PS, PE = 5, 1004   # PO rows
CS, CE = 5, 504    # complaint rows
NV = len(VENDORS)

wb = Workbook()
# ---------- Instructions ----------
ws = wb.active; ws.title = "Instructions"
title(ws, "Vendor Evaluation & Cost Analysis Dashboard", "Weighted vendor scorecard (cost, delivery, quality, complaints) and spend analysis by category, vendor and month")
lines = [
    ("How to use", True),
    ("1. Record every purchase order in PO_Data (columns A–I). Category, amount, price index, delay and on-time flag calculate themselves.", False),
    ("2. Record every vendor complaint in Complaints (columns A–G).", False),
    ("3. Scorecard rates each vendor 0–100 on four criteria and combines them with the weights in Settings into one weighted score, rank and rating.", False),
    ("4. Spend_Analysis breaks spend down by category, vendor and month, with top-5 vendors and charts.", False),
    ("5. Add vendors or items in Settings; change weights or rating thresholds there (weights must total 100%).", False),
    ("", False),
    ("Scoring method (0–100 per criterion)", True),
    ("Cost: each PO's Price Index = Unit Price ÷ benchmark price for that item (1.00 = benchmark). Vendor Cost Score = lowest average index among vendors ÷ this vendor's average index × 100, so the cheapest vendor scores 100.", False),
    ("Delivery: On-time % × 100, where a PO is on time if Delivered Date ≤ Promised Date.", False),
    ("Quality: average Quality Rating (1–5 recorded on receipt) ÷ 5 × 100.", False),
    ("Complaints: 100 − (complaints per 10 POs × penalty points), floored at 0. Penalty is set in Settings (default 15).", False),
    ("Weighted Score = Σ criterion score × weight.  Rating: Preferred ≥ 80, Approved ≥ 65, else Under Review (thresholds in Settings).", False),
    ("", False),
    ("Colour legend", True),
    ("Dark blue header = input column   |   Teal header = calculated   |   Light-yellow cell = setting you may change", False),
    ("", False),
    ("Power BI", True),
    ("The PowerBI folder contains the same data as CSV files plus PowerBI_Guide.md with the data model, DAX measures and report layout, so the dashboard can be rebuilt in Power BI Desktop.", False),
    ("", False),
    ("Note", True),
    ("Vendors, prices and orders here are fictional sample data (Jun–Nov 2026) used to demonstrate the model. Replace them with real records; formulas are pre-filled to 1,000 POs and 500 complaints.", False),
]
for i, (t, b) in enumerate(lines, start=4):
    c = ws.cell(row=i, column=1, value=t); c.font = f(b, 11 if b else 10, NAVY if b else "000000")
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 135

# ---------- Settings ----------
S = wb.create_sheet("Settings")
title(S, "Settings & Master Data", "Yellow cells are editable")
header_row(S, 4, ["Criterion", "Weight"], 1)
for i, (k, w) in enumerate(WEIGHTS):
    S.cell(row=5 + i, column=1, value=k); S.cell(row=5 + i, column=2, value=w)
S["A9"] = "Total (must be 100%)"; S["B9"] = "=SUM(B5:B8)"
S["C9"] = '=IF(ROUND(B9,4)=1,"OK","Weights must total 100%")'
style_body(S, 5, 9, 1, 2, {2: "0%"})
for r in range(5, 9): S.cell(row=r, column=2).fill = fill(INPUT_FILL)
S["A9"].font = f(True); S["B9"].font = f(True); S["C9"].font = f(True, color="C00000")
S["B5"].comment = Comment("Assumption: weights chosen for a facilities buyer (cost-led, then delivery and quality). Adjust to your policy.", "Model")
header_row(S, 11, ["Parameter", "Value"], 1)
params = [("Complaint penalty (points per complaint per 10 POs)", 15), ("Preferred threshold (score ≥)", 80), ("Approved threshold (score ≥)", 65)]
for i, (k, v) in enumerate(params):
    S.cell(row=12 + i, column=1, value=k); S.cell(row=12 + i, column=2, value=v)
style_body(S, 12, 14, 1, 2)
for r in range(12, 15): S.cell(row=r, column=2).fill = fill(INPUT_FILL)
PEN, PREF, APPR = "Settings!$B$12", "Settings!$B$13", "Settings!$B$14"
WC, WD, WQ, WK = "Settings!$B$5", "Settings!$B$6", "Settings!$B$7", "Settings!$B$8"

header_row(S, 4, ["Vendor", "Category"], 4)
for i, v in enumerate(VENDORS):
    S.cell(row=5 + i, column=4, value=v[0]); S.cell(row=5 + i, column=5, value=v[1])
style_body(S, 5, 5 + NV - 1, 4, 5)
VEN_R, VCAT_R = f"Settings!$D$5:$D${4 + NV}", f"Settings!$E$5:$E${4 + NV}"

header_row(S, 4, ["Item", "Category", "Benchmark Price (Rs)"], 7)
r = 5
for cat in CATS:
    for item, b in ITEMS[cat]:
        S.cell(row=r, column=7, value=item); S.cell(row=r, column=8, value=cat); S.cell(row=r, column=9, value=b); r += 1
ITEM_END = r - 1 + 16  # spare rows for new items
style_body(S, 5, ITEM_END, 7, 9, {9: "#,##0"})
for rr in range(5, ITEM_END + 1): S.cell(row=rr, column=9).fill = fill(INPUT_FILL)
S["I4"].comment = Comment("Assumption: reference market price per unit (Rs), e.g. average of 3 quotations or rate contract. Used to compare vendor prices across different items.", "Model")
ITEM_R, ICAT_R, BEN_R = f"Settings!$G$5:$G${ITEM_END}", f"Settings!$H$5:$H${ITEM_END}", f"Settings!$I$5:$I${ITEM_END}"
header_row(S, 4, ["Categories"], 11)
for i, c in enumerate(CATS): S.cell(row=5 + i, column=11, value=c)
style_body(S, 5, 4 + len(CATS), 11, 11)
CAT_R = f"Settings!$K$5:$K${4 + len(CATS)}"
header_row(S, 4, ["Complaint Types"], 13)
for i, c in enumerate(CTYPES): S.cell(row=5 + i, column=13, value=c)
style_body(S, 5, 4 + len(CTYPES), 13, 13)
widths(S, [48, 10, 3, 27, 22, 3, 28, 22, 14, 3, 22, 3, 18])

# ---------- PO_Data ----------
P = wb.create_sheet("PO_Data")
title(P, "Purchase Orders", "Columns A–I are inputs; J–Q are calculated")
header_row(P, 4, ["PO No", "PO Date", "Vendor", "Item", "Qty", "Unit Price (Rs)", "Promised Date", "Delivered Date", "Quality Rating (1-5)"], 1, NAVY)
header_row(P, 4, ["Category", "Amount (Rs)", "Benchmark Price (Rs)", "Price Index", "Delay (days)", "On Time?", "Month", "Vendor Category"], 10, TEAL)
widths(P, [11, 12, 25, 27, 7, 11, 12, 12, 10, 20, 13, 12, 9, 9, 9, 11, 20])
for r in range(PS, PE + 1):
    if r - PS < len(pos):
        for c, v in enumerate(pos[r - PS], start=1): P.cell(row=r, column=c, value=v)
    P[f"J{r}"] = f'=IF($D{r}="","",IFERROR(INDEX({ICAT_R},MATCH($D{r},{ITEM_R},0)),"Unknown item"))'
    P[f"K{r}"] = f'=IF(OR($E{r}="",$F{r}=""),"",$E{r}*$F{r})'
    P[f"L{r}"] = f'=IF($D{r}="","",IFERROR(INDEX({BEN_R},MATCH($D{r},{ITEM_R},0)),""))'
    P[f"M{r}"] = f'=IF(OR($F{r}="",$L{r}="",$L{r}=0),"",$F{r}/$L{r})'
    P[f"N{r}"] = f'=IF(OR($G{r}="",$H{r}=""),"",MAX(0,$H{r}-$G{r}))'
    P[f"O{r}"] = f'=IF($N{r}="","",IF($N{r}=0,"Yes","No"))'
    P[f"P{r}"] = f'=IF($B{r}="","",DATE(YEAR($B{r}),MONTH($B{r}),1))'
    P[f"Q{r}"] = f'=IF($C{r}="","",IFERROR(INDEX({VCAT_R},MATCH($C{r},{VEN_R},0)),""))'
style_body(P, PS, PE, 1, 17, {2: "dd-mmm-yyyy", 6: "#,##0", 7: "dd-mmm-yyyy", 8: "dd-mmm-yyyy",
                              11: "#,##0", 12: "#,##0", 13: "0.00", 14: "0", 16: "mmm-yyyy"})
for rr in range(PS, PE + 1):
    for c in range(10, 18): P.cell(row=rr, column=c).fill = fill(GREY)
for col, rng in [("C", VEN_R), ("D", ITEM_R)]:
    dv = DataValidation(type="list", formula1=f"={rng}", allow_blank=True); dv.add(f"{col}{PS}:{col}{PE}"); P.add_data_validation(dv)
dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True, showErrorMessage=True, error="Rating must be 1 to 5")
dv.add(f"I{PS}:I{PE}"); P.add_data_validation(dv)
P.conditional_formatting.add(f"O{PS}:O{PE}", FormulaRule(formula=[f'$O{PS}="No"'], fill=fill("F8CBAD")))
P.freeze_panes = "B5"; P.auto_filter.ref = f"A4:Q{PE}"
PL = lambda col: f"PO_Data!${col}${PS}:${col}${PE}"

# ---------- Complaints ----------
C = wb.create_sheet("Complaints")
title(C, "Vendor Complaints Log", "One row per complaint")
header_row(C, 4, ["Complaint ID", "Date", "Vendor", "PO No", "Complaint Type", "Severity", "Status"], 1, NAVY)
widths(C, [13, 12, 26, 11, 16, 10, 10])
for r, row in enumerate(comps, start=CS):
    for c, v in enumerate(row, start=1): C.cell(row=r, column=c, value=v)
style_body(C, CS, CE, 1, 7, {2: "dd-mmm-yyyy"})
dv = DataValidation(type="list", formula1=f"={VEN_R}", allow_blank=True); dv.add(f"C{CS}:C{CE}"); C.add_data_validation(dv)
dv = DataValidation(type="list", formula1=f"=Settings!$M$5:$M${4 + len(CTYPES)}", allow_blank=True); dv.add(f"E{CS}:E{CE}"); C.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"Low,Medium,High"', allow_blank=True); dv.add(f"F{CS}:F{CE}"); C.add_data_validation(dv)
dv = DataValidation(type="list", formula1='"Open,Resolved"', allow_blank=True); dv.add(f"G{CS}:G{CE}"); C.add_data_validation(dv)
C.freeze_panes = "B5"; C.auto_filter.ref = f"A4:G{CE}"
CL = lambda col: f"Complaints!${col}${CS}:${col}${CE}"

# ---------- Scorecard ----------
SC = wb.create_sheet("Scorecard", 1)
title(SC, "Weighted Vendor Scorecard", None)
SC["A2"] = '="Weights: Cost "&TEXT(Settings!B5,"0%")&" | Delivery "&TEXT(Settings!B6,"0%")&" | Quality "&TEXT(Settings!B7,"0%")&" | Complaints "&TEXT(Settings!B8,"0%")&"   ("&Settings!C9&")"'
SC["A2"].font = f(italic=True, color="595959")
H = ["Vendor", "Category", "POs", "Total Spend (Rs)", "Avg Price Index", "On-Time %", "Avg Quality (1-5)", "Complaints", "Complaints per 10 POs",
     "Cost Score", "Delivery Score", "Quality Score", "Complaint Score", "Weighted Score", "Rank", "Rating", "key"]
header_row(SC, 4, H[:9], 1, NAVY)
header_row(SC, 4, H[9:16], 10, TEAL)
header_row(SC, 4, ["helper"], 17, "808080")
widths(SC, [27, 21, 6, 14, 10, 10, 10, 11, 11, 9, 9, 9, 10, 10, 7, 13, 6])
R0, R1 = 5, 4 + NV
for i in range(NV):
    r = R0 + i
    SC[f"A{r}"] = f"=Settings!D{5 + i}"
    SC[f"B{r}"] = f"=Settings!E{5 + i}"
    SC[f"C{r}"] = f'=COUNTIF({PL("C")},A{r})'
    SC[f"D{r}"] = f'=SUMIF({PL("C")},A{r},{PL("K")})'
    SC[f"E{r}"] = f'=IFERROR(AVERAGEIF({PL("C")},A{r},{PL("M")}),"")'
    SC[f"F{r}"] = f'=IFERROR(COUNTIFS({PL("C")},A{r},{PL("O")},"Yes")/C{r},"")'
    SC[f"G{r}"] = f'=IFERROR(AVERAGEIF({PL("C")},A{r},{PL("I")}),"")'
    SC[f"H{r}"] = f'=COUNTIF({CL("C")},A{r})'
    SC[f"I{r}"] = f'=IFERROR(H{r}/C{r}*10,"")'
    SC[f"J{r}"] = f'=IFERROR(MIN($E${R0}:$E${R1})/E{r}*100,"")'
    SC[f"K{r}"] = f'=IFERROR(F{r}*100,"")'
    SC[f"L{r}"] = f'=IFERROR(G{r}/5*100,"")'
    SC[f"M{r}"] = f'=IFERROR(MAX(0,100-I{r}*{PEN}),"")'
    SC[f"N{r}"] = f'=IFERROR(J{r}*{WC}+K{r}*{WD}+L{r}*{WQ}+M{r}*{WK},"")'
    SC[f"O{r}"] = f'=IFERROR(RANK(N{r},$N${R0}:$N${R1},0),"")'
    SC[f"P{r}"] = f'=IF(N{r}="","No data",IF(N{r}>={PREF},"Preferred",IF(N{r}>={APPR},"Approved","Under Review")))'
    SC[f"Q{r}"] = f'=IFERROR(N{r}+ROW()/1000000,-1)'
style_body(SC, R0, R1, 1, 17, {4: "#,##0", 5: "0.00", 6: "0.0%", 7: "0.00", 9: "0.0", 10: "0.0", 11: "0.0", 12: "0.0", 13: "0.0", 14: "0.0", 17: "0.000000"})
for r in range(R0, R1 + 1):
    SC[f"N{r}"].font = f(True)
    SC[f"Q{r}"].font = f(color="808080", size=8)
SC.conditional_formatting.add(f"N{R0}:N{R1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, color="5B9BD5"))
SC.conditional_formatting.add(f"P{R0}:P{R1}", FormulaRule(formula=[f'$P{R0}="Preferred"'], fill=fill("C6EFCE"), font=Font(name=FONT, color="006100", bold=True)))
SC.conditional_formatting.add(f"P{R0}:P{R1}", FormulaRule(formula=[f'$P{R0}="Approved"'], fill=fill("FFEB9C")))
SC.conditional_formatting.add(f"P{R0}:P{R1}", FormulaRule(formula=[f'$P{R0}="Under Review"'], fill=fill("F8CBAD"), font=Font(name=FONT, color="9C0006", bold=True)))

# Ranked view + best per category
k0 = R1 + 3
SC[f"A{k0 - 1}"] = "Vendor Ranking (highest weighted score first)"; SC[f"A{k0 - 1}"].font = f(True, 11, NAVY)
header_row(SC, k0, ["Rank", "Vendor", "Category", "Weighted Score", "Rating"], 1)
for i in range(NV):
    r = k0 + 1 + i
    m = f"MATCH(LARGE($Q${R0}:$Q${R1},{i + 1}),$Q${R0}:$Q${R1},0)"
    SC[f"A{r}"] = i + 1
    SC[f"B{r}"] = f"=INDEX($A${R0}:$A${R1},{m})"
    SC[f"C{r}"] = f"=INDEX($B${R0}:$B${R1},{m})"
    SC[f"D{r}"] = f"=INDEX($N${R0}:$N${R1},{m})"
    SC[f"E{r}"] = f"=INDEX($P${R0}:$P${R1},{m})"
style_body(SC, k0 + 1, k0 + NV, 1, 5, {4: "0.0"})
b0 = k0 + NV + 3
SC[f"A{b0 - 1}"] = "Recommended Vendor per Category"; SC[f"A{b0 - 1}"].font = f(True, 11, NAVY)
header_row(SC, b0, ["Category", "Best Vendor", "Best Score", "Vendors Compared", "Category Spend (Rs)"], 1)
for i, cat in enumerate(CATS):
    r = b0 + 1 + i
    SC[f"A{r}"] = f"=Settings!K{5 + i}"
    SC[f"C{r}"] = f'=IFERROR(_xlfn.MAXIFS($N${R0}:$N${R1},$B${R0}:$B${R1},A{r}),"")'
    SC[f"B{r}"] = (f'=IFERROR(INDEX($A${R0}:$A${R1},MATCH(_xlfn.MAXIFS($Q${R0}:$Q${R1},$B${R0}:$B${R1},A{r}),$Q${R0}:$Q${R1},0)),"")')
    SC[f"D{r}"] = f'=COUNTIF($B${R0}:$B${R1},A{r})'
    SC[f"E{r}"] = f'=SUMIF({PL("J")},A{r},{PL("K")})'
style_body(SC, b0 + 1, b0 + len(CATS), 1, 5, {3: "0.0", 5: "#,##0"})
ch = BarChart(); ch.type = "bar"; ch.title = "Weighted Vendor Score (ranked)"; ch.height = 9; ch.width = 15
ch.add_data(Reference(SC, min_col=4, min_row=k0, max_row=k0 + NV), titles_from_data=True)
ch.set_categories(Reference(SC, min_col=2, min_row=k0 + 1, max_row=k0 + NV))
ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 100; ch.legend = None
chart_fix(ch); SC.add_chart(ch, f"G{k0 - 1}")
SC.freeze_panes = "B5"
SC.page_setup.orientation = "landscape"; SC.page_setup.fitToWidth = 1; SC.page_setup.fitToHeight = 0
SC.sheet_properties.pageSetUpPr.fitToPage = True

# ---------- Spend_Analysis ----------
SA = wb.create_sheet("Spend_Analysis", 2)
title(SA, "Spend Analysis", None)
SA["A2"] = '="Total spend Rs "&TEXT(SUM(PO_Data!K5:K1004),"#,##0")&" across "&COUNT(PO_Data!B5:B1004)&" POs"'
SA["A2"].font = f(italic=True, color="595959")
widths(SA, [24, 14, 10, 8, 13, 3] + [12] * 8)
kpi(SA, 4, 1, "Total Spend (Rs)", f'=SUM({PL("K")})', "#,##0")
kpi(SA, 4, 3, "Purchase Orders", f'=COUNT({PL("B")})', "0")
kpi(SA, 4, 5, "On-Time Delivery %", f'=IFERROR(COUNTIF({PL("O")},"Yes")/COUNTIF({PL("O")},"?*"),0)', "0.0%")
kpi(SA, 4, 7, "Avg Price Index", f'=IFERROR(AVERAGE({PL("M")}),0)', "0.00")
kpi(SA, 4, 9, "Complaints", f'=COUNTA({CL("A")})', "0")
SA.column_dimensions["F"].width = 12

c0 = 8
header_row(SA, c0, ["Category", "Spend (Rs)", "Share %", "POs", "Avg PO (Rs)"], 1)
for i in range(len(CATS)):
    r = c0 + 1 + i
    SA[f"A{r}"] = f"=Settings!K{5 + i}"
    SA[f"B{r}"] = f'=SUMIF({PL("J")},A{r},{PL("K")})'
    SA[f"C{r}"] = f'=IFERROR(B{r}/SUM($B${c0 + 1}:$B${c0 + len(CATS)}),0)'
    SA[f"D{r}"] = f'=COUNTIF({PL("J")},A{r})'
    SA[f"E{r}"] = f'=IFERROR(B{r}/D{r},0)'
c1 = c0 + len(CATS)
SA[f"A{c1 + 1}"] = "Total"
SA[f"B{c1 + 1}"] = f"=SUM(B{c0 + 1}:B{c1})"; SA[f"C{c1 + 1}"] = f"=SUM(C{c0 + 1}:C{c1})"
SA[f"D{c1 + 1}"] = f"=SUM(D{c0 + 1}:D{c1})"; SA[f"E{c1 + 1}"] = f"=IFERROR(B{c1 + 1}/D{c1 + 1},0)"
style_body(SA, c0 + 1, c1 + 1, 1, 5, {2: "#,##0", 3: "0.0%", 5: "#,##0"})
for c in range(1, 6): SA.cell(row=c1 + 1, column=c).font = f(True); SA.cell(row=c1 + 1, column=c).fill = fill(LIGHT)

v0 = c1 + 4
SA[f"A{v0 - 1}"] = "Spend by Vendor"; SA[f"A{v0 - 1}"].font = f(True, 11, NAVY)
header_row(SA, v0, ["Vendor", "Spend (Rs)", "Share %", "POs", "Weighted Score"], 1)
for i in range(NV):
    r = v0 + 1 + i
    SA[f"A{r}"] = f"=Scorecard!A{R0 + i}"
    SA[f"B{r}"] = f"=Scorecard!D{R0 + i}"
    SA[f"C{r}"] = f'=IFERROR(B{r}/SUM($B${v0 + 1}:$B${v0 + NV}),0)'
    SA[f"D{r}"] = f"=Scorecard!C{R0 + i}"
    SA[f"E{r}"] = f"=Scorecard!N{R0 + i}"
    SA[f"Z{r}"] = f"=B{r}+ROW()/1000000"   # tie-break helper for top-5
style_body(SA, v0 + 1, v0 + NV, 1, 5, {2: "#,##0", 3: "0.0%", 5: "0.0"})
SA.column_dimensions["Z"].hidden = True
t0 = v0 + NV + 3
SA[f"A{t0 - 1}"] = "Top 5 Vendors by Spend"; SA[f"A{t0 - 1}"].font = f(True, 11, NAVY)
header_row(SA, t0, ["Vendor", "Spend (Rs)", "Share %", "Weighted Score", "Rating"], 1)
for i in range(5):
    r = t0 + 1 + i
    m = f"MATCH(LARGE($Z${v0 + 1}:$Z${v0 + NV},{i + 1}),$Z${v0 + 1}:$Z${v0 + NV},0)"
    SA[f"A{r}"] = f"=INDEX($A${v0 + 1}:$A${v0 + NV},{m})"
    SA[f"B{r}"] = f"=INDEX($B${v0 + 1}:$B${v0 + NV},{m})"
    SA[f"C{r}"] = f"=INDEX($C${v0 + 1}:$C${v0 + NV},{m})"
    SA[f"D{r}"] = f"=INDEX($E${v0 + 1}:$E${v0 + NV},{m})"
    SA[f"E{r}"] = f"=INDEX(Scorecard!$P${R0}:$P${R1},MATCH(A{r},Scorecard!$A${R0}:$A${R1},0))"
style_body(SA, t0 + 1, t0 + 5, 1, 5, {2: "#,##0", 3: "0.0%", 4: "0.0"})

m0 = t0 + 9
SA[f"A{m0 - 1}"] = "Monthly Spend by Category (Rs)"; SA[f"A{m0 - 1}"].font = f(True, 11, NAVY)
header_row(SA, m0, ["Month"] + [f"=Settings!K{5 + i}" for i in range(len(CATS))] + ["Total"], 1)
NM = 6
for i in range(NM):
    r = m0 + 1 + i
    SA[f"A{r}"] = f'=DATE(YEAR(MIN({PL("B")})),MONTH(MIN({PL("B")})),1)' if i == 0 else f"=EDATE(A{r - 1},1)"
    for j in range(len(CATS)):
        col = get_column_letter(2 + j)
        SA[f"{col}{r}"] = f'=SUMIFS({PL("K")},{PL("P")},$A{r},{PL("J")},{col}${m0})'
    tc = get_column_letter(2 + len(CATS))
    SA[f"{tc}{r}"] = f"=SUM(B{r}:{get_column_letter(1 + len(CATS))}{r})"
style_body(SA, m0 + 1, m0 + NM, 1, 2 + len(CATS), {1: "mmm-yyyy", **{c: "#,##0" for c in range(2, 3 + len(CATS))}})

# charts
pie = PieChart(); pie.title = "Spend Share by Category"; pie.height = 8; pie.width = 12
pie.add_data(Reference(SA, min_col=2, min_row=c0, max_row=c1), titles_from_data=True)
pie.set_categories(Reference(SA, min_col=1, min_row=c0 + 1, max_row=c1))
pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
chart_fix(pie, axes=False); SA.add_chart(pie, "G8")
vb = BarChart(); vb.type = "bar"; vb.title = "Spend by Vendor (Rs)"; vb.height = 9; vb.width = 15; vb.legend = None
vb.add_data(Reference(SA, min_col=2, min_row=v0, max_row=v0 + NV), titles_from_data=True)
vb.set_categories(Reference(SA, min_col=1, min_row=v0 + 1, max_row=v0 + NV))
vb.x_axis.scaling.orientation = "maxMin"; vb.y_axis.number_format = "#,##0"
chart_fix(vb); SA.add_chart(vb, f"G{v0 - 1}")
mb = BarChart(); mb.type = "col"; mb.grouping = "stacked"; mb.overlap = 100
mb.title = "Monthly Spend by Category (Rs)"; mb.height = 9; mb.width = 18
mb.add_data(Reference(SA, min_col=2, max_col=1 + len(CATS), min_row=m0, max_row=m0 + NM), titles_from_data=True)
mb.set_categories(Reference(SA, min_col=1, min_row=m0 + 1, max_row=m0 + NM))
mb.x_axis.number_format = "mmm-yy"; mb.y_axis.number_format = "#,##0"
chart_fix(mb); SA.add_chart(mb, f"A{m0 + NM + 2}")
SA.page_setup.orientation = "landscape"; SA.page_setup.fitToWidth = 1; SA.page_setup.fitToHeight = 0
SA.sheet_properties.pageSetUpPr.fitToPage = True

wb.active = 1
wb.save(OUT_XLSX)

# ---------- Power BI CSVs ----------
def wcsv(name, header, rows):
    with open(os.path.join(CSV_DIR, name), "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh); w.writerow(header); w.writerows(rows)
bench = {it: (cat, b) for cat in CATS for it, b in ITEMS[cat]}
wcsv("Vendors.csv", ["Vendor", "Category"], [(v[0], v[1]) for v in VENDORS])
wcsv("Items.csv", ["Item", "Category", "BenchmarkPrice"], [(it, c, b) for it, (c, b) in bench.items()])
wcsv("PurchaseOrders.csv", ["PONo", "PODate", "Vendor", "Item", "Qty", "UnitPrice", "PromisedDate", "DeliveredDate", "QualityRating"],
     [(p[0], p[1].isoformat(), p[2], p[3], p[4], p[5], p[6].isoformat(), p[7].isoformat(), p[8]) for p in pos])
wcsv("Complaints.csv", ["ComplaintID", "ComplaintDate", "Vendor", "PONo", "ComplaintType", "Severity", "Status"],
     [(c[0], c[1].isoformat(), *c[2:]) for c in comps])
wcsv("Weights.csv", ["Criterion", "Weight"], WEIGHTS)
print("saved", len(pos), "POs", len(comps), "complaints")
