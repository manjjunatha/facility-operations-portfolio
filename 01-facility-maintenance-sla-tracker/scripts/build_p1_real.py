"""Project 1 on REAL data: NYC 311 / HPD building-maintenance complaints, ZIP 11226 (Brooklyn), Sep-Nov 2026."""
import os
import sys, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *
import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

SRC, OUT, CLEAN_CSV = sys.argv[1], sys.argv[2], sys.argv[3]
ASOF = dt.datetime(2026, 11, 25)  # date the extract was downloaded (status is as of this date)

# ---------------- clean the real data ----------------
df = pd.read_csv(SRC, low_memory=False)
for c in ["created_date", "closed_date"]:
    df[c] = pd.to_datetime(df[c], errors="coerce")
df = df.sort_values("created_date").reset_index(drop=True)

def outcome(s):
    s = s if isinstance(s, str) else ""
    rules = [
        ("No violation found", ["No violations were issued", "did not violate", "Heat was not required"]),
        ("Violation issued", ["Violations were issued", "Violations were previously issued"]),
        ("Resolved / confirmed corrected", ["had been restored", "condition was corrected", "verified that the following conditions were corrected"]),
        ("No access to inspect", ["not able to gain access", "unable to access"]),
        ("Unable to inspect", ["unable to complete the inspection"]),
        ("Owner notified - open", ["still open"]),
    ]
    for k, pats in rules:
        if any(p in s for p in pats):
            return k
    return "Other / see HPD" if s else "Pending"

df["outcome"] = df["resolution_description"].map(outcome)
df["address"] = df["incident_address"].str.replace(r"\s+", " ", regex=True).str.strip().str.title()
df["descriptor"] = df["descriptor"].fillna("").str.title()
df["channel"] = df["open_data_channel_type"].fillna("UNKNOWN").str.title()
clean = df[["unique_key", "created_date", "address", "complaint_type", "descriptor", "channel", "status", "closed_date", "outcome"]]
clean.to_csv(CLEAN_CSV, index=False)

CATS = sorted(df["complaint_type"].unique(), key=lambda c: -(df["complaint_type"] == c).sum())
PRIO_OF = {"HEAT/HOT WATER": "Critical", "ELEVATOR": "Critical",
           "WATER LEAK": "High", "SAFETY": "High", "ELECTRIC": "High", "PLUMBING": "High",
           "DOOR/WINDOW": "Medium", "UNSANITARY CONDITION": "Medium", "APPLIANCE": "Medium",
           "FLOORING/STAIRS": "Medium", "GENERAL": "Medium",
           "PAINT/PLASTER": "Low", "OUTSIDE BUILDING": "Low"}
PRIOS = [("Critical", 2), ("High", 7), ("Medium", 14), ("Low", 21)]
OUTCOMES = ["Violation issued", "No violation found", "Resolved / confirmed corrected", "Owner notified - open",
            "Unable to inspect", "No access to inspect", "Other / see HPD", "Pending"]
CHANNELS = ["Online", "Phone", "Mobile", "Unknown"]
BLDGS = sorted(df["address"].dropna().unique())
NB = len(BLDGS)
N = len(df)
DS, DE = 5, 5 + N + 1000 - 1     # room for 1,000 more tickets
PS, PE = 5, 104

wb = Workbook()
# ---------------- Instructions ----------------
ws = wb.active; ws.title = "Instructions"
title(ws, "Facility Maintenance & SLA Tracker  (real data edition)",
      "Real building-maintenance complaints: NYC 311 / HPD, ZIP 11226 (Flatbush, Brooklyn), 1 Sep – 6 Nov 2026")
lines = [
    ("Data source", True),
    ("NYC Open Data – '311 Service Requests from 2010 to Present' (dataset erm2-nwe9), filtered to agency HPD (Dept. of Housing Preservation & Development), incident ZIP 11226, created on/after 1-Sep-2026.", False),
    ("Extract: 5,000 complaints (the download limit), covering 1-Sep-2026 to 6-Nov-2026 across 601 residential buildings. Downloaded 25-Nov-2026; Status and Closed Date are as of that day.", False),
    ("Licence: NYC Open Data is free to use (NYC Open Data Terms of Use). Cite as: NYC OpenData, 311 Service Requests, City of New York.", False),
    ("", False),
    ("What was done to the raw data (cleaning)", True),
    ("Kept 9 useful fields of 44; trimmed spaces in addresses; sorted by date. HPD's long resolution text was grouped into 7 short Outcome labels by keyword (e.g. text containing 'Violations were issued' → 'Violation issued').", False),
    ("", False),
    ("Assumptions (editable in Lists)", True),
    ("The 311 data has no priority or SLA, so each Category is mapped to a Priority, and each Priority has an SLA target: Critical 2 days, High 7, Medium 14, Low 21.", False),
    ("Heat/Hot Water and Elevator are Critical (tenant safety in heating season); Water Leak, Safety, Electric, Plumbing are High; Paint/Plaster and Outside Building are Low; the rest Medium.", False),
    ("These targets are this tracker's own service standards, not official NYC/HPD targets. Change them in Lists and every figure updates.", False),
    ("PM_Schedule is an illustrative sample (preventive maintenance is not part of 311 data); its assets are linked to the buildings with the most complaints.", False),
    ("", False),
    ("How it works", True),
    ("Ticket_Log columns A–I are data/inputs; J–P are formulas: Priority (from Category), SLA Target, Due Date, TAT (days), Days Open, SLA Status, Breach flag.", False),
    ("SLA Status: closed tickets are Met if TAT ≤ target else Breached; open tickets are Open - Overdue if the As-of date is past the Due Date, else Open - On Track. As-of date is set in Lists (default 25-Nov-2026, the download date).", False),
    ("SLA Compliance % = Met ÷ (Met + Breached).  Dashboard uses COUNTIFS / AVERAGEIFS / MEDIAN / INDEX-MATCH / LARGE.", False),
    ("To add new tickets, type them below the last row of Ticket_Log; formulas are pre-filled for 1,000 more rows. New buildings must be added to the Lists building register.", False),
    ("", False),
    ("Colour legend", True),
    ("Dark blue header = data/input  |  Teal header = calculated  |  Light-yellow cells (Lists) = settings you can change", False),
]
for i, (t, b) in enumerate(lines, start=4):
    c = ws.cell(row=i, column=1, value=t); c.font = f(b, 11 if b else 10, NAVY if b else "000000")
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 140

# ---------------- Lists ----------------
L = wb.create_sheet("Lists")
title(L, "Lists & Settings", "Yellow cells are editable settings")
header_row(L, 3, ["Category", "Priority"], 1)
for i, c in enumerate(CATS):
    L.cell(row=4 + i, column=1, value=c); L.cell(row=4 + i, column=2, value=PRIO_OF.get(c, "Medium"))
CAT_END = 4 + len(CATS) - 1 + 5
style_body(L, 4, CAT_END, 1, 2)
for r in range(4, CAT_END + 1): L.cell(row=r, column=2).fill = fill(INPUT_FILL)
L["B3"].comment = Comment("Assumption: priority per complaint category, set by this tracker (311 data has no priority).", "Tracker")
header_row(L, 3, ["Priority", "SLA Target (days)"], 4)
for i, (p, d) in enumerate(PRIOS):
    L.cell(row=4 + i, column=4, value=p); L.cell(row=4 + i, column=5, value=d)
style_body(L, 4, 7, 4, 5)
for r in range(4, 8): L.cell(row=r, column=5).fill = fill(INPUT_FILL)
L["E3"].comment = Comment("Assumption: tracker's own SLA targets, not official HPD targets.", "Tracker")
L["D10"] = "As-of date"; L["D10"].font = f(True)
L["E10"] = ASOF; L["E10"].number_format = "dd-mmm-yyyy"; L["E10"].fill = fill(INPUT_FILL); L["E10"].border = BORDER
L["E10"].comment = Comment("Date the data was downloaded; open tickets are judged overdue against this date. Use =TODAY() if you keep the log live.", "Tracker")
L["D11"] = "PM alert window (days)"; L["D11"].font = f(True)
L["E11"] = 7; L["E11"].fill = fill(INPUT_FILL); L["E11"].border = BORDER
header_row(L, 3, ["Status"], 7)
for i, s in enumerate(["Open", "In Progress", "Closed"]): L.cell(row=4 + i, column=7, value=s)
style_body(L, 4, 6, 7, 7)
header_row(L, 3, ["Channel"], 9)
for i, s in enumerate(CHANNELS): L.cell(row=4 + i, column=9, value=s)
style_body(L, 4, 7, 9, 9)
header_row(L, 3, ["Outcome"], 11)
for i, s in enumerate(OUTCOMES): L.cell(row=4 + i, column=11, value=s)
style_body(L, 4, 3 + len(OUTCOMES), 11, 11)
header_row(L, 3, ["Building Register (address)", "Complaints", "rank helper"], 13)
BS, BE = 4, 4 + NB - 1 + 50
for i, b in enumerate(BLDGS): L.cell(row=BS + i, column=13, value=b)
widths(L, [22, 10, 3, 11, 12, 3, 12, 3, 10, 3, 28, 3, 30, 11, 11])
CAT_R, PCAT_R = f"Lists!$A$4:$A${CAT_END}", f"Lists!$B$4:$B${CAT_END}"
PRI_R, SLA_R = "Lists!$D$4:$D$7", "Lists!$E$4:$E$7"
ASOF_C, ALERT = "Lists!$E$10", "Lists!$E$11"
BLD_R = f"Lists!$M${BS}:$M${BE}"

# ---------------- Ticket_Log ----------------
T = wb.create_sheet("Ticket_Log")
title(T, "Maintenance Ticket Log  (NYC 311 / HPD, ZIP 11226)", "Columns A–I = real complaint data; J–P = calculated")
header_row(T, 4, ["Ticket ID (311 SR)", "Raised (date-time)", "Building", "Category", "Sub-Category", "Channel", "Status", "Closed (date-time)", "Outcome"], 1, NAVY)
header_row(T, 4, ["Priority", "SLA Target (days)", "Due Date", "TAT (days)", "Days Open", "SLA Status", "SLA Breach?"], 10, TEAL)
widths(T, [12, 16, 26, 20, 20, 9, 9, 16, 26, 9, 9, 16, 9, 9, 15, 9])
for i, row in enumerate(clean.itertuples(index=False)):
    r = DS + i
    vals = [int(row.unique_key), row.created_date.to_pydatetime(), row.address, row.complaint_type, row.descriptor,
            row.channel, row.status, row.closed_date.to_pydatetime() if pd.notna(row.closed_date) else None, row.outcome]
    for c, v in enumerate(vals, start=1):
        T.cell(row=r, column=c, value=v)
for r in range(DS, DE + 1):
    T[f"J{r}"] = f'=IF($D{r}="","",IFERROR(INDEX({PCAT_R},MATCH($D{r},{CAT_R},0)),"Medium"))'
    T[f"K{r}"] = f'=IF($J{r}="","",IFERROR(INDEX({SLA_R},MATCH($J{r},{PRI_R},0)),""))'
    T[f"L{r}"] = f'=IF(OR($B{r}="",$K{r}=""),"",$B{r}+$K{r})'
    T[f"M{r}"] = f'=IF(OR($B{r}="",$H{r}=""),"",$H{r}-$B{r})'
    T[f"N{r}"] = f'=IF($B{r}="","",IF($H{r}="",{ASOF_C}-$B{r},""))'
    T[f"O{r}"] = (f'=IF(OR($B{r}="",$K{r}=""),"",IF($H{r}<>"",IF($M{r}<=$K{r},"Met","Breached"),'
                  f'IF({ASOF_C}>$L{r},"Open - Overdue","Open - On Track")))')
    T[f"P{r}"] = f'=IF($O{r}="","",IF(OR($O{r}="Breached",$O{r}="Open - Overdue"),"Yes","No"))'
style_body(T, DS, DE, 1, 16, {2: "dd-mmm-yy hh:mm", 8: "dd-mmm-yy hh:mm", 12: "dd-mmm-yy hh:mm", 11: "0", 13: "0.0", 14: "0.0"})
for r in range(DS, DE + 1):
    for c in range(10, 17): T.cell(row=r, column=c).fill = fill(GREY)
for col, rng in [("C", BLD_R), ("D", CAT_R), ("F", "Lists!$I$4:$I$7"), ("G", "Lists!$G$4:$G$6"), ("I", f"Lists!$K$4:$K${3 + len(OUTCOMES)}")]:
    dv = DataValidation(type="list", formula1=f"={rng}", allow_blank=True); dv.add(f"{col}{DS}:{col}{DE}"); T.add_data_validation(dv)
red, green, amber = fill("F8CBAD"), fill("C6EFCE"), fill("FFEB9C")
T.conditional_formatting.add(f"O{DS}:P{DE}", FormulaRule(formula=[f'OR($O{DS}="Breached",$O{DS}="Open - Overdue")'], fill=red, font=Font(name=FONT, color="9C0006", bold=True)))
T.conditional_formatting.add(f"O{DS}:P{DE}", FormulaRule(formula=[f'$O{DS}="Met"'], fill=green))
T.conditional_formatting.add(f"O{DS}:P{DE}", FormulaRule(formula=[f'$O{DS}="Open - On Track"'], fill=amber))
T.freeze_panes = "C5"; T.auto_filter.ref = f"A4:P{DE}"
TL = lambda col: f"Ticket_Log!${col}${DS}:${col}${DE}"

# building counts in Lists (for Top-10)
for r in range(BS, BE + 1):
    L[f"N{r}"] = f'=IF(M{r}="","",COUNTIF({TL("C")},M{r}))'
    L[f"O{r}"] = f'=IF(M{r}="",-1,N{r}+(10000-ROW())/1000000)'
style_body(L, BS, BE, 13, 15, {15: "0.000000"})
for r in range(BS, BE + 1): L[f"O{r}"].font = f(size=8, color="808080")

# ---------------- PM_Schedule (illustrative) ----------------
P = wb.create_sheet("PM_Schedule")
title(P, "Preventive Maintenance Schedule  (illustrative sample – not from 311 data)",
      "Assets in the buildings with the most complaints; enter Last Done after each PM job")
header_row(P, 4, ["PM ID", "Asset", "Building", "PM Task", "Frequency (days)", "Linked Complaint Category", "Last Done"], 1, NAVY)
header_row(P, 4, ["Next Due", "Days to Due", "Alert"], 8, TEAL)
widths(P, [9, 24, 26, 32, 11, 22, 13, 13, 11, 12])
top = df["address"].value_counts().index.tolist()
pm = [
    ("Boiler & burner", top[0], "Burner tune-up, safety controls test", 30, "HEAT/HOT WATER", -6),
    ("Boiler & burner", top[1], "Burner tune-up, safety controls test", 30, "HEAT/HOT WATER", 3),
    ("Boiler & burner", top[2], "Burner tune-up, safety controls test", 30, "HEAT/HOT WATER", 12),
    ("Hot-water heater", top[0], "Flush tank, check thermostat", 90, "HEAT/HOT WATER", 20),
    ("Elevator", top[1], "Monthly service & safety check", 30, "ELEVATOR", -2),
    ("Roof drains & gutters", top[3], "Clear drains before winter", 90, "WATER LEAK", 5),
    ("Common-area plumbing", top[4], "Inspect risers, valves, leaks", 90, "PLUMBING", 40),
    ("Electrical panels", top[0], "Thermal scan & tightening", 180, "ELECTRIC", 60),
    ("Smoke & CO detectors", top[2], "Test all units, replace batteries", 180, "SAFETY", -10),
    ("Pest control", top[0], "Building-wide treatment", 30, "UNSANITARY CONDITION", 4),
    ("Pest control", top[5], "Building-wide treatment", 30, "UNSANITARY CONDITION", 15),
    ("Intercom / buzzer", top[6], "Test all units & door release", 90, "DOOR/WINDOW", -3),
    ("Entrance doors & locks", top[1], "Lubricate, adjust closers", 90, "DOOR/WINDOW", 25),
    ("Stairs & handrails", top[7], "Inspect treads, rails, lighting", 90, "FLOORING/STAIRS", 30),
]
for i, (a, b, t, fq, cat, off) in enumerate(pm):
    r = PS + i
    last = ASOF - dt.timedelta(days=fq - off)
    for c, v in enumerate([f"PM-{101 + i}", a, b, t, fq, cat, last], start=1): P.cell(row=r, column=c, value=v)
for r in range(PS, PE + 1):
    P[f"H{r}"] = f'=IF(OR($E{r}="",$G{r}=""),"",$G{r}+$E{r})'
    P[f"I{r}"] = f'=IF($H{r}="","",$H{r}-TODAY())'
    P[f"J{r}"] = f'=IF($I{r}="","",IF($I{r}<0,"OVERDUE",IF($I{r}<={ALERT},"Due Soon","OK")))'
style_body(P, PS, PE, 1, 10, {7: "dd-mmm-yyyy", 8: "dd-mmm-yyyy", 9: "0", 5: "0"})
for r in range(PS, PE + 1):
    for c in range(8, 11): P.cell(row=r, column=c).fill = fill(GREY)
P.conditional_formatting.add(f"H{PS}:J{PE}", FormulaRule(formula=[f'$J{PS}="OVERDUE"'], fill=red, font=Font(name=FONT, color="9C0006", bold=True)))
P.conditional_formatting.add(f"H{PS}:J{PE}", FormulaRule(formula=[f'$J{PS}="Due Soon"'], fill=amber))
P.conditional_formatting.add(f"H{PS}:J{PE}", FormulaRule(formula=[f'$J{PS}="OK"'], fill=green))
P.freeze_panes = "C5"
PL = lambda col: f"PM_Schedule!${col}${PS}:${col}${PE}"

# ---------------- Dashboard ----------------
D = wb.create_sheet("Dashboard", 1)
title(D, "Facility Maintenance KPI Dashboard — NYC 311 / HPD, ZIP 11226", None)
D["A2"] = (f'="Real complaints "&TEXT(MIN({TL("B")}),"dd-mmm-yyyy")&" to "&TEXT(MAX({TL("B")}),"dd-mmm-yyyy")'
           f'&"  |  status as of "&TEXT({ASOF_C},"dd-mmm-yyyy")&"  |  source: NYC Open Data, 311 Service Requests"')
D["A2"].font = f(italic=True, color="595959")
widths(D, [2, 26, 9, 9, 9, 9, 10, 11, 9, 10, 11, 2] + [9] * 12)
met = f'COUNTIF({TL("O")},"Met")'; br = f'COUNTIF({TL("O")},"Breached")'
kpis = [
    ("Total Complaints", f'=COUNT({TL("B")})', "#,##0"),
    ("Open", f'=COUNTIFS({TL("B")},"<>",{TL("G")},"<>Closed")', "#,##0"),
    ("SLA Compliance %", f'=IFERROR({met}/({met}+{br}),0)', "0.0%"),
    ("Avg TAT (days)", f'=IFERROR(AVERAGE({TL("M")}),0)', "0.0"),
    ("Median TAT (days)", f'=IFERROR(MEDIAN({TL("M")}),0)', "0.0"),
    ("Closed", f'=COUNTIF({TL("G")},"Closed")', "#,##0"),
    ("SLA Breaches (incl. overdue open)", f'=COUNTIF({TL("P")},"Yes")', "#,##0"),
    ("Buildings with Complaints", f'=COUNTIF(Lists!$N${BS}:$N${BE},">0")', "#,##0"),
    ("PM Overdue (sample)", f'=COUNTIF({PL("J")},"OVERDUE")', "0"),
    ("PM Due Soon (sample)", f'=COUNTIF({PL("J")},"Due Soon")', "0"),
]
pos = [(4, 2), (4, 4), (4, 6), (4, 8), (4, 10), (7, 2), (7, 4), (7, 6), (7, 8), (7, 10)]
for (lab, fm, fmt), (r, c) in zip(kpis, pos):
    kpi(D, r, c, lab, fm, fmt); D.row_dimensions[r].height = 28

def breakdown(start_row, label, items, key_col, formula_items=False):
    header_row(D, start_row, [label, "Total", "Open", "Closed", "Met", "Breached", "Compliance %", "Avg TAT (d)", "Median TAT (d)", "Overdue Open"], 2)
    D.row_dimensions[start_row].height = 32
    for i, it in enumerate(items):
        r = start_row + 1 + i; k = f"$B{r}"
        D[f"B{r}"] = it
        D[f"C{r}"] = f'=COUNTIF({TL(key_col)},{k})'
        D[f"D{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("G")},"<>Closed")'
        D[f"E{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("G")},"Closed")'
        D[f"F{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("O")},"Met")'
        D[f"G{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("O")},"Breached")'
        D[f"H{r}"] = f'=IFERROR(F{r}/(F{r}+G{r}),"-")'
        D[f"I{r}"] = f'=IFERROR(AVERAGEIFS({TL("M")},{TL(key_col)},{k}),"-")'
        D[f"J{r}"] = f'=IFERROR(MEDIAN(IF({TL(key_col)}={k},IF({TL("M")}<>"",{TL("M")}))),"-")'
        D[f"K{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("O")},"Open - Overdue")'
    end = start_row + len(items)
    return start_row, end

# MEDIAN(IF()) needs array entry -> use ArrayFormula objects
from openpyxl.worksheet.formula import ArrayFormula
def fix_arrays(r0, r1):
    for r in range(r0 + 1, r1 + 1):
        D[f"J{r}"] = ArrayFormula(f"J{r}", D[f"J{r}"].value)

def total_row(r0, r1):
    tr = r1 + 1
    D[f"B{tr}"] = "Total"
    for col in "CDEFGK": D[f"{col}{tr}"] = f"=SUM({col}{r0 + 1}:{col}{r1})"
    D[f"H{tr}"] = f'=IFERROR(F{tr}/(F{tr}+G{tr}),"-")'
    D[f"I{tr}"] = f'=IFERROR(AVERAGE({TL("M")}),"-")'
    D[f"J{tr}"] = f'=IFERROR(MEDIAN({TL("M")}),"-")'
    style_body(D, r0 + 1, tr, 2, 11, {3: "#,##0", 4: "#,##0", 5: "#,##0", 6: "#,##0", 7: "#,##0", 8: "0.0%", 9: "0.0", 10: "0.0", 11: "#,##0"})
    for c in range(2, 12): D.cell(row=tr, column=c).font = f(True); D.cell(row=tr, column=c).fill = fill(LIGHT)
    return tr

c0, c1 = breakdown(10, "Category", [f"=Lists!A{4 + i}" for i in range(len(CATS))], "D"); fix_arrays(c0, c1); ct = total_row(c0, c1)
D.conditional_formatting.add(f"H{c0 + 1}:H{c1}", ColorScaleRule(start_type="num", start_value=0.5, start_color="F8CBAD", mid_type="num", mid_value=0.8, mid_color="FFEB9C", end_type="num", end_value=1, end_color="C6EFCE"))
p0, p1 = breakdown(ct + 3, "Priority", [p for p, _ in PRIOS], "J"); fix_arrays(p0, p1); pt = total_row(p0, p1)

# Top 10 buildings
b0 = pt + 3
D[f"B{b0 - 1}"] = "Top 10 Buildings by Complaints"; D[f"B{b0 - 1}"].font = f(True, 11, NAVY)
header_row(D, b0, ["Building", "Total", "Open", "Closed", "Met", "Breached", "Compliance %", "Avg TAT (d)", "Heat/Hot Water", "Overdue Open"], 2)
D.row_dimensions[b0].height = 32
for i in range(10):
    r = b0 + 1 + i; k = f"$B{r}"
    D[f"B{r}"] = f"=INDEX(Lists!$M${BS}:$M${BE},MATCH(LARGE(Lists!$O${BS}:$O${BE},{i + 1}),Lists!$O${BS}:$O${BE},0))"
    D[f"C{r}"] = f'=COUNTIF({TL("C")},{k})'
    D[f"D{r}"] = f'=COUNTIFS({TL("C")},{k},{TL("G")},"<>Closed")'
    D[f"E{r}"] = f'=COUNTIFS({TL("C")},{k},{TL("G")},"Closed")'
    D[f"F{r}"] = f'=COUNTIFS({TL("C")},{k},{TL("O")},"Met")'
    D[f"G{r}"] = f'=COUNTIFS({TL("C")},{k},{TL("O")},"Breached")'
    D[f"H{r}"] = f'=IFERROR(F{r}/(F{r}+G{r}),"-")'
    D[f"I{r}"] = f'=IFERROR(AVERAGEIFS({TL("M")},{TL("C")},{k}),"-")'
    D[f"J{r}"] = f'=COUNTIFS({TL("C")},{k},{TL("D")},"HEAT/HOT WATER")'
    D[f"K{r}"] = f'=COUNTIFS({TL("C")},{k},{TL("O")},"Open - Overdue")'
b1 = b0 + 10
D[f"B{b1 + 1}"] = '="Top 10 = "&TEXT(SUM(C' + str(b0 + 1) + ':C' + str(b1) + ')/COUNT(' + TL("B") + '),"0.0%")&" of all complaints"'
D[f"B{b1 + 1}"].font = f(italic=True, color="595959", size=9)
style_body(D, b0 + 1, b1, 2, 11, {8: "0.0%", 9: "0.0"})

# SLA status mix, outcome mix, channel mix, monthly
s0 = b1 + 4
header_row(D, s0, ["SLA Status", "Tickets", "Share"], 2)
for i, s in enumerate(["Met", "Breached", "Open - On Track", "Open - Overdue"]):
    r = s0 + 1 + i; D[f"B{r}"] = s; D[f"C{r}"] = f'=COUNTIF({TL("O")},B{r})'; D[f"D{r}"] = f"=IFERROR(C{r}/SUM($C${s0 + 1}:$C${s0 + 4}),0)"
style_body(D, s0 + 1, s0 + 4, 2, 4, {4: "0.0%"})
header_row(D, s0, ["Outcome (HPD resolution)", "Tickets", "Share"], 6)
D.merge_cells(start_row=s0, start_column=6, end_row=s0, end_column=7)
for i, o in enumerate(OUTCOMES):
    r = s0 + 1 + i
    D.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    D[f"F{r}"] = f"=Lists!K{4 + i}"; D[f"H{r}"] = f'=COUNTIF({TL("I")},F{r})'; D[f"I{r}"] = f"=IFERROR(H{r}/SUM($H${s0 + 1}:$H${s0 + len(OUTCOMES)}),0)"
style_body(D, s0 + 1, s0 + len(OUTCOMES), 6, 9, {9: "0.0%"})
D[f"H{s0}"] = "Tickets"; D[f"I{s0}"] = "Share"
for c in ("H", "I"):
    D[f"{c}{s0}"].font = f(True, 10, "FFFFFF"); D[f"{c}{s0}"].fill = fill(NAVY); D[f"{c}{s0}"].alignment = Alignment(horizontal="center", vertical="center"); D[f"{c}{s0}"].border = BORDER
ch0 = s0 + 7
header_row(D, ch0, ["Channel", "Tickets", "Share"], 2)
for i, c in enumerate(CHANNELS[:3]):
    r = ch0 + 1 + i; D[f"B{r}"] = c; D[f"C{r}"] = f'=COUNTIF({TL("F")},B{r})'; D[f"D{r}"] = f"=IFERROR(C{r}/SUM($C${ch0 + 1}:$C${ch0 + 3}),0)"
style_body(D, ch0 + 1, ch0 + 3, 2, 4, {4: "0.0%"})
m0 = ch0 + 6
header_row(D, m0, ["Month", "Raised", "Closed", "Breached", "Compliance %"], 2)
for i in range(3):
    r = m0 + 1 + i
    D[f"B{r}"] = f'=DATE(YEAR(MIN({TL("B")})),MONTH(MIN({TL("B")})),1)' if i == 0 else f"=EDATE(B{r - 1},1)"
    D[f"C{r}"] = f'=COUNTIFS({TL("B")},">="&B{r},{TL("B")},"<"&EDATE(B{r},1))'
    D[f"D{r}"] = f'=COUNTIFS({TL("H")},">="&B{r},{TL("H")},"<"&EDATE(B{r},1))'
    D[f"E{r}"] = f'=COUNTIFS({TL("B")},">="&B{r},{TL("B")},"<"&EDATE(B{r},1),{TL("O")},"Breached")'
    D[f"F{r}"] = (f'=IFERROR(COUNTIFS({TL("B")},">="&B{r},{TL("B")},"<"&EDATE(B{r},1),{TL("O")},"Met")/'
                  f'(COUNTIFS({TL("B")},">="&B{r},{TL("B")},"<"&EDATE(B{r},1),{TL("O")},"Met")+E{r}),"-")')
style_body(D, m0 + 1, m0 + 3, 2, 6, {2: "mmm-yyyy", 6: "0.0%"})
D[f"B{m0 + 4}"] = "Nov-2026 is partial: the 5,000-row extract ends on 6-Nov-2026."
D[f"B{m0 + 4}"].font = f(italic=True, color="595959", size=9)

# Charts (right side, column M onwards)
def addbar(title_, r0, r1, cols, anchor, h=8, w=17, catcol=2, stacked=False):
    ch = BarChart(); ch.type = "bar"; ch.title = title_; ch.height = h; ch.width = w; ch.style = 10
    for c in cols: ch.add_data(Reference(D, min_col=c, min_row=r0, max_row=r1), titles_from_data=True)
    ch.set_categories(Reference(D, min_col=catcol, min_row=r0 + 1, max_row=r1))
    ch.x_axis.scaling.orientation = "maxMin"
    if stacked: ch.grouping = "stacked"; ch.overlap = 100
    chart_fix(ch); D.add_chart(ch, anchor)
addbar("Complaints by Category (Met vs Breached vs Open)", c0, c1, [6, 7, 4], "M4", h=10, stacked=True)
addbar("Top 10 Buildings by Complaints", b0, b1, [3], "M25", h=8)
pie = PieChart(); pie.title = "SLA Status Mix"; pie.height = 7.5; pie.width = 10
pie.add_data(Reference(D, min_col=3, min_row=s0, max_row=s0 + 4), titles_from_data=True)
pie.set_categories(Reference(D, min_col=2, min_row=s0 + 1, max_row=s0 + 4))
pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
for a in ("showVal", "showCatName", "showSerName", "showLegendKey"): setattr(pie.dataLabels, a, False)
chart_fix(pie, axes=False); D.add_chart(pie, "M42")
addbar("Inspection Outcomes", s0, s0 + len(OUTCOMES), [8], "M57", h=7.5, catcol=6)
mc = BarChart(); mc.type = "col"; mc.title = "Monthly Complaints Raised vs Closed"; mc.height = 7.5; mc.width = 14
for c in (3, 4): mc.add_data(Reference(D, min_col=c, min_row=m0, max_row=m0 + 3), titles_from_data=True)
mc.set_categories(Reference(D, min_col=2, min_row=m0 + 1, max_row=m0 + 3)); mc.x_axis.number_format = "mmm-yy"
chart_fix(mc); D.add_chart(mc, "M72")
D.freeze_panes = "A4"
D.page_setup.orientation = "landscape"; D.page_setup.fitToWidth = 1; D.page_setup.fitToHeight = 0
D.sheet_properties.pageSetUpPr.fitToPage = True

wb.move_sheet("Lists", 2)
wb.active = 1
wb.save(OUT)
print("rows", N, "buildings", NB, "cats", CATS)
