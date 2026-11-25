"""Project 1: Facility Maintenance and SLA Tracker (MS Excel)."""
import os
import random, sys, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

random.seed(42)
OUT = sys.argv[1]
TODAY = dt.date(2026, 11, 25)

LOCS = ["Block A - Ground Floor", "Block A - First Floor", "Block B - Labs", "Block C - Library",
        "Admin Office", "Cafeteria", "Hostel", "Parking & Grounds"]
CATS = ["Electrical", "Plumbing", "HVAC", "Housekeeping", "IT / Network", "Carpentry", "Civil"]
PRIOS = [("Critical", 1), ("High", 2), ("Medium", 4), ("Low", 7)]
VENDORS = {"Electrical": "BrightSpark Electricals", "Plumbing": "AquaFix Plumbing Services",
           "HVAC": "CoolAir HVAC Solutions", "Housekeeping": "CleanPro Facility Services",
           "IT / Network": "NetLink IT Services", "Carpentry": "WoodCraft Interiors",
           "Civil": "SolidBuild Civil Works"}
STATUS = ["Open", "In Progress", "On Hold", "Closed"]
DESC = {
    "Electrical": ["Tube light not working", "Power socket sparking", "MCB tripping repeatedly", "Corridor lights flickering", "UPS backup failure"],
    "Plumbing": ["Tap leaking in washroom", "Blocked drain", "Low water pressure", "Flush tank not working", "Water cooler leaking"],
    "HVAC": ["AC not cooling", "AC water dripping", "Exhaust fan noisy", "AC remote/thermostat fault", "Air handling unit odour"],
    "Housekeeping": ["Washroom cleaning required", "Spill in corridor", "Waste bins overflowing", "Pest control request", "Deep cleaning before event"],
    "IT / Network": ["Wi-Fi not reachable", "LAN port dead", "Projector not detecting laptop", "Printer jam", "CCTV camera offline"],
    "Carpentry": ["Door hinge broken", "Chair repair", "Cupboard lock jammed", "Window latch broken", "Table wobbling"],
    "Civil": ["Wall seepage", "Floor tile cracked", "Ceiling plaster falling", "Ramp damaged", "Paint peeling"],
}
CAT_W = [22, 16, 14, 14, 14, 10, 10]
PRIO_W = [10, 25, 45, 20]

DS, DE = 5, 504      # ticket log data rows
PS, PE = 5, 104      # PM data rows

wb = Workbook()

# ---------------- Instructions ----------------
ws = wb.active
ws.title = "Instructions"
title(ws, "Facility Maintenance & SLA Tracker", "Maintenance ticket log with automated turnaround time, SLA-breach flags, preventive maintenance alerts and a KPI dashboard")
lines = [
    ("How it works", True),
    ("1. Log every complaint as a row in Ticket_Log (columns A–I). Columns J–O calculate themselves.", False),
    ("2. When a job is finished, set Status = Closed and enter the Closed Date. TAT and SLA status update automatically.", False),
    ("3. Keep PM_Schedule current: enter the Last Done date after each preventive job. Next Due and the alert update automatically.", False),
    ("4. Dashboard shows open tickets, SLA compliance %, average TAT and issues by category, location and priority.", False),
    ("5. Change SLA targets, locations, categories, vendors or the PM alert window in the Lists sheet; every sheet follows.", False),
    ("", False),
    ("Colour legend", True),
    ("Dark blue header = you type here (input)   |   Teal header = calculated, do not overwrite", False),
    ("Light-yellow cells in Lists = settings you can change", False),
    ("", False),
    ("Formula logic", True),
    ("SLA Target (days) = looked up from the ticket's Priority in Lists (INDEX/MATCH).", False),
    ("Due Date = Raised Date + SLA Target.   TAT (days) = Closed Date − Raised Date.", False),
    ("SLA Status: closed tickets are 'Met' if TAT ≤ target, else 'Breached'; open tickets are 'Open - Overdue' once today passes the Due Date, else 'Open - On Track'.", False),
    ("SLA Compliance % = Met ÷ (Met + Breached), i.e. measured on closed tickets.", False),
    ("PM alert: 'OVERDUE' if Next Due is past, 'Due Soon' within the alert window (default 7 days), else 'OK'.", False),
    ("", False),
    ("Note", True),
    ("The rows supplied are sample data for a fictional campus (Sep–Nov 2026) to demonstrate the tracker. Delete rows 5 onward in Ticket_Log and PM_Schedule and enter real data; formulas are pre-filled to row 504 / 104.", False),
    ("Open tickets age with TODAY(), so the open/overdue counts change each day the file is opened.", False),
]
for i, (t, b) in enumerate(lines, start=4):
    c = ws.cell(row=i, column=1, value=t)
    c.font = f(b, 11 if b else 10, NAVY if b else "000000")
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 130

# ---------------- Lists ----------------
L = wb.create_sheet("Lists")
title(L, "Lists & Settings", "Master lists used by dropdowns and formulas — edit the yellow cells")
header_row(L, 3, ["Locations"], 1)
header_row(L, 3, ["Categories"], 3)
header_row(L, 3, ["Priority", "SLA Target (days)"], 5)
header_row(L, 3, ["Vendors"], 8)
header_row(L, 3, ["Status"], 10)
for i, v in enumerate(LOCS): L.cell(row=4 + i, column=1, value=v)
for i, v in enumerate(CATS): L.cell(row=4 + i, column=3, value=v)
for i, (p, d) in enumerate(PRIOS):
    L.cell(row=4 + i, column=5, value=p)
    L.cell(row=4 + i, column=6, value=d)
for i, v in enumerate(VENDORS.values()): L.cell(row=4 + i, column=8, value=v)
for i, v in enumerate(STATUS): L.cell(row=4 + i, column=10, value=v)
for col, n in [(1, 15), (3, 15), (5, 4), (6, 4), (8, 15), (10, 4)]:
    for r in range(4, 4 + n):
        c = L.cell(row=r, column=col)
        c.font = f(); c.border = BORDER; c.fill = fill(INPUT_FILL)
L["F3"].comment = Comment("Assumption: typical facility SLA targets by priority. Change to match your organisation's contract.", "Tracker")
L["E10"] = "PM alert window (days)"; L["E10"].font = f(True)
L["F10"] = 7; L["F10"].fill = fill(INPUT_FILL); L["F10"].border = BORDER; L["F10"].font = f()
L["E11"] = "Rows 4–18 hold list items; add new items in the blank yellow cells."; L["E11"].font = f(italic=True, color="595959")
widths(L, [26, 3, 18, 3, 14, 12, 3, 30, 3, 14])
LOC_R = "Lists!$A$4:$A$18"; CAT_R = "Lists!$C$4:$C$18"; PRI_R = "Lists!$E$4:$E$7"
SLA_R = "Lists!$F$4:$F$7"; VEN_R = "Lists!$H$4:$H$18"; ST_R = "Lists!$J$4:$J$7"
ALERT = "Lists!$F$10"

# ---------------- Ticket_Log ----------------
T = wb.create_sheet("Ticket_Log")
title(T, "Maintenance Ticket Log", "Columns A–I are inputs; J–O are calculated")
heads_in = ["Ticket ID", "Raised Date", "Location", "Category", "Priority", "Issue Description", "Assigned Vendor", "Status", "Closed Date"]
heads_calc = ["SLA Target (days)", "Due Date", "TAT (days)", "Days Open", "SLA Status", "SLA Breach?"]
header_row(T, 4, heads_in, 1, NAVY)
header_row(T, 4, heads_calc, 10, TEAL)
widths(T, [10, 12, 22, 14, 10, 30, 27, 12, 12, 11, 12, 10, 10, 16, 11])

tickets = []
start = dt.date(2026, 9, 1)
for i in range(64):
    raised = start + dt.timedelta(days=int(85 * (i / 64)) + random.randint(0, 1))
    if raised > dt.date(2026, 11, 24): raised = dt.date(2026, 11, 24)
    cat = random.choices(CATS, CAT_W)[0]
    pr, sla = random.choices(PRIOS, PRIO_W)[0]
    loc = random.choice(LOCS)
    age = (TODAY - raised).days
    closed = None
    if age > 12 or random.random() < 0.45:
        within = random.random() < 0.82
        dur = random.randint(0, sla) if within else sla + random.randint(1, 4)
        cd = raised + dt.timedelta(days=dur)
        if cd <= dt.date(2026, 11, 24):
            closed = cd
    status = "Closed" if closed else random.choice(["Open", "In Progress", "In Progress", "On Hold"])
    tickets.append([f"TKT-{1001 + i}", raised, loc, cat, pr, random.choice(DESC[cat]), VENDORS[cat], status, closed])

for r in range(DS, DE + 1):
    if r - DS < len(tickets):
        for c, v in enumerate(tickets[r - DS], start=1):
            T.cell(row=r, column=c, value=v)
    T[f"J{r}"] = f'=IF($B{r}="","",IFERROR(INDEX({SLA_R},MATCH($E{r},{PRI_R},0)),""))'
    T[f"K{r}"] = f'=IF(OR($B{r}="",$J{r}=""),"",$B{r}+$J{r})'
    T[f"L{r}"] = f'=IF(OR($B{r}="",$I{r}=""),"",$I{r}-$B{r})'
    T[f"M{r}"] = f'=IF($B{r}="","",IF($I{r}="",TODAY()-$B{r},""))'
    T[f"N{r}"] = (f'=IF(OR($B{r}="",$J{r}=""),"",IF($I{r}<>"",IF($L{r}<=$J{r},"Met","Breached"),'
                  f'IF(TODAY()>$K{r},"Open - Overdue","Open - On Track")))')
    T[f"O{r}"] = f'=IF($N{r}="","",IF(OR($N{r}="Breached",$N{r}="Open - Overdue"),"Yes","No"))'
style_body(T, DS, DE, 1, 15, {2: "dd-mmm-yyyy", 9: "dd-mmm-yyyy", 11: "dd-mmm-yyyy", 10: "0", 12: "0", 13: "0"})
for r in range(DS, DE + 1):
    for c in range(10, 16):
        T.cell(row=r, column=c).fill = fill(GREY)
for col, rng in [("C", LOC_R), ("D", CAT_R), ("E", PRI_R), ("G", VEN_R), ("H", ST_R)]:
    dv = DataValidation(type="list", formula1=f"={rng}", allow_blank=True)
    dv.add(f"{col}{DS}:{col}{DE}"); T.add_data_validation(dv)
dvd = DataValidation(type="date", operator="greaterThan", formula1="DATE(2000,1,1)", allow_blank=True,
                     showErrorMessage=True, error="Enter a valid date")
dvd.add(f"B{DS}:B{DE}"); dvd.add(f"I{DS}:I{DE}"); T.add_data_validation(dvd)
red, green, amber = fill("F8CBAD"), fill("C6EFCE"), fill("FFEB9C")
rng = f"N{DS}:O{DE}"
T.conditional_formatting.add(rng, FormulaRule(formula=[f'OR($N{DS}="Breached",$N{DS}="Open - Overdue")'], fill=red, font=Font(name=FONT, color="9C0006", bold=True)))
T.conditional_formatting.add(rng, FormulaRule(formula=[f'$N{DS}="Met"'], fill=green))
T.conditional_formatting.add(rng, FormulaRule(formula=[f'$N{DS}="Open - On Track"'], fill=amber))
T.freeze_panes = "B5"
T.auto_filter.ref = f"A4:O{DE}"

# ---------------- PM_Schedule ----------------
P = wb.create_sheet("PM_Schedule")
title(P, "Preventive Maintenance Schedule", "Enter Last Done after each PM job; Next Due and alert update automatically")
header_row(P, 4, ["PM ID", "Asset", "Location", "PM Task", "Frequency (days)", "Responsible Vendor", "Last Done"], 1, NAVY)
header_row(P, 4, ["Next Due", "Days to Due", "Alert"], 8, TEAL)
widths(P, [9, 26, 22, 34, 11, 27, 13, 13, 11, 12])
pm = [
    ("DG Set 125 kVA", "Parking & Grounds", "Oil, filter & load test", 90, "BrightSpark Electricals"),
    ("Main LT Panel", "Admin Office", "Thermal scan & tightening", 180, "BrightSpark Electricals"),
    ("UPS 20 kVA", "Block B - Labs", "Battery health check", 90, "BrightSpark Electricals"),
    ("Split ACs - Block A", "Block A - First Floor", "Filter cleaning", 30, "CoolAir HVAC Solutions"),
    ("Split ACs - Library", "Block C - Library", "Filter cleaning", 30, "CoolAir HVAC Solutions"),
    ("Central AHU", "Block B - Labs", "Belt, coil & drain service", 90, "CoolAir HVAC Solutions"),
    ("Overhead Water Tank", "Hostel", "Tank cleaning", 180, "AquaFix Plumbing Services"),
    ("RO Water Purifiers", "Cafeteria", "Filter & membrane change", 60, "AquaFix Plumbing Services"),
    ("Water Pumps", "Parking & Grounds", "Pump & valve inspection", 90, "AquaFix Plumbing Services"),
    ("Fire Extinguishers", "Block A - Ground Floor", "Pressure & expiry check", 30, "SolidBuild Civil Works"),
    ("Lift - Block A", "Block A - Ground Floor", "AMC service visit", 30, "SolidBuild Civil Works"),
    ("Network Rack", "Admin Office", "Switch & cable audit", 90, "NetLink IT Services"),
    ("CCTV System", "Parking & Grounds", "Camera & DVR check", 30, "NetLink IT Services"),
    ("Pest Control", "Cafeteria", "General pest treatment", 30, "CleanPro Facility Services"),
    ("Washrooms - All Blocks", "Block A - Ground Floor", "Deep cleaning", 14, "CleanPro Facility Services"),
    ("Classroom Furniture", "Block A - First Floor", "Inspection & repair", 180, "WoodCraft Interiors"),
]
for i, (asset, loc, task, freq, ven) in enumerate(pm):
    r = PS + i
    # last done chosen so we get a mix of OK / due soon / overdue
    off = random.choice([-5, -2, 2, 5, 10, 20, 40])
    last = TODAY - dt.timedelta(days=max(1, freq + off if off < 0 else freq - off))
    for c, v in enumerate([f"PM-{101 + i}", asset, loc, task, freq, ven, last], start=1):
        P.cell(row=r, column=c, value=v)
for r in range(PS, PE + 1):
    P[f"H{r}"] = f'=IF(OR($E{r}="",$G{r}=""),"",$G{r}+$E{r})'
    P[f"I{r}"] = f'=IF($H{r}="","",$H{r}-TODAY())'
    P[f"J{r}"] = f'=IF($I{r}="","",IF($I{r}<0,"OVERDUE",IF($I{r}<={ALERT},"Due Soon","OK")))'
style_body(P, PS, PE, 1, 10, {7: "dd-mmm-yyyy", 8: "dd-mmm-yyyy", 9: "0", 5: "0"})
for r in range(PS, PE + 1):
    for c in range(8, 11): P.cell(row=r, column=c).fill = fill(GREY)
dv = DataValidation(type="list", formula1=f"={LOC_R}", allow_blank=True); dv.add(f"C{PS}:C{PE}"); P.add_data_validation(dv)
dv = DataValidation(type="list", formula1=f"={VEN_R}", allow_blank=True); dv.add(f"F{PS}:F{PE}"); P.add_data_validation(dv)
P.conditional_formatting.add(f"H{PS}:J{PE}", FormulaRule(formula=[f'$J{PS}="OVERDUE"'], fill=red, font=Font(name=FONT, color="9C0006", bold=True)))
P.conditional_formatting.add(f"H{PS}:J{PE}", FormulaRule(formula=[f'$J{PS}="Due Soon"'], fill=amber))
P.conditional_formatting.add(f"H{PS}:J{PE}", FormulaRule(formula=[f'$J{PS}="OK"'], fill=green))
P.freeze_panes = "C5"
P.auto_filter.ref = f"A4:J{PE}"

# ---------------- Dashboard ----------------
D = wb.create_sheet("Dashboard", 1)
title(D, "Facility Maintenance KPI Dashboard", None)
D["A2"] = '="As of "&TEXT(TODAY(),"dd-mmm-yyyy")&"  |  all figures update from Ticket_Log and PM_Schedule"'
D["A2"].font = f(italic=True, color="595959")
widths(D, [2, 22, 10, 10, 10, 10, 12, 10, 10, 2] + [9] * 12)
TL = lambda col: f"Ticket_Log!${col}${DS}:${col}${DE}"
PL = lambda col: f"PM_Schedule!${col}${PS}:${col}${PE}"
met = f'COUNTIF({TL("N")},"Met")'; br = f'COUNTIF({TL("N")},"Breached")'
kpis = [
    ("Total Tickets", f'=COUNT({TL("B")})', "0"),
    ("Open Tickets", f'=COUNTIFS({TL("B")},"<>",{TL("H")},"<>Closed")', "0"),
    ("SLA Compliance %", f'=IFERROR({met}/({met}+{br}),0)', "0.0%"),
    ("Avg TAT (days)", f'=IFERROR(AVERAGE({TL("L")}),0)', "0.0"),
    ("Closed Tickets", f'=COUNTIFS({TL("H")},"Closed")', "0"),
    ("SLA Breaches (incl. overdue open)", f'=COUNTIF({TL("O")},"Yes")', "0"),
    ("PM Tasks Overdue", f'=COUNTIF({PL("J")},"OVERDUE")', "0"),
    ("PM Due Within Window", f'=COUNTIF({PL("J")},"Due Soon")', "0"),
]
pos = [(4, 2), (4, 4), (4, 6), (4, 8), (7, 2), (7, 4), (7, 6), (7, 8)]
# columns B..I used as 4 tiles x 2 cols
for (lab, fm, fmt), (r, c) in zip(kpis, pos):
    kpi(D, r, c, lab, fm, fmt)
    D.row_dimensions[r].height = 28

def breakdown(start_row, label, items, key_col):
    header_row(D, start_row, [label, "Total", "Open", "Closed", "Met", "Breached", "Compliance %", "Avg TAT", "Overdue Open"], 2)
    for i, it in enumerate(items):
        r = start_row + 1 + i
        k = f"$B{r}"
        D[f"B{r}"] = it
        D[f"C{r}"] = f'=COUNTIFS({TL(key_col)},{k})'
        D[f"D{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("H")},"<>Closed")'
        D[f"E{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("H")},"Closed")'
        D[f"F{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("N")},"Met")'
        D[f"G{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("N")},"Breached")'
        D[f"H{r}"] = f'=IFERROR(F{r}/(F{r}+G{r}),"-")'
        D[f"I{r}"] = f'=IFERROR(AVERAGEIFS({TL("L")},{TL(key_col)},{k}),"-")'
        D[f"J{r}"] = f'=COUNTIFS({TL(key_col)},{k},{TL("N")},"Open - Overdue")'
    end = start_row + len(items)
    tr = end + 1
    D[f"B{tr}"] = "Total"
    for col in "CDEFGJ":
        D[f"{col}{tr}"] = f"=SUM({col}{start_row + 1}:{col}{end})"
    D[f"H{tr}"] = f'=IFERROR(F{tr}/(F{tr}+G{tr}),"-")'
    D[f"I{tr}"] = f'=IFERROR(AVERAGE({TL("L")}),"-")'
    style_body(D, start_row + 1, tr, 2, 10, {8: "0.0%", 9: "0.0"})
    for c in range(2, 11):
        D.cell(row=tr, column=c).font = f(True); D.cell(row=tr, column=c).fill = fill(LIGHT)
    return start_row, end

D.column_dimensions["J"].width = 10
cat0, cat1 = breakdown(10, "Category", CATS, "D")
loc0, loc1 = breakdown(cat1 + 3, "Location", LOCS, "C")
pri0, pri1 = breakdown(loc1 + 3, "Priority", [p for p, _ in PRIOS], "E")

# SLA status mix + monthly trend
s0 = pri1 + 3
header_row(D, s0, ["SLA Status", "Tickets"], 2)
for i, s in enumerate(["Met", "Breached", "Open - On Track", "Open - Overdue"]):
    r = s0 + 1 + i
    D[f"B{r}"] = s; D[f"C{r}"] = f'=COUNTIF({TL("N")},B{r})'
style_body(D, s0 + 1, s0 + 4, 2, 3)
m0 = s0 + 7
header_row(D, m0, ["Month", "Raised", "Closed", "Breached"], 2)
for i in range(4):
    r = m0 + 1 + i
    D[f"B{r}"] = f'=DATE(YEAR(MIN({TL("B")})),MONTH(MIN({TL("B")})),1)' if i == 0 else f"=EDATE(B{r - 1},1)"
    D[f"C{r}"] = f'=COUNTIFS({TL("B")},">="&B{r},{TL("B")},"<"&EDATE(B{r},1))'
    D[f"D{r}"] = f'=COUNTIFS({TL("I")},">="&B{r},{TL("I")},"<"&EDATE(B{r},1))'
    D[f"E{r}"] = f'=COUNTIFS({TL("B")},">="&B{r},{TL("B")},"<"&EDATE(B{r},1),{TL("O")},"Yes")'
style_body(D, m0 + 1, m0 + 4, 2, 5, {2: "mmm-yyyy"})

# Charts
def bar(title_, r0, r1, cols, anchor, horiz=True, h=7.5):
    ch = BarChart(); ch.type = "bar" if horiz else "col"; ch.title = title_
    ch.style = 10; ch.height = h; ch.width = 16
    for c in cols:
        data = Reference(D, min_col=c, min_row=r0, max_row=r1)
        ch.add_data(data, titles_from_data=True)
    ch.set_categories(Reference(D, min_col=2, min_row=r0 + 1, max_row=r1))
    if horiz: ch.x_axis.scaling.orientation = "maxMin"
    chart_fix(ch); D.add_chart(ch, anchor)

bar("Issues by Category (Total vs Breached)", cat0, cat1, [3, 7], "L4", h=8)
bar("Issues by Location", loc0, loc1, [3, 4], "L21", h=8)
pie = PieChart(); pie.title = "SLA Status Mix"; pie.height = 7.5; pie.width = 10
pie.add_data(Reference(D, min_col=3, min_row=s0, max_row=s0 + 4), titles_from_data=True)
pie.set_categories(Reference(D, min_col=2, min_row=s0 + 1, max_row=s0 + 4))
pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
chart_fix(pie, axes=False); D.add_chart(pie, "L38")
line = BarChart(); line.type = "col"; line.title = "Monthly Tickets Raised vs Closed"; line.height = 7.5; line.width = 16
for c in [3, 4]:
    line.add_data(Reference(D, min_col=c, min_row=m0, max_row=m0 + 4), titles_from_data=True)
line.set_categories(Reference(D, min_col=2, min_row=m0 + 1, max_row=m0 + 4))
line.x_axis.number_format = "mmm-yy"
chart_fix(line); D.add_chart(line, "L54")
D.freeze_panes = "A4"

wb.move_sheet("Instructions", -10)
wb.active = 1
wb.save(OUT)
print("saved", OUT, len(tickets), "tickets")
