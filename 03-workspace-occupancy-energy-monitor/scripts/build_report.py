"""Project 3: Workspace Occupancy and Energy Monitor - Excel energy-waste report (+ sample CSV)."""
import os
import random, sys, csv, datetime as dt
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment

random.seed(11)
OUT, SAMPLE_CSV = sys.argv[1], sys.argv[2]
INT = 5  # minutes between sample rows
ROOMS = [("Room 101 - Classroom", 200), ("Lab 2 - Electronics", 240), ("Faculty Room", 120)]
DAYS = [dt.date(2026, 11, 16) + dt.timedelta(days=i) for i in range(5)]  # Mon-Fri

def h(x):  # "9:30" -> minutes
    a, b = x.split(":"); return int(a) * 60 + int(b)

def in_blocks(m, blocks): return any(s <= m < e for s, e in blocks)

# ---------- simulate (same rules as firmware, at 5-min resolution) ----------
rows = []
for di, day in enumerate(DAYS):
    wd = day.weekday()
    plan = {}
    # Classroom: 3 lecture blocks; lights/AC sometimes left on after last class
    cls = [(h("9:00"), h("10:00")), (h("10:15"), h("12:15")), (h("13:30"), h("15:30"))]
    lag_l = [0, 150, 30, 180, 15][di]; lag_a = [0, 120, 0, 150, 0][di]
    plan[0] = dict(occ=cls, light=[(h("8:55"), cls[-1][1] + lag_l)],
                   ac=[(s, e) for s, e in cls[:-1]] + [(cls[-1][0], cls[-1][1] + lag_a)])
    # Lab: full sessions Mon/Wed/Fri, short Tue/Thu; AC left on overnight Wednesday
    lab = [(h("10:00"), h("13:00")), (h("14:00"), h("17:00"))] if wd in (0, 2, 4) else [(h("14:00"), h("16:00"))]
    ac_lab = [(h("10:00"), h("17:30"))] if wd in (0, 2, 4) else [(h("14:00"), h("16:30"))]
    if wd == 2: ac_lab = [(h("10:00"), 24 * 60)]
    if wd == 3: ac_lab = [(0, h("8:00"))] + ac_lab
    plan[1] = dict(occ=lab, light=[(h("9:45"), h("17:30"))] if wd in (0, 2, 4) else [(h("13:45"), h("16:45"))], ac=ac_lab)
    # Faculty room: in use most of the day, lights on until housekeeping leaves at 19:00
    plan[2] = dict(occ=[(h("8:30"), h("13:00")), (h("14:00"), h("17:30"))],
                   light=[(h("8:15"), h("19:00"))], ac=[(h("9:00"), h("17:30"))])
    for ri, (room, _) in enumerate(ROOMS):
        p = plan[ri]
        prev_motion = False; empty_on = 0
        for m in range(0, 24 * 60, INT):
            motion = in_blocks(m, p["occ"]) and random.random() > (0.08 if ri < 2 else 0.25)
            occupied = motion or (prev_motion and in_blocks(m, p["occ"]))  # 5-min PIR hold-off
            prev_motion = motion
            light = in_blocks(m, p["light"])
            ac = in_blocks(m, p["ac"])
            amps = round(random.uniform(5.6, 6.6), 2) if ac else round(random.uniform(0.0, 0.15), 2)
            if not occupied and (light or ac): empty_on += 1
            else: empty_on = 0
            waste = 1 if empty_on * INT > 10 else 0
            ts = dt.datetime.combine(day, dt.time()) + dt.timedelta(minutes=m)
            rows.append([ts, room, int(occupied), int(light), int(ac), amps, waste])
rows.sort(key=lambda r: (r[0], [x for x, _ in ROOMS].index(r[1])))
with open(SAMPLE_CSV, "w", newline="") as fh:
    w = csv.writer(fh); w.writerow(["Timestamp", "Room", "Occupied", "Light_On", "AC_On", "AC_Amps", "Waste_Flag"])
    for r in rows: w.writerow([r[0].strftime("%Y-%m-%d %H:%M:%S")] + r[1:])

DS, DE = 5, 30004
wb = Workbook()
# ---------- Instructions ----------
ws = wb.active; ws.title = "Instructions"
title(ws, "Workspace Occupancy & Energy Report", "Turns the ESP32 monitor's log into occupancy, utilisation and energy-waste figures per room")
lines = [
    ("How to use", True),
    ("1. Run tools/serial_logger.py while the ESP32 is connected; it writes occupancy_log.csv.", False),
    ("2. Open the CSV, copy all data rows (not the header) and paste into Raw_Data starting at cell A5, columns A–G. Delete the sample rows first.", False),
    ("3. In Settings, set Log Interval to the firmware's LOG_INTERVAL (1 minute by default), and enter each room's name exactly as in the firmware with its total lighting load.", False),
    ("4. Dashboard updates automatically: utilisation, hours lights/AC ran in empty rooms, kWh and ₹ wasted, hour-by-hour heatmaps and daily trend.", False),
    ("", False),
    ("How the numbers are worked out", True),
    ("Each Raw_Data row is one reading for one room and represents one log interval of time.", False),
    ("Occupied = PIR saw motion in the last 5 minutes. Light_On from the LDR, AC_On and AC_Amps from the CT clamp. Waste_Flag = 1 once a room has been empty with lights/AC on for over 10 minutes (the firmware's alert).", False),
    ("Lights wasted (kWh) = intervals where room empty and lights on × interval × room lighting load (W) ÷ 1000.", False),
    ("AC wasted (kWh) = Σ measured AC current while room empty × mains voltage × power factor × interval ÷ 1000.", False),
    ("Working-hours utilisation = share of readings between the working-hour start and end in which the room was occupied. Rooms below the under-use threshold are flagged.", False),
    ("", False),
    ("Colour legend", True),
    ("Dark blue header = input column   |   Teal header = calculated   |   Light-yellow cell = setting you may change", False),
    ("", False),
    ("Note", True),
    ("The rows supplied are SIMULATED sample data (16–20 Nov 2026, 5-minute interval) generated with the same rules as the firmware, to demonstrate the report before real data is collected. Replace them with your own logged data.", False),
    ("Raw_Data has helper formulas pre-filled to row 30,004 (about 7 days of 1-minute logging for 3 rooms). For longer logs, copy column H further down.", False),
]
for i, (t, b) in enumerate(lines, start=4):
    c = ws.cell(row=i, column=1, value=t); c.font = f(b, 11 if b else 10, NAVY if b else "000000")
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 135

# ---------- Settings ----------
S = wb.create_sheet("Settings")
title(S, "Settings", "Yellow cells are editable")
header_row(S, 4, ["Parameter", "Value", "Note"], 1)
params = [
    ("Log interval (minutes)", INT, "Sample data uses 5 min. Set to 1 for data from the firmware (LOG_INTERVAL_MS = 60 s)."),
    ("Working hours start (hour, 0-23)", 8, "Used for utilisation %"),
    ("Working hours end (hour, 0-23)", 18, "Readings with hour < end are counted"),
    ("Electricity tariff (Rs per kWh)", 8, "Assumption: typical Indian commercial/institutional rate; replace with your bill's rate"),
    ("Mains voltage (V)", 230, "India standard single-phase"),
    ("AC power factor", 0.9, "Assumption: typical inverter split AC"),
    ("Under-use threshold (utilisation %)", 0.4, "Rooms below this in working hours are flagged Under-used"),
    ("Days for projection", 30, "Projected waste = waste per logged day × this"),
]
for i, (k, v, n) in enumerate(params):
    S.cell(row=5 + i, column=1, value=k); S.cell(row=5 + i, column=2, value=v); S.cell(row=5 + i, column=3, value=n)
style_body(S, 5, 12, 1, 3)
S["B11"].number_format = "0%"
for r in range(5, 13): S.cell(row=r, column=2).fill = fill(INPUT_FILL); S.cell(row=r, column=3).font = f(italic=True, color="595959")
INTV, WS_, WE_, TAR, VOLT, PF, UTH, PDAYS = [f"Settings!$B${r}" for r in range(5, 13)]
header_row(S, 15, ["Room (must match firmware name)", "Lighting Load (W)"], 1)
for i, (rn, wt) in enumerate(ROOMS):
    S.cell(row=16 + i, column=1, value=rn); S.cell(row=16 + i, column=2, value=wt)
style_body(S, 16, 18, 1, 2)
for r in range(16, 19): S.cell(row=r, column=1).fill = fill(INPUT_FILL); S.cell(row=r, column=2).fill = fill(INPUT_FILL)
S["B15"].comment = Comment("Assumption: number of tube lights × 20 W (LED). Classroom 10, Lab 12, Faculty room 6.", "Report")
widths(S, [36, 12, 90])

# ---------- Raw_Data ----------
R = wb.create_sheet("Raw_Data")
title(R, "Raw Data (from serial_logger.py)", "Paste logged rows into A5:G — column H is calculated")
header_row(R, 4, ["Timestamp", "Room", "Occupied", "Light_On", "AC_On", "AC_Amps", "Waste_Flag"], 1, NAVY)
header_row(R, 4, ["Hour"], 8, TEAL)
widths(R, [18, 22, 9, 9, 8, 9, 10, 7])
for i, row in enumerate(rows):
    r = DS + i
    for c, v in enumerate(row, start=1): R.cell(row=r, column=c, value=v)
    R.cell(row=r, column=1).number_format = "dd-mmm-yyyy hh:mm"
for r in range(DS, DE + 1):
    R[f"H{r}"] = f'=IF(A{r}="","",HOUR(A{r}))'
R.freeze_panes = "A5"; R.auto_filter.ref = f"A4:H{DE}"
RL = lambda c: f"Raw_Data!${c}${DS}:${c}${DE}"

# ---------- Dashboard ----------
D = wb.create_sheet("Dashboard", 1)
title(D, "Occupancy & Energy Waste Dashboard", None)
D["A2"] = (f'="Data from "&TEXT(MIN({RL("A")}),"dd-mmm-yyyy hh:mm")&" to "&TEXT(MAX({RL("A")}),"dd-mmm-yyyy hh:mm")&'
           f'"  |  "&COUNT({RL("A")})&" readings"')
D["A2"].font = f(italic=True, color="595959")
widths(D, [24, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13])
DAYSLOG = f'((MAX({RL("A")})-MIN({RL("A")}))+{INTV}/1440)'

# per-room table
t0 = 9
H = ["Room", "Logged (h)", "Occupied (h)", "Working-hrs Utilisation %", "Lights ON in Empty Room (h)", "AC ON in Empty Room (h)",
     "Waste Alert Time (h)", "Lights Wasted (kWh)", "AC Wasted (kWh)", "Total Wasted (kWh)", "Waste Cost (Rs)", "Projected Cost per Period (Rs)", "Status"]
header_row(D, t0, H, 1)
D.row_dimensions[t0].height = 45
for i in range(3):
    r = t0 + 1 + i; k = f"$A{r}"
    D[f"A{r}"] = f"=Settings!A{16 + i}"
    D[f"B{r}"] = f'=COUNTIF({RL("B")},{k})*{INTV}/60'
    D[f"C{r}"] = f'=COUNTIFS({RL("B")},{k},{RL("C")},1)*{INTV}/60'
    D[f"D{r}"] = (f'=IFERROR(COUNTIFS({RL("B")},{k},{RL("C")},1,{RL("H")},">="&{WS_},{RL("H")},"<"&{WE_})/'
                  f'COUNTIFS({RL("B")},{k},{RL("H")},">="&{WS_},{RL("H")},"<"&{WE_}),0)')
    D[f"E{r}"] = f'=COUNTIFS({RL("B")},{k},{RL("C")},0,{RL("D")},1)*{INTV}/60'
    D[f"F{r}"] = f'=COUNTIFS({RL("B")},{k},{RL("C")},0,{RL("E")},1)*{INTV}/60'
    D[f"G{r}"] = f'=COUNTIFS({RL("B")},{k},{RL("G")},1)*{INTV}/60'
    D[f"H{r}"] = f"=E{r}*Settings!B{16 + i}/1000"
    D[f"I{r}"] = f'=SUMIFS({RL("F")},{RL("B")},{k},{RL("C")},0,{RL("E")},1)*{VOLT}*{PF}*{INTV}/60/1000'
    D[f"J{r}"] = f"=H{r}+I{r}"
    D[f"K{r}"] = f"=J{r}*{TAR}"
    D[f"L{r}"] = f"=IFERROR(K{r}/{DAYSLOG}*{PDAYS},0)"
    D[f"M{r}"] = (f'=IF(B{r}=0,"No data",IF(AND(D{r}<{UTH},J{r}>0),"Under-used + Waste",IF(D{r}<{UTH},"Under-used",'
                  f'IF(J{r}>0,"Energy waste","OK"))))')
tr = t0 + 4
D[f"A{tr}"] = "Total"
for col in "BCEFGHIJKL":
    D[f"{col}{tr}"] = f"=SUM({col}{t0 + 1}:{col}{t0 + 3})"
D[f"D{tr}"] = (f'=IFERROR(COUNTIFS({RL("C")},1,{RL("H")},">="&{WS_},{RL("H")},"<"&{WE_})/'
               f'COUNTIFS({RL("B")},"<>",{RL("H")},">="&{WS_},{RL("H")},"<"&{WE_}),0)')
style_body(D, t0 + 1, tr, 1, 13, {**{c: "0.0" for c in (2, 3, 5, 6, 7)}, 4: "0.0%", 8: "0.00", 9: "0.00", 10: "0.00", 11: "#,##0", 12: "#,##0"})
for c in range(1, 14): D.cell(row=tr, column=c).font = f(True); D.cell(row=tr, column=c).fill = fill(LIGHT)
D.conditional_formatting.add(f"M{t0 + 1}:M{t0 + 3}", FormulaRule(formula=[f'$M{t0 + 1}<>"OK"'], fill=fill("F8CBAD"), font=Font(name=FONT, color="9C0006", bold=True)))
D.conditional_formatting.add(f"M{t0 + 1}:M{t0 + 3}", FormulaRule(formula=[f'$M{t0 + 1}="OK"'], fill=fill("C6EFCE")))
D[f"A{tr + 1}"] = '="Projection period = "&Settings!B12&" days at the same daily pattern; tariff Rs "&Settings!B8&"/kWh (Settings)."'
D[f"A{tr + 1}"].font = f(italic=True, color="595959", size=9)

# KPI tiles
kpi(D, 4, 1, "Energy Wasted (kWh)", f"=J{tr}", "0.0")
kpi(D, 4, 3, "Waste Cost (Rs)", f"=K{tr}", "#,##0")
kpi(D, 4, 5, "Projected Waste (Rs/period)", f"=L{tr}", "#,##0")
kpi(D, 4, 7, "Working-hrs Utilisation", f"=D{tr}", "0.0%")
kpi(D, 4, 9, "Waste Alert Hours", f"=G{tr}", "0.0")
kpi(D, 4, 11, "Days Logged", f"=IFERROR({DAYSLOG},0)", "0.0")

# Heatmaps
hm = tr + 4
D[f"A{hm - 1}"] = "Occupancy % by Hour of Day"; D[f"A{hm - 1}"].font = f(True, 11, NAVY)
D[f"F{hm - 1}"] = "Lights/AC ON while Empty, % of readings"; D[f"F{hm - 1}"].font = f(True, 11, NAVY)
header_row(D, hm, ["Hour"] + [f"=$A${t0 + 1 + i}" for i in range(3)], 1)
header_row(D, hm, ["Hour"] + [f"=$A${t0 + 1 + i}" for i in range(3)], 6)
for hr in range(24):
    r = hm + 1 + hr
    D[f"A{r}"] = hr; D[f"F{r}"] = hr
    for i in range(3):
        rc = f"{get_column_letter(2 + i)}${hm}"
        tot = f'COUNTIFS({RL("B")},{rc},{RL("H")},$A{r})'
        D[f"{get_column_letter(2 + i)}{r}"] = f'=IFERROR(COUNTIFS({RL("B")},{rc},{RL("H")},$A{r},{RL("C")},1)/{tot},"")'
        rc2 = f"{get_column_letter(7 + i)}${hm}"
        tot2 = f'COUNTIFS({RL("B")},{rc2},{RL("H")},$F{r})'
        D[f"{get_column_letter(7 + i)}{r}"] = (f'=IFERROR((COUNTIFS({RL("B")},{rc2},{RL("H")},$F{r},{RL("C")},0,{RL("D")},1)+'
                                               f'COUNTIFS({RL("B")},{rc2},{RL("H")},$F{r},{RL("C")},0,{RL("D")},0,{RL("E")},1))/{tot2},"")')
style_body(D, hm + 1, hm + 24, 1, 4, {2: "0%", 3: "0%", 4: "0%", 1: '00":00"'})
style_body(D, hm + 1, hm + 24, 6, 9, {7: "0%", 8: "0%", 9: "0%", 6: '00":00"'})
D.conditional_formatting.add(f"B{hm + 1}:D{hm + 24}", ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="num", end_value=1, end_color="2E75B6"))
D.conditional_formatting.add(f"G{hm + 1}:I{hm + 24}", ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="num", end_value=1, end_color="E06666"))

# Daily waste
dd = hm + 28
D[f"A{dd - 1}"] = "Daily Energy Wasted (kWh)"; D[f"A{dd - 1}"].font = f(True, 11, NAVY)
header_row(D, dd, ["Date"] + [f"=$A${t0 + 1 + i}" for i in range(3)] + ["Total"], 1)
ND = 7
for j in range(ND):
    r = dd + 1 + j
    D[f"A{r}"] = f'=INT(MIN({RL("A")}))' if j == 0 else f"=A{r - 1}+1"
    for i in range(3):
        col = get_column_letter(2 + i); rc = f"{col}${dd}"
        win = f'{RL("A")},">="&$A{r},{RL("A")},"<"&$A{r}+1'
        D[f"{col}{r}"] = (f'=COUNTIFS({RL("B")},{rc},{win},{RL("C")},0,{RL("D")},1)*{INTV}/60*Settings!$B${16 + i}/1000'
                          f'+SUMIFS({RL("F")},{RL("B")},{rc},{win},{RL("C")},0,{RL("E")},1)*{VOLT}*{PF}*{INTV}/60/1000')
    D[f"E{r}"] = f"=SUM(B{r}:D{r})"
style_body(D, dd + 1, dd + ND, 1, 5, {1: "ddd dd-mmm", 2: "0.00", 3: "0.00", 4: "0.00", 5: "0.00"})

# charts
c1 = BarChart(); c1.type = "col"; c1.grouping = "stacked"; c1.overlap = 100
c1.title = "Energy Wasted by Room (kWh)"; c1.height = 7.5; c1.width = 14
c1.add_data(Reference(D, min_col=8, max_col=9, min_row=t0, max_row=t0 + 3), titles_from_data=True)
c1.set_categories(Reference(D, min_col=1, min_row=t0 + 1, max_row=t0 + 3))
chart_fix(c1); D.add_chart(c1, f"K{hm - 1}")
c2 = BarChart(); c2.type = "bar"; c2.title = "Working-hours Utilisation %"; c2.height = 6; c2.width = 14; c2.legend = None
c2.add_data(Reference(D, min_col=4, min_row=t0, max_row=t0 + 3), titles_from_data=True)
c2.set_categories(Reference(D, min_col=1, min_row=t0 + 1, max_row=t0 + 3))
c2.y_axis.scaling.min = 0; c2.y_axis.scaling.max = 1; c2.y_axis.number_format = "0%"
chart_fix(c2); D.add_chart(c2, f"K{hm + 15}")
c3 = BarChart(); c3.type = "col"; c3.grouping = "stacked"; c3.overlap = 100
c3.title = "Daily Energy Wasted (kWh)"; c3.height = 7.5; c3.width = 16
c3.add_data(Reference(D, min_col=2, max_col=4, min_row=dd, max_row=dd + ND), titles_from_data=True)
c3.set_categories(Reference(D, min_col=1, min_row=dd + 1, max_row=dd + ND))
c3.x_axis.number_format = "dd-mmm"
chart_fix(c3); D.add_chart(c3, f"G{dd - 1}")
D.page_setup.orientation = "landscape"; D.page_setup.fitToWidth = 1; D.page_setup.fitToHeight = 0
D.sheet_properties.pageSetUpPr.fitToPage = True

wb.active = 1
wb.save(OUT)
print("saved", len(rows), "rows")
