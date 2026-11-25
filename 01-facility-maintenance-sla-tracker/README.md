# Facility Maintenance and SLA Tracker
**MS Excel · Python (data cleaning) · real data: NYC 311 / HPD building-maintenance complaints**

An Excel tool for a facilities desk. It logs maintenance tickets, calculates turnaround time (TAT), flags SLA breaches automatically, tracks a preventive-maintenance (PM) schedule with due-date alerts, and shows everything on a KPI dashboard.

## Files
| File | Contents |
|---|---|
| `Facility_Maintenance_SLA_Tracker_NYC311.xlsx` | **Main version.** Built on 5,000 real complaints |
| `Facility_Maintenance_SLA_Tracker_Template.xlsx` | Template version with 64 sample tickets, a starting point for a new site |
| `data/nyc311_hpd_11226_clean.csv` | The cleaned real data used in the workbook |
| `media/Facility_SLA_Tracker_Explainer_Voice.mp4` | 3-minute explainer video (What · Why · How · Where) |
| `scripts/` | Python scripts that build the workbooks (`build_p1_real.py <raw csv> <out.xlsx> <clean.csv>`) and the video scenes |

## Data source
- **NYC Open Data, "311 Service Requests from 2010 to Present"** (dataset `erm2-nwe9`), City of New York.
- **Filter:** agency = HPD (Dept. of Housing Preservation & Development), ZIP 11226 (Flatbush, Brooklyn), created on or after 1 Sep 2026.
- **Extract:** 5,000 complaints (the download limit), raised 1 Sep – 6 Nov 2026, across 601 residential buildings. Downloaded 25 Nov 2026; status is as of that date.
- **Cleaning:** kept 9 of 44 fields, tidied addresses, and grouped HPD's long resolution text into 7 outcome labels.

## Assumptions
The 311 data has no priority or SLA, so the tracker sets its own. These are editable in the **Lists** sheet:

| Priority | Categories | SLA target |
|---|---|---|
| Critical | Heat/Hot Water, Elevator | 2 days |
| High | Water Leak, Safety, Electric, Plumbing | 7 days |
| Medium | Unsanitary Condition, Door/Window, General, Flooring/Stairs, Appliance | 14 days |
| Low | Paint/Plaster, Outside Building | 21 days |

These are this tracker's own service standards, not official HPD targets. The **PM_Schedule** sheet is an illustrative sample, because preventive maintenance isn't in 311 data.

## Sheets
- **Dashboard:** 10 KPI tiles; breakdowns by category, priority and top-10 buildings; SLA status mix; inspection outcomes; complaint channel; monthly trend; 5 charts.
- **Ticket_Log:** the 5,000 real tickets. Formulas add priority, SLA target, due date, TAT, days open, SLA status and breach flag.
- **Lists:** category→priority map, SLA targets, as-of date, building register (601 addresses), dropdown lists.
- **PM_Schedule:** preventive-maintenance assets with OVERDUE / Due Soon / OK alerts.
- **Instructions:** source, cleaning steps, assumptions and formula logic.

Excel features used: `COUNTIFS`, `AVERAGEIFS`, `MEDIAN(IF())` array formulas, `INDEX/MATCH`, `LARGE`, `EDATE`, data validation, conditional formatting, charts.

## Key findings (real data, 1 Sep – 6 Nov 2026)
- **5,000 complaints:** 4,513 closed and 487 open. Average TAT is 7.9 days; median 5.2 days.
- **SLA compliance is 68.2%.** Of the open complaints, 483 of 487 are already past their target.
- **Heat/Hot Water is the largest category (1,116, 22%).** It is resolved fastest (median 1.5 days) but meets the 2-day target only 65.9% of the time.
- **Weakest areas:** Elevator (19% compliance) and Water Leak, Electric and Plumbing (47–49%).
- **Concentration:** the top 10 buildings generate 19.9% of all complaints. **80 Lenox Road** alone has 278, 126 of them Heat/Hot Water, which points to a boiler problem.
- **Volume trend:** complaints rose 41% from September (1,927) to October (2,724) as heating season started.
- **Inspection outcomes:** violations were issued on 22% of complaints. 25% of the time HPD could not inspect or get access.

All dashboard figures were cross-checked against an independent Python (pandas) calculation. They match exactly.

## Limitations
- SLA targets are assumptions, not HPD's official targets.
- "Closed" in 311 means HPD closed the complaint, not always that a repair was made.
- The extract stops on 6 Nov 2026 (5,000-row limit), so November is partial.
- The PM schedule uses sample assets and dates.
