# SLA Report Bot

Turns **any** ticket, complaint or work-order file (CSV or Excel) into a rule-based SLA report in Excel. It follows a written rulebook (`sla_rules.json`) exactly, checks data quality, and shows every rule it applied.

## How to use (Windows)
1. Drag your CSV or Excel file onto **`SLA_Bot.bat`**.
2. A report named `SLA_Report_<your file>_<date>.xlsx` appears next to your file.

Command line: `python sla_bot.py tickets.xlsx [--rules my_rules.json] [--as-of 2026-11-25] [--out report.xlsx]`

It needs Python with pandas and openpyxl. Anaconda already has both, and `SLA_Bot.bat` uses Anaconda automatically.

## What the bot does
1. **Finds the columns** by name, whatever your file calls them. For example, the created date can be named *created, opened, raised on, reported, date logged…*, and the same goes for the closed date, status, priority, category, location and assignee. If it guesses wrong, type the exact column name in `sla_rules.json → columns`.
2. **Reads dates** in either format: it decides automatically between Indian/UK style 25/11/2026 and US style 11/25/2026.
3. **Sets each ticket's priority**:
   - from the file's priority column, understanding P1–P4, Sev1, Urgent, High/Low and similar;
   - or, if there is no priority column, from category rules and keywords ("fire", "gas", "leak"…).
4. **Applies the SLA targets.** Defaults: Critical 2, High 7, Medium 14, Low 21 days. The clock can count calendar days or business days, with weekends and holidays skipped.
5. **Gives every ticket a status**:
   - **Met** or **Breached**: closed tickets, within or past their target;
   - **Open - On Track**, **Open - At Risk** or **Open - Overdue**: open tickets, judged against the as-of date;
   - **Excluded**: cancelled, duplicate and similar tickets;
   - **Data issue**: tickets that can't be scored.
6. **Checks data quality**: missing or unreadable dates, a closed date before the created date, "Closed" status with no closed date, duplicate IDs, unknown priority values, dates in the future. Nothing is silently dropped.
7. **Writes the report.** The workbook has these sheets:
   - **Summary**: RAG status, KPIs, auto-written findings and a chart;
   - **By Priority**, **By Category**, **By Location** and **By Assignee**;
   - **Monthly Trend**;
   - **Open Overdue**: the to-do list;
   - **Breaches**;
   - **Ticket Detail**: every row, with the reason for its result;
   - **Data Quality**;
   - **Rules Applied**: an audit trail of every setting and decision.

## Rules you can change (`sla_rules.json`)
| Setting | Default |
|---|---|
| SLA targets by priority | Critical 2 · High 7 · Medium 14 · Low 21 days |
| Clock | calendar (or `business` + weekend days + holidays) |
| Compliance target / RAG | 90 % target; Amber within 10 % |
| At-risk warning | open ticket has used 75 % of its target |
| As-of date | today (or `latest`, or a fixed date) |
| Closed / excluded statuses | closed, resolved, done… / cancelled, duplicate, invalid… |

## Tested on
- **Real NYC 311 data**: 5,000 HPD complaints. The bot reproduced the hand-built tracker's numbers exactly: 68.2 % compliance, 3,079 met, 1,434 breached and 483 open overdue (`Sample_SLA_Report_NYC311.xlsx`).
- **A campus file shaped differently**: different column names, dd/mm/yyyy dates, P1–P4 priorities and planted problems: a missing date, a closed date before the created date, a closed status with no date, a duplicate ID and a cancelled ticket. All of them were caught, and the 95.2 % compliance matched an independent calculation (`Sample_SLA_Report_Campus.xlsx`).
