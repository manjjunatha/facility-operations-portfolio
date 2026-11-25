# Facility Operations Portfolio
**Manjunatha P** · B.Tech ECE, NMAM Institute of Technology (2027) · [github.com/manjjunatha](https://github.com/manjjunatha)

Four projects on facility operations: maintenance SLA tracking, vendor evaluation, energy monitoring, and an automated SLA reporting bot. They use MS Excel, Power BI, Python and an ESP32 microcontroller.

| # | Project | Tools | What it does |
|---|---|---|---|
| 1 | [Facility Maintenance & SLA Tracker](01-facility-maintenance-sla-tracker/) | Excel, Python | Logs maintenance tickets, flags SLA breaches automatically, tracks preventive maintenance, and reports KPIs on a dashboard. **Built on 5,000 real NYC 311 complaints.** Includes a 3-minute explainer video. |
| 2 | [Vendor Evaluation & Cost Analysis Dashboard](02-vendor-evaluation-dashboard/) | Excel, Power BI | Weighted vendor scorecard (cost, delivery, quality, complaints), spend analysis by category, vendor and month, plus a Power BI build guide with DAX. |
| 3 | [Workspace Occupancy & Energy Monitor](03-workspace-occupancy-energy-monitor/) | ESP32, sensors, Python, Excel | PIR / light / current sensors flag lights or AC left on in empty rooms; an Excel report shows energy wasted and under-used rooms. |
| 4 | [SLA Report Bot](04-sla-report-bot/) | Python (pandas, openpyxl) | Drop in **any** ticket file (CSV or Excel) and get a rule-based SLA report. It auto-detects columns, follows an editable rulebook, checks data quality and keeps an audit trail. |

## Highlights
- **Real data:** Project 1 runs on 5,000 NYC Open Data maintenance complaints. It found 68.2% SLA compliance, and one building, 80 Lenox Road, produced 278 complaints.
- **Verified numbers:** dashboard figures were cross-checked against independent Python calculations.
- **Automation:** the SLA Report Bot reproduces Project 1's results from the raw file in seconds.

## Data and honesty notes
- Project 1 uses real public data (NYC Open Data, 311 Service Requests). Its SLA targets are the tracker's own assumptions.
- Projects 2 and 3 use clearly labelled **sample/simulated data** to demonstrate the models. Project 3's firmware still needs to be built and tested on hardware.
- Each project's README lists its assumptions and limitations.

## Requirements
- Excel 2016 or later (or LibreOffice) for the workbooks.
- Python 3.9+ with `pandas` and `openpyxl` for the scripts and the bot. Anaconda includes both.
- Arduino IDE with the ESP32 core for Project 3's firmware.
