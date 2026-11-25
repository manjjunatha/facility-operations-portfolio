"""Scene definitions for the 3-minute explainer. Each scene: id, seconds, chapter, bg class, caption, html body, anim seconds."""

CSS = r"""
:root{--dark:#14213D;--light:#F6F5F1;--card:#FCFCFA;--accent:#2F6FED;--warn:#D9661F;--body:#4A5568;--muted:#7B8597;--line:#DAD8D0;--ondark:#C3CCDC;}
*{box-sizing:border-box}
html,body{margin:0;width:1920px;height:1080px;overflow:hidden}
body{font-family:'Inter',Arial,sans-serif;color:var(--dark);background:var(--light)}
.stage{position:relative;width:1920px;height:1080px;padding:150px 140px 0 140px}
.dark{background:var(--dark);color:var(--light)}
.blue{background:var(--accent);color:var(--light)}
h1,h2,h3,.num{font-family:'Inter Display','Inter',Arial,sans-serif;margin:0}
h1{font-size:120px;font-weight:700;line-height:1.02;letter-spacing:-0.02em}
h2{font-size:72px;font-weight:700;line-height:1.08;letter-spacing:-0.015em;margin-bottom:48px;max-width:1550px}
h3{font-size:34px;font-weight:600;line-height:1.2;margin-bottom:14px}
p{margin:0}
.eyebrow{font-size:24px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--accent);margin-bottom:24px}
.dark .eyebrow{color:#8FB0FF}.blue .eyebrow{color:#DCE6FF}
.lead{font-size:36px;line-height:1.45;color:var(--body);max-width:1450px}
.dark .lead{color:var(--ondark)}.blue .lead{color:#E8EEFF}
.body{font-size:26px;line-height:1.5;color:var(--body)}
.dark .body{color:var(--ondark)}
/* chrome */
.top{position:absolute;left:140px;right:140px;top:56px;display:flex;justify-content:space-between;align-items:center;font-size:22px;color:var(--muted)}
.dark .top,.blue .top{color:#9AA6BF}.blue .top{color:#DCE6FF}
.chips{display:flex;gap:10px}
.chip{padding:8px 18px;border-radius:999px;border:1.5px solid currentColor;font-weight:600;letter-spacing:.08em;font-size:18px;opacity:.55}
.chip.on{background:var(--accent);border-color:var(--accent);color:#fff;opacity:1}
.blue .chip.on{background:#fff;color:var(--accent);border-color:#fff}
.cap{position:absolute;left:50%;bottom:56px;transform:translateX(-50%);max-width:1500px;background:rgba(20,33,61,.92);color:#F6F5F1;font-size:32px;line-height:1.4;padding:18px 34px;border-radius:14px;text-align:center}
.dark .cap{background:rgba(246,245,241,.10)}
.blue .cap{background:rgba(20,33,61,.55)}
/* blocks */
.grid{display:grid;gap:28px}.g2{grid-template-columns:repeat(2,1fr)}.g3{grid-template-columns:repeat(3,1fr)}.g4{grid-template-columns:repeat(4,1fr)}
.card{background:var(--card);border:1px solid var(--line);border-top:6px solid var(--accent);border-radius:14px;padding:36px}
.card.warn{border-top-color:var(--warn)}
.tag{font-size:22px;font-weight:600;color:var(--accent);margin-bottom:16px;letter-spacing:.04em}
.card.warn .tag{color:var(--warn)}
.stat{border-left:4px solid var(--line);padding:6px 0 6px 30px}
.dark .stat{border-left-color:#33456B}
.num{font-size:118px;font-weight:700;line-height:1;letter-spacing:-0.02em;margin-bottom:14px}
.num.b{color:var(--accent)}.dark .num.b{color:#8FB0FF}
.label{font-size:26px;line-height:1.4;color:var(--body)}.dark .label{color:var(--ondark)}
.step{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:30px 30px}
.step .n{font-family:'Inter Display',Arial,sans-serif;font-size:48px;font-weight:700;color:var(--accent);margin-bottom:8px}
.step.key{background:var(--accent);border-color:var(--accent)}.step.key .n,.step.key h3,.step.key .body{color:#fff}
table{border-collapse:collapse;width:100%}
th{font-size:24px;font-weight:600;text-align:left;color:var(--muted);padding:0 24px 16px 0;border-bottom:2px solid var(--dark)}
td{font-size:30px;padding:22px 24px 22px 0;border-bottom:1px solid var(--line)}
.r{text-align:right}
tr.hl td{color:var(--accent);font-weight:600}
.brow{display:flex;align-items:center;gap:24px;margin-bottom:10px}
.brow .nm{width:340px;text-align:right;font-size:26px}
.brow .tr{flex:1;height:32px;background:#ECEAE3;border-radius:4px}
.brow .fl{height:32px;background:#A9B0BC;border-radius:4px;transform-origin:left center;animation:grow .9s cubic-bezier(.2,.7,.2,1) both}
.brow.low .fl{background:var(--accent)}
.brow .vl{width:120px;text-align:right;font-size:26px;font-weight:600}
.brow.low .nm,.brow.low .vl{color:var(--accent);font-weight:600}
.chap{display:flex;flex-direction:column;justify-content:center;height:830px}
.chap .big{font-family:'Inter Display',Arial,sans-serif;font-size:220px;font-weight:700;letter-spacing:-0.03em;line-height:1}
.chap .no{font-size:30px;font-weight:600;color:#8FB0FF;letter-spacing:.16em;margin-bottom:20px}
.chap .sub{font-size:44px;color:var(--ondark);margin-top:24px}
.shotwrap{width:1640px;height:640px;overflow:hidden;border-radius:14px;border:1px solid var(--line);background:#fff}
.shotwrap img{width:1640px;height:auto;transform-origin:0 0;animation:pan 12s ease-in-out both}
.uses .card{border-top-color:var(--accent);padding:30px}
.uses h3{font-size:32px;margin:0}
/* motion */
@keyframes rise{from{opacity:0;transform:translateY(30px)}to{opacity:1;transform:none}}
@keyframes grow{from{transform:scaleX(0)}to{transform:scaleX(1)}}
@keyframes fadein{from{opacity:0}to{opacity:1}}
@keyframes pan{0%{transform:scale(1) translate(0,0)}35%{transform:scale(1) translate(0,0)}100%{transform:scale(1.55) translate(-30px,-40px)}}
@keyframes wipe{from{transform:scaleX(0)}to{transform:scaleX(1)}}
.rise{animation:rise .7s cubic-bezier(.2,.7,.2,1) both}
.fade{animation:fadein .8s ease both}
.d1{animation-delay:.15s}.d2{animation-delay:.3s}.d3{animation-delay:.45s}.d4{animation-delay:.6s}.d5{animation-delay:.75s}.d6{animation-delay:.9s}.d7{animation-delay:1.05s}
.cap{animation:fadein .6s ease .4s both}
.accbar{width:180px;height:8px;background:var(--accent);margin-bottom:44px;transform-origin:left;animation:wipe .8s ease both}
"""

CH = ["WHAT", "WHY", "HOW", "WHERE"]


def top(chapter):
    chips = "".join(f'<span class="chip{" on" if c == chapter else ""}">{c}</span>' for c in CH)
    return f'<div class="top"><span>Facility Maintenance &amp; SLA Tracker</span><div class="chips">{chips}</div></div>'


def chap(no, word, sub):
    return f'<div class="chap"><p class="no rise">{no}</p><p class="big rise d1">{word}</p><p class="sub rise d2">{sub}</p></div>'


SCENES = [
    dict(id="s01", sec=8, ch=None, bg="dark", anim=2.2,
         cap="A 3-minute walkthrough of my Facility Maintenance &amp; SLA Tracker.",
         html='<div style="padding-top:110px"><p class="eyebrow rise">Project 1 · MS Excel · Facility operations</p><div class="accbar"></div>'
              '<h1 class="rise d1" style="max-width:1500px">Facility Maintenance &amp; SLA Tracker</h1>'
              '<p class="lead rise d2" style="margin-top:40px">Built and tested on 5,000 real NYC maintenance complaints</p>'
              '<p class="body rise d3" style="margin-top:60px;font-size:28px">Manjunatha P · B.Tech ECE, NMAM Institute of Technology</p></div>'),
    dict(id="s02", sec=4, ch="WHAT", bg="dark", anim=1.4, cap="", html=chap("01", "What", "What the project is")),
    dict(id="s03", sec=14, ch="WHAT", bg="", anim=2.0,
         cap="What is it? An Excel tool that logs complaints, checks deadlines and reports KPIs.",
         html='<p class="eyebrow rise">What</p><h2 class="rise d1">An Excel tool that shows a facility team which repairs are late</h2>'
              '<p class="lead rise d2">It logs every maintenance complaint, checks each one against a service deadline (SLA), and turns the results into a live dashboard. No manual checking.</p>'),
    dict(id="s04", sec=11, ch="WHAT", bg="", anim=2.0,
         cap="Four parts: a ticket log, automatic SLA flags, maintenance alerts and a dashboard.",
         html='<p class="eyebrow rise">What it does</p><h2 class="rise d1">Four parts, one workbook</h2><div class="grid g4">'
              '<div class="card rise d2"><p class="tag">01</p><h3>Ticket log</h3><p class="body">Every complaint with building, category, status and dates.</p></div>'
              '<div class="card rise d3"><p class="tag">02</p><h3>SLA flags</h3><p class="body">Turnaround time and Met, Breached or Overdue by formula.</p></div>'
              '<div class="card rise d4"><p class="tag">03</p><h3>PM alerts</h3><p class="body">Preventive jobs marked Overdue or Due Soon.</p></div>'
              '<div class="card rise d5"><p class="tag">04</p><h3>Dashboard</h3><p class="body">10 KPIs, 5 tables and 5 charts, always up to date.</p></div></div>'),
    dict(id="s05", sec=4, ch="WHY", bg="dark", anim=1.4, cap="", html=chap("02", "Why", "Why it matters")),
    dict(id="s06", sec=14, ch="WHY", bg="", anim=2.0,
         cap="Without one log, late repairs stay hidden until tenants complain again.",
         html='<p class="eyebrow rise">The problem</p><h2 class="rise d1">Late repairs hide in scattered complaint records</h2><div class="grid g3">'
              '<div class="card warn rise d2"><p class="tag">No single log</p><h3>Complaints live in calls, emails and registers</h3><p class="body">Follow-up depends on memory.</p></div>'
              '<div class="card warn rise d3"><p class="tag">No deadline tracking</p><h3>Nobody knows a job is late</h3><p class="body">Turnaround time is never measured against a target.</p></div>'
              '<div class="card warn rise d4"><p class="tag">No pattern view</p><h3>Repeat problems go unnoticed</h3><p class="body">Preventive maintenance becomes guesswork.</p></div></div>'),
    dict(id="s07", sec=12, ch="WHY", bg="", anim=2.2,
         cap="So the goal: log, automate, prevent, measure and decide.",
         html='<p class="eyebrow rise">The objectives</p><h2 class="rise d1">Five objectives</h2>'
              '<table><tr class="rise d1"><td style="width:120px;color:var(--accent);font-weight:700">01</td><td><b>Log</b> every complaint in one register</td></tr>'
              '<tr class="rise d2"><td style="color:var(--accent);font-weight:700">02</td><td><b>Automate</b> turnaround time and SLA breach flags</td></tr>'
              '<tr class="rise d3"><td style="color:var(--accent);font-weight:700">03</td><td><b>Prevent</b> breakdowns with preventive-maintenance alerts</td></tr>'
              '<tr class="rise d4"><td style="color:var(--accent);font-weight:700">04</td><td><b>Measure</b> compliance %, turnaround time and open work</td></tr>'
              '<tr class="rise d5"><td style="color:var(--accent);font-weight:700">05</td><td><b>Decide</b> which categories and buildings need action</td></tr></table>'),
    dict(id="s08", sec=4, ch="HOW", bg="dark", anim=1.4, cap="", html=chap("03", "How", "How it works")),
    dict(id="s09", sec=14, ch="HOW", bg="", anim=2.0,
         cap="It runs on real data: 5,000 NYC 311 maintenance complaints from Brooklyn, Sep to Nov 2026.",
         html='<p class="eyebrow rise">The data</p><h2 class="rise d1">Real complaints, not sample data</h2><div class="grid g4" style="margin-bottom:56px">'
              '<div class="stat rise d2"><p class="num b">5,000</p><p class="label">maintenance complaints</p></div>'
              '<div class="stat rise d3"><p class="num">601</p><p class="label">residential buildings</p></div>'
              '<div class="stat rise d4"><p class="num">67</p><p class="label">days, 1 Sep – 6 Nov 2026</p></div>'
              '<div class="stat rise d5"><p class="num">13</p><p class="label">complaint categories</p></div></div>'
              '<p class="body rise d6" style="font-size:28px">Source: NYC Open Data, 311 Service Requests (HPD), ZIP 11226. Cleaned in Python: 9 of 44 fields kept.</p>'),
    dict(id="s10", sec=14, ch="HOW", bg="", anim=2.2,
         cap="Each complaint flows from the log to an SLA formula, then to the dashboard, automatically.",
         html='<p class="eyebrow rise">The flow</p><h2 class="rise d1">From complaint to action in six steps</h2><div class="grid g3">'
              '<div class="step rise d1"><p class="n">1</p><h3>Complaint raised</h3><p class="body">Call, app or form</p></div>'
              '<div class="step rise d2"><p class="n">2</p><h3>Ticket_Log</h3><p class="body">Nine fields, dropdowns</p></div>'
              '<div class="step rise d3"><p class="n">3</p><h3>Lists lookup</h3><p class="body">Category → priority → SLA days</p></div>'
              '<div class="step rise d4"><p class="n">4</p><h3>SLA formulas</h3><p class="body">Due date, TAT, status flag</p></div>'
              '<div class="step key rise d5"><p class="n">5</p><h3>Dashboard</h3><p class="body">Updates instantly</p></div>'
              '<div class="step rise d6"><p class="n">6</p><h3>Team acts</h3><p class="body">Overdue and repeat sites first</p></div></div>'),
    dict(id="s11", sec=12, ch="HOW", bg="", anim=2.0,
         cap="Each category gets a priority and a deadline: 2 days for heating, up to 21 days for paint.",
         html='<p class="eyebrow rise">SLA rules</p><h2 class="rise d1">Every category gets a priority and a deadline</h2>'
              '<table class="rise d2"><tr><th>Priority</th><th>Example categories</th><th class="r">SLA target</th></tr>'
              '<tr><td><b>Critical</b></td><td>Heat/Hot Water, Elevator</td><td class="r">2 days</td></tr>'
              '<tr><td><b>High</b></td><td>Water Leak, Electric, Plumbing, Safety</td><td class="r">7 days</td></tr>'
              '<tr><td><b>Medium</b></td><td>Door/Window, Appliance, Unsanitary</td><td class="r">14 days</td></tr>'
              '<tr><td><b>Low</b></td><td>Paint/Plaster, Outside Building</td><td class="r">21 days</td></tr></table>'
              '<p class="body rise d3" style="margin-top:28px">Targets are the tracker’s own standards and can be changed in one sheet.</p>'),
    dict(id="s12", sec=12, ch="HOW", bg="", anim=12.0,
         cap="The dashboard updates itself: KPIs, category breakdowns and charts, all from formulas.",
         html='<p class="eyebrow rise">The dashboard</p><h2 class="rise d1" style="margin-bottom:32px">Built entirely with Excel formulas</h2>'
              '<div class="shotwrap rise d2"><img src="dash.png"></div>'),
    dict(id="s13", sec=4, ch="WHERE", bg="dark", anim=1.4, cap="", html=chap("04", "Where", "Where it shows results, and where it fits")),
    dict(id="s14", sec=12, ch="WHERE", bg="dark", anim=2.0,
         cap="Result: 68.2% of complaints met their SLA, and 483 of 487 open jobs were already late.",
         html='<p class="eyebrow rise">Results on real data</p><h2 class="rise d1">Two in three complaints met their SLA</h2><div class="grid g4">'
              '<div class="stat rise d2"><p class="num b">68.2%</p><p class="label">SLA compliance</p></div>'
              '<div class="stat rise d3"><p class="num">7.9</p><p class="label">days average turnaround</p></div>'
              '<div class="stat rise d4"><p class="num">483</p><p class="label">of 487 open jobs past target</p></div>'
              '<div class="stat rise d5"><p class="num">+41%</p><p class="label">complaints, Sep to Oct</p></div></div>'),
    dict(id="s15", sec=12, ch="WHERE", bg="", anim=2.2,
         cap="Elevator, leak, electric and plumbing jobs were fixed on time less than half the time.",
         html='<p class="eyebrow rise">Where SLAs are missed</p><h2 class="rise d1" style="margin-bottom:36px">SLA compliance by category</h2>'
              + "".join(f'<div class="brow{" low" if v < 50 else ""}"><p class="nm">{n}</p><div class="tr"><div class="fl" style="width:{v}%;animation-delay:{0.25 + i * 0.06:.2f}s"></div></div><p class="vl">{v:.1f}%</p></div>'
                        for i, (n, v) in enumerate([("Elevator", 19.0), ("Water Leak", 47.3), ("Electric", 47.5), ("Plumbing", 48.9), ("Safety", 53.6), ("Heat/Hot Water", 65.9), ("General", 70.6), ("Door/Window", 73.0), ("Unsanitary", 77.1), ("Paint/Plaster", 91.9)]))),
    dict(id="s16", sec=10, ch="WHERE", bg="blue", anim=2.0,
         cap="One building logged 278 complaints, 126 about heat. That is where to act first.",
         html='<div style="padding-top:60px"><p class="eyebrow rise">Biggest single problem</p><h1 class="rise d1" style="font-size:150px">80 Lenox Road</h1>'
              '<p class="lead rise d2" style="margin-top:36px;font-size:46px;line-height:1.3">278 complaints in 67 days, 126 of them for heat or hot water. A clear case for boiler preventive maintenance.</p></div>'),
    dict(id="s17", sec=10, ch="WHERE", bg="", anim=2.2,
         cap="The same tracker works anywhere a facility team handles repair requests.",
         html='<p class="eyebrow rise">Where it can be used</p><h2 class="rise d1">Anywhere a facility team handles repairs</h2><div class="grid g3 uses">'
              '<div class="card rise d2"><h3>College campuses</h3></div><div class="card rise d3"><h3>Corporate offices</h3></div>'
              '<div class="card rise d4"><h3>Warehouses &amp; fulfilment centres</h3></div><div class="card rise d5"><h3>Hospitals</h3></div>'
              '<div class="card rise d6"><h3>Hotels</h3></div><div class="card rise d7"><h3>Apartment societies</h3></div></div>'),
    dict(id="s18", sec=9, ch=None, bg="dark", anim=2.0,
         cap="Thank you. Manjunatha P · github.com/manjjunatha",
         html='<div style="padding-top:110px"><p class="eyebrow rise">Summary</p><div class="accbar"></div><h2 class="rise d1">A simple Excel tool that turns raw complaints into decisions</h2>'
              '<p class="lead rise d2">Excel formulas &amp; dashboards · Python data cleaning · SLA monitoring · KPI reporting</p>'
              '<p class="body rise d3" style="margin-top:56px;font-size:28px">Manjunatha P · manjjunathp18@gmail.com · github.com/manjjunatha</p></div>'),
]

assert sum(s["sec"] for s in SCENES) == 180, sum(s["sec"] for s in SCENES)


def page(s):
    cap = f'<p class="cap">{s["cap"]}</p>' if s["cap"] else ""
    tp = top(s["ch"]) if s["ch"] else ""
    return (f"<html><head><meta charset='utf-8'><style>{CSS}</style></head>"
            f"<body class='{s['bg']}'><div class='stage {s['bg']}'>{tp}{s['html']}{cap}</div></body></html>")
