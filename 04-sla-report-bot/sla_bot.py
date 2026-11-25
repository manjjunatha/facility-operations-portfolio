#!/usr/bin/env python3
"""
SLA Report Bot
==============
Give it any ticket / complaint / work-order file (CSV or Excel). It:
  1. finds the columns it needs (ID, created date, closed date, status, priority,
     category, location, assignee) by their names, whatever they are called,
  2. applies the SLA rulebook in sla_rules.json exactly, ticket by ticket,
  3. checks data quality and never silently drops a ticket,
  4. writes a formatted Excel SLA report next to the input file.

Usage:
    python sla_bot.py <input file> [--rules sla_rules.json] [--out report.xlsx] [--as-of 2026-11-25]
Or on Windows: drag the file onto SLA_Bot.bat
"""
import argparse
import datetime as dt
import json
import os
import re
import sys

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

VERSION = "1.0"
PRIOS = ["Critical", "High", "Medium", "Low"]

# ----------------------------------------------------------------------------- column detection
SYNONYMS = {
    "id": ["ticket id", "ticket no", "ticket number", "unique key", "sr number", "request id", "case id",
           "incident id", "complaint id", "work order", "wo number", "ticket", "id", "number", "ref", "reference"],
    "created": ["created date", "created at", "created on", "creation date", "created", "date opened", "opened date",
                "open date", "opened", "raised date", "raised on", "raised", "reported date", "reported on", "reported",
                "date logged", "logged date", "logged", "submitted date", "submitted", "request date", "start date",
                "received date", "date"],
    "closed": ["closed date", "closed at", "closed on", "date closed", "closed", "resolved date", "resolved at",
               "resolved on", "resolved", "resolution date", "completed date", "completion date", "completed",
               "end date", "finish date", "fixed date"],
    "status": ["ticket status", "current status", "status", "state"],
    "priority": ["priority", "severity", "urgency", "impact"],
    "category": ["complaint type", "issue type", "problem type", "request type", "category", "sub category",
                 "type", "service", "trade", "department", "issue", "problem"],
    "location": ["building", "site", "incident address", "address", "location", "block", "facility", "area",
                 "zone", "room", "branch", "floor", "city"],
    "assignee": ["assigned vendor", "assigned to", "assignee", "vendor", "contractor", "technician", "engineer",
                 "owner", "team", "agency name", "agency", "resolver group"],
}


def gapply(gb, fn):
    """groupby.apply that works on old and new pandas."""
    try:
        return gb.apply(fn, include_groups=False)
    except TypeError:
        return gb.apply(fn)


def to_dt(s, day_first):
    try:
        return pd.to_datetime(s, errors="coerce", dayfirst=day_first, format="mixed")
    except (TypeError, ValueError):
        return pd.to_datetime(s, errors="coerce", dayfirst=day_first)


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def looks_like_coords(series):
    v = series.dropna().astype(str).head(200)
    if v.empty:
        return False
    pat = re.compile(r"^\(?\s*-?\d{1,3}\.\d+\s*,\s*-?\d{1,3}\.\d+\s*\)?$|^\{|POINT\s*\(")
    return v.map(lambda x: bool(pat.search(x))).mean() > 0.5


def date_parse_rate(series, day_first):
    s = series.dropna()
    if s.empty:
        return 0.0
    if pd.api.types.is_datetime64_any_dtype(s):
        return 1.0
    sample = s.astype(str).head(500)
    if sample.str.fullmatch(r"\d+(\.\d+)?").mean() > 0.8:  # plain numbers are not dates here
        return 0.0
    return to_dt(sample, day_first).notna().mean()


def detect_columns(df, overrides):
    found, log = {}, {}
    used = set()
    headers = {c: norm(c) for c in df.columns}
    for role in ["created", "closed", "id", "status", "priority", "category", "location", "assignee"]:
        want = (overrides or {}).get(role, "")
        if want:
            if want in df.columns:
                found[role] = want; used.add(want); log[role] = f"'{want}' (set in rules file)"
            else:
                log[role] = f"'{want}' set in rules file but NOT in the data"
            continue
        best = None
        for c, h in headers.items():
            if c in used:
                continue
            for i, syn in enumerate(SYNONYMS[role]):
                score = None
                if h == syn:
                    score = (3, -i)
                elif re.search(rf"(^| ){re.escape(syn)}( |$)", h):
                    score = (2, -i)
                if score is None:
                    continue
                col = df[c]
                if role in ("created", "closed") and date_parse_rate(col, False) < 0.5 and date_parse_rate(col, True) < 0.5:
                    continue
                if role in ("category", "location", "assignee", "status", "priority"):
                    if col.dropna().empty or looks_like_coords(col):
                        continue
                    if role == "location" and pd.api.types.is_numeric_dtype(col):
                        continue
                if best is None or score > best[0]:
                    best = (score, c)
                break
        if best:
            found[role] = best[1]; used.add(best[1])
            log[role] = f"'{best[1]}' (auto-detected)"
        else:
            log[role] = "not found"
    # last resort for created date: first column that is mostly dates
    if "created" not in found:
        for c in df.columns:
            if c not in used and max(date_parse_rate(df[c], False), date_parse_rate(df[c], True)) > 0.8:
                found["created"] = c; used.add(c); log["created"] = f"'{c}' (first date-like column)"
                break
    return found, log


def parse_dates(series, day_first_setting):
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce"), "native dates"
    s = series.astype("string")
    if day_first_setting in (True, False):
        return to_dt(s, day_first_setting), f"day_first={day_first_setting}"
    a = to_dt(s, False)
    b = to_dt(s, True)
    # prefer the reading where more values parse; tie -> the one with fewer dates in the future
    now = pd.Timestamp.now()
    score_a = (a.notna().sum(), -(a > now).sum())
    score_b = (b.notna().sum(), -(b > now).sum())
    return (b, "auto: day-first") if score_b > score_a else (a, "auto: month-first / ISO")


# ----------------------------------------------------------------------------- rules
def load_rules(path):
    here = os.path.dirname(os.path.abspath(__file__))
    path = path or os.path.join(here, "sla_rules.json")
    with open(path, encoding="utf-8") as fh:
        return json.load(fh), path


def build_priority_mapper(rules):
    alias = {}
    for p, vals in rules.get("priority_aliases", {}).items():
        if p.startswith("_"):
            continue
        for v in vals:
            alias[norm(v)] = p
        alias[norm(p)] = p
    exact = {norm(k): v for k, v in rules.get("category_priority", {}).get("exact", {}).items()}
    kw = [(p, norm(w)) for p, words in rules.get("category_priority", {}).get("keywords", {}).items() for w in words]
    return alias, exact, kw


# ----------------------------------------------------------------------------- engine
def run(input_path, rules, as_of_cli=None, out_path=None):
    t0 = dt.datetime.now()
    # ---- load
    ext = os.path.splitext(input_path)[1].lower()
    sheet_used = None
    if ext in (".xlsx", ".xlsm", ".xls"):
        book = pd.read_excel(input_path, sheet_name=None)
        sheet_used, raw = max(book.items(), key=lambda kv: len(kv[1]))
    else:
        raw = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                raw = pd.read_csv(input_path, low_memory=False, encoding=enc)
                break
            except UnicodeDecodeError:
                continue
    raw = raw.dropna(how="all")
    raw.columns = [str(c).strip() for c in raw.columns]
    n_raw = len(raw)

    cols, col_log = detect_columns(raw, rules.get("columns", {}))
    if "created" not in cols:
        raise SystemExit("Could not find a 'created / opened / raised' date column. Set columns.created in sla_rules.json.")

    df = pd.DataFrame(index=raw.index)
    df["Ticket ID"] = raw[cols["id"]].astype(str) if "id" in cols else [f"ROW-{i + 2}" for i in range(n_raw)]
    df["Source Row"] = np.arange(2, n_raw + 2)
    day_first = rules.get("date_format", {}).get("day_first", "auto")
    df["Created"], created_how = parse_dates(raw[cols["created"]], day_first)
    if "closed" in cols:
        df["Closed"], closed_how = parse_dates(raw[cols["closed"]], day_first)
    else:
        df["Closed"], closed_how = pd.NaT, "no closed-date column"
    for role, label in [("status", "Status"), ("category", "Category"), ("location", "Location"), ("assignee", "Assignee")]:
        df[label] = raw[cols[role]].astype("string").str.replace(r"\s+", " ", regex=True).str.strip() if role in cols else pd.NA
    df["Priority (source)"] = raw[cols["priority"]].astype("string").str.strip() if "priority" in cols else pd.NA

    # ---- as-of date
    as_of_rule = as_of_cli or rules.get("as_of", {}).get("value", "today")
    if str(as_of_rule).lower() == "today":
        as_of = pd.Timestamp(dt.datetime.now().replace(microsecond=0))
    elif str(as_of_rule).lower() == "latest":
        as_of = max(df["Created"].max(), df["Closed"].max() if df["Closed"].notna().any() else df["Created"].max())
    else:
        as_of = pd.Timestamp(as_of_rule)
        if len(str(as_of_rule).strip()) <= 10:  # date only -> end of that day
            as_of = as_of + pd.Timedelta(hours=23, minutes=59, seconds=59)

    # ---- priority
    alias, exact, kw = build_priority_mapper(rules)
    default_p = rules.get("default_priority", "Medium")
    pri_rule_count = {"from priority column": 0, "category exact rule": 0, "category keyword rule": 0, "default priority": 0}
    dq = []  # (row index, severity, issue)

    def map_priority(i, src, cat):
        if "priority" in cols:
            if pd.notna(src) and norm(src) in alias:
                pri_rule_count["from priority column"] += 1
                return alias[norm(src)], "Priority column"
            dq.append((i, "Warning", f"Priority value '{src}' not recognised; default '{default_p}' used"))
            pri_rule_count["default priority"] += 1
            return default_p, "Default (unrecognised value)"
        c = norm(cat) if pd.notna(cat) else ""
        if c in exact:
            pri_rule_count["category exact rule"] += 1
            return exact[c], "Category rule"
        for p, w in kw:
            if w and re.search(rf"(^| ){re.escape(w)}( |$)", c):
                pri_rule_count["category keyword rule"] += 1
                return p, f"Keyword '{w}'"
        pri_rule_count["default priority"] += 1
        return default_p, "Default priority"

    pr = [map_priority(i, s, c) for i, s, c in zip(df.index, df["Priority (source)"], df["Category"])]
    df["Priority"] = [p for p, _ in pr]
    df["Priority Rule"] = [r for _, r in pr]
    targets = {k: float(v) for k, v in rules["sla_targets_days"].items()}
    df["SLA Target (days)"] = df["Priority"].map(targets)

    # ---- clock
    clock = rules.get("clock", {})
    mode = clock.get("mode", "calendar")
    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    weekend = set(clock.get("weekend_days", ["Sat", "Sun"]))
    weekmask = [0 if d in weekend else 1 for d in days]
    holidays = np.array([np.datetime64(h) for h in clock.get("holidays", [])], dtype="datetime64[D]")

    def elapsed(start, end):
        if pd.isna(start) or pd.isna(end):
            return np.nan
        if mode == "business":
            return float(np.busday_count(start.date(), end.date(), weekmask=weekmask, holidays=holidays))
        return (end - start).total_seconds() / 86400.0

    closed_set = {norm(s) for s in rules.get("closed_statuses", [])}
    excl_set = {norm(s) for s in rules.get("excluded_statuses", [])}
    at_risk = float(rules.get("at_risk_threshold", 0.75))

    # duplicates
    if "id" in cols:
        dups = df["Ticket ID"].duplicated(keep="first")
        for i in df.index[dups]:
            dq.append((i, "Error", "Duplicate ticket ID; later copy excluded"))
    else:
        dups = pd.Series(False, index=df.index)

    sla, tat, age, due, why = [], [], [], [], []
    for i, r in df.iterrows():
        st = norm(r["Status"]) if pd.notna(r["Status"]) else ""
        cr, cl, tgt = r["Created"], r["Closed"], r["SLA Target (days)"]
        due.append(cr + pd.Timedelta(days=tgt) if pd.notna(cr) and mode == "calendar" else pd.NaT)
        if dups.loc[i]:
            sla.append("Excluded"); tat.append(np.nan); age.append(np.nan); why.append("Duplicate ID"); continue
        if st and st in excl_set:
            sla.append("Excluded"); tat.append(np.nan); age.append(np.nan); why.append(f"Status '{r['Status']}' is excluded"); continue
        if pd.isna(cr):
            dq.append((i, "Error", "Created date missing or unreadable")); sla.append("Data issue")
            tat.append(np.nan); age.append(np.nan); why.append("No created date"); continue
        if cr > as_of:
            dq.append((i, "Warning", "Created date is after the as-of date"))
        if pd.notna(cl):
            if cl < cr:
                dq.append((i, "Error", "Closed date is before created date")); sla.append("Data issue")
                tat.append(np.nan); age.append(np.nan); why.append("Closed before created"); continue
            if st and st not in closed_set:
                dq.append((i, "Warning", f"Has a closed date but status is '{r['Status']}'; treated as closed"))
            e = elapsed(cr, cl)
            tat.append(e); age.append(np.nan)
            sla.append("Met" if e <= tgt else "Breached")
            why.append(f"Closed in {e:.1f} days vs target {tgt:g}")
        else:
            if st and st in closed_set:
                dq.append((i, "Error", "Status says closed but closed date is missing")); sla.append("Data issue")
                tat.append(np.nan); age.append(np.nan); why.append("Closed status without date"); continue
            e = elapsed(cr, as_of)
            tat.append(np.nan); age.append(e)
            if e > tgt:
                sla.append("Open - Overdue")
            elif e >= at_risk * tgt:
                sla.append("Open - At Risk")
            else:
                sla.append("Open - On Track")
            why.append(f"Open {e:.1f} days vs target {tgt:g}")
    df["Due Date"] = due
    df["TAT (days)"] = tat
    df["Days Open"] = age
    df["SLA Status"] = sla
    df["Reason"] = why
    df["Overrun (days)"] = np.where(df["SLA Status"] == "Breached", df["TAT (days)"] - df["SLA Target (days)"],
                                   np.where(df["SLA Status"] == "Open - Overdue", df["Days Open"] - df["SLA Target (days)"], np.nan))
    df["Month"] = df["Created"].dt.to_period("M").dt.to_timestamp()

    # ---- metrics
    def summarize(g):
        met = (g["SLA Status"] == "Met").sum()
        br = (g["SLA Status"] == "Breached").sum()
        return pd.Series({
            "Tickets": int((~g["SLA Status"].isin(["Excluded"])).sum()),
            "Closed": int(met + br),
            "Open": int(g["SLA Status"].str.startswith("Open").sum()),
            "Met": int(met), "Breached": int(br),
            "Compliance %": met / (met + br) if met + br else np.nan,
            "Avg TAT (days)": g["TAT (days)"].mean(),
            "Median TAT (days)": g["TAT (days)"].median(),
            "Open Overdue": int((g["SLA Status"] == "Open - Overdue").sum()),
            "Open At Risk": int((g["SLA Status"] == "Open - At Risk").sum()),
            "Data Issues": int((g["SLA Status"] == "Data issue").sum()),
        })

    overall = summarize(df)
    target = float(rules.get("compliance_target", 0.9))
    margin = float(rules.get("rag_amber_margin", 0.1))
    comp = overall["Compliance %"]
    rag = "No data" if pd.isna(comp) else ("GREEN" if comp >= target else ("AMBER" if comp >= target - margin else "RED"))

    breakdowns = {}
    labels = {"priority": "Priority", "category": "Category", "location": "Location", "assignee": "Assignee"}
    top_n = int(rules.get("top_n", 15))
    for b in rules.get("breakdowns", ["priority", "category", "location", "assignee"]):
        col = labels.get(b)
        if not col or (b != "priority" and b not in cols):
            continue
        sub = df[df["SLA Status"] != "Excluded"].copy()
        sub[col] = sub[col].fillna("(blank)")
        tab = gapply(sub.groupby(col, dropna=False), summarize)
        if b == "priority":
            tab = tab.reindex([p for p in PRIOS if p in tab.index] + [p for p in tab.index if p not in PRIOS])
        else:
            tab = tab.sort_values("Tickets", ascending=False)
            if len(tab) > top_n:
                keep = tab.index[:top_n]
                rest = sub[~sub[col].isin(keep)]
                tab = pd.concat([tab.loc[keep], summarize(rest).to_frame(f"Other ({len(tab) - top_n} more)").T])
        breakdowns[labels[b]] = tab

    monthly = None
    if df["Month"].notna().any():
        m = gapply(df[df["SLA Status"] != "Excluded"].groupby("Month"), summarize)
        closed_in = df[df["Closed"].notna()].groupby(df["Closed"].dt.to_period("M").dt.to_timestamp()).size()
        m["Closed in Month"] = closed_in.reindex(m.index).fillna(0).astype(int)
        monthly = m

    # ---- findings (plain-English, generated from the numbers)
    findings = []
    if pd.notna(comp):
        findings.append(f"Overall SLA compliance is {comp:.1%} against a target of {target:.0%}: {rag}. "
                        f"{int(overall['Met']):,} closed tickets met their SLA and {int(overall['Breached']):,} breached it.")
    if overall["Open"]:
        oldest = df.loc[df["SLA Status"].str.startswith("Open"), "Days Open"].max()
        findings.append(f"{int(overall['Open']):,} tickets are open: {int(overall['Open Overdue']):,} already overdue and "
                        f"{int(overall['Open At Risk']):,} at risk (≥{at_risk:.0%} of target used). Oldest open ticket: {oldest:.0f} days.")
    pt = breakdowns.get("Priority")
    if pt is not None and pt["Compliance %"].notna().any():
        w = pt["Compliance %"].idxmin()
        findings.append(f"Weakest priority: {w}, at {pt.loc[w, 'Compliance %']:.1%} compliance "
                        f"(target {targets.get(w, float('nan')):g} days).")
    min_n = max(10, int(0.01 * overall["Closed"]))
    for key in ("Category", "Location", "Assignee"):
        t = breakdowns.get(key)
        if t is None:
            continue
        t2 = t[(t["Closed"] >= min_n) & ~t.index.astype(str).str.startswith("Other (")]
        if len(t2) < 2:
            continue
        if not t2.empty and t2["Compliance %"].notna().any():
            w = t2["Compliance %"].idxmin()
            findings.append(f"Weakest {key.lower()} (min. {min_n} closed tickets): {w}, at {t2.loc[w, 'Compliance %']:.1%} "
                            f"compliance over {int(t2.loc[w, 'Closed']):,} closed tickets.")
        if key == "Location" and not t.empty:
            top = t[~t.index.astype(str).str.startswith("Other (")]
            if not top.empty:
                share = top.iloc[0]["Tickets"] / overall["Tickets"]
                findings.append(f"Busiest location: {top.index[0]} with {int(top.iloc[0]['Tickets']):,} tickets ({share:.1%} of all).")
    n_err = sum(1 for _, s, _ in dq if s == "Error")
    n_warn = sum(1 for _, s, _ in dq if s == "Warning")
    findings.append(f"Data quality: {n_err:,} errors and {n_warn:,} warnings found; "
                    f"{int(overall['Data Issues']):,} tickets could not be scored and are listed on the Data Quality sheet.")

    meta = dict(input=os.path.abspath(input_path), sheet=sheet_used, rows=n_raw, cols=cols, col_log=col_log,
                created_how=created_how, closed_how=closed_how, as_of=as_of, as_of_rule=as_of_rule, mode=mode,
                weekend=sorted(weekend, key=days.index), holidays=clock.get("holidays", []), targets=targets,
                target=target, margin=margin, rag=rag, at_risk=at_risk, pri_rule_count=pri_rule_count,
                closed_statuses=rules.get("closed_statuses", []), excluded_statuses=rules.get("excluded_statuses", []),
                default_p=default_p, generated=t0, title=rules.get("report_title", "SLA Performance Report"),
                org=rules.get("organisation", ""), excluded=int((df["SLA Status"] == "Excluded").sum()))
    out_path = out_path or os.path.join(os.path.dirname(os.path.abspath(input_path)),
                                        f"SLA_Report_{os.path.splitext(os.path.basename(input_path))[0]}_{t0:%Y%m%d_%H%M}.xlsx")
    write_report(out_path, df, overall, breakdowns, monthly, findings, dq, meta)
    return out_path, overall, rag, findings


# ----------------------------------------------------------------------------- excel writer
NAVY, TEAL, LIGHT, GREY = "1F3864", "2E75B6", "DDEBF7", "F2F2F2"
FONT = "Arial"
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
RAG_FILL = {"GREEN": "C6EFCE", "AMBER": "FFEB9C", "RED": "F8CBAD", "No data": "EDEDED"}
RAG_FONT = {"GREEN": "006100", "AMBER": "7F6000", "RED": "9C0006", "No data": "595959"}


def F(bold=False, size=10, color="000000", italic=False):
    return Font(name=FONT, bold=bold, size=size, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def header(ws, row, names, col=1, color=NAVY):
    for j, n in enumerate(names):
        c = ws.cell(row=row, column=col + j, value=n)
        c.font = F(True, 10, "FFFFFF"); c.fill = fill(color); c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30


def title(ws, text, sub=None):
    ws.sheet_view.showGridLines = False
    ws["A1"] = text; ws["A1"].font = F(True, 16, NAVY)
    if sub:
        ws["A2"] = sub; ws["A2"].font = F(False, 9, "595959", True)


def clean(v):
    if v is None:
        return None
    if isinstance(v, float) and np.isnan(v):
        return None
    if v is pd.NaT or (isinstance(v, pd.Timestamp) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return None if np.isnan(v) else float(v)
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


def write_table(ws, r0, df_, fmts=None, index_name=None, col0=1):
    names = ([index_name] if index_name else []) + list(df_.columns)
    header(ws, r0, names, col0)
    for i, (idx, row) in enumerate(df_.iterrows()):
        vals = ([idx] if index_name else []) + list(row.values)
        for j, v in enumerate(vals):
            c = ws.cell(row=r0 + 1 + i, column=col0 + j, value=clean(v))
            c.font = F(); c.border = BORDER
            if fmts and names[j] in fmts:
                c.number_format = fmts[names[j]]
    return r0 + len(df_)


BFMT = {"Compliance %": "0.0%", "Avg TAT (days)": "0.0", "Median TAT (days)": "0.0", "Tickets": "#,##0",
        "Closed": "#,##0", "Open": "#,##0", "Met": "#,##0", "Breached": "#,##0", "Open Overdue": "#,##0",
        "Open At Risk": "#,##0", "Data Issues": "#,##0", "Closed in Month": "#,##0"}


def add_compliance_format(ws, rng, target, margin):
    first = rng.split(":")[0]
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first}),{first}>={target})'], fill=fill("C6EFCE")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first}),{first}<{target},{first}>={target - margin})'], fill=fill("FFEB9C")))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(ISNUMBER({first}),{first}<{target - margin})'], fill=fill("F8CBAD")))


def write_report(path, df, overall, breakdowns, monthly, findings, dq, m):
    wb = Workbook()
    # ---------------- Summary
    ws = wb.active; ws.title = "Summary"
    title(ws, m["title"] + (f" — {m['org']}" if m["org"] else ""),
          f"Source: {os.path.basename(m['input'])}" + (f" (sheet '{m['sheet']}')" if m["sheet"] else "") +
          f"  |  As of {m['as_of']:%d-%b-%Y %H:%M}  |  Generated {m['generated']:%d-%b-%Y %H:%M} by SLA Report Bot v{VERSION}")
    ws.column_dimensions["A"].width = 34; ws.column_dimensions["B"].width = 16
    for c in "CDEFGHIJ":
        ws.column_dimensions[c].width = 13
    ws["A4"] = "SLA status"; ws["A4"].font = F(True, 11)
    ws["B4"] = m["rag"]; ws["B4"].font = Font(name=FONT, bold=True, size=14, color=RAG_FONT[m["rag"]])
    ws["B4"].fill = fill(RAG_FILL[m["rag"]]); ws["B4"].alignment = Alignment(horizontal="center")
    ws["C4"] = f"Target {m['target']:.0%} · Amber within {m['margin']:.0%}"; ws["C4"].font = F(italic=True, color="595959")
    kpis = [("Tickets in scope", overall["Tickets"], "#,##0"), ("Closed (scored)", overall["Closed"], "#,##0"),
            ("Met SLA", overall["Met"], "#,##0"), ("Breached SLA", overall["Breached"], "#,##0"),
            ("SLA compliance", overall["Compliance %"], "0.0%"), ("Average TAT (days)", overall["Avg TAT (days)"], "0.0"),
            ("Median TAT (days)", overall["Median TAT (days)"], "0.0"), ("Open tickets", overall["Open"], "#,##0"),
            ("Open – overdue", overall["Open Overdue"], "#,##0"), ("Open – at risk", overall["Open At Risk"], "#,##0"),
            ("Data issues (not scored)", overall["Data Issues"], "#,##0"), ("Excluded (rules)", m["excluded"], "#,##0")]
    header(ws, 6, ["Key figure", "Value"])
    for i, (k, v, f_) in enumerate(kpis):
        a = ws.cell(row=7 + i, column=1, value=k); a.font = F(); a.border = BORDER
        b = ws.cell(row=7 + i, column=2, value=clean(v)); b.font = F(True); b.border = BORDER; b.number_format = f_
    ws.cell(row=11, column=2).fill = fill(RAG_FILL[m["rag"]])
    r = 7 + len(kpis) + 1
    ws.cell(row=r, column=1, value="Key findings").font = F(True, 12, NAVY)
    for i, t in enumerate(findings):
        c = ws.cell(row=r + 1 + i, column=1, value=f"{i + 1}. {t}")
        c.font = F(); c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r + 1 + i, start_column=1, end_row=r + 1 + i, end_column=10)
        ws.row_dimensions[r + 1 + i].height = 30
    r = r + len(findings) + 2
    pt = breakdowns.get("Priority")
    if pt is not None:
        ws.cell(row=r, column=1, value="Compliance by priority").font = F(True, 12, NAVY)
        cols_ = ["Tickets", "Closed", "Met", "Breached", "Compliance %", "Median TAT (days)", "Open Overdue"]
        t = pt[cols_].copy(); t.insert(0, "SLA Target (days)", [m["targets"].get(p) for p in t.index])
        end = write_table(ws, r + 1, t, {**BFMT, "SLA Target (days)": "0"}, "Priority")
        add_compliance_format(ws, f"F{r + 2}:F{end + 1}", m["target"], m["margin"])
        ch = BarChart(); ch.type = "bar"; ch.title = "SLA compliance by priority"; ch.height = 6; ch.width = 14
        ch.add_data(Reference(ws, min_col=6, min_row=r + 1, max_row=end + 1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=r + 2, max_row=end + 1))
        ch.x_axis.scaling.orientation = "maxMin"; ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1
        ch.y_axis.number_format = "0%"; ch.legend = None; ch.roundedCorners = False
        ch.x_axis.delete = False; ch.y_axis.delete = False
        ws.add_chart(ch, f"A{end + 4}")

    # ---------------- breakdown sheets
    for name, tab in breakdowns.items():
        if name == "Priority":
            continue
        s = wb.create_sheet(f"By {name}")
        title(s, f"SLA performance by {name.lower()}", f"Sorted by ticket count; compliance = Met ÷ (Met + Breached). Colours vs target {m['target']:.0%}.")
        end = write_table(s, 4, tab, BFMT, name)
        add_compliance_format(s, f"G5:G{end + 1}", m["target"], m["margin"])
        s.column_dimensions["A"].width = 34
        for j in range(2, 14):
            s.column_dimensions[get_column_letter(j)].width = 12
        s.freeze_panes = "B5"
        n = min(len(tab), 15)
        ch = BarChart(); ch.type = "bar"; ch.title = f"Tickets by {name.lower()} (met / breached / open)"
        ch.grouping = "stacked"; ch.overlap = 100; ch.height = 9; ch.width = 18
        for colname in ("Met", "Breached", "Open"):
            ci = list(tab.columns).index(colname) + 2
            ch.add_data(Reference(s, min_col=ci, min_row=4, max_row=4 + n), titles_from_data=True)
        ch.set_categories(Reference(s, min_col=1, min_row=5, max_row=4 + n))
        ch.x_axis.scaling.orientation = "maxMin"; ch.roundedCorners = False; ch.x_axis.delete = False; ch.y_axis.delete = False
        s.add_chart(ch, f"A{end + 3}")
    if "Priority" in breakdowns:
        s = wb.create_sheet("By Priority", 1)
        title(s, "SLA performance by priority", "Targets from the rulebook; compliance = Met ÷ (Met + Breached).")
        t = breakdowns["Priority"].copy(); t.insert(0, "SLA Target (days)", [m["targets"].get(p) for p in t.index])
        end = write_table(s, 4, t, {**BFMT, "SLA Target (days)": "0"}, "Priority")
        add_compliance_format(s, f"H5:H{end + 1}", m["target"], m["margin"])
        s.column_dimensions["A"].width = 16
        for j in range(2, 15):
            s.column_dimensions[get_column_letter(j)].width = 12

    # ---------------- monthly trend
    if monthly is not None and len(monthly):
        s = wb.create_sheet("Monthly Trend")
        title(s, "Monthly trend (by month raised)", "Closed in Month counts tickets by their closed date.")
        t = monthly[["Tickets", "Closed in Month", "Met", "Breached", "Compliance %", "Median TAT (days)", "Open"]]
        end = write_table(s, 4, t, {**BFMT, "Month": "mmm-yyyy"}, "Month")
        for rr in range(5, end + 1):
            s.cell(row=rr, column=1).number_format = "mmm-yyyy"
        add_compliance_format(s, f"F5:F{end}", m["target"], m["margin"])
        for j in range(1, 9):
            s.column_dimensions[get_column_letter(j)].width = 14
        ch = BarChart(); ch.type = "col"; ch.title = "Tickets raised vs closed per month"; ch.height = 7; ch.width = 16
        for ci in (2, 3):
            ch.add_data(Reference(s, min_col=ci, min_row=4, max_row=end), titles_from_data=True)
        ch.set_categories(Reference(s, min_col=1, min_row=5, max_row=end)); ch.x_axis.number_format = "mmm-yy"
        ch.roundedCorners = False; ch.x_axis.delete = False; ch.y_axis.delete = False
        s.add_chart(ch, f"J4")

    # ---------------- action lists
    show = ["Ticket ID", "Created", "Closed", "Priority", "Category", "Location", "Assignee", "Status",
            "SLA Target (days)", "TAT (days)", "Days Open", "Overrun (days)", "SLA Status"]
    dfmt = {"Created": "dd-mmm-yyyy hh:mm", "Closed": "dd-mmm-yyyy hh:mm", "TAT (days)": "0.0", "Days Open": "0.0",
            "Overrun (days)": "0.0", "SLA Target (days)": "0", "Due Date": "dd-mmm-yyyy hh:mm"}
    present = [c for c in show if c in df.columns and (df[c].notna().any() or c in ("Closed", "TAT (days)", "Days Open"))]
    s = wb.create_sheet("Open Overdue")
    title(s, "Open tickets past their SLA — act on these first", "Sorted by days overdue (largest first). Includes 'At Risk' tickets below the overdue ones.")
    od = df[df["SLA Status"].isin(["Open - Overdue", "Open - At Risk"])].copy()
    od["_o"] = od["SLA Status"].eq("Open - Overdue")
    od = od.sort_values(["_o", "Overrun (days)", "Days Open"], ascending=[False, False, False])[[c for c in present if c not in ("Closed", "TAT (days)")]]
    write_table(s, 4, od.reset_index(drop=True), dfmt)
    _widths(s, od.columns)
    s = wb.create_sheet("Breaches")
    title(s, "Closed tickets that breached SLA", "Sorted by overrun (days beyond target).")
    bd = df[df["SLA Status"] == "Breached"].sort_values("Overrun (days)", ascending=False)[[c for c in present if c != "Days Open"]]
    write_table(s, 4, bd.reset_index(drop=True), dfmt)
    _widths(s, bd.columns)

    # ---------------- detail
    s = wb.create_sheet("Ticket Detail")
    title(s, "Every ticket with the rule result", "One row per input row. 'Reason' explains each decision; 'Source Row' is the row in your file.")
    det_cols = ["Source Row", "Ticket ID", "Created", "Closed", "Status", "Category", "Location", "Assignee",
                "Priority (source)", "Priority", "Priority Rule", "SLA Target (days)", "Due Date", "TAT (days)",
                "Days Open", "SLA Status", "Reason"]
    det_cols = [c for c in det_cols if c in df.columns and (df[c].notna().any() or c in ("Closed", "TAT (days)", "Days Open"))]
    write_table(s, 4, df[det_cols].reset_index(drop=True), dfmt)
    _widths(s, det_cols)
    last = 4 + len(df)
    sc = get_column_letter(det_cols.index("SLA Status") + 1)
    rng = f"{sc}5:{sc}{last}"
    for val, colr in [("Met", "C6EFCE"), ("Breached", "F8CBAD"), ("Open - Overdue", "F8CBAD"), ("Open - At Risk", "FFEB9C"),
                      ("Data issue", "D9D9D9")]:
        s.conditional_formatting.add(rng, FormulaRule(formula=[f'${sc}5="{val}"'], fill=fill(colr)))
    s.freeze_panes = "C5"; s.auto_filter.ref = f"A4:{get_column_letter(len(det_cols))}{last}"

    # ---------------- data quality
    s = wb.create_sheet("Data Quality")
    title(s, "Data quality checks", "Errors stop a ticket from being scored; warnings are scored but should be fixed at source.")
    dqdf = pd.DataFrame(dq, columns=["idx", "Severity", "Issue"])
    if len(dqdf):
        cnt = dqdf.groupby(["Severity", "Issue"]).size().rename("Rows").reset_index()
        cnt["Issue"] = cnt["Issue"].str.replace(r"'[^']*'", "'…'", regex=True)
        cnt = cnt.groupby(["Severity", "Issue"], as_index=False)["Rows"].sum().sort_values(["Severity", "Rows"], ascending=[True, False])
        end = write_table(s, 4, cnt.reset_index(drop=True), {"Rows": "#,##0"})
        r0 = end + 3
        s.cell(row=r0 - 1, column=1, value="Rows with issues").font = F(True, 11, NAVY)
        det = dqdf.merge(df[["Source Row", "Ticket ID"]], left_on="idx", right_index=True).drop(columns="idx")
        write_table(s, r0, det[["Source Row", "Ticket ID", "Severity", "Issue"]].reset_index(drop=True))
    else:
        s["A4"] = "No data quality issues found."; s["A4"].font = F(True, 11, "006100")
    for col_, w in zip("ABCD", (12, 22, 12, 70)):
        s.column_dimensions[col_].width = w

    # ---------------- rules applied
    s = wb.create_sheet("Rules Applied")
    title(s, "Rules applied in this report", "Everything the bot decided, so the report can be audited and repeated.")
    rows = [("Input file", m["input"]), ("Rows read", m["rows"]), ("Sheet used", m["sheet"] or "-")]
    rows += [(f"Column: {k}", v) for k, v in m["col_log"].items()]
    rows += [("Created dates read as", m["created_how"]), ("Closed dates read as", m["closed_how"]),
             ("As-of date", f"{m['as_of']:%d-%b-%Y %H:%M} (rule: {m['as_of_rule']})"),
             ("Clock", "Calendar days (24 h = 1 day)" if m["mode"] == "calendar" else
              f"Business days only; weekend = {', '.join(m['weekend'])}; holidays = {', '.join(m['holidays']) or 'none'}")]
    rows += [(f"SLA target – {p}", f"{d:g} days") for p, d in m["targets"].items()]
    rows += [("Default priority", m["default_p"]), ("Closed statuses", ", ".join(m["closed_statuses"])),
             ("Excluded statuses", ", ".join(m["excluded_statuses"])),
             ("Met rule", "Closed ticket with TAT ≤ SLA target"), ("Breached rule", "Closed ticket with TAT > SLA target"),
             ("Overdue rule", "Open ticket whose age at the as-of date > SLA target"),
             ("At-risk rule", f"Open ticket that has used ≥ {m['at_risk']:.0%} of its SLA target"),
             ("Compliance", "Met ÷ (Met + Breached); open, excluded and data-issue tickets are not in the ratio"),
             ("RAG", f"GREEN ≥ {m['target']:.0%}; AMBER ≥ {m['target'] - m['margin']:.0%}; RED below")]
    rows += [(f"Priority source – {k}", f"{v:,} tickets") for k, v in m["pri_rule_count"].items()]
    header(s, 4, ["Rule / setting", "Value"])
    for i, (k, v) in enumerate(rows):
        a = s.cell(row=5 + i, column=1, value=k); a.font = F(True); a.border = BORDER
        b = s.cell(row=5 + i, column=2, value=clean(v)); b.font = F(); b.border = BORDER
        b.alignment = Alignment(wrap_text=True)
    s.column_dimensions["A"].width = 32; s.column_dimensions["B"].width = 100
    for sh in wb.worksheets:
        sh.page_setup.orientation = "landscape"; sh.page_setup.fitToWidth = 1; sh.page_setup.fitToHeight = 0
        sh.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(path)


def _widths(ws, cols):
    w = {"Ticket ID": 14, "Created": 17, "Closed": 17, "Due Date": 17, "Priority": 10, "Category": 24, "Location": 28,
         "Assignee": 22, "Status": 12, "SLA Status": 16, "Reason": 34, "Priority Rule": 20, "Priority (source)": 14,
         "Source Row": 10}
    for j, c in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w.get(c, 12)
    ws.freeze_panes = "B5"


def main():
    ap = argparse.ArgumentParser(description="SLA Report Bot: turns any ticket file into a rule-based SLA report.")
    ap.add_argument("input", nargs="?", help="CSV or Excel file with tickets")
    ap.add_argument("--rules", help="rules JSON (default: sla_rules.json next to this script)")
    ap.add_argument("--out", help="output .xlsx path")
    ap.add_argument("--as-of", help="override the as-of date, e.g. 2026-11-25")
    a = ap.parse_args()
    if not a.input:
        a.input = input("Path to your ticket file (CSV or Excel): ").strip().strip('"')
    if not os.path.exists(a.input):
        sys.exit(f"File not found: {a.input}")
    rules, rp = load_rules(a.rules)
    print(f"SLA Report Bot v{VERSION}\n  input : {a.input}\n  rules : {rp}")
    out, overall, rag, findings = run(a.input, rules, a.as_of, a.out)
    c = overall['Compliance %']
    print(f"\n  SLA compliance: {'no closed tickets yet' if pd.isna(c) else f'{c:.1%}'}  ->  {rag}")
    for f_ in findings:
        print("  - " + f_)
    print(f"\n  Report saved: {out}")


if __name__ == "__main__":
    main()
