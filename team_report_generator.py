#!/usr/bin/env python3
"""
Team Daily Report Generator v2
- Multi-lane Gantt timeline (no overlapping)
- Date range filters: From/To, Yesterday, Monthly
- All dates from Google Sheet embedded for client-side filtering
"""

import pandas as pd
import json
import os
import sys
import argparse
from datetime import datetime
from io import BytesIO

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

SHEET_ID = "1kv8kHLdoMuvfoewIpyAx9wSmIbgJdhWtMdpRqSsQ_F0"

CAT_COLORS = {
    "Meetings":             "#5c6bc0",
    "Break / Leave":        "#b0bec5",
    "Reporting & Review":   "#00897b",
    "Communication":        "#ef6c00",
    "Digital & Tech":       "#0288d1",
    "Finance & Accounts":   "#388e3c",
    "Operations":           "#f57c00",
    "General Work":         "#8e24aa",
}

EMPLOYEE_COLORS = [
    "#5c6bc0","#00897b","#ef6c00","#0288d1",
    "#388e3c","#8e24aa","#f57c00","#d32f2f",
    "#0097a7","#7b1fa2","#c62828","#2e7d32",
]

def fetch_data(local_file=None):
    if local_file:
        return pd.read_excel(local_file, sheet_name=None)
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    resp.raise_for_status()
    return pd.read_excel(BytesIO(resp.content), sheet_name=None)


def categorize(title):
    t = str(title).lower()
    if any(k in t for k in ["meeting", "hod", "daily meet", "morning meet", "sk bz", "sk and head", "all hod"]):
        return "Meetings"
    if any(k in t for k in ["lunch", "break", "leave", "half day"]):
        return "Break / Leave"
    if any(k in t for k in ["tally","bank","account","financial","salary","suspense","bill","invoice","fee","reco"]):
        return "Finance & Accounts"
    if any(k in t for k in ["report","prepare","audit","review","update","sheet","ppt","checklist","entry","fill","verify","register","mark"]):
        return "Reporting & Review"
    if any(k in t for k in ["call","follow","coordinate","email","whatsapp","revert","forward","send","share","ping","msg"]):
        return "Communication"
    if any(k in t for k in ["ai","seo","digital","social","website","content","agent","blog","podcast","caption","post","lead post"]):
        return "Digital & Tech"
    if any(k in t for k in ["renovation","construction","building","hoarding","camera"]):
        return "Operations"
    return "General Work"


def assign_lanes(tasks):
    """Greedy lane assignment so overlapping tasks don't stack on same row."""
    sorted_tasks = sorted(tasks, key=lambda t: t["start_mins"])
    lane_ends = []  # track end_mins of last task in each lane
    for task in sorted_tasks:
        placed = False
        for i, end in enumerate(lane_ends):
            if end <= task["start_mins"]:
                task["lane"] = i
                lane_ends[i] = task["end_mins"]
                placed = True
                break
        if not placed:
            task["lane"] = len(lane_ends)
            lane_ends.append(task["end_mins"])
    return len(lane_ends)


def process_sheet(df, name):
    """Returns {date_str: [task_dict, ...]} for one employee."""
    df = df.copy()
    df.columns = ["Date", "Event Title", "Start Time", "End Time", "Description", "Fetched At"]
    df = df.iloc[1:].reset_index(drop=True)
    df = df[df["Event Title"].notna()].copy()
    if df.empty:
        return {}

    df["Start Time"] = pd.to_datetime(df["Start Time"], errors="coerce")
    df["End Time"]   = pd.to_datetime(df["End Time"],   errors="coerce")
    df["Date"]       = pd.to_datetime(df["Date"],        errors="coerce")
    df = df.dropna(subset=["Start Time", "End Time", "Date"])
    df["Duration_mins"] = (df["End Time"] - df["Start Time"]).dt.total_seconds() / 60
    # Filter out all-day/placeholder events (>= 480 mins = 8 hours)
    df = df[(df["Duration_mins"] > 0) & (df["Duration_mins"] < 480)].copy()
    if df.empty:
        return {}

    df["Category"] = df["Event Title"].apply(categorize)

    DAY_START = 9 * 60   # 9:00 AM
    DAY_END   = 19 * 60  # 7:00 PM
    SPAN      = DAY_END - DAY_START

    by_date = {}
    for date_val, grp in df.groupby(df["Date"].dt.date):
        date_str = str(date_val)
        tasks = []
        for _, row in grp.iterrows():
            s_mins = row["Start Time"].hour * 60 + row["Start Time"].minute
            e_mins = row["End Time"].hour   * 60 + row["End Time"].minute
            left  = max(0.0, min(100.0, (s_mins - DAY_START) / SPAN * 100))
            width = max(0.3, min(100.0 - left, (e_mins - s_mins) / SPAN * 100))
            tasks.append({
                "title":     str(row["Event Title"]).strip(),
                "start":     row["Start Time"].strftime("%H:%M"),
                "end":       row["End Time"].strftime("%H:%M"),
                "duration":  int(row["Duration_mins"]),
                "category":  row["Category"],
                "left_pct":  round(left, 2),
                "width_pct": round(width, 2),
                "start_mins": s_mins,
                "end_mins":   e_mins,
                "lane":       0,
            })
        num_lanes = assign_lanes(tasks)
        # Summary stats for this date
        work = [t for t in tasks if t["category"] != "Break / Leave"]
        cats = {}
        for t in tasks:
            cats[t["category"]] = cats.get(t["category"], 0) + t["duration"]
        by_date[date_str] = {
            "tasks": tasks,
            "num_lanes": num_lanes,
            "stats": {
                "total_tasks": len(tasks),
                "work_tasks":  len(work),
                "work_hours":  round(sum(t["duration"] for t in work) / 60, 1),
                "start_time":  min(t["start"] for t in tasks) if tasks else "N/A",
                "end_time":    max(t["end"]   for t in tasks) if tasks else "N/A",
                "categories":  cats,
            }
        }
    return by_date


def build_html(all_employees, all_dates):
    emp_json   = json.dumps(all_employees, ensure_ascii=False)
    dates_json = json.dumps(sorted(all_dates), ensure_ascii=False)
    cat_colors_json = json.dumps(CAT_COLORS, ensure_ascii=False)
    emp_colors_json = json.dumps(EMPLOYEE_COLORS, ensure_ascii=False)

    today      = datetime.now().strftime("%Y-%m-%d")
    latest_date = sorted(all_dates)[-1] if all_dates else today

    global_legend = "".join(
        f'<span class="gl-item"><span class="gl-dot" style="background:{c}"></span>{k}</span>'
        for k, c in CAT_COLORS.items()
    )

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>GCS Team Daily Report</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Inter',sans-serif;background:#eef2ff;color:#1e293b;min-height:100vh}}

/* HEADER */
.site-header{{background:linear-gradient(135deg,#4f46e5 0%,#7c3aed 55%,#db2777 100%);color:#fff;padding:28px 40px 22px}}
.header-top{{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:14px}}
.brand{{display:flex;align-items:center;gap:12px}}
.brand-icon{{width:46px;height:46px;background:rgba(255,255,255,.2);border-radius:13px;display:flex;align-items:center;justify-content:center;font-size:22px}}
.brand h1{{font-size:21px;font-weight:700}}
.brand p{{font-size:12px;opacity:.8;margin-top:2px}}
.stat-row{{display:flex;gap:16px;margin-top:20px;flex-wrap:wrap}}
.stat-box{{background:rgba(255,255,255,.18);border-radius:13px;padding:13px 20px;min-width:120px;backdrop-filter:blur(6px)}}
.stat-box .sval{{font-size:26px;font-weight:700}}
.stat-box .slbl{{font-size:11px;opacity:.85;margin-top:2px}}

/* DATE FILTER BAR */
.filter-bar{{background:#fff;border-bottom:1px solid #e2e8f0;padding:14px 40px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;position:sticky;top:0;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.06)}}
.filter-bar label{{font-size:12px;font-weight:600;color:#64748b;white-space:nowrap}}
.date-input{{padding:7px 11px;border-radius:9px;border:1.5px solid #e2e8f0;font-size:12px;font-family:inherit;outline:none;color:#1e293b;transition:border .2s}}
.date-input:focus{{border-color:#4f46e5}}
.quick-btn{{padding:7px 14px;border-radius:20px;border:1.5px solid #e2e8f0;background:#f8fafc;font-size:12px;font-weight:600;cursor:pointer;color:#475569;transition:all .2s;white-space:nowrap}}
.quick-btn:hover,.quick-btn.active{{background:#4f46e5;color:#fff;border-color:#4f46e5}}
.apply-btn{{padding:7px 18px;border-radius:20px;border:none;background:#4f46e5;color:#fff;font-size:12px;font-weight:600;cursor:pointer;transition:background .2s}}
.apply-btn:hover{{background:#4338ca}}
.search-wrap{{margin-left:auto;display:flex;align-items:center;gap:6px}}
.search-wrap input{{padding:7px 13px;border-radius:20px;border:1.5px solid #e2e8f0;font-size:12px;outline:none;width:180px;transition:border .2s}}
.search-wrap input:focus{{border-color:#4f46e5}}

/* LEGEND BAR */
.legend-bar{{background:#fff;border-bottom:1px solid #e2e8f0;padding:10px 40px;display:flex;gap:14px;flex-wrap:wrap;align-items:center}}
.gl-item{{display:flex;align-items:center;gap:5px;font-size:11px;color:#64748b}}
.gl-dot{{width:9px;height:9px;border-radius:50%;flex-shrink:0}}

/* GRID */
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(460px,1fr));gap:22px;padding:28px 40px;max-width:1600px;margin:0 auto}}

/* CARD */
.emp-card{{background:#fff;border-radius:18px;box-shadow:0 4px 20px rgba(79,70,229,.07),0 1px 4px rgba(0,0,0,.04);overflow:hidden;transition:transform .2s,box-shadow .2s}}
.emp-card:hover{{transform:translateY(-3px);box-shadow:0 8px 30px rgba(79,70,229,.13)}}
.card-header{{display:flex;align-items:center;gap:13px;padding:16px 20px}}
.avatar{{width:46px;height:46px;border-radius:13px;color:#fff;font-weight:700;font-size:15px;display:flex;align-items:center;justify-content:center;flex-shrink:0}}
.emp-info h3{{font-size:15px;font-weight:700}}
.emp-meta{{display:flex;gap:7px;align-items:center;margin-top:4px;flex-wrap:wrap}}
.badge{{padding:3px 9px;border-radius:20px;font-size:11px;font-weight:600}}
.badge.grey{{background:#f1f5f9;color:#94a3b8}}
.stars{{font-size:12px;color:#f59e0b;letter-spacing:1px}}
.emp-time{{font-size:11px;color:#94a3b8;margin-top:2px}}

/* GANTT TIMELINE */
.timeline-section{{padding:10px 20px 4px}}
.tl-header{{display:flex;justify-content:space-between;margin-bottom:4px}}
.tl-title{{font-size:11px;font-weight:600;color:#64748b}}
.tl-hours-row{{position:relative;height:18px;margin-bottom:2px}}
.tl-hour-mark{{position:absolute;font-size:9px;color:#94a3b8;transform:translateX(-50%)}}
.gantt-wrap{{position:relative;background:#f8fafc;border-radius:8px;overflow:hidden}}
.gantt-lane{{position:relative;height:26px;border-bottom:1px solid #f1f5f9}}
.gantt-lane:last-child{{border-bottom:none}}
.gantt-seg{{position:absolute;top:3px;height:20px;border-radius:5px;cursor:default;display:flex;align-items:center;padding:0 5px;transition:opacity .15s;min-width:2px}}
.gantt-seg:hover{{opacity:.82;z-index:10}}
.gantt-seg .seg-lbl{{font-size:9px;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;pointer-events:none}}

/* CAT TAGS */
.cat-legend{{padding:7px 20px 10px;display:flex;gap:5px;flex-wrap:wrap}}
.cat-tag{{padding:3px 9px;border-radius:20px;font-size:11px;font-weight:500}}

/* TABS */
.card-tabs{{display:flex;border-top:1px solid #f1f5f9;border-bottom:1px solid #f1f5f9}}
.tab-btn{{flex:1;padding:9px;font-size:12px;font-weight:600;background:none;border:none;cursor:pointer;color:#94a3b8;transition:all .15s}}
.tab-btn.active{{color:#4f46e5;border-bottom:2px solid #4f46e5}}
.tab-btn:hover{{background:#f8fafc}}
.tab-pane{{padding:12px 20px 14px;max-height:240px;overflow-y:auto;display:none}}
.tab-pane.active{{display:block}}

/* TASK LIST */
.task-row{{display:flex;align-items:flex-start;gap:8px;padding:6px 0;border-bottom:1px solid #f8fafc}}
.task-row:last-child{{border-bottom:none}}
.task-dot{{width:7px;height:7px;border-radius:50%;margin-top:4px;flex-shrink:0}}
.task-content{{flex:1;min-width:0}}
.task-time{{font-size:10px;color:#94a3b8}}
.task-title{{font-size:12px;font-weight:500;margin-top:1px;display:block}}
.task-dur{{font-size:11px;color:#64748b;font-weight:600;white-space:nowrap;flex-shrink:0}}

/* INSIGHTS */
.insights-box{{display:flex;flex-direction:column;gap:7px}}
.insight-item{{display:flex;gap:10px;align-items:flex-start;background:#f8fafc;border-radius:9px;padding:9px 11px;font-size:12px}}

/* NO DATA */
.no-data-msg{{padding:22px;text-align:center;color:#94a3b8;font-size:13px}}

/* EMPTY STATE */
.empty-state{{grid-column:1/-1;text-align:center;padding:60px;color:#94a3b8}}
.empty-state div{{font-size:48px;margin-bottom:12px}}
.empty-state p{{font-size:15px}}

footer{{text-align:center;padding:24px;color:#94a3b8;font-size:11px;border-top:1px solid #e2e8f0;background:#fff;margin-top:8px}}

::-webkit-scrollbar{{width:4px;height:4px}}
::-webkit-scrollbar-track{{background:#f1f5f9}}
::-webkit-scrollbar-thumb{{background:#c7d2fe;border-radius:10px}}

@media(max-width:600px){{
  .site-header,.filter-bar,.legend-bar{{padding-left:16px;padding-right:16px}}
  .grid{{padding:14px 16px;grid-template-columns:1fr}}
}}
</style>
</head>
<body>

<header class="site-header">
  <div class="header-top">
    <div class="brand">
      <div class="brand-icon">📊</div>
      <div>
        <h1>GCS Team Daily Report</h1>
        <p>Calendar-based productivity dashboard</p>
      </div>
    </div>
  </div>
  <div class="stat-row">
    <div class="stat-box"><div class="sval" id="hd-members">—</div><div class="slbl">Total Members</div></div>
    <div class="stat-box"><div class="sval" id="hd-active">—</div><div class="slbl">Active Today</div></div>
    <div class="stat-box"><div class="sval" id="hd-avghrs">—</div><div class="slbl">Avg Work Hours</div></div>
    <div class="stat-box"><div class="sval" id="hd-date">—</div><div class="slbl">Viewing Period</div></div>
  </div>
</header>

<div class="filter-bar">
  <label>📅 From</label>
  <input type="date" class="date-input" id="fromDate"/>
  <label>To</label>
  <input type="date" class="date-input" id="toDate"/>
  <button class="quick-btn" id="btn-yesterday" onclick="setPreset('yesterday',this)">Yesterday</button>
  <button class="quick-btn active" id="btn-today" onclick="setPreset('today',this)">Today</button>
  <button class="quick-btn" id="btn-week" onclick="setPreset('week',this)">This Week</button>
  <button class="quick-btn" id="btn-month" onclick="setPreset('month',this)">Monthly</button>
  <button class="apply-btn" onclick="applyFilter()">Apply ▶</button>
  <div class="search-wrap">
    🔍<input type="text" id="searchInput" placeholder="Search employee…" oninput="searchCards(this.value)"/>
  </div>
</div>

<div class="legend-bar">
  <span style="font-size:11px;font-weight:600;color:#64748b">Categories:</span>
  {global_legend}
</div>

<div class="grid" id="grid"></div>

<footer id="footer">
  GCS Group · Report generated on {datetime.now().strftime("%d %b %Y, %I:%M %p")} IST · Auto-runs daily at 1:00 PM IST
</footer>

<script>
const ALL_DATA    = {emp_json};
const ALL_DATES   = {dates_json};
const CAT_COLORS  = {cat_colors_json};
const EMP_COLORS  = {emp_colors_json};

// ── DATE HELPERS ─────────────────────────────────────────────────────────────
function today()     {{ return new Date().toISOString().slice(0,10); }}
function yesterday() {{
  const d = new Date(); d.setDate(d.getDate()-1);
  return d.toISOString().slice(0,10);
}}
function weekStart() {{
  const d = new Date();
  d.setDate(d.getDate() - d.getDay() + (d.getDay()===0?-6:1));
  return d.toISOString().slice(0,10);
}}
function monthStart() {{
  const d = new Date();
  return `${{d.getFullYear()}}-${{String(d.getMonth()+1).padStart(2,'0')}}-01`;
}}

function setPreset(type, btn) {{
  document.querySelectorAll('.quick-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const t = today(), y = yesterday();
  if (type==='yesterday') {{ document.getElementById('fromDate').value=y; document.getElementById('toDate').value=y; }}
  if (type==='today')     {{ document.getElementById('fromDate').value=t; document.getElementById('toDate').value=t; }}
  if (type==='week')      {{ document.getElementById('fromDate').value=weekStart(); document.getElementById('toDate').value=t; }}
  if (type==='month')     {{ document.getElementById('fromDate').value=monthStart(); document.getElementById('toDate').value=t; }}
  applyFilter();
}}

// ── FILTER LOGIC ─────────────────────────────────────────────────────────────
function getDatesInRange(from, to) {{
  return ALL_DATES.filter(d => d >= from && d <= to);
}}

function mergeTasksForRange(empDates, from, to) {{
  const merged = [];
  const inRange = getDatesInRange(from, to);
  for (const d of inRange) {{
    if (empDates[d]) merged.push(...empDates[d].tasks);
  }}
  return merged;
}}

function calcStats(tasks) {{
  const work = tasks.filter(t => t.category !== 'Break / Leave');
  const cats = {{}};
  tasks.forEach(t => cats[t.category] = (cats[t.category]||0) + t.duration);
  return {{
    total_tasks: tasks.length,
    work_tasks:  work.length,
    work_hours:  Math.round(work.reduce((a,t)=>a+t.duration,0)/60*10)/10,
    start_time:  tasks.length ? tasks.map(t=>t.start).sort()[0]   : 'N/A',
    end_time:    tasks.length ? tasks.map(t=>t.end).sort().pop()   : 'N/A',
    categories:  cats,
  }};
}}

// Greedy lane assignment in JS
function assignLanes(tasks) {{
  const sorted = [...tasks].sort((a,b)=>a.start_mins-b.start_mins);
  const laneEnds = [];
  const result = [];
  for (const t of sorted) {{
    let placed = false;
    for (let i=0;i<laneEnds.length;i++) {{
      if (laneEnds[i]<=t.start_mins) {{ t.lane=i; laneEnds[i]=t.end_mins; placed=true; break; }}
    }}
    if (!placed) {{ t.lane=laneEnds.length; laneEnds.push(t.end_mins); }}
    result.push(t);
  }}
  return {{ tasks: result, numLanes: laneEnds.length }};
}}

// ── RENDER ───────────────────────────────────────────────────────────────────
function generateInsights(stats) {{
  const ins = [];
  if (stats.work_hours>=7)       ins.push(['✅',`Excellent — ${{stats.work_hours}}h of productive work`]);
  else if (stats.work_hours>=5)  ins.push(['👍',`Good effort — ${{stats.work_hours}}h of work logged`]);
  else if (stats.work_hours>0)   ins.push(['⚠️',`Low activity — only ${{stats.work_hours}}h logged`]);
  else                           ins.push(['❌','No work activities recorded']);
  const workCats = Object.fromEntries(Object.entries(stats.categories).filter(([k])=>k!=='Break / Leave'));
  if (Object.keys(workCats).length) {{
    const top = Object.entries(workCats).sort((a,b)=>b[1]-a[1])[0];
    const total = Object.values(workCats).reduce((a,b)=>a+b,0);
    ins.push(['🎯',`Primary focus: ${{top[0]}} (${{Math.round(top[1]/total*100)}}% of work time)`]);
  }}
  const mtg = stats.categories['Meetings']||0;
  if (mtg>90) ins.push(['📅',`High meeting load: ${{Math.round(mtg)}} mins`]);
  if (stats.start_time!=='N/A' && parseInt(stats.start_time)<10) ins.push(['🌅','Early start — punctual and proactive']);
  return ins;
}}

function renderCard(emp, idx, from, to) {{
  const color = EMP_COLORS[idx % EMP_COLORS.length];
  const initials = emp.name.split(' ').slice(0,2).map(p=>p[0].toUpperCase()).join('');

  const empDates = emp.dates||{{}};
  const allTasks = mergeTasksForRange(empDates, from, to);

  if (!allTasks.length) {{
    return `<div class="emp-card no-data" data-name="${{emp.name}}">
      <div class="card-header" style="background:${{color}}15;border-left:4px solid ${{color}}">
        <div class="avatar" style="background:${{color}}">${{initials}}</div>
        <div class="emp-info"><h3>${{emp.name}}</h3><div class="emp-meta"><span class="badge grey">No activity</span></div></div>
      </div>
      <div class="no-data-msg">📭 No calendar events for selected period</div>
    </div>`;
  }}

  const laneResult = assignLanes(allTasks.map(t=>(Object.assign({{}},t))));
  const lanedTasks = laneResult.tasks, numLanes = laneResult.numLanes;
  const stats = calcStats(allTasks);
  const insights = generateInsights(stats);
  const stars = '★'.repeat(Math.min(5,Math.max(1,Math.round(stats.work_hours/8*5)))) + '☆'.repeat(5-Math.min(5,Math.max(1,Math.round(stats.work_hours/8*5))));

  // Hour marks 9-19
  let hrMarks = '';
  for (let h=9;h<=19;h+=2) {{
    const pos = (h*60-9*60)/(10*60)*100;
    hrMarks += `<span class="tl-hour-mark" style="left:${{pos.toFixed(1)}}%">${{String(h).padStart(2,'0')}}:00</span>`;
  }}

  // Gantt lanes
  const laneH = 26;
  const ganttH = numLanes * laneH;
  let lanes = '';
  for (let l=0;l<numLanes;l++) {{
    const segs = lanedTasks.filter(t=>t.lane===l);
    const segsHtml = segs.map(t => {{
      const clr = CAT_COLORS[t.category]||'#999';
      const dur = t.duration<60 ? t.duration+'m' : Math.round(t.duration/60*10)/10+'h';
      return `<div class="gantt-seg" style="left:${{t.left_pct}}%;width:${{t.width_pct}}%;background:${{clr}}"
               title="${{t.title}}&#10;${{t.start}} – ${{t.end}} (${{dur}})&#10;${{t.category}}">
               <span class="seg-lbl">${{t.title}}</span></div>`;
    }}).join('');
    lanes += `<div class="gantt-lane">${{segsHtml}}</div>`;
  }}

  // Category legend
  const workCats = Object.entries(stats.categories).filter(([k])=>k!=='Break / Leave').sort((a,b)=>b[1]-a[1]);
  const catTags = workCats.map(([k,v]) => {{
    const clr = CAT_COLORS[k]||'#999';
    return `<span class="cat-tag" style="background:${{clr}}22;color:${{clr}};border:1px solid ${{clr}}55">${{k}} · ${{Math.round(v/60*10)/10}}h</span>`;
  }}).join('');

  // Task list
  const taskRows = allTasks.sort((a,b)=>a.start.localeCompare(b.start)).map(t => {{
    const clr = CAT_COLORS[t.category]||'#999';
    const dur = t.duration<60 ? t.duration+'m' : Math.round(t.duration/60*10)/10+'h';
    return `<div class="task-row">
      <span class="task-dot" style="background:${{clr}}"></span>
      <div class="task-content">
        <span class="task-time">${{t.start}} – ${{t.end}}</span>
        <span class="task-title">${{t.title}}</span>
      </div>
      <span class="task-dur">${{dur}}</span>
    </div>`;
  }}).join('');

  const insightRows = insights.map(([ic,tx])=>
    `<div class="insight-item"><span>${{ic}}</span><span>${{tx}}</span></div>`
  ).join('');

  const cid = `c${{idx}}`;
  return `<div class="emp-card" data-name="${{emp.name}}" id="${{cid}}">
    <div class="card-header" style="background:${{color}}15;border-left:4px solid ${{color}}">
      <div class="avatar" style="background:${{color}}">${{initials}}</div>
      <div class="emp-info">
        <h3>${{emp.name}}</h3>
        <div class="emp-meta">
          <span class="badge" style="background:${{color}}22;color:${{color}}">${{stats.work_tasks}} tasks</span>
          <span class="badge" style="background:${{color}}22;color:${{color}}">${{stats.work_hours}}h work</span>
          <span class="stars">${{stars}}</span>
        </div>
        <div class="emp-time">🕐 ${{stats.start_time}} – ${{stats.end_time}}</div>
      </div>
    </div>

    <div class="timeline-section">
      <div class="tl-hours-row">${{hrMarks}}</div>
      <div class="gantt-wrap" style="height:${{ganttH}}px">${{lanes}}</div>
    </div>

    <div class="cat-legend">${{catTags}}</div>

    <div class="card-tabs">
      <button class="tab-btn active" onclick="switchTab(this,'${{cid}}-tasks')">📋 Tasks</button>
      <button class="tab-btn" onclick="switchTab(this,'${{cid}}-insights')">💡 Analysis</button>
    </div>
    <div class="tab-pane active" id="${{cid}}-tasks"><div class="task-list">${{taskRows}}</div></div>
    <div class="tab-pane" id="${{cid}}-insights"><div class="insights-box">${{insightRows}}</div></div>
  </div>`;
}}

function renderAll(from, to) {{
  const grid = document.getElementById('grid');
  let html = '';
  let active=0, totalWh=0;

  ALL_DATA.forEach((emp,i) => {{
    const empDates = emp.dates||{{}};
    const tasks = mergeTasksForRange(empDates, from, to);
    if (tasks.length) {{ active++; totalWh += calcStats(tasks).work_hours; }}
    html += renderCard(emp, i, from, to);
  }});

  grid.innerHTML = html || '<div class="empty-state"><div>📭</div><p>No data for selected period</p></div>';

  // Update header stats
  document.getElementById('hd-members').textContent = ALL_DATA.length;
  document.getElementById('hd-active').textContent   = active;
  document.getElementById('hd-avghrs').textContent   = active ? (Math.round(totalWh/active*10)/10)+'h' : '—';
  document.getElementById('hd-date').textContent     = from===to ? from.slice(5).replace('-','/') : from.slice(5).replace('-','/')+'–'+to.slice(5).replace('-','/');
}}

function applyFilter() {{
  const from = document.getElementById('fromDate').value;
  const to   = document.getElementById('toDate').value;
  if (!from || !to) return;
  renderAll(from, to);
}}

function switchTab(btn, paneId) {{
  const card = btn.closest('.emp-card');
  card.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  card.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById(paneId).classList.add('active');
}}

function searchCards(q) {{
  q = q.toLowerCase().trim();
  document.querySelectorAll('.emp-card').forEach(c => {{
    const name = (c.dataset.name||'').toLowerCase();
    c.style.display = (!q || name.includes(q)) ? '' : 'none';
  }});
}}

// ── INIT ─────────────────────────────────────────────────────────────────────
(function init() {{
  const latestDate = ALL_DATES.length ? ALL_DATES[ALL_DATES.length-1] : today();
  document.getElementById('fromDate').value = latestDate;
  document.getElementById('toDate').value   = latestDate;
  renderAll(latestDate, latestDate);
}})();
</script>
</body>
</html>"""
    return html


def process_all(sheets):
    """Return list of {name, dates: {date_str: {tasks, stats}}} and set of all dates."""
    all_employees = []
    all_dates = set()
    for name, df in sheets.items():
        by_date = process_sheet(df, name)
        all_dates.update(by_date.keys())
        # Convert task dicts to be JSON-safe
        clean_dates = {}
        for d, v in by_date.items():
            clean_dates[d] = {"tasks": v["tasks"], "stats": v["stats"]}
        all_employees.append({"name": name, "dates": clean_dates})
    return all_employees, sorted(all_dates)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--local",  help="Path to local Excel file", default=None)
    parser.add_argument("--output", help="Output HTML file path",    default=None)
    args = parser.parse_args()

    print(f"[{datetime.now():%H:%M:%S}] Fetching data…")
    sheets = fetch_data(local_file=args.local)

    print(f"[{datetime.now():%H:%M:%S}] Processing {len(sheets)} employees…")
    employees, all_dates = process_all(sheets)

    html = build_html(employees, all_dates)

    if args.output:
        out = args.output
    else:
        fname = f"Team_Daily_Report_{datetime.now():%Y-%m-%d}.html"
        out   = os.path.join(os.path.expanduser("~"), "Desktop", fname)

    os.makedirs(os.path.dirname(out) if os.path.dirname(out) else ".", exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[{datetime.now():%H:%M:%S}] ✅ Saved: {out}")
    return out


if __name__ == "__main__":
    main()
