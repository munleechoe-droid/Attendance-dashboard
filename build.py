#!/usr/bin/env python3
"""
Attendance Dashboard Builder
=============================
Usage: python3 build.py [path_to_xlsx]
Default file: 0_DailyAttendanceReport_Master.xlsx (place in same folder)
Output: attendance_dashboard.html
"""

import pandas as pd
import json
import os
import sys
import warnings
warnings.filterwarnings('ignore')

# ── CONFIG ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
DATA_FILE   = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SCRIPT_DIR, '0_DailyAttendanceReport_Master.xlsx')
SHEET_NAME  = 'Master_Daily Attendance'
OUTPUT_FILE = os.path.join(SCRIPT_DIR, 'attendance_dashboard.html')

MONTH_MAP = {
    '01-26':'Jan 26','02-26':'Feb 26','03-26':'Mar 26','04-26':'Apr 26',
    '05-26':'May 26','06-26':'Jun 26','07-26':'Jul 26','08-26':'Aug 26',
    '09-26':'Sep 26','10-26':'Oct 26','11-26':'Nov 26','12-26':'Dec 26',
    '01-25':'Jan 25','02-25':'Feb 25','03-25':'Mar 25',
}

print(f"Reading: {DATA_FILE}")
df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
df['Date'] = pd.to_datetime(df['Date'])
df['Employee Id'] = df['Employee Id'].astype(str).str.strip()
df['Direct Manager Employee Id'] = df['Direct Manager Employee Id'].astype(str).str.strip()

def clean(s):
    return str(s).replace('\u2800','').replace('\u200e','').strip() if pd.notna(s) else ''

# Date dimensions
df['Year']        = df['Date'].dt.year
df['Quarter']     = df['Date'].dt.quarter
df['WeekNum']     = df['Date'].dt.isocalendar().week.astype(int)
df['Month_Label'] = df['Month'].map(MONTH_MAP).fillna(df['Month'])

def sc(status, is_late, wd):
    s = str(status) if pd.notna(status) else ''
    if 'Present' in s: return 'PL' if is_late == 'Yes' else 'P'
    if 'Single Punch' in s: return 'SP'
    if 'Leave' in s: return 'L'
    if wd == 1: return 'A'
    return 'N'

def shift_code(s):
    if pd.isna(s): return ''
    s = str(s)
    if 'S1' in s: return 'S1'
    if 'S2' in s: return 'S2'
    if 'RPH' in s: return 'RPH'
    if 'RMD' in s: return 'RMD'
    return 'S'

print("Building summary...")
work = df[df['Working Days'] == 1].copy()
summary = work.groupby(['Current Department','Branch','Current Designation','Month_Label','Year','Quarter','WeekNum']).agg(
    emp=('Employee Id','nunique'),
    wd=('Working Days','sum'),
    present=('Status', lambda x: x.str.contains('Present', na=False).sum()),
    late=('Is Late ', lambda x: (x=='Yes').sum()),
    sp=('Single Punch','sum'),
    leave=('On Leave','sum'),
).reset_index()
summary['att'] = (summary['present'] / summary['wd'] * 100).round(1)
summary_rows = summary.to_dict('records')

print("Building detail...")
detail_rows = work.groupby(['Employee Id','Month_Label','Year','Quarter','WeekNum']).agg(
    name=('Employee Name', lambda x: clean(x.iloc[0])),
    dept=('Current Department','first'),
    branch=('Branch','first'),
    desig=('Current Designation','first'),
    mgr_name=('Direct Manager Name', lambda x: clean(x.iloc[0])),
    mgr_id=('Direct Manager Employee Id','first'),
    wd=('Working Days','sum'),
    present=('Status', lambda x: x.str.contains('Present', na=False).sum()),
    late=('Is Late ', lambda x: (x=='Yes').sum()),
    sp=('Single Punch','sum'),
    leave=('On Leave','sum'),
).reset_index()
detail_rows['att'] = (detail_rows['present'] / detail_rows['wd'] * 100).round(1)
detail_list = detail_rows.to_dict('records')

print("Building heatmap...")
detail_heat = {}
for _, row in df.iterrows():
    branch   = str(row['Branch']).strip()
    month    = str(row['Month_Label']) if pd.notna(row.get('Month_Label')) else ''
    if not month or month == 'nan': continue
    emp_id   = str(row['Employee Id']).strip()
    emp_name = clean(row['Employee Name'])
    dept     = str(row['Current Department']).strip()
    desig    = str(row['Current Designation']).strip()
    mgr_id   = str(row['Direct Manager Employee Id']).strip()
    day      = row['Date'].strftime('%d') if pd.notna(row['Date']) else ''
    if not day: continue
    status_c = sc(row['Status'], row['Is Late '], row['Working Days'])
    sh       = shift_code(row['Shift'])
    if branch not in detail_heat: detail_heat[branch] = {}
    if month not in detail_heat[branch]: detail_heat[branch][month] = {}
    if emp_id not in detail_heat[branch][month]:
        detail_heat[branch][month][emp_id] = {'name':emp_name,'dept':dept,'desig':desig,'mgr_id':mgr_id,'days':{}}
    detail_heat[branch][month][emp_id]['days'][day] = {'sc': status_c, 'sh': sh}

print("Building org...")
emp_profile = df.groupby('Employee Id').agg(
    name=('Employee Name', lambda x: clean(x.iloc[0])),
    dept=('Current Department','first'),
    desig=('Current Designation','first'),
    branch=('Branch','first'),
    mgr_id=('Direct Manager Employee Id','first'),
    mgr_name=('Direct Manager Name', lambda x: clean(x.iloc[0])),
).reset_index()
work_stats = work.groupby('Employee Id').agg(
    wd=('Working Days','sum'),
    present=('Status', lambda x: x.str.contains('Present', na=False).sum()),
    late=('Is Late ', lambda x: (x=='Yes').sum()),
).reset_index()
emp_profile = emp_profile.merge(work_stats, on='Employee Id', how='left')
emp_profile['att']        = (emp_profile['present'] / emp_profile['wd'] * 100).round(1).fillna(0)
emp_profile['Employee Id'] = emp_profile['Employee Id'].astype(str)
emp_profile['mgr_id']      = emp_profile['mgr_id'].astype(str)
org_list = emp_profile.to_dict('records')

branches     = sorted(df['Branch'].dropna().unique().tolist())
departments  = sorted(df['Current Department'].dropna().unique().tolist())
designations = sorted(df['Current Designation'].dropna().unique().tolist())
month_order  = [m for m in ['Jan 25','Feb 25','Mar 25','Apr 25','May 25','Jun 25','Jul 25','Aug 25','Sep 25','Oct 25','Nov 25','Dec 25',
                             'Jan 26','Feb 26','Mar 26','Apr 26','May 26','Jun 26','Jul 26','Aug 26','Sep 26','Oct 26','Nov 26','Dec 26']
                if m in df['Month_Label'].unique().tolist()]
years        = sorted(df['Year'].unique().tolist())
quarters     = sorted(df['Quarter'].unique().tolist())
weeks        = sorted(df['WeekNum'].unique().tolist())

# Week labels: Week N (dd Mon - dd Mon)
week_labels = {}
for w in weeks:
    wdf = df[df['WeekNum']==w]['Date']
    label = f"W{w} ({wdf.min().strftime('%d %b')} – {wdf.max().strftime('%d %b')})"
    week_labels[str(w)] = label

# Month day counts for heatmap
import calendar
month_cfg = {}
for ml in month_order:
    parts = ml.split(' ')
    mon_num = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}[parts[0]]
    yr = 2000 + int(parts[1])
    days_in = calendar.monthrange(yr, mon_num)[1]
    # Jan 1 2026 = Thursday = 3 (0=Mon..6=Sun → Sun=6 → need 0=Sun)
    first_dow = (calendar.weekday(yr, mon_num, 1) + 1) % 7  # 0=Sun
    month_cfg[ml] = {'days': days_in, 'start': first_dow}

out = {
    'summary': summary_rows, 'detail': detail_list,
    'heat': detail_heat, 'org': org_list,
    'branches': branches, 'departments': departments, 'designations': designations,
    'months': month_order, 'years': years, 'quarters': quarters,
    'weeks': weeks, 'week_labels': week_labels, 'month_cfg': month_cfg,
    'generated': pd.Timestamp.now().strftime('%d %b %Y %H:%M'),
}

print("Serialising data...")
data_js = json.dumps(out)
print(f"Data size: {len(data_js)//1024} KB")

# ── HTML ───────────────────────────────────────────────────────────────────────
print("Building HTML...")

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Attendance Dashboard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;background:#f5f7fa;color:#1a2332;min-height:100vh}
.header{background:#1a2332;color:#fff;padding:13px 24px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.header h1{font-size:15px;font-weight:600}
.header-right{font-size:11px;opacity:.45}
.tabs{background:#243447;display:flex;padding:0 20px;overflow-x:auto;gap:0}
.tab{padding:10px 18px;font-size:12px;color:#8ab;cursor:pointer;border-bottom:3px solid transparent;white-space:nowrap;flex-shrink:0}
.tab:hover{color:#fff}.tab.active{color:#fff;border-bottom-color:#2196F3}
/* Filter bar */
.filter-bar{background:#fff;padding:8px 16px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;border-bottom:1px solid #e9ecef}
.flt{display:flex;align-items:center;gap:5px}
.flt label{font-size:11px;color:#868e96;font-weight:500;white-space:nowrap}
select.single{font-size:11px;padding:4px 7px;border-radius:6px;border:1px solid #dee2e6;background:#fff;color:#1a2332;cursor:pointer;outline:none}
select.single:focus{border-color:#2196F3}
/* Multi-select */
.ms-wrap{position:relative}
.ms-btn{font-size:11px;padding:4px 8px;border-radius:6px;border:1px solid #dee2e6;background:#fff;color:#1a2332;cursor:pointer;display:flex;align-items:center;gap:5px;min-width:120px;max-width:185px;user-select:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.ms-btn:hover,.ms-btn.active{border-color:#2196F3}
.ms-arrow{margin-left:auto;font-size:9px;opacity:.5;flex-shrink:0}
.ms-badge{background:#2196F3;color:#fff;border-radius:10px;font-size:9px;padding:1px 5px;font-weight:700;flex-shrink:0}
.ms-dropdown{display:none;position:absolute;top:calc(100% + 3px);left:0;background:#fff;border:1px solid #dee2e6;border-radius:8px;z-index:200;min-width:210px;max-width:280px;box-shadow:0 4px 16px rgba(0,0,0,.12);overflow:hidden}
.ms-dropdown.open{display:block}
.ms-search{padding:7px 9px;border-bottom:1px solid #f1f3f5}
.ms-search input{width:100%;font-size:11px;padding:3px 7px;border:1px solid #dee2e6;border-radius:4px;outline:none}
.ms-search input:focus{border-color:#2196F3}
.ms-list{max-height:200px;overflow-y:auto}
.ms-item{display:flex;align-items:center;gap:7px;padding:5px 10px;cursor:pointer;font-size:11px;color:#343a40}
.ms-item:hover{background:#f8fafc}
.ms-item.select-all{border-bottom:1px solid #f1f3f5;font-weight:500;color:#495057}
.ms-item input[type=checkbox]{cursor:pointer;accent-color:#2196F3}
/* Date filter tabs */
.date-tabs{display:flex;gap:4px;align-items:center}
.date-tab{padding:3px 10px;border-radius:4px;font-size:11px;cursor:pointer;border:1px solid #dee2e6;background:#fff;color:#868e96}
.date-tab.active{background:#2196F3;color:#fff;border-color:#2196F3}
.date-subfilter{display:none;align-items:center;gap:5px}
.date-subfilter.visible{display:flex}
/* Buttons */
.btn{font-size:11px;padding:4px 10px;border-radius:6px;border:1px solid #dee2e6;background:#fff;color:#868e96;cursor:pointer;white-space:nowrap}
.btn:hover{background:#f8fafc;color:#1a2332}
.btn-export{background:#1a2332;color:#fff;border-color:#1a2332;display:flex;align-items:center;gap:4px}
.btn-export:hover{background:#243447}
.export-menu{position:relative}
.export-dd{display:none;position:absolute;top:calc(100%+3px);right:0;background:#fff;border:1px solid #dee2e6;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.12);z-index:200;min-width:140px;overflow:hidden}
.export-dd.open{display:block}
.exp-item{padding:8px 13px;font-size:11px;color:#343a40;cursor:pointer}
.exp-item:hover{background:#f8fafc}
.exp-item+.exp-item{border-top:1px solid #f1f3f5}
/* Toggle */
.toggle-wrap{display:flex;align-items:center;gap:5px;font-size:11px;color:#868e96}
.toggle-track{position:relative;width:32px;height:18px;background:#dee2e6;border-radius:9px;cursor:pointer;transition:background .2s;flex-shrink:0}
.toggle-track.on{background:#2196F3}
.toggle-thumb{position:absolute;top:2px;left:2px;width:14px;height:14px;background:#fff;border-radius:50%;transition:transform .2s;box-shadow:0 1px 3px rgba(0,0,0,.2)}
.toggle-track.on .toggle-thumb{transform:translateX(14px)}
/* Legend */
.legend{background:#f8fafc;padding:6px 20px;display:flex;gap:12px;flex-wrap:wrap;border-bottom:1px solid #e9ecef;align-items:center}
.leg{display:flex;align-items:center;gap:3px;font-size:11px;color:#495057}
.leg-dot{width:11px;height:11px;border-radius:2px;flex-shrink:0}
/* Content */
.content{padding:14px 20px}
.pane{display:none}.pane.active{display:block}
/* KPIs */
.kpi-row{display:flex;gap:9px;margin-bottom:13px;flex-wrap:wrap}
.kpi{background:#fff;border-radius:8px;padding:10px 13px;flex:1;min-width:95px;border:.5px solid #e9ecef}
.kpi-val{font-size:20px;font-weight:700}
.kpi-lbl{font-size:10px;color:#868e96;margin-top:2px}
/* Summary table */
.tbl-wrap{background:#fff;border-radius:8px;border:.5px solid #e9ecef;overflow:hidden}
.tbl-scroll{overflow-x:auto}
table.stbl{width:100%;border-collapse:collapse;font-size:12px}
table.stbl th{padding:8px 11px;text-align:left;background:#f8fafc;font-weight:600;font-size:11px;color:#495057;border-bottom:1px solid #dee2e6;white-space:nowrap;cursor:pointer;user-select:none}
table.stbl th:hover{background:#f1f3f5}
table.stbl td{padding:7px 11px;border-bottom:.5px solid #f1f3f5;color:#343a40;white-space:nowrap}
table.stbl tr:last-child td{border-bottom:none}
table.stbl tr:hover td{background:#f8fafc}
.ap{display:inline-block;padding:2px 7px;border-radius:20px;font-size:10px;font-weight:600}
.hi{background:#e8f5e9;color:#2e7d32}.mid{background:#fff8e1;color:#f57f17}.lo{background:#ffebee;color:#c62828}
.sa{font-size:9px;margin-left:2px;opacity:.5}
/* Heatmap */
.heat-wrap{background:#fff;border-radius:8px;border:.5px solid #e9ecef;overflow:hidden}
.heat-scroll{overflow-x:auto}
table.htbl{border-collapse:collapse;font-size:10px;white-space:nowrap}
.he{min-width:150px;width:150px;padding:6px 9px;font-size:11px;font-weight:600;color:#495057;background:#f8fafc;position:sticky;left:0;z-index:3;border-right:2px solid #dee2e6;border-bottom:1px solid #dee2e6}
.hd{min-width:120px;width:120px;padding:6px 7px;font-size:11px;font-weight:600;color:#495057;background:#f8fafc;position:sticky;left:150px;z-index:3;border-right:2px solid #dee2e6;border-bottom:1px solid #dee2e6}
.hdc{min-width:32px;width:32px;text-align:center;padding:3px 1px;background:#f8fafc;border-bottom:1px solid #dee2e6}
.hdn{font-size:10px;font-weight:600;color:#343a40}
.hdw{font-size:8px;color:#adb5bd}
.hdc.wk{background:#f1f3f5}.hdc.wk .hdn{color:#ced4da}
.tde{padding:3px 9px;font-size:10px;color:#343a40;background:#fff;position:sticky;left:0;z-index:1;border-right:1.5px solid #dee2e6;border-bottom:.5px solid #f1f3f5;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tdd{padding:3px 7px;font-size:9px;color:#495057;background:#fff;position:sticky;left:150px;z-index:1;border-right:2px solid #dee2e6;border-bottom:.5px solid #f1f3f5;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.tds{font-size:8px;color:#adb5bd;display:block;overflow:hidden;text-overflow:ellipsis}
.tdc{width:32px;min-width:32px;height:25px;text-align:center;vertical-align:middle;border-bottom:.5px solid #f1f3f5;border-right:.5px solid rgba(0,0,0,.03)}
tr:hover .tde,tr:hover .tdd{background:#f8fafc}
tr:hover .tdc{filter:brightness(.91)}
.ci{display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:1px}
.sl{font-size:9px;font-weight:700;line-height:1}
.hl{font-size:7px;line-height:1;opacity:.9;display:none}
.P{background:#4CAF50;color:#fff}.PL{background:#FFC107;color:#5d3a00}
.SP{background:#FF9800;color:#fff}.L{background:#2196F3;color:#fff}
.A{background:#F44336;color:#fff}.N{background:#f8fafc;color:#ced4da}.N.wk{background:#f1f3f5}
.no-data{text-align:center;padding:40px;color:#868e96;font-size:13px}
/* Org Chart */
.org-prompt{text-align:center;padding:40px;color:#868e96;font-size:13px}
.org-tree-wrap{overflow:auto;padding:16px 0}
/* Family tree */
.tree-node{display:flex;flex-direction:column;align-items:center;position:relative}
.tree-children{display:flex;gap:12px;position:relative;padding-top:24px;justify-content:center}
.tree-children::before{content:'';position:absolute;top:0;left:50%;transform:translateX(-50%);width:1px;height:24px;background:#dee2e6}
.tree-children::after{content:'';position:absolute;top:0;left:0;right:0;height:1px;background:#dee2e6}
.tree-children .tree-node:first-child::before,.tree-children .tree-node:last-child::before{content:'';position:absolute;top:-1px;width:50%;height:1px;background:#f5f7fa}
.tree-children .tree-node:first-child::before{left:0}
.tree-children .tree-node:last-child::before{right:0}
.tree-children .tree-node{padding-top:0;position:relative}
.tree-children .tree-node::before{content:'';position:absolute;top:-24px;left:50%;transform:translateX(-50%);width:1px;height:24px;background:#dee2e6}
.org-card{background:#fff;border:.5px solid #dee2e6;border-radius:10px;padding:10px 12px;min-width:140px;max-width:180px;text-align:center;cursor:default;transition:box-shadow .15s;position:relative}
.org-card:hover{box-shadow:0 3px 12px rgba(0,0,0,.1)}
.org-card.has-children{cursor:pointer}
.org-card .oc-name{font-size:10px;font-weight:600;color:#1a2332;line-height:1.3;margin-bottom:2px}
.org-card .oc-desig{font-size:8px;color:#868e96;margin-bottom:5px;line-height:1.3}
.org-card .oc-att{font-size:14px;font-weight:700;margin-bottom:1px}
.org-card .oc-meta{display:flex;gap:5px;justify-content:center;flex-wrap:wrap}
.org-card .oc-badge{font-size:8px;padding:1px 5px;border-radius:8px;font-weight:500}
.org-card.lv0{border-top:3px solid #9C27B0}
.org-card.lv1{border-top:3px solid #2196F3}
.org-card.lv2{border-top:3px solid #4CAF50}
.org-card.lv3{border-top:3px solid #FF9800}
.oc-att.hi{color:#4CAF50}.oc-att.mid{color:#FF9800}.oc-att.lo{color:#F44336}
.collapse-btn{position:absolute;bottom:-9px;left:50%;transform:translateX(-50%);width:18px;height:18px;border-radius:50%;background:#2196F3;color:#fff;font-size:10px;display:flex;align-items:center;justify-content:center;cursor:pointer;z-index:2;line-height:1}
.collapsed .tree-children{display:none}
.org-scroll{overflow-x:auto;padding-bottom:8px}
.footer{text-align:center;padding:14px;font-size:11px;color:#ced4da}
</style>
</head>
<body>

<div class="header">
  <h1>📅 Attendance Dashboard</h1>
  <div class="header-right" id="gen-time"></div>
</div>

<div class="tabs">
  <div class="tab active" onclick="switchTab('summary',this)">📈 Summary</div>
  <div class="tab" onclick="switchTab('daily',this)">📅 Daily Heatmap</div>
  <div class="tab" onclick="switchTab('org',this)">🏢 Org Chart</div>
</div>

<!-- Filter bar -->
<div class="filter-bar" id="filter-bar">
  <!-- Date mode tabs -->
  <div class="flt">
    <label>Period</label>
    <div class="date-tabs">
      <div class="date-tab active" onclick="setDateMode('month',this)">Month</div>
      <div class="date-tab" onclick="setDateMode('quarter',this)">Quarter</div>
      <div class="date-tab" onclick="setDateMode('week',this)">Week</div>
      <div class="date-tab" onclick="setDateMode('year',this)">Year</div>
    </div>
  </div>

  <!-- Year always visible -->
  <div class="flt" id="year-flt">
    <label>Year</label>
    <select class="single" id="f-year" onchange="applyFilters()"></select>
  </div>

  <!-- Month subfilter -->
  <div class="date-subfilter visible" id="sf-month">
    <label style="font-size:11px;color:#868e96;font-weight:500">Month</label>
    <select class="single" id="f-month" onchange="applyFilters()">
      <option value="">(All)</option>
    </select>
  </div>

  <!-- Quarter subfilter -->
  <div class="date-subfilter" id="sf-quarter">
    <label style="font-size:11px;color:#868e96;font-weight:500">Quarter</label>
    <select class="single" id="f-quarter" onchange="applyFilters()">
      <option value="">(All)</option>
      <option value="1">Q1</option>
      <option value="2">Q2</option>
      <option value="3">Q3</option>
      <option value="4">Q4</option>
    </select>
  </div>

  <!-- Week subfilter -->
  <div class="date-subfilter" id="sf-week">
    <label style="font-size:11px;color:#868e96;font-weight:500">Week</label>
    <select class="single" id="f-week" onchange="applyFilters()">
      <option value="">(All)</option>
    </select>
  </div>

  <!-- Department multi -->
  <div class="flt">
    <label>Department</label>
    <div class="ms-wrap" id="ms-dept">
      <div class="ms-btn" onclick="toggleMs('ms-dept')">
        <span id="ms-dept-label">All departments</span>
        <span class="ms-badge" id="ms-dept-badge" style="display:none">0</span>
        <span class="ms-arrow">▾</span>
      </div>
      <div class="ms-dropdown" id="ms-dept-dd">
        <div class="ms-search"><input type="text" placeholder="Search..." oninput="filterMsSearch(this,'ms-dept')"></div>
        <div class="ms-list" id="ms-dept-list"></div>
      </div>
    </div>
  </div>

  <!-- Branch multi -->
  <div class="flt">
    <label>Branch</label>
    <div class="ms-wrap" id="ms-branch">
      <div class="ms-btn" onclick="toggleMs('ms-branch')">
        <span id="ms-branch-label">All branches</span>
        <span class="ms-badge" id="ms-branch-badge" style="display:none">0</span>
        <span class="ms-arrow">▾</span>
      </div>
      <div class="ms-dropdown" id="ms-branch-dd">
        <div class="ms-search"><input type="text" placeholder="Search..." oninput="filterMsSearch(this,'ms-branch')"></div>
        <div class="ms-list" id="ms-branch-list"></div>
      </div>
    </div>
  </div>

  <!-- Designation multi -->
  <div class="flt">
    <label>Designation</label>
    <div class="ms-wrap" id="ms-desig">
      <div class="ms-btn" onclick="toggleMs('ms-desig')">
        <span id="ms-desig-label">All designations</span>
        <span class="ms-badge" id="ms-desig-badge" style="display:none">0</span>
        <span class="ms-arrow">▾</span>
      </div>
      <div class="ms-dropdown" id="ms-desig-dd">
        <div class="ms-search"><input type="text" placeholder="Search..." oninput="filterMsSearch(this,'ms-desig')"></div>
        <div class="ms-list" id="ms-desig-list"></div>
      </div>
    </div>
  </div>

  <button class="btn" onclick="resetFilters()">Reset</button>

  <!-- Export -->
  <div class="export-menu" style="margin-left:auto" id="export-wrap">
    <button class="btn btn-export" onclick="toggleExportMenu()">⬇ Export ▾</button>
    <div class="export-dd" id="export-dd">
      <div class="exp-item" onclick="doExport('csv')">📄 CSV</div>
      <div class="exp-item" onclick="doExport('xlsx')">📊 Excel</div>
    </div>
  </div>

  <!-- Shift toggle (daily only) -->
  <div class="toggle-wrap" id="shift-wrap" style="display:none">
    <span>Shift</span>
    <div class="toggle-track" id="shift-tog" onclick="toggleShift()"><div class="toggle-thumb"></div></div>
    <span id="shift-lbl">Off</span>
  </div>
</div>

<!-- Heatmap legend -->
<div class="legend" id="heat-legend" style="display:none">
  <div class="leg"><div class="leg-dot P"></div>Present</div>
  <div class="leg"><div class="leg-dot PL"></div>Present+Late</div>
  <div class="leg"><div class="leg-dot SP"></div>Single punch</div>
  <div class="leg"><div class="leg-dot L"></div>On leave</div>
  <div class="leg"><div class="leg-dot A"></div>Absent</div>
  <div class="leg"><div class="leg-dot N"></div>Non-working</div>
  <div class="leg" style="margin-left:auto;font-size:10px;color:#adb5bd">Hover cell for details</div>
</div>

<div class="content">

  <!-- SUMMARY -->
  <div class="pane active" id="pane-summary">
    <div class="kpi-row">
      <div class="kpi"><div class="kpi-val" id="k-emp">—</div><div class="kpi-lbl">Employees</div></div>
      <div class="kpi"><div class="kpi-val" id="k-att" style="color:#4CAF50">—</div><div class="kpi-lbl">Avg Attendance</div></div>
      <div class="kpi"><div class="kpi-val" id="k-late" style="color:#FFC107">—</div><div class="kpi-lbl">Late incidents</div></div>
      <div class="kpi"><div class="kpi-val" id="k-sp" style="color:#FF9800">—</div><div class="kpi-lbl">Single punch</div></div>
      <div class="kpi"><div class="kpi-val" id="k-leave" style="color:#2196F3">—</div><div class="kpi-lbl">On leave</div></div>
      <div class="kpi"><div class="kpi-val" id="k-wd">—</div><div class="kpi-lbl">Working days</div></div>
    </div>
    <div class="tbl-wrap"><div class="tbl-scroll">
      <table class="stbl">
        <thead><tr>
          <th onclick="sortTable('dept')">Department<span class="sa" id="s-dept">▼</span></th>
          <th onclick="sortTable('Branch')">Branch<span class="sa" id="s-Branch"></span></th>
          <th onclick="sortTable('desig')">Designation<span class="sa" id="s-desig"></span></th>
          <th onclick="sortTable('Month_Label')">Period<span class="sa" id="s-Month_Label"></span></th>
          <th onclick="sortTable('emp')">Emp<span class="sa" id="s-emp"></span></th>
          <th onclick="sortTable('wd')">Working Days<span class="sa" id="s-wd"></span></th>
          <th onclick="sortTable('present')">Present<span class="sa" id="s-present"></span></th>
          <th onclick="sortTable('late')">Late<span class="sa" id="s-late"></span></th>
          <th onclick="sortTable('sp')">Single Punch<span class="sa" id="s-sp"></span></th>
          <th onclick="sortTable('leave')">On Leave<span class="sa" id="s-leave"></span></th>
          <th onclick="sortTable('att')">Att Rate<span class="sa" id="s-att"></span></th>
        </tr></thead>
        <tbody id="sum-body"></tbody>
      </table>
    </div></div>
  </div>

  <!-- DAILY HEATMAP -->
  <div class="pane" id="pane-daily">
    <div class="heat-wrap"><div class="heat-scroll">
      <table class="htbl">
        <thead id="heat-head"></thead>
        <tbody id="heat-body"></tbody>
      </table>
    </div></div>
  </div>

  <!-- ORG CHART -->
  <div class="pane" id="pane-org">
    <div id="org-out"><div class="org-prompt">Select a <strong>Branch</strong> (single) to view the org chart for that branch.</div></div>
  </div>

</div>
<div class="footer">Attendance Dashboard · Data: <span id="gen-footer"></span></div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
var D = DATA_PLACEHOLDER;

var shiftOn=false, curTab='summary', sortCol='dept', sortAsc=true, dateMode='month';
var selDepts=[], selBranches=[], selDesigs=[];
var DOW=['Su','Mo','Tu','We','Th','Fr','Sa'];
var SC_LBL={P:'P',PL:'P*',SP:'SP',L:'L',A:'A',N:''};

// Init header
document.getElementById('gen-time').textContent = 'Data: '+D.generated;
document.getElementById('gen-footer').textContent = D.generated;

// Populate dropdowns
(function(){
  // Year
  var ys=document.getElementById('f-year');
  ys.innerHTML='<option value="">(All)</option>';
  D.years.forEach(function(y){ var o=document.createElement('option'); o.value=y; o.textContent=y; ys.appendChild(o); });
  // Month
  var ms=document.getElementById('f-month');
  D.months.forEach(function(m){ var o=document.createElement('option'); o.value=m; o.textContent=m; ms.appendChild(o); });
  // Week
  var ws=document.getElementById('f-week');
  D.weeks.forEach(function(w){ var o=document.createElement('option'); o.value=w; o.textContent=D.week_labels[String(w)]||('Week '+w); ws.appendChild(o); });
  // Multi-selects
  buildMsList('ms-dept', D.departments, selDepts, 'departments');
  buildMsList('ms-branch', D.branches, selBranches, 'branches');
  buildMsList('ms-desig', D.designations, selDesigs, 'designations');
})();

function buildMsList(wrpId, items, selArr, type){
  var list=document.getElementById(wrpId+'-list');
  var html='<div class="ms-item select-all" onclick="msToggleAll(\''+wrpId+'\',\''+type+'\')"><input type="checkbox" id="'+wrpId+'-all"> Select all</div>';
  items.forEach(function(item,i){
    html+='<div class="ms-item" onclick="msToggle(\''+wrpId+'\','+i+',\''+type+'\')"><input type="checkbox" id="'+wrpId+'-'+i+'" value="'+item.replace(/'/g,"\\'")+'"> '+item+'</div>';
  });
  list.innerHTML=html;
}

function getMsItems(type){ return type==='departments'?D.departments:type==='branches'?D.branches:D.designations; }
function getMsSel(type){ return type==='departments'?selDepts:type==='branches'?selBranches:selDesigs; }
function msToggle(wrpId,idx,type){
  var items=getMsItems(type), selArr=getMsSel(type), val=items[idx];
  var cb=document.getElementById(wrpId+'-'+idx); cb.checked=!cb.checked;
  if(cb.checked){ if(selArr.indexOf(val)<0) selArr.push(val); }
  else { var i=selArr.indexOf(val); if(i>=0) selArr.splice(i,1); }
  updateMsBtn(wrpId,selArr,type); applyFilters();
}
function msToggleAll(wrpId,type){
  var items=getMsItems(type), selArr=getMsSel(type);
  var allCb=document.getElementById(wrpId+'-all'); allCb.checked=!allCb.checked;
  items.forEach(function(item,i){ var cb=document.getElementById(wrpId+'-'+i); if(cb) cb.checked=allCb.checked; });
  selArr.length=0; if(allCb.checked) items.forEach(function(x){ selArr.push(x); });
  updateMsBtn(wrpId,selArr,type); applyFilters();
}
function updateMsBtn(wrpId,selArr,type){
  var lbl=document.getElementById(wrpId+'-label'), badge=document.getElementById(wrpId+'-badge');
  var btn=document.querySelector('#'+wrpId+' .ms-btn');
  var names={'departments':'departments','branches':'branches','designations':'designations'};
  if(!selArr.length){ lbl.textContent='All '+names[type]; badge.style.display='none'; btn.classList.remove('active'); }
  else if(selArr.length===1){ lbl.textContent=selArr[0]; badge.style.display='none'; btn.classList.add('active'); }
  else { lbl.textContent=selArr.length+' selected'; badge.textContent=selArr.length; badge.style.display=''; btn.classList.add('active'); }
}
function filterMsSearch(input,wrpId){
  var q=input.value.toLowerCase();
  document.querySelectorAll('#'+wrpId+'-list .ms-item:not(.select-all)').forEach(function(el){
    el.style.display=el.textContent.toLowerCase().includes(q)?'':'none';
  });
}
function toggleMs(wrpId){
  var dd=document.getElementById(wrpId+'-dd'), isOpen=dd.classList.contains('open');
  closeAllDropdowns();
  if(!isOpen) dd.classList.add('open');
}
function toggleExportMenu(){
  var dd=document.getElementById('export-dd'), isOpen=dd.classList.contains('open');
  closeAllDropdowns(); if(!isOpen) dd.classList.add('open');
}
function closeAllDropdowns(){
  document.querySelectorAll('.ms-dropdown,.export-dd').forEach(function(d){ d.classList.remove('open'); });
}
document.addEventListener('click',function(e){
  if(!e.target.closest('.ms-wrap')&&!e.target.closest('.export-menu')) closeAllDropdowns();
});

function setDateMode(mode, el){
  dateMode=mode;
  document.querySelectorAll('.date-tab').forEach(function(t){ t.classList.remove('active'); });
  el.classList.add('active');
  document.querySelectorAll('.date-subfilter').forEach(function(sf){ sf.classList.remove('visible'); });
  if(mode!=='year'){ var sf=document.getElementById('sf-'+mode); if(sf) sf.classList.add('visible'); }
  applyFilters();
}

function getF(){
  var yr=document.getElementById('f-year').value;
  var mo=document.getElementById('f-month').value;
  var qt=document.getElementById('f-quarter').value;
  var wk=document.getElementById('f-week').value;
  return { year:yr?parseInt(yr):null, month:mo, quarter:qt?parseInt(qt):null,
           week:wk?parseInt(wk):null, depts:selDepts.slice(), branches:selBranches.slice(), desigs:selDesigs.slice() };
}

function matchF(f, r){
  if(f.year     && r.Year    !==f.year)     return false;
  if(f.month    && r.Month_Label!==f.month) return false;
  if(f.quarter  && r.Quarter !==f.quarter)  return false;
  if(f.week     && r.WeekNum !==f.week)     return false;
  if(f.depts.length  && r.dept   && f.depts.indexOf(r.dept)<0)    return false;
  if(f.branches.length&&(r.branch||r.Branch)&&f.branches.indexOf(r.branch||r.Branch)<0) return false;
  if(f.desigs.length && r.desig  && f.desigs.indexOf(r.desig)<0)  return false;
  return true;
}

function resetFilters(){
  document.getElementById('f-year').value='';
  document.getElementById('f-month').value='';
  document.getElementById('f-quarter').value='';
  document.getElementById('f-week').value='';
  selDepts.length=0; selBranches.length=0; selDesigs.length=0;
  ['ms-dept','ms-branch','ms-desig'].forEach(function(w){
    var type=w==='ms-dept'?'departments':w==='ms-branch'?'branches':'designations';
    document.querySelectorAll('#'+w+'-list input[type=checkbox]').forEach(function(cb){ cb.checked=false; });
    updateMsBtn(w,[],type);
  });
  applyFilters();
}

function applyFilters(){
  if(curTab==='summary') renderSummary();
  else if(curTab==='daily') renderHeatmap();
  else renderOrg();
}

function switchTab(tab,el){
  curTab=tab;
  document.querySelectorAll('.tab').forEach(function(t){ t.classList.remove('active'); });
  el.classList.add('active');
  document.querySelectorAll('.pane').forEach(function(p){ p.classList.remove('active'); });
  document.getElementById('pane-'+tab).classList.add('active');
  document.getElementById('shift-wrap').style.display   = tab==='daily'?'flex':'none';
  document.getElementById('heat-legend').style.display  = tab==='daily'?'flex':'none';
  document.getElementById('export-wrap').style.display  = tab==='org'?'none':'';
  applyFilters();
}

// ── SUMMARY ───────────────────────────────────────────────────────────────────
function sortTable(col){
  if(sortCol===col) sortAsc=!sortAsc; else{sortCol=col;sortAsc=true;}
  document.querySelectorAll('.sa').forEach(function(e){e.textContent='';});
  var a=document.getElementById('s-'+col); if(a) a.textContent=sortAsc?'▲':'▼';
  renderSummary();
}
function renderSummary(){
  var f=getF();
  var rows=D.summary.filter(function(r){ return matchF(f,r); });
  rows.sort(function(a,b){
    var av=a[sortCol],bv=b[sortCol];
    return typeof av==='string'?(sortAsc?String(av).localeCompare(String(bv)):String(bv).localeCompare(String(av))):(sortAsc?av-bv:bv-av);
  });
  var totWD=0,totP=0,totL=0,totSP=0,totLv=0,totEmp=0;
  rows.forEach(function(r){totWD+=r.wd;totP+=r.present;totL+=r.late;totSP+=r.sp;totLv+=r.leave;totEmp+=r.emp;});
  document.getElementById('k-emp').textContent   = totEmp.toLocaleString();
  document.getElementById('k-att').textContent   = totWD>0?(totP/totWD*100).toFixed(1)+'%':'—';
  document.getElementById('k-late').textContent  = totL.toLocaleString();
  document.getElementById('k-sp').textContent    = totSP.toLocaleString();
  document.getElementById('k-leave').textContent = totLv.toLocaleString();
  document.getElementById('k-wd').textContent    = totWD.toLocaleString();
  var html='';
  if(!rows.length) html='<tr><td colspan="11" class="no-data">No data for selected filters.</td></tr>';
  else rows.forEach(function(r){
    var cls=r.att>=80?'hi':r.att>=60?'mid':'lo';
    var period = dateMode==='week'?(D.week_labels[String(r.WeekNum)]||('W'+r.WeekNum)):
                 dateMode==='quarter'?('Q'+r.Quarter+' '+r.Year):
                 dateMode==='year'?String(r.Year):r.Month_Label;
    html+='<tr><td>'+r.dept+'</td><td>'+r.Branch+'</td><td style="max-width:200px;overflow:hidden;text-overflow:ellipsis">'+r.desig+'</td><td>'+period+'</td><td>'+r.emp+'</td><td>'+r.wd+'</td><td>'+r.present+'</td><td>'+r.late+'</td><td>'+r.sp+'</td><td>'+r.leave+'</td><td><span class="ap '+cls+'">'+r.att+'%</span></td></tr>';
  });
  document.getElementById('sum-body').innerHTML=html;
}

// ── HEATMAP ───────────────────────────────────────────────────────────────────
function toggleShift(){
  shiftOn=!shiftOn;
  document.getElementById('shift-tog').classList.toggle('on',shiftOn);
  document.getElementById('shift-lbl').textContent=shiftOn?'On':'Off';
  document.querySelectorAll('.hl').forEach(function(el){ el.style.display=shiftOn?'block':'none'; });
}

function renderHeatmap(){
  var f=getF();
  // Heatmap needs a single month — default to first available
  var month=f.month;
  if(!month){
    // pick first month matching year/quarter/week if set
    month=D.months[0];
    if(f.year||f.quarter||f.week){
      for(var i=0;i<D.summary.length;i++){
        var r=D.summary[i];
        if((!f.year||r.Year===f.year)&&(!f.quarter||r.Quarter===f.quarter)&&(!f.week||r.WeekNum===f.week)){
          month=r.Month_Label; break;
        }
      }
    }
    document.getElementById('f-month').value=month;
  }
  var cfg=D.month_cfg[month]||{days:31,start:0};

  var emps=[];
  Object.keys(D.heat).forEach(function(branch){
    if(f.branches.length && f.branches.indexOf(branch)<0) return;
    var mData=D.heat[branch][month]; if(!mData) return;
    Object.keys(mData).forEach(function(id){
      var e=mData[id];
      if(f.depts.length   && f.depts.indexOf(e.dept)<0)  return;
      if(f.desigs.length  && f.desigs.indexOf(e.desig)<0) return;
      emps.push({id:id,name:e.name,dept:e.dept,desig:e.desig,branch:branch,days:e.days});
    });
  });
  emps.sort(function(a,b){ return a.name.localeCompare(b.name); });

  var hh='<tr><th class="he">Employee</th><th class="hd">Dept · Designation</th>';
  for(var d=1;d<=cfg.days;d++){
    var dow=(cfg.start+d-1)%7, wk=dow===0||dow===6;
    hh+='<th class="hdc'+(wk?' wk':'')+'"><div class="hdn">'+('0'+d).slice(-2)+'</div><div class="hdw">'+DOW[dow]+'</div></th>';
  }
  hh+='</tr>';
  document.getElementById('heat-head').innerHTML=hh;

  var bh='';
  if(!emps.length){
    bh='<tr><td colspan="'+(cfg.days+2)+'" class="no-data">No employees found. Select a branch to view the heatmap.</td></tr>';
  } else {
    emps.forEach(function(emp){
      bh+='<tr><td class="tde" title="'+emp.name+'">'+emp.name+'</td>';
      bh+='<td class="tdd">'+emp.dept+'<span class="tds">'+emp.desig+'</span></td>';
      for(var d=1;d<=cfg.days;d++){
        var ds=('0'+d).slice(-2), dow=(cfg.start+d-1)%7, wk=dow===0||dow===6;
        var day=emp.days[ds]||{sc:'N',sh:''}, s=day.sc||'N', sh=day.sh||'', lbl=SC_LBL[s]||'';
        bh+='<td class="tdc '+s+(wk&&s==='N'?' wk':'')+'" title="'+ds+' '+month+' · '+s+(sh?' · '+sh:'')+'">';
        bh+='<div class="ci"><span class="sl">'+lbl+'</span><span class="hl" style="display:'+(shiftOn?'block':'none')+'">'+sh+'</span></div></td>';
      }
      bh+='</tr>';
    });
  }
  document.getElementById('heat-body').innerHTML=bh;
}

// ── ORG CHART (family tree) ───────────────────────────────────────────────────
var orgMap={};
D.org.forEach(function(e){ orgMap[e['Employee Id']]=e; });

function buildChildren(parentId, allIds){
  return D.org.filter(function(e){ return e.mgr_id===parentId && e['Employee Id']!==parentId && allIds.has(e['Employee Id']); });
}

function renderOrgCard(emp, level, allIds){
  var children=buildChildren(emp['Employee Id'], allIds);
  var hasChildren=children.length>0;
  var ac=emp.att>=80?'hi':emp.att>=60?'mid':'lo';
  var lv=Math.min(level,3);

  var html='<div class="tree-node" id="node-'+emp['Employee Id']+'">';
  html+='<div class="org-card lv'+lv+(hasChildren?' has-children':'')+'" '+(hasChildren?'onclick="toggleNode(\''+emp['Employee Id']+'\')"':'')+' title="'+emp['Employee Id']+' · '+emp.desig+'">';
  html+='<div class="oc-name">'+emp.name+'</div>';
  html+='<div class="oc-desig">'+emp.desig+'</div>';
  html+='<div class="oc-att '+ac+'">'+emp.att+'%</div>';
  html+='<div class="oc-meta">';
  html+='<span class="oc-badge" style="background:#f1f3f5;color:#495057">'+emp.late+' late</span>';
  html+='<span class="oc-badge" style="background:#f1f3f5;color:#495057">'+emp.wd+' days</span>';
  html+='</div>';
  if(hasChildren) html+='<div class="collapse-btn" title="Expand/Collapse">'+children.length+'</div>';
  html+='</div>';

  if(hasChildren){
    html+='<div class="tree-children">';
    children.sort(function(a,b){ return a.name.localeCompare(b.name); });
    children.forEach(function(child){ html+=renderOrgCard(child, level+1, allIds); });
    html+='</div>';
  }
  html+='</div>';
  return html;
}

function toggleNode(empId){
  var node=document.getElementById('node-'+empId);
  if(node) node.classList.toggle('collapsed');
}

function renderOrg(){
  var f=getF();
  var out=document.getElementById('org-out');

  // Filter employees
  var filtered=D.org.filter(function(e){
    if(f.branches.length && f.branches.indexOf(e.branch)<0) return false;
    if(f.depts.length    && f.depts.indexOf(e.dept)<0)      return false;
    if(f.desigs.length   && f.desigs.indexOf(e.desig)<0)    return false;
    return true;
  });

  if(!filtered.length){
    out.innerHTML='<div class="org-prompt">No employees found for selected filters.</div>'; return;
  }
  if(filtered.length>200){
    out.innerHTML='<div class="org-prompt">Too many employees ('+filtered.length+') to display as a tree. Please filter by <strong>Branch</strong> or <strong>Department</strong> to narrow down.</div>'; return;
  }

  var allIds=new Set(filtered.map(function(e){ return e['Employee Id']; }));

  // Find roots: employees whose manager is not in the filtered set
  var roots=filtered.filter(function(e){ return !allIds.has(e.mgr_id); });
  roots.sort(function(a,b){ return a.name.localeCompare(b.name); });

  var html='<div class="org-scroll"><div style="display:flex;gap:24px;flex-wrap:wrap;justify-content:flex-start;padding:8px 0">';
  roots.forEach(function(root){ html+=renderOrgCard(root, 0, allIds); });
  html+='</div></div>';
  out.innerHTML=html;
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
function doExport(fmt){
  closeAllDropdowns();
  var f=getF();
  var rows=D.summary.filter(function(r){ return matchF(f,r); });
  var data=[['Department','Branch','Designation','Month','Year','Quarter','Week','Employees','Working Days','Present','Late','Single Punch','On Leave','Att Rate %']];
  rows.forEach(function(r){
    data.push([r.dept,r.Branch,r.desig,r.Month_Label,r.Year,r.Quarter,r.WeekNum,r.emp,r.wd,r.present,r.late,r.sp,r.leave,r.att]);
  });
  if(fmt==='csv'){
    var csv=data.map(function(row){ return row.map(function(c){ return '"'+String(c).replace(/"/g,'""')+'"'; }).join(','); }).join('\\n');
    var blob=new Blob([csv],{type:'text/csv'}); var url=URL.createObjectURL(blob);
    var a=document.createElement('a'); a.href=url; a.download='summary_export.csv'; a.click(); URL.revokeObjectURL(url);
  } else {
    if(typeof XLSX==='undefined'){alert('Loading Excel library, please try again.');return;}
    var ws=XLSX.utils.aoa_to_sheet(data), wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,'Summary'); XLSX.writeFile(wb,'summary_export.xlsx');
  }
}

// Init
renderSummary();
</script>
</body>
</html>"""

html_out = HTML.replace('DATA_PLACEHOLDER', data_js)

with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
    f.write(html_out)

print(f"✅ Dashboard saved: {OUTPUT_FILE}")
print(f"   Size: {os.path.getsize(OUTPUT_FILE)//1024} KB")
print(f"   Generated: {out['generated']}")
