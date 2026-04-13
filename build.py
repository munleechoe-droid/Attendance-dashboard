#!/usr/bin/env python3
"""
Attendance Dashboard Builder
Usage: python3 build.py [path_to_xlsx]
Output: data.js  (upload to GitHub alongside index.html)
"""
import pandas as pd, json, os, sys, warnings, calendar
warnings.filterwarnings('ignore')

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
DATA_FILE   = sys.argv[1] if len(sys.argv)>1 else os.path.join(SCRIPT_DIR,'0_DailyAttendanceReport_Master.xlsx')
SHEET_NAME  = 'Master_Daily Attendance'
OUTPUT_FILE = os.path.join(SCRIPT_DIR,'data.js')
MONTH_MAP   = {'01-26':'Jan 26','02-26':'Feb 26','03-26':'Mar 26','04-26':'Apr 26',
               '05-26':'May 26','06-26':'Jun 26','07-26':'Jul 26','08-26':'Aug 26',
               '09-26':'Sep 26','10-26':'Oct 26','11-26':'Nov 26','12-26':'Dec 26'}

print(f"Reading: {DATA_FILE}")
df = pd.read_excel(DATA_FILE, sheet_name=SHEET_NAME)
df['Date'] = pd.to_datetime(df['Date'])
df['Employee Id'] = df['Employee Id'].astype(str).str.strip()
df['Direct Manager Employee Id'] = df['Direct Manager Employee Id'].astype(str).str.strip()

def clean(s): return str(s).replace('\u2800','').replace('\u200e','').strip() if pd.notna(s) else ''

df['Month_Label'] = df['Month'].map(MONTH_MAP).fillna(df['Month'])
df['Year']    = df['Date'].dt.year
df['Quarter'] = df['Date'].dt.quarter
df['WeekNum'] = df['Date'].dt.isocalendar().week.astype(int)
df['DayStr']  = df['Date'].dt.strftime('%d')
work = df[df['Working Days']==1].copy()

def sc(status, is_late, wd):
    s = str(status) if pd.notna(status) else ''
    if 'Present' in s: return 'PL' if is_late=='Yes' else 'P'
    if 'Single Punch' in s: return 'SP'
    if 'Leave' in s: return 'L'
    if wd==1: return 'A'
    return 'N'

def shift_code(s):
    if pd.isna(s): return ''
    s=str(s); return 'S1' if 'S1' in s else 'S2' if 'S2' in s else 'S'

# Summary with weekly granularity
print("Building summary...")
summary = work.groupby(['Current Department','Branch','Current Designation','Month_Label','Year','Quarter','WeekNum']).agg(
    emp=('Employee Id','nunique'), wd=('Working Days','sum'),
    present=('Status', lambda x: x.str.contains('Present',na=False).sum()),
    late=('Is Late ',lambda x:(x=='Yes').sum()), sp=('Single Punch','sum'), leave=('On Leave','sum'),
).reset_index()
summary['att'] = (summary['present']/summary['wd']*100).round(1)
summary_rows = summary.rename(columns={'Current Department':'dept','Branch':'branch','Current Designation':'desig'}).to_dict('records')

# Week-day map
week_day_map = {}
for (month, week), grp in df.groupby(['Month_Label','WeekNum']):
    if month not in week_day_map: week_day_map[month]={}
    week_day_map[month][str(week)] = sorted(grp['DayStr'].unique().tolist())

# Heatmap
print("Building heatmap...")
detail_heat = {}
for _, row in df.iterrows():
    branch=str(row['Branch']).strip(); month=str(row['Month_Label'])
    if not month or month=='nan': continue
    emp_id=str(row['Employee Id']).strip(); day=row['DayStr']
    status_c=sc(row['Status'],row['Is Late '],row['Working Days']); sh=shift_code(row['Shift'])
    if status_c=='N': continue
    if branch not in detail_heat: detail_heat[branch]={}
    if month not in detail_heat[branch]: detail_heat[branch][month]={}
    if emp_id not in detail_heat[branch][month]:
        detail_heat[branch][month][emp_id]={'name':clean(row['Employee Name']),'dept':str(row['Current Department']).strip(),
            'desig':str(row['Current Designation']).strip(),'mgr_id':str(row['Direct Manager Employee Id']).strip(),'days':{}}
    entry={'sc':status_c}
    if sh: entry['sh']=sh
    detail_heat[branch][month][emp_id]['days'][day]=entry

# Org
print("Building org...")
ep = df.groupby('Employee Id').agg(name=('Employee Name',lambda x:clean(x.iloc[0])),
    dept=('Current Department','first'),desig=('Current Designation','first'),branch=('Branch','first'),
    mgr_id=('Direct Manager Employee Id','first'),mgr_name=('Direct Manager Name',lambda x:clean(x.iloc[0]))).reset_index()
ws_agg = work.groupby('Employee Id').agg(wd=('Working Days','sum'),
    present=('Status',lambda x:x.str.contains('Present',na=False).sum()),late=('Is Late ',lambda x:(x=='Yes').sum())).reset_index()
ep = ep.merge(ws_agg,on='Employee Id',how='left')
ep['att']=(ep['present']/ep['wd']*100).round(1).fillna(0)
ep['Employee Id']=ep['Employee Id'].astype(str); ep['mgr_id']=ep['mgr_id'].astype(str)
org_list=ep.to_dict('records')

# Filters
month_order=[m for m in ['Jan 26','Feb 26','Mar 26','Apr 26','May 26','Jun 26','Jul 26','Aug 26','Sep 26','Oct 26','Nov 26','Dec 26'] if m in df['Month_Label'].unique()]
months_num={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
month_cfg={}
for ml in month_order:
    p=ml.split(' '); mn=months_num[p[0]]; yr=2000+int(p[1])
    month_cfg[ml]={'days':calendar.monthrange(yr,mn)[1],'start':(calendar.weekday(yr,mn,1)+1)%7}
weeks=sorted(df['WeekNum'].unique().tolist())
week_labels={str(w):f"W{w} ({df[df['WeekNum']==w]['Date'].min().strftime('%d %b')} – {df[df['WeekNum']==w]['Date'].max().strftime('%d %b')})" for w in weeks}

out={'summary':summary_rows,'heat':detail_heat,'org':org_list,
    'branches':sorted(df['Branch'].dropna().unique().tolist()),
    'departments':sorted(df['Current Department'].dropna().unique().tolist()),
    'designations':sorted(df['Current Designation'].dropna().unique().tolist()),
    'months':month_order,'years':sorted(df['Year'].unique().tolist()),
    'quarters':sorted(df['Quarter'].unique().tolist()),'weeks':weeks,
    'week_labels':week_labels,'week_day_map':week_day_map,'month_cfg':month_cfg,
    'generated':pd.Timestamp.now().strftime('%d %b %Y %H:%M')}

with open(OUTPUT_FILE,'w') as f:
    f.write('var D = '+json.dumps(out)+';')
print(f"✅ data.js saved: {os.path.getsize(OUTPUT_FILE)//1024} KB")
