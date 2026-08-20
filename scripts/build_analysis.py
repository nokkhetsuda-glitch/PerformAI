import pandas as pd
import numpy as np
import calendar
import re

pd.set_option('display.max_columns', None)
pd.set_option('display.width', 250)

BASE = 'Perform AI Team'

# ---------- Load ----------
master = pd.read_excel(f'{BASE}/Master_รายชื่อพนักงานปัจจุบัน-PDE.xlsx')
master['full_name'] = (master['ชื่อ'].str.strip() + ' ' + master['สกุล'].str.strip()).str.replace(r'\s+', ' ', regex=True)

mapping = pd.read_excel(f'{BASE}/JobRole_Mapping_PCC_Database_R.0.xlsx', sheet_name='Mapping Result')
pde_map = mapping[mapping['บริษัท (Code)'] == 'PDE'].copy()

master_full = master.merge(
    pde_map[['รหัสพนักงาน', 'Job Code (PCC Database)', 'ชื่อตำแหน่งใน Job Role Database',
             'ชื่อตำแหน่ง (TH)', 'Job Family', 'Pipeline Level', 'Job Level ของ Pipeline',
             'ผลรวม JE', 'คะแนนความใกล้เคียง (%)', 'สถานะการแมส']],
    left_on='Employee Code', right_on='รหัสพนักงาน', how='left'
)

ts = pd.read_csv('all_timesheet.csv')
ts['employee_name'] = ts['employee_name'].str.replace(r'\s+', ' ', regex=True).str.strip()
ts = ts[ts['employee_name'] != 'TEMP'].copy()

# ---------- Categorize activity ----------
def categorize(job_no, desc):
    j = str(job_no).upper()
    d = str(desc)
    if 'LEAVE' in j or 'ลาป่วย' in d or 'ลากิจ' in d or 'ลาพักร้อน' in d:
        return 'Leave / ลา'
    if 'ADMIN' in j:
        return 'Admin / งานธุรการ'
    if re.search(r'IT SUPPORT', d.upper()):
        return 'IT Support (ticket)'
    if j.startswith('PJ-') or j.startswith('DBA-') or j.startswith('MA-') or 'GRP-RPM' in j or 'ICT-' in j:
        return 'Project work'
    if re.match(r'^\d{3}-', j):
        return 'Project / Job code'
    return 'Other / ไม่ระบุหมวด'

ts['category'] = ts.apply(lambda r: categorize(r['job_no'], r['description']), axis=1)

# ---------- Standard working hours per month (weekdays * 8, proxy) ----------
def std_hours(year, month):
    days_in_month = calendar.monthrange(year, month)[1]
    weekdays = sum(1 for d in range(1, days_in_month + 1) if calendar.weekday(year, month, d) < 5)
    return weekdays * 8

month_std = {}
for (y, m) in ts[['year', 'month']].drop_duplicates().itertuples(index=False):
    month_std[(y, m)] = std_hours(y, m)

# ---------- Per employee-month summary ----------
emp_month = ts.groupby(['employee_name', 'year', 'month', 'month_name']).agg(
    total_hours=('hours', 'sum'),
    n_entries=('hours', 'count'),
    n_distinct_jobs=('job_no', pd.Series.nunique),
).reset_index()
emp_month['std_hours'] = emp_month.apply(lambda r: month_std[(r['year'], r['month'])], axis=1)
emp_month['fill_ratio_%'] = (emp_month['total_hours'] / emp_month['std_hours'] * 100).round(1)

cat_month = ts.groupby(['employee_name', 'year', 'month', 'category'])['hours'].sum().reset_index()
cat_pivot = cat_month.pivot_table(index=['employee_name', 'year', 'month'], columns='category', values='hours', fill_value=0).reset_index()

emp_month = emp_month.merge(cat_pivot, on=['employee_name', 'year', 'month'], how='left')
cat_cols = [c for c in cat_pivot.columns if c not in ('employee_name', 'year', 'month')]
for c in cat_cols:
    emp_month[f'{c} %'] = (emp_month[c] / emp_month['total_hours'] * 100).round(1)

# ---------- Overall per-employee summary (Jan-Jul) ----------
emp_overall = ts.groupby('employee_name').agg(
    months_active=('month', 'nunique'),
    total_hours_7mo=('hours', 'sum'),
    n_entries=('hours', 'count'),
    n_distinct_jobs=('job_no', pd.Series.nunique),
).reset_index()
cat_overall = ts.groupby(['employee_name', 'category'])['hours'].sum().unstack(fill_value=0)
cat_overall_pct = cat_overall.div(cat_overall.sum(axis=1), axis=0).mul(100).round(1)
cat_overall_pct.columns = [f'{c} %' for c in cat_overall_pct.columns]
emp_overall = emp_overall.merge(cat_overall.reset_index(), on='employee_name', how='left')
emp_overall = emp_overall.merge(cat_overall_pct.reset_index(), on='employee_name', how='left')

emp_overall = emp_overall.merge(
    master_full[['full_name', 'Employee Code', 'Job Role', 'Position', 'Employee Level',
                 'ชื่อหน่วยงาน (ระดับ 4)', 'ชื่อตำแหน่งใน Job Role Database', 'Job Family',
                 'ผลรวม JE', 'คะแนนความใกล้เคียง (%)', 'สถานะการแมส']],
    left_on='employee_name', right_on='full_name', how='left'
).drop(columns=['full_name'])

# ---------- Department rollup ----------
dept_month = ts.merge(master_full[['full_name', 'ชื่อหน่วยงาน (ระดับ 4)']], left_on='employee_name', right_on='full_name', how='left')
dept_summary = dept_month.groupby('ชื่อหน่วยงาน (ระดับ 4)').agg(
    headcount=('employee_name', 'nunique'),
    total_hours=('hours', 'sum'),
).reset_index().sort_values('total_hours', ascending=False)

# ---------- Mapping quality summary ----------
map_quality = master_full['สถานะการแมส'].value_counts(dropna=False).reset_index()
map_quality.columns = ['สถานะการแมส Job Role', 'จำนวนพนักงาน']

# ---------- Trend by month ----------
trend = ts.groupby(['year', 'month', 'month_name']).agg(
    headcount=('employee_name', 'nunique'),
    total_hours=('hours', 'sum'),
).reset_index().sort_values(['year', 'month'])
trend['std_hours_total'] = trend.apply(lambda r: month_std[(r['year'], r['month'])] * r['headcount'], axis=1)
trend['fill_ratio_%'] = (trend['total_hours'] / trend['std_hours_total'] * 100).round(1)

# ---------- Save intermediate ----------
emp_month.to_csv('emp_month_summary.csv', index=False)
emp_overall.to_csv('emp_overall_summary.csv', index=False)
dept_summary.to_csv('dept_summary.csv', index=False)
map_quality.to_csv('map_quality.csv', index=False)
trend.to_csv('trend.csv', index=False)

print('=== Trend by month ===')
print(trend)
print('\n=== Dept summary ===')
print(dept_summary)
print('\n=== Mapping quality ===')
print(map_quality)
print('\n=== Category overall totals ===')
print(ts.groupby('category')['hours'].sum().sort_values(ascending=False))
print('\n=== emp_overall sample ===')
print(emp_overall.head(10))
