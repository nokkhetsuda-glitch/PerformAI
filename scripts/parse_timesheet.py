import pandas as pd
import re
import glob
import os

COL_NO = 3
COL_NAME = 4
COL_DATE = 12
COL_JOBNO = 13
COL_DESC = 20
COL_HRS = 22

MONTH_MAP = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'April': 4, 'May': 5, 'June': 6, 'July': 7,
}

def parse_file(path):
    fname = os.path.basename(path)
    m = re.search(r'Timesheet PDE_(\w+) (\d{4})', fname)
    month_name, year = m.group(1), int(m.group(2))
    month_num = MONTH_MAP[month_name]

    df = pd.read_excel(path, header=None)
    n_rows, n_cols = df.shape

    records = []
    cur_emp = None
    cur_emp_no = None
    cur_emp_total = None

    for i in range(n_rows):
        row = df.iloc[i]
        no_val = row[COL_NO] if COL_NO < n_cols else None
        name_val = row[COL_NAME] if COL_NAME < n_cols else None
        date_val = row[COL_DATE] if COL_DATE < n_cols else None

        is_emp_header = (
            pd.notna(no_val) and isinstance(no_val, (int, float)) and
            pd.notna(name_val) and isinstance(name_val, str) and name_val.strip() != ''
        )
        is_detail = (
            pd.notna(date_val) and isinstance(date_val, str) and '/' in date_val
        )

        if is_emp_header:
            cur_emp = name_val.strip()
            cur_emp_no = int(no_val)
            # total hours often in last non-null column of this row
            last_val = None
            for c in range(n_cols - 1, COL_DATE, -1):
                v = row[c]
                if pd.notna(v) and isinstance(v, (int, float)):
                    last_val = v
                    break
            cur_emp_total = last_val
            continue

        if is_detail and cur_emp is not None:
            job_no = row[COL_JOBNO] if COL_JOBNO < n_cols else None
            desc = row[COL_DESC] if COL_DESC < n_cols else None
            hrs = row[COL_HRS] if COL_HRS < n_cols else None
            records.append({
                'year': year,
                'month': month_num,
                'month_name': month_name,
                'emp_seq_no': cur_emp_no,
                'employee_name': cur_emp,
                'emp_month_total_hrs': cur_emp_total,
                'date_str': date_val,
                'job_no': job_no.strip() if isinstance(job_no, str) else job_no,
                'description': desc.strip() if isinstance(desc, str) else desc,
                'hours': hrs,
            })

    return pd.DataFrame(records)


if __name__ == '__main__':
    path = '/tmp/analysis/Perform AI Team/Daily Report PDE/Timesheet PDE_July 2026.xls'
    d = parse_file(path)
    print(d.shape)
    print(d.head(20))
    print('unique employees:', d['employee_name'].nunique())
    print(d['employee_name'].unique()[:20])
