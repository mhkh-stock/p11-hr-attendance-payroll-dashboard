# P11 — Data dictionary

**Source:** generated SAMPLE data (`src/generate_data.py`, seed 11)
for the fictional "CityWorks Pvt Ltd". 24 employees in 4 departments,
October 2026, 22 working days (Mon–Fri, 1–30 Oct 2026). No real
company, no real employees, no real salaries.

## `data/raw/attendance_SAMPLE.csv` — one row per employee per working day (528 rows)

| Column | Type | Meaning |
|---|---|---|
| date | date | YYYY-MM-DD, a Mon–Fri working day in Oct 2026 |
| employee_id | text | `CW-001` … `CW-024` |
| employee_name | text | Fictional employee name |
| department | text | Operations, Sales, Accounts, HR & Admin |
| status | text | Present, Late, Half-day, Leave, Absent |

## `data/raw/employees_SAMPLE.csv` — one row per employee (24 rows)

| Column | Type | Meaning |
|---|---|---|
| employee_id | text | `CW-001` … `CW-024`, matches the attendance log |
| employee_name | text | Fictional employee name |
| department | text | Operations, Sales, Accounts, HR & Admin |
| designation | text | Job title (Manager, Supervisor, Officer, Executive, etc.) |
| basic_salary_pkr | int | Monthly basic salary in PKR (32,000–120,000) |
| transport_allowance_pkr | int | Monthly transport allowance in PKR |
| medical_allowance_pkr | int | Monthly medical allowance in PKR |
| overtime_hours | int | Overtime hours worked in October (0–16; managers 0) |

## Rules baked into the data

- Attendance is generated per employee with behaviour weights:
  most staff are regular (~88% present); Tariq Aziz (CW-007) is
  chronically late, Shahid Iqbal (CW-008) chronically absent,
  Imran Shah (CW-005) takes a longer leave spell. The patterns are
  in `src/generate_data.py` (`BEHAVIOUR`).
- Leave is paid leave: it counts toward attendance but never
  deducts pay.
- Managers work 0 overtime hours by design; field and office staff
  carry the overtime.
- Salaries are realistic PKR ranges for a small Pakistani company,
  not real salaries.

## Payroll rules (applied by the workbook's formulas, not the data)

- Daily wage = basic / 26; hourly rate = basic / 208; overtime at 1.5×.
- Every 3 lates = 1 day deducted; half-day = 0.5 day; absence = 1 day.
- EOBI = 1% of basic. Income tax on gross: 0% to PKR 50,000,
  5% on PKR 50,001–100,000, 10% above PKR 100,000.
- All money rounded to whole rupees.
