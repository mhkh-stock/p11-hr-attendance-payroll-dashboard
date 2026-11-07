# P11 — HR Attendance & Payroll Dashboard: one workbook that runs the month

**Company:** CityWorks Pvt Ltd (fictional) · **Period:** October 2026 ·
**Status: SAMPLE delivery on DEMO data**
**Tools:** Excel (live formulas + native chart) · Python (openpyxl builder)
**Deliverable:** `outputs/hr_attendance_payroll_dashboard_SAMPLE.xlsx`
(Dashboard + Attendance + Payroll + Data, 24 employees, 4 departments, 22 working days)

> ⚠️ **What this is and is not.** "CityWorks Pvt Ltd" is a fictional
> company and all 24 employees are generated SAMPLE data
> (`src/generate_data.py`, seed 11 — re-running reproduces the
> identical file). What is real is the workbook: every number on
> Dashboard is a live formula over the Attendance and Payroll sheets,
> every attendance count is a COUNTIFS over the Data sheet, and every
> pay component — overtime, late/absence deductions, EOBI, tax — is a
> formula, so changing a status or a salary re-runs the whole month's
> payroll with nothing re-typed. No real employee, salary or attendance
> record is in this data.

## Business question it answers

> "Who actually showed up this month, who didn't, and what does
> everyone get paid — with every deduction explained, not just typed in?"

A small company owner usually has attendance in one register and pay
in another, and the link between them is manual arithmetic. This
delivery joins them: one workbook where attendance drives payroll,
so a late day or an absence shows up in the pay slip automatically,
and the dashboard shows at a glance where attendance — and money —
is leaking.

## What is inside

| File | What it does |
|---|---|
| `outputs/hr_attendance_payroll_dashboard_SAMPLE.xlsx` | The deliverable: Dashboard (18 KPIs + by-department table + attendance-rate chart, all formulas) + Attendance (24 per-employee summaries via COUNTIFS) + Payroll (22 columns, every pay component formula-driven) + Data (528-row attendance log + 24-row salary table as Excel Tables) |
| `data/raw/attendance_SAMPLE.csv` | 528 SAMPLE attendance rows: one per employee per working day |
| `data/raw/employees_SAMPLE.csv` | 24 SAMPLE employees: department, designation, basic salary, allowances, overtime hours |
| `src/generate_data.py` | Generates the sample CSVs (deterministic, seed 11) |
| `src/build_workbook.py` | Builds the workbook: 4 sheets, live formulas, native chart |
| `src/checks.py` | Data-quality checks as code — 64 checks, all passing |
| `requirements.txt` | Pinned dependency: `openpyxl==3.1.2` (verified with Python 3.12) |

## How it fits together

```
data/raw/attendance_SAMPLE.csv ─┐
                                ├─▶ src/build_workbook.py ──▶ outputs/hr_attendance_payroll_dashboard_SAMPLE.xlsx
data/raw/employees_SAMPLE.csv ──┘      (openpyxl)             Dashboard / Attendance / Payroll / Data
                                                              ▲ src/checks.py verifies all of it
```

## Payroll rules (written on the Payroll sheet itself)

- **Daily wage** = basic salary / 26 · **Hourly rate** = basic / 208 (26 days × 8 hours)
- **Overtime** paid at 1.5× the hourly rate
- **Late:** every 3 lates = 1 day deducted (1–2 lates cost nothing)
- **Half-day** = 0.5 day deducted · **Absence** = 1 full day deducted · **Leave** is paid, no deduction
- **EOBI** = 1% of basic salary
- **Income tax** on gross pay: 0% to PKR 50,000, 5% on PKR 50,001–100,000, 10% above PKR 100,000
- All money rounded to whole rupees

## 3 key findings (from this sample)

1. **Operations carries the attendance problem.** Operations runs at
   87.8% attendance against Accounts' 96.0% — and 10 of the month's
   14 absences sit in Operations (8 people) while Accounts (4 people)
   had zero. *So what:* the attendance conversation is a one-department
   conversation, not a company-wide one — the Dashboard's
   by-department table exists so the owner sees that in one glance.
2. **One employee drives the absence story.** Shahid Iqbal (Driver,
   Operations) was absent 6 of 22 days — 68.2% attendance, the lowest
   in the company — and lost PKR 8,772 of pay to absences alone,
   taking home PKR 38,040, the lowest net pay of the month.
   *So what:* the Dashboard names the top absentee by formula, so the
   name updates itself next month instead of being typed into a report.
3. **Attendance, not tax, is the biggest deduction.** Attendance-related
   deductions (absence PKR 28,810 + half-day PKR 12,136 + late
   PKR 8,961) total PKR 49,907 — 51.9% of all PKR 96,248 deducted,
   ahead of income tax (PKR 32,011) and EOBI (PKR 14,330). *So what:*
   fixing attendance pays more than renegotiating any allowance —
   and the late rule is lenient by design: 34 late days produced only
   PKR 8,961 of deductions because the first 2 lates per employee are free.

## Verified numbers

| KPI | Value |
|---|---|
| Employees | 24 (Operations 8, Sales 6, Accounts 4, HR & Admin 6) |
| Working days (Oct 2026) | 22 (Mon–Fri, 1–30 Oct) |
| Employee-days | 528 |
| Attendance rate | 92.71% |
| Present / Late / Half-day / Leave / Absent | 450 / 34 / 11 / 19 / 14 |
| Total basic salaries | PKR 1,433,000 |
| Total allowances | PKR 241,000 |
| Total overtime pay | PKR 52,205 |
| Total gross pay | PKR 1,726,205 |
| Total deductions | PKR 96,248 |
| Total net pay | PKR 1,629,957 |
| Average net pay | PKR 67,915 |
| Highest net pay | Ahmed Raza (Operations Manager), PKR 127,800 |
| Lowest net pay | Shahid Iqbal (Driver), PKR 38,040 |
| Top absentee | Shahid Iqbal, 6 absences |
| Most late | Tariq Aziz (Driver), 5 lates |
| Best department (attendance) | Accounts, 96.0% |
| Weakest department (attendance) | Operations, 87.8% |
| Net pay by department | Operations PKR 552,621 · Sales PKR 411,541 · Accounts PKR 322,428 · HR & Admin PKR 343,367 |
| Checks | 64 passing (`python3 src/checks.py`) |

Verification: `python3 src/checks.py` (64 checks pass — every KPI
recomputed independently from the CSVs, plus live-formula assertions
on all four sheets) **and** the workbook was opened in LibreOffice,
which calculated every formula: Dashboard shows 92.71% attendance,
PKR 1,726,205 gross, PKR 1,629,957 net, top absentee Shahid Iqbal —
identical to the Python recompute.

## Limits — what this workbook does not do

- **One month, one fictional company.** October 2026 only, PKR only,
  24 employees in 4 departments. A real company replaces the two
  CSVs (same columns) and re-runs the builder.
- **The tax slabs are the sample's simplified rule** (0% / 5% / 10%
  on monthly gross). Real payroll must use the current FBR slabs and
  the company's actual EOBI registration — change the two formulas
  in `src/build_workbook.py` and re-run.
- **Leave is paid and uncapped here.** No leave balance is tracked;
  an employee on 22 days of leave would still be paid in full.
- **Overtime is hours × rate only.** No night-shift or weekend
  premium, no approval workflow — the OT Hours column is the input.
- **Open it in Excel or Google Sheets to see the numbers.** The
  workbook carries live formulas (openpyxl writes formulas, not
  cached values), so they calculate when the file is opened — a raw
  file reader that does not calculate formulas will show blanks
  instead of KPI values.

## How to run it (under 5 steps)

First time only: `pip install -r requirements.txt` (openpyxl 3.1.2,
verified with Python 3.12).

1. Open `outputs/hr_attendance_payroll_dashboard_SAMPLE.xlsx` — that
   is the deliverable.
2. To regenerate the sample data: `python3 src/generate_data.py`.
3. To rebuild the workbook: `python3 src/build_workbook.py`.
4. To check the data: `python3 src/checks.py`.
5. To reuse it for a real client: replace the two CSVs in `data/raw/`
   (same columns, one row per employee per working day) and re-run
   steps 2–3 — Dashboard, attendance counts and payroll rebuild
   themselves from the new lines.

## Tools

- Python 3.12 · openpyxl 3.1.2 (pinned in `requirements.txt`)
- Excel / Google Sheets (the client's tool, not a new tool to learn)

## Data source note

**SAMPLE data, generated 7 Nov 2026.** No real company or employee
is involved. The workbook, the formulas and the payroll rules are
the product; the numbers exist to prove them.

## Repo standard note

This paid project follows the portfolio repo standard (Master Plan
§3.2), slimmed for an Excel deliverable: README + data dictionary
(`data/README.md`) + the deliverable workbook + reproducible scripts
+ a Fiverr-ready package note (`PACKAGE-NOTE.md`). The workbook is
the deliverable a buyer opens; the scripts are the proof it was
built, not typed.
