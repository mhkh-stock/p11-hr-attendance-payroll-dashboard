"""Checks as code for P11. Recomputes every headline number from the raw
CSVs (independent of the workbook) and verifies the workbook structure
(live formulas, tables, chart, sheets). openpyxl does not calculate
formulas, so formula cells are asserted by their formula strings, never
by cached values — the payroll math itself is re-derived from the CSVs
in Python below, using the same rules (Excel ROUND = half away from
zero, implemented here as rhu).

Run: python3 src/checks.py
"""
import csv
import datetime
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parents[1]
ATT_CSV = BASE / "data" / "raw" / "attendance_SAMPLE.csv"
EMP_CSV = BASE / "data" / "raw" / "employees_SAMPLE.csv"
XLSX = BASE / "outputs" / "hr_attendance_payroll_dashboard_SAMPLE.xlsx"
fails = []


def check(name, cond, detail=""):
    print(("PASS " if cond else "FAIL ") + name + (f" — {detail}" if detail else ""))
    if not cond:
        fails.append(name)


def rhu(x):
    """Round half away from zero (Excel ROUND) for non-negative x."""
    return int(x + 0.5)


att = list(csv.DictReader(open(ATT_CSV)))
emp = list(csv.DictReader(open(EMP_CSV)))
per = defaultdict(Counter)
for r in att:
    per[r["employee_id"]][r["status"]] += 1

# ---- data-quality checks on the CSVs ----
check("24 employees", len(emp) == 24, str(len(emp)))
check("4 departments", len({e["department"] for e in emp}) == 4)
check("dept sizes 8/6/4/6",
      Counter(e["department"] for e in emp)
      == {"Operations": 8, "Sales": 6, "Accounts": 4, "HR & Admin": 6})
check("528 attendance rows (24 x 22)", len(att) == 528, str(len(att)))
check("22 working days, 2026-10-01 to 2026-10-30",
      len({r["date"] for r in att}) == 22
      and min(r["date"] for r in att) == "2026-10-01"
      and max(r["date"] for r in att) == "2026-10-30")
check("no weekend dates in attendance",
      all(datetime.date.fromisoformat(r["date"]).weekday() < 5 for r in att))
check("statuses only Present/Late/Half-day/Leave/Absent",
      set(r["status"] for r in att)
      == {"Present", "Late", "Half-day", "Leave", "Absent"})
check("no duplicate (employee, date) rows",
      len({(r["employee_id"], r["date"]) for r in att}) == 528)
check("every employee has exactly 22 days",
      all(sum(per[eid].values()) == 22 for eid in per))
check("no nulls in key columns",
      all(r[k] for r in att for k in ("date", "employee_id", "status"))
      and all(e[k] for e in emp for k in ("employee_id", "employee_name", "basic_salary_pkr")))
check("basic salaries all positive, 32,000-120,000",
      all(32000 <= int(e["basic_salary_pkr"]) <= 120000 for e in emp))
check("overtime hours 0-16, managers 0",
      all(0 <= int(e["overtime_hours"]) <= 16 for e in emp)
      and all(int(e["overtime_hours"]) == 0
              for e in emp if "Manager" in e["designation"]))

status = Counter(r["status"] for r in att)
check("Present 450 / Late 34 / Half-day 11 / Leave 19 / Absent 14",
      dict(status) == {"Present": 450, "Late": 34, "Half-day": 11,
                       "Leave": 19, "Absent": 14}, str(dict(status)))
att_rate = (status["Present"] + status["Late"] + 0.5 * status["Half-day"]) / 528
check("attendance rate 92.71%",
      round(att_rate * 100, 2) == 92.71, f"{att_rate * 100:.2f}%")
check("top absentee Shahid Iqbal (CW-008), 6 absences",
      max(per.items(), key=lambda kv: kv[1]["Absent"])[0] == "CW-008"
      and per["CW-008"]["Absent"] == 6)
check("most late Tariq Aziz (CW-007), 5 lates",
      max(per.items(), key=lambda kv: kv[1]["Late"])[0] == "CW-007"
      and per["CW-007"]["Late"] == 5)
check("lowest attendance Shahid Iqbal, 68.2%",
      round((per["CW-008"]["Present"] + per["CW-008"]["Late"]
             + 0.5 * per["CW-008"]["Half-day"]) / 22 * 100, 1) == 68.2)

# ---- independent payroll recompute from the CSVs ----
def payroll(e, c):
    basic = int(e["basic_salary_pkr"])
    daily = rhu(basic / 26)
    hourly = rhu(basic / 208 * 100) / 100
    ot_pay = rhu(int(e["overtime_hours"]) * hourly * 1.5)
    gross = basic + int(e["transport_allowance_pkr"]) \
        + int(e["medical_allowance_pkr"]) + ot_pay
    late_ded = (c["Late"] // 3) * daily
    half_ded = rhu(c["Half-day"] * 0.5 * daily)
    abs_ded = c["Absent"] * daily
    eobi = rhu(basic * 0.01)
    tax = rhu(0 if gross <= 50000 else
              ((gross - 50000) * 0.05 if gross <= 100000
               else 2500 + (gross - 100000) * 0.1))
    ded = late_ded + half_ded + abs_ded + eobi + tax
    return {"gross": gross, "ot": ot_pay, "late": late_ded, "half": half_ded,
            "abs": abs_ded, "eobi": eobi, "tax": tax, "ded": ded,
            "net": gross - ded}

pay = {e["employee_id"]: payroll(e, per[e["employee_id"]]) for e in emp}
tot = {k: sum(p[k] for p in pay.values()) for k in
       ("gross", "ot", "late", "half", "abs", "eobi", "tax", "ded", "net")}
check("total basic PKR 1,433,000",
      sum(int(e["basic_salary_pkr"]) for e in emp) == 1433000)
check("total allowances PKR 241,000",
      sum(int(e["transport_allowance_pkr"]) + int(e["medical_allowance_pkr"])
          for e in emp) == 241000)
check("total overtime pay PKR 52,205", tot["ot"] == 52205, str(tot["ot"]))
check("total gross pay PKR 1,726,205", tot["gross"] == 1726205, str(tot["gross"]))
check("late deductions PKR 8,961", tot["late"] == 8961, str(tot["late"]))
check("half-day deductions PKR 12,136", tot["half"] == 12136, str(tot["half"]))
check("absence deductions PKR 28,810", tot["abs"] == 28810, str(tot["abs"]))
check("EOBI total PKR 14,330", tot["eobi"] == 14330, str(tot["eobi"]))
check("income tax total PKR 32,011", tot["tax"] == 32011, str(tot["tax"]))
check("total deductions PKR 96,248", tot["ded"] == 96248, str(tot["ded"]))
check("total net pay PKR 1,629,957", tot["net"] == 1629957, str(tot["net"]))
check("deduction components sum to total deductions",
      tot["late"] + tot["half"] + tot["abs"] + tot["eobi"] + tot["tax"] == tot["ded"])
check("gross - deductions = net", tot["gross"] - tot["ded"] == tot["net"])
check("every net pay is a whole rupee",
      all(p["net"] == int(p["net"]) for p in pay.values()))
check("Ahmed Raza net PKR 127,800 (highest)", pay["CW-001"]["net"] == 127800)
check("Shahid Iqbal net PKR 38,040 (lowest)", pay["CW-008"]["net"] == 38040)
check("Sana Malik net PKR 114,669", pay["CW-009"]["net"] == 114669)
check("Fatima Zahra net PKR 98,242", pay["CW-019"]["net"] == 98242)
check("no net pay exceeds gross", all(p["net"] <= p["gross"] for p in pay.values()))
check("dept net pay: Operations 552,621 / Sales 411,541 / Accounts 322,428 / HR & Admin 343,367",
      {d: sum(pay[e["employee_id"]]["net"] for e in emp if e["department"] == d)
       for d in ("Operations", "Sales", "Accounts", "HR & Admin")}
      == {"Operations": 552621, "Sales": 411541, "Accounts": 322428,
          "HR & Admin": 343367})

# ---- workbook structure + live formulas ----
wb = load_workbook(XLSX)
check("sheets = Dashboard, Data, Attendance, Payroll",
      wb.sheetnames == ["Dashboard", "Data", "Attendance", "Payroll"],
      str(wb.sheetnames))

d = wb["Data"]
check("Data attendance log rows = 529 incl header", d.max_row >= 529,
      str(d.max_row))
check("Data has AttendanceLog + SalaryTable tables",
      "AttendanceLog" in d.tables and "SalaryTable" in d.tables,
      str(list(d.tables)))
check("Data salary table rows = 25 incl header",
      d["G1"].value == "Employee ID" and d["G25"].value == "CW-024",
      f"{d['G1'].value} / {d['G25'].value}")

a = wb["Attendance"]
check("Attendance rows = 29 (title + header + 24 + total + note)", a.max_row == 29,
      str(a.max_row))
check("Attendance Present col is a COUNTIFS formula (row 3)",
      a["E3"].value == '=COUNTIFS(Data!$B$2:$B$529,$A3,Data!$E$2:$E$529,"Present")',
      str(a["E3"].value))
check("Attendance counts are COUNTIFS in all 24 rows x 5 statuses",
      all(isinstance(a.cell(row=r, column=c).value, str)
          and a.cell(row=r, column=c).value.startswith("=COUNTIFS")
          for r in range(3, 27) for c in range(5, 10)))
check("Attendance % is a live formula in all 24 rows",
      all(isinstance(a.cell(row=r, column=11).value, str)
          and a.cell(row=r, column=11).value.startswith("=(E")
          for r in range(3, 27)))
check("Attendance total row is SUM formulas",
      a["E27"].value == "=SUM(E3:E26)" and a["J27"].value == "=SUM(J3:J26)",
      f"{a['E27'].value} / {a['J27'].value}")

p = wb["Payroll"]
check("Payroll rows = 29 (title + header + 24 + total + note)", p.max_row == 29,
      str(p.max_row))
check("Payroll Basic is INDEX/MATCH into Data (row 3)",
      p["E3"].value == "=INDEX(Data!$K:$K,MATCH($A3,Data!$G:$G,0))",
      str(p["E3"].value))
check("Payroll OT pay is ROUND(hours x rate x 1.5) (row 3)",
      p["J3"].value == "=ROUND(H3*I3*1.5,0)", str(p["J3"].value))
check("Payroll Late count is INDEX/MATCH into Attendance (row 3)",
      p["L3"].value == "=INDEX(Attendance!F:F,MATCH($A3,Attendance!$A:$A,0))",
      str(p["L3"].value))
check("Payroll late deduction is FLOOR(late/3) x daily wage (row 3)",
      p["P3"].value == "=FLOOR(L3/3,1)*O3", str(p["P3"].value))
check("Payroll tax formula has the 50k/100k slabs (row 3)",
      isinstance(p["T3"].value, str) and "50000" in p["T3"].value
      and "100000" in p["T3"].value, str(p["T3"].value))
check("Payroll net pay is Gross - Total Deductions in all 24 rows",
      all(p.cell(row=r, column=22).value == f"=K{r}-U{r}" for r in range(3, 27)))
check("Payroll total row is SUM formulas",
      p["K27"].value == "=SUM(K3:K26)" and p["V27"].value == "=SUM(V3:V26)",
      f"{p['K27'].value} / {p['V27'].value}")
check("Payroll gross = basic + allowances + OT in all rows",
      all(p.cell(row=r, column=11).value == f"=E{r}+F{r}+G{r}+J{r}"
          for r in range(3, 27)))

db = wb["Dashboard"]
check("Dashboard employees is a COUNTA formula",
      db["B5"].value == "=COUNTA(Attendance!A3:A26)", str(db["B5"].value))
check("Dashboard employee-days is a COUNTA formula",
      db["B7"].value == "=COUNTA(Data!B2:B529)", str(db["B7"].value))
check("Dashboard attendance rate is a live formula",
      isinstance(db["B8"].value, str) and "SUM(Attendance!E3:E26)" in db["B8"].value
      and "/B7" in db["B8"].value, str(db["B8"].value))
check("Dashboard gross/net are SUM formulas over Payroll",
      db["B14"].value == "=SUM(Payroll!K3:K26)"
      and db["B16"].value == "=SUM(Payroll!V3:V26)",
      f"{db['B14'].value} / {db['B16'].value}")
check("Dashboard top absentee is INDEX/MATCH on MAX over the same range",
      db["B17"].value == "=INDEX(Attendance!B3:B26,MATCH(MAX(Attendance!I3:I26),"
                         "Attendance!I3:I26,0))",
      str(db["B17"].value))
check("Dashboard most-late is INDEX/MATCH on MAX over the same range",
      db["B19"].value == "=INDEX(Attendance!B3:B26,MATCH(MAX(Attendance!F3:F26),"
                         "Attendance!F3:F26,0))",
      str(db["B19"].value))
check("Dashboard lowest-attendance name is INDEX/MATCH on MIN over the same range",
      db["B21"].value == "=INDEX(Attendance!B3:B26,MATCH(MIN(Attendance!K3:K26),"
                         "Attendance!K3:K26,0))",
      str(db["B21"].value))
check("Dashboard dept table Present is a COUNTIFS formula",
      isinstance(db["C27"].value, str) and "COUNTIFS" in db["C27"].value
      and '"Present"' in db["C27"].value, str(db["C27"].value))
check("Dashboard has a bar chart", len(db._charts) == 1,
      f"{len(db._charts)} charts")

print(f"\n{len(fails)} failures" if fails else "\nALL CHECKS PASS")
sys.exit(1 if fails else 0)
