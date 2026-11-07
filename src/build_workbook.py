"""Build the P11 Excel workbook (openpyxl, live formulas — no hard-coded totals).

Sheets: Dashboard (KPIs + by-department table + chart, all formulas),
Attendance (per-employee daily summary via COUNTIFS over Data),
Payroll (every pay component formula-driven, INDEX/MATCH into Data and
Attendance), Data (attendance log + salary table, the source tables).

Run: python3 src/build_workbook.py  (after src/generate_data.py)
Output: outputs/hr_attendance_payroll_dashboard_SAMPLE.xlsx
"""

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "outputs" / "hr_attendance_payroll_dashboard_SAMPLE.xlsx"

ATT_ROWS = 528          # 24 employees x 22 days
ATT_FIRST, ATT_LAST = 2, 529
EMP_FIRST, EMP_LAST = 2, 25   # salary table rows on Data
SUM_FIRST, SUM_LAST = 3, 26   # Attendance/Payroll data rows (title row 1, header row 2)

TITLE_FILL = PatternFill("solid", fgColor="1F4E79")
TITLE_FONT = Font(bold=True, color="FFFFFF", size=11)
HDR_FILL = PatternFill("solid", fgColor="2E75B6")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TOTAL_FILL = PatternFill("solid", fgColor="D9E2F3")
TOTAL_FONT = Font(bold=True, size=11)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "#,##0"
PCT = "0.0%"


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def style_total_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = TOTAL_FONT
        cell.border = BOX


def title_row(ws, text, ncols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    cell = ws.cell(row=1, column=1, value=text)
    cell.font = Font(bold=True, size=13, color="1F4E79")
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22


def build_data_sheet(wb):
    ws = wb.create_sheet("Data")
    # --- attendance log (A:E) ---
    att = list(csv.DictReader(open(ROOT / "data" / "raw" / "attendance_SAMPLE.csv")))
    assert len(att) == ATT_ROWS
    headers = ["Date", "Employee ID", "Employee Name", "Department", "Status"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    for r, row in enumerate(att, ATT_FIRST):
        ws.cell(row=r, column=1, value=row["date"])
        ws.cell(row=r, column=2, value=row["employee_id"])
        ws.cell(row=r, column=3, value=row["employee_name"])
        ws.cell(row=r, column=4, value=row["department"])
        ws.cell(row=r, column=5, value=row["status"])
    style_header(ws, 1, 5)
    ws.add_table(Table(displayName="AttendanceLog", ref=f"A1:E{ATT_LAST}",
                       tableStyleInfo=TableStyleInfo(name="TableStyleMedium9",
                                                     showRowStripes=True)))
    # --- salary table (G:N) ---
    sal = list(csv.DictReader(open(ROOT / "data" / "raw" / "employees_SAMPLE.csv")))
    assert len(sal) == 24
    sal_headers = ["Employee ID", "Employee Name", "Department", "Designation",
                   "Basic Salary (PKR)", "Transport Allowance (PKR)",
                   "Medical Allowance (PKR)", "Overtime Hours"]
    for c, h in enumerate(sal_headers, 7):
        ws.cell(row=1, column=c, value=h)
    for r, row in enumerate(sal, EMP_FIRST):
        ws.cell(row=r, column=7, value=row["employee_id"])
        ws.cell(row=r, column=8, value=row["employee_name"])
        ws.cell(row=r, column=9, value=row["department"])
        ws.cell(row=r, column=10, value=row["designation"])
        ws.cell(row=r, column=11, value=int(row["basic_salary_pkr"]))
        ws.cell(row=r, column=12, value=int(row["transport_allowance_pkr"]))
        ws.cell(row=r, column=13, value=int(row["medical_allowance_pkr"]))
        ws.cell(row=r, column=14, value=int(row["overtime_hours"]))
    for c in range(7, 15):
        cell = ws.cell(row=1, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
    ws.add_table(Table(displayName="SalaryTable", ref=f"G1:N{EMP_LAST}",
                       tableStyleInfo=TableStyleInfo(name="TableStyleMedium9",
                                                     showRowStripes=True)))
    for c in range(11, 14):
        for r in range(EMP_FIRST, EMP_LAST + 1):
            ws.cell(row=r, column=c).number_format = MONEY
    ws.cell(row=ATT_LAST + 2, column=1,
           value="SAMPLE data — fictional CityWorks Pvt Ltd, October 2026. "
                 "Attendance log (528 rows) + salary table (24 rows). "
                 "Nothing here is a real employee.")
    widths = {"A": 12, "B": 13, "C": 18, "D": 13, "E": 11, "F": 3,
              "G": 13, "H": 18, "I": 13, "J": 19, "K": 14, "L": 15, "M": 15, "N": 13}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return sal


def build_attendance_sheet(wb, sal):
    ws = wb.create_sheet("Attendance")
    title_row(ws, "Attendance Summary — October 2026 (CityWorks Pvt Ltd, SAMPLE)", 11)
    headers = ["Employee ID", "Employee Name", "Department", "Designation",
               "Present", "Late", "Half-day", "Leave", "Absent",
               "Total Days", "Attendance %"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=2, column=c, value=h)
    style_header(ws, 2, 11)
    ws.row_dimensions[2].height = 30
    for i, row in enumerate(sal):
        r = SUM_FIRST + i
        ws.cell(row=r, column=1, value=row["employee_id"])
        ws.cell(row=r, column=2, value=row["employee_name"])
        ws.cell(row=r, column=3, value=row["department"])
        ws.cell(row=r, column=4, value=row["designation"])
        for col, status in [(5, "Present"), (6, "Late"), (7, "Half-day"),
                            (8, "Leave"), (9, "Absent")]:
            ws.cell(row=r, column=col,
                   value=f'=COUNTIFS(Data!$B${ATT_FIRST}:$B${ATT_LAST},$A{r},'
                         f'Data!$E${ATT_FIRST}:$E${ATT_LAST},"{status}")')
        ws.cell(row=r, column=10, value=f"=SUM(E{r}:I{r})")
        ws.cell(row=r, column=11, value=f"=(E{r}+F{r}+0.5*G{r})/J{r}")
        ws.cell(row=r, column=11).number_format = PCT
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = BOX
    tot = SUM_LAST + 1  # row 27
    ws.cell(row=tot, column=1, value="TOTAL")
    for col in range(5, 11):
        letter = get_column_letter(col)
        ws.cell(row=tot, column=col, value=f"=SUM({letter}{SUM_FIRST}:{letter}{SUM_LAST})")
    ws.cell(row=tot, column=11, value=f"=(E{tot}+F{tot}+0.5*G{tot})/J{tot}")
    ws.cell(row=tot, column=11).number_format = PCT
    style_total_row(ws, tot, 11)
    ws.cell(row=tot + 2, column=1,
           value="Attendance % = (Present + Late + 0.5 x Half-day) / Total Days. "
                 "Leave is paid leave — no pay deduction. Every count is a "
                 "live COUNTIFS over the Data sheet.")
    widths = [13, 18, 13, 19, 9, 9, 10, 9, 9, 11, 13]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def build_payroll_sheet(wb, sal):
    ws = wb.create_sheet("Payroll")
    title_row(ws, "Payroll — October 2026 (CityWorks Pvt Ltd, SAMPLE)", 22)
    headers = ["Employee ID", "Employee Name", "Department", "Designation",
               "Basic Salary", "Transport Allow.", "Medical Allow.", "OT Hours",
               "Hourly Rate", "Overtime Pay", "Gross Pay",
               "Late (days)", "Half-day (days)", "Absent (days)", "Daily Wage",
               "Late Deduction", "Half-day Deduction", "Absence Deduction",
               "EOBI (1%)", "Income Tax", "Total Deductions", "Net Pay"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=2, column=c, value=h)
    style_header(ws, 2, 22)
    ws.row_dimensions[2].height = 36
    for i, row in enumerate(sal):
        r = SUM_FIRST + i
        ws.cell(row=r, column=1, value=row["employee_id"])
        ws.cell(row=r, column=2, value=row["employee_name"])
        ws.cell(row=r, column=3, value=row["department"])
        ws.cell(row=r, column=4, value=row["designation"])
        # pay inputs pulled from the Data salary table by employee ID
        ws.cell(row=r, column=5, value=f"=INDEX(Data!$K:$K,MATCH($A{r},Data!$G:$G,0))")
        ws.cell(row=r, column=6, value=f"=INDEX(Data!$L:$L,MATCH($A{r},Data!$G:$G,0))")
        ws.cell(row=r, column=7, value=f"=INDEX(Data!$M:$M,MATCH($A{r},Data!$G:$G,0))")
        ws.cell(row=r, column=8, value=f"=INDEX(Data!$N:$N,MATCH($A{r},Data!$G:$G,0))")
        # hourly rate = basic / 208 (26 days x 8 hours); OT paid at 1.5x
        ws.cell(row=r, column=9, value=f"=ROUND(E{r}/208,2)")
        ws.cell(row=r, column=10, value=f"=ROUND(H{r}*I{r}*1.5,0)")
        ws.cell(row=r, column=11, value=f"=E{r}+F{r}+G{r}+J{r}")
        # attendance counts pulled from the Attendance sheet by employee ID
        ws.cell(row=r, column=12, value=f"=INDEX(Attendance!F:F,MATCH($A{r},Attendance!$A:$A,0))")
        ws.cell(row=r, column=13, value=f"=INDEX(Attendance!G:G,MATCH($A{r},Attendance!$A:$A,0))")
        ws.cell(row=r, column=14, value=f"=INDEX(Attendance!I:I,MATCH($A{r},Attendance!$A:$A,0))")
        # daily wage = basic / 26; every 3 lates = 1 day deducted
        ws.cell(row=r, column=15, value=f"=ROUND(E{r}/26,0)")
        ws.cell(row=r, column=16, value=f"=FLOOR(L{r}/3,1)*O{r}")
        ws.cell(row=r, column=17, value=f"=ROUND(M{r}*0.5*O{r},0)")
        ws.cell(row=r, column=18, value=f"=N{r}*O{r}")
        ws.cell(row=r, column=19, value=f"=ROUND(E{r}*0.01,0)")
        # tax on gross: 0 to 50k; 5% of 50k-100k; 10% above 100k
        ws.cell(row=r, column=20,
               value=f"=ROUND(IF(K{r}<=50000,0,IF(K{r}<=100000,(K{r}-50000)*0.05,"
                     f"2500+(K{r}-100000)*0.1)),0)")
        ws.cell(row=r, column=21, value=f"=SUM(P{r}:T{r})")
        ws.cell(row=r, column=22, value=f"=K{r}-U{r}")
        for c in (5, 6, 7, 10, 11, 15, 16, 17, 18, 19, 20, 21, 22):
            ws.cell(row=r, column=c).number_format = MONEY
        ws.cell(row=r, column=9).number_format = "#,##0.00"
        for c in range(1, 23):
            ws.cell(row=r, column=c).border = BOX
    tot = SUM_LAST + 1  # row 27
    ws.cell(row=tot, column=1, value="TOTAL")
    for col in (5, 6, 7, 8, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 22):
        letter = get_column_letter(col)
        ws.cell(row=tot, column=col, value=f"=SUM({letter}{SUM_FIRST}:{letter}{SUM_LAST})")
        ws.cell(row=tot, column=col).number_format = MONEY
    style_total_row(ws, tot, 22)
    ws.cell(row=tot + 2, column=1,
           value="Rules: daily wage = basic / 26; hourly rate = basic / 208; overtime at 1.5x; "
                 "every 3 lates = 1 day deducted; half-day = 0.5 day; absence = 1 day; "
                 "leave is paid. EOBI = 1% of basic. Tax on gross: 0% to PKR 50,000, "
                 "5% on PKR 50,001-100,000, 10% above PKR 100,000. All money rounded to whole rupees.")
    widths = [13, 18, 13, 19, 12, 12, 12, 9, 11, 12, 12, 10, 12, 11, 11,
              12, 13, 13, 11, 11, 13, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def build_dashboard_sheet(wb):
    ws = wb.create_sheet("Dashboard", 0)
    ws.merge_cells("A1:B1")
    ws["A1"] = "CityWorks Pvt Ltd — HR Attendance & Payroll Dashboard"
    ws["A1"].font = Font(bold=True, size=15, color="1F4E79")
    ws.merge_cells("A2:B2")
    ws["A2"] = "October 2026  ·  24 employees  ·  4 departments  ·  22 working days  ·  SAMPLE data"
    ws["A2"].font = Font(size=11, color="595959")
    ws["A4"] = "Headline KPIs"
    ws["A4"].font = Font(bold=True, size=12, color="1F4E79")

    kpis = [
        ("Total employees", f"=COUNTA(Attendance!A{SUM_FIRST}:A{SUM_LAST})", None),
        ("Working days (Oct 2026)", 22, None),
        ("Employee-days", f"=COUNTA(Data!B{ATT_FIRST}:B{ATT_LAST})", None),
        ("Attendance rate",
         f"=(SUM(Attendance!E{SUM_FIRST}:E{SUM_LAST})"
         f"+SUM(Attendance!F{SUM_FIRST}:F{SUM_LAST})"
         f"+0.5*SUM(Attendance!G{SUM_FIRST}:G{SUM_LAST}))/B7", PCT),
        ("Present days", f"=SUM(Attendance!E{SUM_FIRST}:E{SUM_LAST})", MONEY),
        ("Late days", f"=SUM(Attendance!F{SUM_FIRST}:F{SUM_LAST})", MONEY),
        ("Half-days", f"=SUM(Attendance!G{SUM_FIRST}:G{SUM_LAST})", MONEY),
        ("Leave days (paid)", f"=SUM(Attendance!H{SUM_FIRST}:H{SUM_LAST})", MONEY),
        ("Absent days", f"=SUM(Attendance!I{SUM_FIRST}:I{SUM_LAST})", MONEY),
        ("Total gross pay (PKR)", f"=SUM(Payroll!K{SUM_FIRST}:K{SUM_LAST})", MONEY),
        ("Total deductions (PKR)", f"=SUM(Payroll!U{SUM_FIRST}:U{SUM_LAST})", MONEY),
        ("Total net pay (PKR)", f"=SUM(Payroll!V{SUM_FIRST}:V{SUM_LAST})", MONEY),
        ("Top absentee",
         f"=INDEX(Attendance!B{SUM_FIRST}:B{SUM_LAST},MATCH(MAX(Attendance!I{SUM_FIRST}:I{SUM_LAST}),"
         f"Attendance!I{SUM_FIRST}:I{SUM_LAST},0))", None),
        ("Top absentee — absent days",
         f"=MAX(Attendance!I{SUM_FIRST}:I{SUM_LAST})", MONEY),
        ("Most late employee",
         f"=INDEX(Attendance!B{SUM_FIRST}:B{SUM_LAST},MATCH(MAX(Attendance!F{SUM_FIRST}:F{SUM_LAST}),"
         f"Attendance!F{SUM_FIRST}:F{SUM_LAST},0))", None),
        ("Most late — late days",
         f"=MAX(Attendance!F{SUM_FIRST}:F{SUM_LAST})", MONEY),
        ("Lowest attendance %",
         f"=INDEX(Attendance!B{SUM_FIRST}:B{SUM_LAST},MATCH(MIN(Attendance!K{SUM_FIRST}:K{SUM_LAST}),"
         f"Attendance!K{SUM_FIRST}:K{SUM_LAST},0))", None),
        ("Lowest attendance % — value",
         f"=MIN(Attendance!K{SUM_FIRST}:K{SUM_LAST})", PCT),
    ]
    for i, (label, value, fmt) in enumerate(kpis):
        r = 5 + i
        ws.cell(row=r, column=1, value=label).font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=value)
        c.font = Font(bold=True, size=11)
        if fmt:
            c.number_format = fmt
        ws.cell(row=r, column=1).border = BOX
        c.border = BOX

    # --- by-department table ---
    base = 5 + len(kpis) + 2  # row 25
    ws.cell(row=base, column=1, value="By department").font = Font(bold=True, size=12, color="1F4E79")
    hdr = base + 1
    dept_headers = ["Department", "Employees", "Present", "Late", "Half-day",
                    "Leave", "Absent", "Attendance %"]
    for c, h in enumerate(dept_headers, 1):
        ws.cell(row=hdr, column=c, value=h)
    style_header(ws, hdr, 8)
    depts = ["Operations", "Sales", "Accounts", "HR & Admin"]
    for i, dept in enumerate(depts):
        r = hdr + 1 + i
        ws.cell(row=r, column=1, value=dept).font = Font(bold=True)
        ws.cell(row=r, column=2,
               value=f"=COUNTIF(Attendance!$C${SUM_FIRST}:$C${SUM_LAST},$A{r})")
        for col, status in [(3, "Present"), (4, "Late"), (5, "Half-day"),
                            (6, "Leave"), (7, "Absent")]:
            ws.cell(row=r, column=col,
                   value=f'=COUNTIFS(Data!$D${ATT_FIRST}:$D${ATT_LAST},$A{r},'
                         f'Data!$E${ATT_FIRST}:$E${ATT_LAST},"{status}")')
        ws.cell(row=r, column=8, value=f"=(C{r}+D{r}+0.5*E{r})/(B{r}*22)")
        ws.cell(row=r, column=8).number_format = PCT
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = BOX
    tot = hdr + 1 + len(depts)
    ws.cell(row=tot, column=1, value="TOTAL")
    for col in range(2, 8):
        letter = get_column_letter(col)
        ws.cell(row=tot, column=col,
               value=f"=SUM({letter}{hdr + 1}:{letter}{hdr + len(depts)})")
    ws.cell(row=tot, column=8, value=f"=(C{tot}+D{tot}+0.5*E{tot})/(B{tot}*22)")
    ws.cell(row=tot, column=8).number_format = PCT
    style_total_row(ws, tot, 8)

    ws.cell(row=tot + 2, column=1,
           value="Every number on this sheet is a live formula over Attendance, "
                 "Payroll and Data — change a status or a salary and the "
                 "dashboard updates. Open in Excel or Google Sheets to see values.")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 22
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 13

    # --- chart: attendance rate by department ---
    chart = BarChart()
    chart.type = "col"
    chart.title = "Attendance rate by department"
    chart.y_axis.title = "Attendance rate"
    chart.y_axis.scaling.min = 0.8
    chart.y_axis.number_format = "0%"
    chart.style = 10
    chart.height = 7.5
    chart.width = 14
    data = Reference(ws, min_col=8, min_row=hdr, max_row=hdr + len(depts))
    cats = Reference(ws, min_col=1, min_row=hdr + 1, max_row=hdr + len(depts))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.shape = 4
    ws.add_chart(chart, f"A{tot + 4}")
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)
    sal = build_data_sheet(wb)
    build_attendance_sheet(wb, sal)
    build_payroll_sheet(wb, sal)
    build_dashboard_sheet(wb)
    wb.save(OUT)
    print(f"workbook written: {OUT}")
    print("sheets:", wb.sheetnames)


if __name__ == "__main__":
    main()
