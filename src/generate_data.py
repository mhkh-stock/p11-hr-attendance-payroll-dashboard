"""Generate SAMPLE attendance + employee data for P11 (seed 11).

Fictional company "CityWorks Pvt Ltd", October 2026, 24 employees in
4 departments, 22 working days (Mon-Fri). Deterministic: re-running
produces identical CSVs. Nothing here is real employee data.
"""

import calendar
import csv
import random
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SEED = 11

# emp_id, name, department, designation, basic salary (PKR),
# transport allowance, medical allowance, overtime hours
EMPLOYEES = [
    ("CW-001", "Ahmed Raza", "Operations", "Operations Manager", 120000, 10000, 5000, 0),
    ("CW-002", "Bilal Hussain", "Operations", "Supervisor", 78000, 7000, 4000, 6),
    ("CW-003", "Usman Tariq", "Operations", "Supervisor", 75000, 7000, 4000, 8),
    ("CW-004", "Faisal Mehmood", "Operations", "Field Officer", 48000, 5000, 3000, 14),
    ("CW-005", "Imran Shah", "Operations", "Field Officer", 46000, 5000, 3000, 12),
    ("CW-006", "Naveed Akhtar", "Operations", "Field Officer", 45000, 5000, 3000, 16),
    ("CW-007", "Tariq Aziz", "Operations", "Driver", 40000, 4000, 3000, 10),
    ("CW-008", "Shahid Iqbal", "Operations", "Driver", 38000, 4000, 3000, 8),
    ("CW-009", "Sana Malik", "Sales", "Sales Manager", 110000, 10000, 5000, 0),
    ("CW-010", "Ayesha Khan", "Sales", "Sales Executive", 58000, 7000, 4000, 4),
    ("CW-011", "Hamza Yousaf", "Sales", "Sales Executive", 56000, 7000, 4000, 6),
    ("CW-012", "Zainab Ali", "Sales", "Sales Executive", 55000, 7000, 4000, 5),
    ("CW-013", "Omar Farooq", "Sales", "Sales Assistant", 42000, 5000, 3000, 6),
    ("CW-014", "Maryam Siddiqui", "Sales", "Sales Assistant", 40000, 5000, 3000, 4),
    ("CW-015", "Kamran Ahmed", "Accounts", "Finance Manager", 105000, 10000, 5000, 0),
    ("CW-016", "Saad Mirza", "Accounts", "Accountant", 68000, 7000, 4000, 10),
    ("CW-017", "Hina Raza", "Accounts", "Accountant", 65000, 7000, 4000, 9),
    ("CW-018", "Waqas Ali", "Accounts", "Cashier", 44000, 5000, 3000, 5),
    ("CW-019", "Fatima Zahra", "HR & Admin", "HR Manager", 95000, 10000, 5000, 0),
    ("CW-020", "Ali Raza", "HR & Admin", "Admin Officer", 60000, 7000, 4000, 3),
    ("CW-021", "Rabia Anwar", "HR & Admin", "Receptionist", 42000, 5000, 3000, 2),
    ("CW-022", "Hassan Javed", "HR & Admin", "Office Assistant", 36000, 5000, 3000, 6),
    ("CW-023", "Daniyal Khan", "HR & Admin", "Office Assistant", 35000, 5000, 3000, 5),
    ("CW-024", "Muhammad Younas", "HR & Admin", "Cleaner", 32000, 4000, 3000, 4),
]

# Per-employee attendance behaviour: (present, late, half-day, leave, absent)
# weights. Most staff are regular; a few named patterns give the data a
# story a manager can act on (chronic late, chronic absent, long leave).
BEHAVIOUR = {
    "CW-007": (0.62, 0.30, 0.03, 0.02, 0.03),  # Tariq Aziz: chronically late
    "CW-008": (0.78, 0.05, 0.02, 0.05, 0.10),  # Shahid Iqbal: chronically absent
    "CW-005": (0.80, 0.04, 0.02, 0.12, 0.02),  # Imran Shah: long leave spell
}
DEFAULT = (0.88, 0.06, 0.02, 0.03, 0.01)
STATUSES = ["Present", "Late", "Half-day", "Leave", "Absent"]


def working_days(year=2026, month=10):
    """All Mon-Fri dates of the month as YYYY-MM-DD strings."""
    days = []
    for day in range(1, calendar.monthrange(year, month)[1] + 1):
        if calendar.weekday(year, month, day) < 5:
            days.append(f"{year}-{month:02d}-{day:02d}")
    return days


def main():
    rng = random.Random(SEED)
    days = working_days()
    assert len(days) == 22, f"expected 22 working days, got {len(days)}"

    emp_path = ROOT / "data" / "raw" / "employees_SAMPLE.csv"
    with open(emp_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["employee_id", "employee_name", "department", "designation",
                    "basic_salary_pkr", "transport_allowance_pkr",
                    "medical_allowance_pkr", "overtime_hours"])
        w.writerows(EMPLOYEES)

    att_path = ROOT / "data" / "raw" / "attendance_SAMPLE.csv"
    with open(att_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["date", "employee_id", "employee_name", "department", "status"])
        for emp_id, name, dept, *_ in EMPLOYEES:
            weights = BEHAVIOUR.get(emp_id, DEFAULT)
            for date in days:
                w.writerow([date, emp_id, name, dept, rng.choices(STATUSES, weights)[0]])

    print(f"employees: {len(EMPLOYEES)} -> {emp_path}")
    print(f"attendance rows: {len(EMPLOYEES) * len(days)} over {len(days)} days -> {att_path}")


if __name__ == "__main__":
    main()
