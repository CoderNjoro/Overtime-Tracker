# Overtime Calculator — Biometric Payroll System

A Python GUI desktop application that reads biometric check-in/out data
from Excel or PDF files and calculates overtime per employee.

---

## ⚡ Quick Start

### 1. Install Python
Download from https://python.org (version 3.10 or newer)

### 2. Install dependencies
Open a terminal / command prompt in this folder and run:

```
pip install -r requirements.txt
```

### 3. Run the app

```
python overtime_calculator.py
```

---

## 📋 How to Use

1. **Set Configuration** (left panel):
   - Regular Hours/Day: default 9 hrs (Mon–Sat)
   - Sunday Regular Hours: default 4.5 hrs (half day)
   - OT Rate Multiplier: default 1.5×
   - Hourly Rate (KES): set your base rate

2. **Load File** — click the upload area and select:
   - Excel file: `.xlsx` or `.xls`
   - PDF export: `.pdf`

3. **Click "⚡ Calculate Overtime"**

4. **View Results** in the table:
   - 🔥 Red = High OT (>20 hrs)
   - ⚡ Yellow = Mid OT (>10 hrs)
   - ✓ Green = Low OT (>0 hrs)

5. **Double-click any row** to see a day-by-day breakdown

6. **Export to Excel** — saves a full report with:
   - OT Summary sheet (all employees)
   - Daily Breakdown sheet (every day per employee)

---

## 📊 Expected File Format

The app auto-detects the correct sheet and columns.
Your biometric file should have columns similar to:

| Employee Staff ID | Full Name           | Date                | Record Type  |
|-------------------|---------------------|---------------------|--------------|
| 13333134          | JOYCE WAMBUI NJOROGE| 2026/02/23 07:41:56 | Check In/Out |

Column names are flexible — the parser looks for keywords like
"staff", "employee", "name", "date", "time".

---

## 🧮 Overtime Rules

- **Mon–Sat**: Hours worked beyond the configured regular hours = OT
- **Sunday**: Hours worked beyond half-day (default 4.5 hrs) = OT
- **OT Pay** = Total OT hours × Hourly Rate × OT Rate Multiplier
- First punch of day = Check-In, Last punch of day = Check-Out
