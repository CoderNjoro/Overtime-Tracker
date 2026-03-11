"""
Overtime Calculator — Biometric Payroll System
pip install openpyxl pandas pdfplumber xlsxwriter xlrd
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, os, re, traceback, sys
from datetime import datetime, date, timedelta
from collections import defaultdict

try:    import pandas as pd
except: raise SystemExit("Missing: pip install pandas")
try:    import openpyxl
except: raise SystemExit("Missing: pip install openpyxl")
try:    import pdfplumber
except: raise SystemExit("Missing: pip install pdfplumber")
try:    import xlsxwriter
except: raise SystemExit("Missing: pip install xlsxwriter")

import json, pathlib

def _fmt_hm(hours):
    """Convert decimal hours to e.g. 1hr 30mins"""
    if hours <= 0:
        return "—"
    total_min = round(hours * 60)
    h = total_min // 60
    m = total_min % 60
    if h > 0 and m > 0:
        return f"{h}hr {m}mins"
    elif h > 0:
        return f"{h}hr"
    else:
        return f"{m}mins"

# Config file lives next to the script/executable so it persists across runs,
# whether running as a .py file or a frozen EXE (PyInstaller, etc.)
if getattr(sys, "frozen", False):
    _base_dir = pathlib.Path(sys.executable).parent
else:
    _base_dir = pathlib.Path(__file__).parent
CONFIG_FILE = _base_dir / "config.json"

DEFAULT_CONFIG = {
    "shifts": [
        {"name":"Shift 06-14",  "start":"06:00","end":"14:00","reg":"8"},
        {"name":"Shift 14-22",  "start":"14:00","end":"22:00","reg":"8"},
        {"name":"Shift 08-17",  "start":"08:00","end":"17:00","reg":"8"},
        {"name":"Shift 20-06",  "start":"20:00","end":"06:00","reg":"8"},
        {"name":"Shift 18-02",  "start":"18:00","end":"02:00","reg":"8"},
        {"name":"Shift 19-04",  "start":"19:00","end":"04:00","reg":"8"},
    ],
    "sun_reg_hrs": "6",      # kept for compatibility but no longer used
    "late_deduct": True,
    "hourly_rate": "200",
}

def load_config():
    try:
        if CONFIG_FILE.exists():
            with open(CONFIG_FILE,"r") as f:
                data = json.load(f)
                # Merge with defaults so new keys are always present
                for k,v in DEFAULT_CONFIG.items():
                    data.setdefault(k,v)
                return data
    except Exception:
        pass
    return dict(DEFAULT_CONFIG)

def save_config(data):
    try:
        with open(CONFIG_FILE,"w") as f:
            json.dump(data, f, indent=2)
        return True
    except Exception as ex:
        return False

DARK={
    "bg":"#0f1523","surface":"#162032","card":"#1c2a3e","border":"#243448",
    "accent":"#00d4aa","adk":"#009e7f","orange":"#ff6b35","purple":"#7c3aed",
    "yellow":"#f59e0b","text":"#e2e8f0","muted":"#64748b","green":"#10b981",
    "red":"#ef4444","white":"#ffffff","blue":"#3b82f6",
}
LIGHT={
    "bg":"#f0f4f8","surface":"#ffffff","card":"#e8edf2","border":"#cbd5e1",
    "accent":"#00917a","adk":"#007a66","orange":"#e85d04","purple":"#6d28d9",
    "yellow":"#b45309","text":"#1e293b","muted":"#64748b","green":"#059669",
    "red":"#dc2626","white":"#1e293b","blue":"#2563eb",
}
C=dict(DARK)
FT=("Segoe UI",20,"bold"); FH=("Segoe UI",11,"bold"); FB=("Segoe UI",10)
FS=("Segoe UI",9); FM=("Consolas",9); FL=("Segoe UI",9,"bold")

# ══════════════════════════════════════════════════════════
#  PARSING
# ══════════════════════════════════════════════════════════
def _detect_engine(fp):
    with open(fp,"rb") as f: h=f.read(4)
    if h[:2]==b"PK": return "openpyxl"
    if h[:4]==bytes([0xD0,0xCF,0x11,0xE0]): return "xlrd"
    return "xlrd" if fp.lower().endswith(".xls") else "openpyxl"

def _parse_dt(raw):
    if isinstance(raw,datetime): return raw
    if isinstance(raw,date): return datetime(raw.year,raw.month,raw.day)
    s=str(raw).strip()
    if not s or s in("nan","NaT",""): return None
    for fmt in ("%Y/%m/%d %H:%M:%S","%Y-%m-%d %H:%M:%S","%Y/%m/%d %H:%M",
                "%Y-%m-%d %H:%M","%d/%m/%Y %H:%M:%S","%d-%m-%Y %H:%M:%S",
                "%Y-%m-%dT%H:%M:%S","%m/%d/%Y %H:%M:%S","%m/%d/%Y %H:%M"):
        try: return datetime.strptime(s,fmt)
        except: pass
    return None

def _rows_to_records(df, hr):
    """Convert dataframe to records given header row index."""
    headers=[str(v).strip().lower() for v in df.iloc[hr]]
    data=df.iloc[hr+1:].reset_index(drop=True)
    data.columns=range(len(data.columns))
    ci=cn=cd=None
    for i,h in enumerate(headers):
        if ci is None and any(k in h for k in("employee","staff","id","no")): ci=i
        if cn is None and "name" in h and "company" not in h: cn=i
        if cd is None and any(k in h for k in("date","time","datetime","punch")): cd=i
    ci=ci if ci is not None else 0
    cn=cn if cn is not None else 1
    cd=cd if cd is not None else 2
    recs=[]
    for _,row in data.iterrows():
        eid=str(row.get(ci,"")).strip()
        name=str(row.get(cn,"")).strip()
        dt=_parse_dt(row.get(cd))
        if not eid or eid in("nan","") or dt is None: continue
        if name in("nan",""): name=eid
        recs.append({"id":eid,"name":name,"dt":dt})
    return recs

def _guess_columns_and_parse(df):
    """
    Heuristically guess which columns are ID, Name, and Datetime when no header keywords match.
    Returns records list or empty list.
    """
    # Take a sample of rows (first 50) for guessing
    sample = df.head(50).astype(str).replace('nan', '')
    n_rows = len(sample)
    if n_rows == 0:
        return []

    scores = {'id': [], 'name': [], 'dt': []}
    for col in sample.columns:
        col_vals = sample[col].tolist()
        # ID: mostly numeric strings of reasonable length (4-12 digits)
        id_score = sum(1 for v in col_vals if re.fullmatch(r'\d{4,12}', v.strip()))
        scores['id'].append((col, id_score))
        # Name: contains letters and spaces, not purely numeric
        name_score = sum(1 for v in col_vals if re.search(r'[A-Za-z]{2,}', v) and not re.fullmatch(r'\d+', v.strip()))
        scores['name'].append((col, name_score))
        # Datetime: can be parsed by _parse_dt
        dt_score = sum(1 for v in col_vals if _parse_dt(v) is not None)
        scores['dt'].append((col, dt_score))

    # Pick columns with highest scores (if score > 0)
    def best_col(score_list):
        # Sort by score descending, then by column index ascending
        sorted_cols = sorted(score_list, key=lambda x: (-x[1], x[0]))
        if sorted_cols and sorted_cols[0][1] > 0:
            return sorted_cols[0][0]
        return None

    id_col = best_col(scores['id'])
    name_col = best_col(scores['name'])
    dt_col = best_col(scores['dt'])

    # Fallback to first few columns if nothing found
    if id_col is None:
        id_col = df.columns[0] if len(df.columns) > 0 else None
    if name_col is None:
        name_col = df.columns[1] if len(df.columns) > 1 else id_col
    if dt_col is None:
        dt_col = df.columns[2] if len(df.columns) > 2 else (df.columns[0] if len(df.columns) > 0 else None)

    if id_col is None or dt_col is None:
        return []  # cannot proceed

    # Check if dt_col actually contains combined date+time; if not, maybe separate date and time columns exist
    # Try to find two columns that together form datetime (date and time)
    if dt_col is not None:
        # If _parse_dt fails on most values, look for separate date/time columns
        sample_dt = df[dt_col].astype(str).head(20)
        if sum(1 for v in sample_dt if _parse_dt(v) is None) > 15:  # >75% fail
            # Look for a column with 'date' in name (if any) and another with 'time'
            date_col = time_col = None
            for col in df.columns:
                col_name = str(col).lower()
                if 'date' in col_name:
                    date_col = col
                elif 'time' in col_name:
                    time_col = col
            if date_col is not None and time_col is not None:
                # Combine date and time
                combined = []
                for idx, row in df.iterrows():
                    d = row[date_col]
                    t = row[time_col]
                    if pd.notna(d) and pd.notna(t):
                        try:
                            dt_str = f"{d} {t}"
                            parsed = _parse_dt(dt_str)
                            combined.append(parsed)
                        except:
                            combined.append(None)
                    else:
                        combined.append(None)
                df['_combined_dt'] = combined
                dt_col = '_combined_dt'

    # Build records
    records = []
    for _, row in df.iterrows():
        eid = str(row.get(id_col, '')).strip()
        name = str(row.get(name_col, '')).strip()
        dt_val = row.get(dt_col)
        dt = _parse_dt(dt_val) if dt_val is not None else None
        if not eid or eid.lower() in ('nan', '') or dt is None:
            continue
        if not name or name.lower() in ('nan', ''):
            name = eid
        records.append({"id": eid, "name": name, "dt": dt})
    return records

def parse_excel(fp):
    ext = os.path.splitext(fp)[1].lower()
    if ext == ".csv":
        return parse_csv(fp)

    engine = _detect_engine(fp)

    # Try to read as a real Excel file first
    try:
        xl = pd.ExcelFile(fp, engine=engine)
    except Exception as e:
        # If it's an .xls file and the error is due to missing xlrd or a BOF record,
        # attempt to parse as HTML (common with web reports saved as .xls)
        if engine == 'xlrd' and ext == '.xls':
            err_str = str(e)
            # Missing xlrd or typical .xls corruption/BOF error
            if 'xlrd' in err_str or 'BOF record' in err_str:
                try:
                    html_dfs = pd.read_html(fp)
                    # Try each table until we get some records
                    for df in html_dfs:
                        # Clean up: drop rows/cols that are all NaN
                        df = df.dropna(how='all').dropna(axis=1, how='all')
                        if df.empty:
                            continue
                        # First, try to find a header row with keywords
                        for i in range(min(10, len(df))):
                            row = df.iloc[i].astype(str).str.lower().str.strip()
                            has_id = any(k in x for k in ('employee', 'staff', 'id', 'no', 'code', 'emp') for x in row)
                            has_name = any(k in x for k in ('name', 'employee name', 'full name') for x in row)
                            has_date = any(k in x for k in ('date', 'time', 'datetime', 'punch', 'check', 'clock') for x in row)
                            if has_id and (has_name or has_date):
                                records = _rows_to_records(df, i)
                                if records:
                                    return records
                        # If no keyword header found, guess columns by content
                        records = _guess_columns_and_parse(df)
                        if records:
                            return records
                    # If none worked, raise error
                    raise ValueError("No valid records found in any HTML table.")
                except Exception as html_err:
                    # If HTML parsing also fails, raise a helpful error about missing xlrd
                    if 'xlrd' in err_str:
                        raise ImportError(
                            "The file appears to be a real .xls spreadsheet, but 'xlrd' is not installed.\n"
                            "Please install it with: pip install xlrd>=2.0.1"
                        ) from html_err
                    else:
                        raise e from html_err
        # If we couldn't handle it, re-raise the original exception
        raise

    # --- original Excel parsing continues unchanged ---
    sheet = xl.sheet_names[0]
    for s in xl.sheet_names:
        if any(k in s.lower() for k in ("biometric", "original", "punch", "attendance")):
            sheet = s
            break
    df = xl.parse(sheet, header=None)

    # Find header row
    for i in range(min(10, len(df))):
        v = [str(x).lower().strip() for x in df.iloc[i]]
        has_id = any(k in x for k in ("employee", "staff", "id", "no") for x in v)
        has_name = any("name" in x for x in v)
        has_time = any(k in x for k in ("date", "time", "datetime", "punch") for x in v)
        if has_id and (has_name or has_time):
            return _rows_to_records(df, i)
    return _rows_to_records(df, min(3, len(df)-2))

def parse_csv(fp):
    try: df=pd.read_csv(fp,header=None,encoding="utf-8")
    except: df=pd.read_csv(fp,header=None,encoding="latin-1")
    for i in range(min(5,len(df))):
        v=[str(x).lower().strip() for x in df.iloc[i]]
        if any(k in x for k in("employee","staff") for x in v):
            return _rows_to_records(df,i)
    return _rows_to_records(df,0)

def parse_pdf(fp):
    recs=[]
    dp=re.compile(r"\d{4}[/-]\d{2}[/-]\d{2}\s+\d{1,2}:\d{2}:\d{2}")
    ip=re.compile(r"\b\d{4,12}\b")
    with pdfplumber.open(fp) as pdf:
        for page in pdf.pages:
            tables=page.extract_tables()
            if tables:
                for tbl in tables:
                    for row in tbl:
                        if not row: continue
                        cells=[str(c or "").strip() for c in row]
                        eid=nm=dt=None
                        for cell in cells:
                            if re.match(r"^\d{4,12}$",cell) and not eid: eid=cell
                            elif dp.match(cell) and not dt: dt=_parse_dt(cell)
                            elif re.search(r"[A-Za-z]{2,}",cell) and not nm: nm=cell
                        if eid and dt: recs.append({"id":eid,"name":nm or eid,"dt":dt})
            else:
                for line in (page.extract_text() or "").splitlines():
                    dm=dp.search(line); im=ip.search(line)
                    if dm and im:
                        dt=_parse_dt(dm.group())
                        if dt: recs.append({"id":im.group(),"name":"","dt":dt})
    return recs

# ══════════════════════════════════════════════════════════
#  OVERTIME ENGINE
# ══════════════════════════════════════════════════════════

def _match_shift(cin, shifts):
    """Return the best matching shift for this check-in time.
    
    Automatically handles employees working multiple shifts in a week:
    - Finds shift whose start time is closest to check-in
    - Only matches if check-in is within 3 hours of shift start
    - Falls back to closest match if none within 3hrs (avoids None)
    """
    if not shifts:
        return None
    best=None; best_diff=1e9
    within_window=None; window_diff=1e9
    for sh in shifts:
        sh_start=cin.replace(hour=sh["start_h"],minute=sh["start_m"],second=0,microsecond=0)
        diff=abs((cin-sh_start).total_seconds())
        # Track best overall match
        if diff < best_diff:
            best_diff=diff; best=sh
        # Track best match within 3hr window
        if diff <= 3*3600 and diff < window_diff:
            window_diff=diff; within_window=sh
    # Prefer a match within 3hrs; fall back to closest if none found
    return within_window if within_window else best

def _shift_sched_out_dt(cin, sh):
    """Return the expected checkout datetime for a given shift and check-in."""
    cout_dt=cin.replace(hour=sh["end_h"],minute=sh["end_m"],second=0,microsecond=0)
    # Night shift: end time is next day
    if sh["end_h"] < sh["start_h"]:
        if cout_dt <= cin:
            cout_dt += timedelta(days=1)
    return cout_dt

def calculate_overtime(records, shifts, sun_reg_hrs, late_deduct, emp_overrides=None):
    """
    Rules:
    - Mon-Fri : OT after reg_hours (8hrs excl. lunch)
    - Saturday: day off -- any hours worked = 100% OT
    - Sunday  : OT only if any Mon-Sat work that week; otherwise all hours regular
    - Early arrivals clipped to shift start (no early bonus)
    - Late arrivals: deduct lateness from OT earned
    """
    emp_map=defaultdict(lambda:{"name":"","punches":[]})
    for rec in records:
        eid=rec["id"]
        if rec["name"] and not emp_map[eid]["name"]:
            emp_map[eid]["name"]=rec["name"]
        emp_map[eid]["punches"].append(rec["dt"])

    emp_overrides = emp_overrides or {}

    # Build a lookup: shift_name (lowercase) -> shift dict
    shift_by_name = {sh["name"].lower(): sh for sh in shifts}

    results=[]
    for eid,emp in emp_map.items():
        all_punches=sorted(emp["punches"])
        total_reg=weekday_ot=saturday_ot=sunday_ot=total_late=0.0
        days_worked=0
        breakdown=[]

        # Group punches using shifts: check-in = first punch, check-out = last punch
        # within the expected shift window. Next-day punches are only borrowed if they
        # fall before the shift's scheduled end time + 2hrs, leaving the rest for their
        # own day. This correctly handles night shifts in a 24hr punch system.
        by_day = defaultdict(list)
        for p in all_punches:
            by_day[p.date()].append(p)

        used_punch_ids = set()
        daily_pairs = []

        # Resolve override shift once for this employee
        _ov = emp_overrides.get(eid, {})
        if isinstance(_ov, dict):
            override_name = _ov.get("shift", "").lower()
            override_skip_first = _ov.get("skip_first", False)
        else:
            override_name = str(_ov).lower()
            override_skip_first = False
        override_sh = shift_by_name.get(override_name) if override_name else None

        for d in sorted(by_day.keys()):
            day_punches = [p for p in sorted(by_day[d]) if id(p) not in used_punch_ids]
            if not day_punches:
                continue

            if override_sh and override_skip_first and len(day_punches) > 1:
                # Skip first punch (prev night's checkout), use second as check-in
                candidate = day_punches[1]
                # Sanity check: candidate must be within 3hrs of shift start
                sh_start_check = datetime(d.year, d.month, d.day,
                                          override_sh["start_h"], override_sh["start_m"], 0)
                if abs((candidate - sh_start_check).total_seconds()) <= 3 * 3600:
                    cin = candidate
                else:
                    # Candidate is too far from shift start — fall back to first punch
                    cin = day_punches[0]
                best_sh = override_sh
            elif override_sh:
                cin = day_punches[0]
                best_sh = override_sh
            else:
                cin = day_punches[0]
                # Auto-detect best shift for this specific day's check-in
                # This handles employees who rotate between shifts across the week
                best_sh = _match_shift(cin, shifts)

            used_punch_ids.add(id(cin))

            # Determine cutoff based on shift end
            if best_sh:
                sh_end_dt = cin.replace(
                    hour=best_sh["end_h"], minute=best_sh["end_m"], second=0, microsecond=0)
                if best_sh["end_h"] < best_sh["start_h"]:  # night shift crosses midnight
                    sh_end_dt += timedelta(days=1)
                cutoff = sh_end_dt + timedelta(hours=2)  # 2hr grace after shift end
            else:
                cutoff = cin + timedelta(hours=14)  # fallback: 14hr window

            # All same-day punches from cin onwards
            session_punches = [cin] + [p for p in day_punches if id(p) not in used_punch_ids and p >= cin]

            # Borrow next-day punches as checkout
            next_d = d + timedelta(days=1)
            if next_d in by_day:
                next_day_punches = sorted(by_day[next_d])
                if override_sh and override_skip_first:
                    # Night shift: checkout = first punch of next day only
                    first_next = next(
                        (p for p in next_day_punches if id(p) not in used_punch_ids), None)
                    if first_next:
                        session_punches.append(first_next)
                        used_punch_ids.add(id(first_next))
                else:
                    # Normal: borrow next-day punches only within shift cutoff
                    for p in next_day_punches:
                        if id(p) not in used_punch_ids and p <= cutoff:
                            session_punches.append(p)
                            used_punch_ids.add(id(p))

            session_punches = sorted(session_punches)
            cout = session_punches[-1] if len(session_punches) > 1 else None
            if cout:
                used_punch_ids.add(id(cout))
            extra = max(0, len(session_punches) - 2)
            daily_pairs.append((cin, cout, extra))

        # First pass: collect worked dates for Sunday OT rule
        week_days_worked = defaultdict(set)
        for cin, cout, _ in daily_pairs:
            if cout:
                week_key = (cin.isocalendar()[0], cin.isocalendar()[1])
                week_days_worked[week_key].add(cin.weekday())

        # Second pass: full calculation
        for cin, cout, extra_punches in daily_pairs:

            if cout is None:
                breakdown.append({
                    "date":cin.date().strftime("%Y-%m-%d"),
                    "day":cin.date().strftime("%a"),
                    "shift":"—","sched_in":"—","sched_out":"—",
                    "check_in":cin.strftime("%H:%M:%S"),"check_out":"—",
                    "check_in_full":cin.strftime("%Y-%m-%d %H:%M:%S"),
                    "check_out_full":"—",
                    "worked":0.0,"early_min":0.0,"late_min":0.0,"ot":0.0,
                    "is_sunday":cin.weekday()==6,"is_saturday":cin.weekday()==5,
                    "note":"⚠ Unmatched check-in (no checkout found)",
                })
                continue

            worked_raw=(cout-cin).total_seconds()/3600

            if worked_raw > 24:
                breakdown.append({
                    "date":cin.date().strftime("%Y-%m-%d"),
                    "day":cin.date().strftime("%a"),
                    "shift":"—","sched_in":"—","sched_out":"—",
                    "check_in":cin.strftime("%H:%M:%S"),"check_out":"—",
                    "check_in_full":cin.strftime("%Y-%m-%d %H:%M:%S"),
                    "check_out_full":"—",
                    "worked":0.0,"early_min":0.0,"late_min":0.0,"ot":0.0,
                    "is_sunday":cin.weekday()==6,"is_saturday":cin.weekday()==5,
                    "note":"⚠ Missing checkout (gap >24h)",
                })
                continue

            if worked_raw < 0.05:
                continue

            extra_note = f" ({extra_punches} extra punches ignored)" if extra_punches > 0 else ""


            # Match shift — use per-employee override if set
            _ov2 = emp_overrides.get(eid, {})
            _ovname2 = (_ov2.get("shift","") if isinstance(_ov2, dict) else str(_ov2)).lower()
            if _ovname2 and _ovname2 in shift_by_name:
                best_sh = shift_by_name[_ovname2]
            else:
                best_sh = _match_shift(cin, shifts)
            if best_sh:
                sh_start=cin.replace(hour=best_sh["start_h"],minute=best_sh["start_m"],second=0,microsecond=0)
                reg_hrs   =best_sh["reg_hours"]
                shift_name=best_sh["name"]
                sched_in  =f"{best_sh['start_h']:02d}:{best_sh['start_m']:02d}"
                sched_out_dt=_shift_sched_out_dt(cin, best_sh)
                sched_out=sched_out_dt.strftime("%H:%M")
                if best_sh["end_h"] < best_sh["start_h"]:
                    sched_out += " (+1)"
            else:
                sh_start=cin.replace(hour=6,minute=0,second=0,microsecond=0)
                reg_hrs=8.0; shift_name="Day"
                sched_in="06:00"; sched_out="15:00"

            # Clip early arrival
            early_min=max(0.0,(sh_start-cin).total_seconds()/60)
            late_min =max(0.0,(cin-sh_start).total_seconds()/60)
            total_late+=late_min
            effective_cin=sh_start if early_min>0 else cin
            worked_eff=(cout-effective_cin).total_seconds()/3600

            shift_date=cin.date()
            weekday=shift_date.weekday()   # 0=Mon … 5=Sat, 6=Sun
            is_sunday  =(weekday==6)
            is_saturday=(weekday==5)

            # Determine OT threshold for this day
            week_key=(cin.isocalendar()[0], cin.isocalendar()[1])
            worked_days_this_week=week_days_worked[week_key]

            notes=[]
            if early_min>0:
                notes.append(f"Early {int(early_min)}m (clipped)")

            if is_sunday:
                # NEW RULE: Sunday is OT if any Mon–Sat work this week; otherwise regular
                mon_to_sat = {0,1,2,3,4,5}
                worked_mon_sat = worked_days_this_week & mon_to_sat
                if worked_mon_sat:
                    reg = 0.0
                    ot = worked_eff
                    notes.append("Sun: all OT (worked Mon–Sat this week)")
                else:
                    reg = worked_eff
                    ot = 0.0
                    notes.append("Sun: all regular (no Mon–Sat work this week)")
                sunday_ot += ot
            elif is_saturday:
                # Saturday: OT after 6hrs (5hrs work + 1hr lunch)
                sat_threshold = 6.0
                reg = min(worked_eff, sat_threshold)
                ot  = max(0.0, worked_eff - sat_threshold)
                saturday_ot += ot
                if ot > 0:
                    notes.append(f"Sat: {reg:.2f}h regular + {ot:.2f}h OT")
                else:
                    notes.append(f"Sat: {reg:.2f}h regular (no OT)")
            else:
                # Mon–Fri: OT after 9hrs
                weekday_threshold = 9.0
                reg = min(worked_eff, weekday_threshold)
                ot  = max(0.0, worked_eff - weekday_threshold)
                weekday_ot += ot

            # Late deducted from effective hours (already reflected in worked_eff)
            if late_min > 0:
                late_h = int(late_min) // 60
                late_m = int(late_min) % 60
                late_str = f"{late_h}h {late_m}m" if late_h > 0 else f"{late_m}m"
                notes.append(f"Late {late_str}")

            checkout_display=cout.strftime("%H:%M:%S")
            if cout.date()>cin.date():
                checkout_display+=f" ({cout.strftime('%a')} +1)"

            total_reg+=reg
            days_worked+=1
            breakdown.append({
                "date":          shift_date.strftime("%Y-%m-%d"),
                "day":           shift_date.strftime("%a"),
                "shift":         shift_name,
                "sched_in":      sched_in,
                "sched_out":     sched_out,
                "check_in":      cin.strftime("%H:%M:%S"),
                "check_out":     checkout_display,
                "check_in_full": cin.strftime("%Y-%m-%d %H:%M:%S"),
                "check_out_full":cout.strftime("%Y-%m-%d %H:%M:%S"),
                "worked":        round(worked_eff,2),
                "actual_worked": round(worked_raw,2),
                "early_min":     round(early_min,1),
                "late_min":      round(late_min,1),
                "ot":            round(ot,2),
                "is_sunday":     is_sunday,
                "is_saturday":   is_saturday,
                "note":          "  |  ".join(notes),
            })

        total_ot=weekday_ot+saturday_ot+sunday_ot
        results.append({
            "id":eid,"name":emp["name"] or eid,
            "days_worked":days_worked,
            "regular_hours":round(total_reg,2),
            "weekday_ot":round(weekday_ot,2),
            "saturday_ot":round(saturday_ot,2),
            "sunday_ot":round(sunday_ot,2),
            "total_ot":round(total_ot,2),
            "total_late_min":round(total_late,1),
            "ot_pay":0.0,
            "breakdown":breakdown,
        })

    results.sort(key=lambda r:r["total_ot"],reverse=True)
    return results

# ══════════════════════════════════════════════════════════
#  EXPORT
# ══════════════════════════════════════════════════════════
def export_to_excel(results,filepath):
    with xlsxwriter.Workbook(filepath) as wb:
        hdr=wb.add_format({"bold":True,"bg_color":"#0f1523","font_color":"#00d4aa","border":1,"align":"center"})
        ttl=wb.add_format({"bold":True,"font_size":13,"font_color":"#0f1523","bg_color":"#00d4aa","align":"center"})
        num=wb.add_format({"num_format":"0.00","align":"right"})
        kes=wb.add_format({"num_format":'"KES "#,##0.00',"align":"right"})
        otn=wb.add_format({"num_format":"0.00","align":"right","font_color":"#ff6b35","bold":True})
        tot=wb.add_format({"bold":True,"bg_color":"#162032","font_color":"#00d4aa","border":1})
        totn=wb.add_format({"bold":True,"bg_color":"#162032","font_color":"#00d4aa","num_format":"0.00","align":"right"})
        totk=wb.add_format({"bold":True,"bg_color":"#162032","font_color":"#00d4aa","num_format":'"KES "#,##0.00',"align":"right"})

        ws=wb.add_worksheet("OT Summary")
        ws.set_column("A:A",14); ws.set_column("B:B",28); ws.set_column("C:I",15)
        ws.merge_range("A1:I1","OVERTIME SUMMARY REPORT",ttl)
        ws.write("A2",f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        for c,h in enumerate(["Staff ID","Full Name","Days","Regular Hrs","Mon–Fri OT","Sat OT","Sun OT","Total OT","Late (min)","OT Pay (KES)"]):
            ws.write(3,c,h,hdr)
        for r,e in enumerate(results,4):
            ws.write(r,0,e["id"]); ws.write(r,1,e["name"])
            ws.write(r,2,e["days_worked"],num); ws.write(r,3,e["regular_hours"],num)
            ws.write(r,4,e["weekday_ot"],otn); ws.write(r,5,e.get("saturday_ot",0),otn)
            ws.write(r,6,e["sunday_ot"],otn);  ws.write(r,7,e["total_ot"],otn)
            ws.write(r,8,e["total_late_min"],num); ws.write(r,9,e["ot_pay"],kes)
        tr=len(results)+4
        ws.write(tr,0,"TOTAL",tot); ws.write(tr,1,f"{len(results)} employees",tot)
        for c,k in enumerate(["days_worked","regular_hours","weekday_ot","saturday_ot","sunday_ot","total_ot","total_late_min"],2):
            ws.write(tr,c,round(sum(e.get(k,0) for e in results),2),totn)
        ws.write(tr,9,round(sum(e["ot_pay"] for e in results),2),totk)

        ws2=wb.add_worksheet("Daily Breakdown")
        ws2.set_column("A:B",14); ws2.set_column("C:C",26); ws2.set_column("D:M",13)
        for c,h in enumerate(["Staff ID","Name","Full Name","Date","Day","Shift","Sched In","Sched Out",
                               "Check-In","Check-Out","Hrs Worked","Late (min)","OT Hrs","Note"]):
            ws2.write(0,c,h,hdr)
        row=1
        for e in results:
            for d in e["breakdown"]:
                ws2.write(row,0,e["id"]); ws2.write(row,1,e["name"]); ws2.write(row,2,e["name"])
                ws2.write(row,3,d["date"]); ws2.write(row,4,d["day"]); ws2.write(row,5,d["shift"])
                ws2.write(row,6,d["sched_in"]); ws2.write(row,7,d["sched_out"])
                ws2.write(row,8,d["check_in"]); ws2.write(row,9,d["check_out"])
                ws2.write(row,10,d["worked"],num); ws2.write(row,11,d["late_min"],num)
                ws2.write(row,12,d["ot"],otn); ws2.write(row,13,d["note"])
                row+=1

# ══════════════════════════════════════════════════════════
#  TREEVIEW STYLE
# ══════════════════════════════════════════════════════════
def _setup_style():
    s=ttk.Style(); s.theme_use("clam")
    s.configure("OT.Treeview",background=C["card"],foreground=C["text"],
                fieldbackground=C["card"],rowheight=28,font=FB)
    s.configure("OT.Treeview.Heading",background=C["surface"],
                foreground=C["accent"],font=FL,relief="flat")
    s.map("OT.Treeview",background=[("selected",C["border"])],
          foreground=[("selected",C["white"])])

def _make_tree(parent, cols_def):
    """cols_def = list of (id, heading, width, anchor). Returns (tree, vsb, hsb)."""
    cols=[c[0] for c in cols_def]
    tree=ttk.Treeview(parent,columns=cols,show="headings",style="OT.Treeview",selectmode="browse")
    for cid,htxt,w,anc in cols_def:
        tree.heading(cid,text=htxt)
        tree.column(cid,width=w,anchor=anc,stretch=tk.YES)
    vsb=ttk.Scrollbar(parent,orient=tk.VERTICAL,  command=tree.yview)
    hsb=ttk.Scrollbar(parent,orient=tk.HORIZONTAL,command=tree.xview)
    tree.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
    tree.grid(row=0,column=0,sticky="nsew")
    vsb.grid(row=0,column=1,sticky="ns")
    hsb.grid(row=1,column=0,sticky="ew")
    parent.grid_rowconfigure(0,weight=1); parent.grid_columnconfigure(0,weight=1)
    return tree,vsb,hsb

# ══════════════════════════════════════════════════════════
#  MAIN APP
# ══════════════════════════════════════════════════════════
class OvertimeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Overtime Calculator — Biometric Payroll")
        self.geometry("1240x860"); self.minsize(980,660)
        self.configure(bg=C["bg"])
        _setup_style()
        self._records=[]; self._results=[]; self._filepath=""
        self._dark_mode=True
        self._emp_overrides={}  # staff_id -> shift_name
        self._overrides=[]  # populated by _build_emp_overrides
        self._sort_col=""; self._sort_rev=True; self._tree=None
        self._vars={}; self._shifts=[]
        self._config=load_config()          # ← load saved config
        self._build_ui()
        self.update_idletasks()
        w,h=self.winfo_width(),self.winfo_height()
        self.geometry(f"{w}x{h}+{(self.winfo_screenwidth()-w)//2}+{(self.winfo_screenheight()-h)//2}")

    # ── Top bar ───────────────────────────────────────────
    def _build_ui(self):
        top=tk.Frame(self,bg=C["surface"],height=58); top.pack(fill=tk.X); top.pack_propagate(False)
        tk.Label(top,text="OT",bg=C["accent"],fg=C["bg"],font=("Segoe UI",16,"bold"),
                 width=3,padx=4).pack(side=tk.LEFT,padx=14,pady=10)
        tk.Label(top,text="Overtime Calculator",font=FT,bg=C["surface"],fg=C["white"]).pack(side=tk.LEFT,pady=8)
        tk.Label(top,text="Biometric Payroll System",font=FS,bg=C["surface"],fg=C["muted"]).pack(side=tk.LEFT,padx=8,pady=18)
        self._theme_btn=tk.Button(top,text="☀ Light",font=FS,bg=C["border"],fg=C["text"],
            relief=tk.FLAT,bd=0,cursor="hand2",padx=10,pady=4,
            command=self._toggle_theme)
        self._theme_btn.pack(side=tk.RIGHT,padx=14,pady=14)

        main=tk.Frame(self,bg=C["bg"]); main.pack(fill=tk.BOTH,expand=True,padx=14,pady=10)

        # Left: fixed-width scrollable sidebar
        lo=tk.Frame(main,bg=C["bg"],width=330); lo.pack(side=tk.LEFT,fill=tk.Y,padx=(0,12)); lo.pack_propagate(False)

        # Action buttons pinned to bottom of left
        bf=tk.Frame(lo,bg=C["bg"]); bf.pack(side=tk.BOTTOM,fill=tk.X,pady=(6,0))
        self._calc_btn=tk.Button(bf,text="⚡  CALCULATE OVERTIME",font=FH,
            bg=C["accent"],fg=C["bg"],activebackground=C["adk"],
            relief=tk.FLAT,bd=0,cursor="hand2",command=self._run_calc,state=tk.DISABLED)
        self._calc_btn.pack(fill=tk.X,ipady=11,pady=(0,4))
        self._export_btn=tk.Button(bf,text="↓  Export to Excel",font=FB,
            bg=C["surface"],fg=C["accent"],activebackground=C["border"],
            relief=tk.FLAT,bd=0,cursor="hand2",command=self._export,state=tk.DISABLED)
        self._export_btn.pack(fill=tk.X,ipady=8)

        # Scrollable canvas + visible scrollbar
        scroll_area=tk.Frame(lo,bg=C["bg"]); scroll_area.pack(fill=tk.BOTH,expand=True)
        lc=tk.Canvas(scroll_area,bg=C["bg"],highlightthickness=0)
        vsb=ttk.Scrollbar(scroll_area,orient=tk.VERTICAL,command=lc.yview)
        lc.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT,fill=tk.Y)
        lc.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)

        lf=tk.Frame(lc,bg=C["bg"])
        # Keep inner frame width = canvas width
        def _on_canvas_resize(e):
            lc.itemconfig(win_id,width=e.width)
        win_id=lc.create_window((0,0),window=lf,anchor=tk.NW)
        lc.bind("<Configure>",_on_canvas_resize)
        lf.bind("<Configure>",lambda e:lc.configure(scrollregion=lc.bbox("all")))

        # Mouse-wheel scrolling (Windows & Mac)
        def _on_wheel(ev):
            lc.yview_scroll(int(-1*(ev.delta/120)),"units")
        lc.bind("<Enter>",lambda e:lc.bind_all("<MouseWheel>",_on_wheel))
        lc.bind("<Leave>",lambda e:lc.unbind_all("<MouseWheel>"))

        # Right panel with tabs
        right=tk.Frame(main,bg=C["bg"]); right.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)

        nb=ttk.Notebook(right); nb.pack(fill=tk.BOTH,expand=True)
        self._nb=nb
        tab1=tk.Frame(nb,bg=C["bg"]); nb.add(tab1,text="  📋 Results  ")
        tab2=tk.Frame(nb,bg=C["bg"]); nb.add(tab2,text="  📊 Charts  ")

        self._build_shifts(lf)
        self._build_general(lf)
        self._build_upload(lf)
        self._build_emp_overrides(lf)
        self._build_results(tab1)
        self._build_charts(tab2)

    # ── Shift rows ────────────────────────────────────────
    def _build_shifts(self,parent):
        outer=tk.Frame(parent,bg=C["bg"]); outer.pack(fill=tk.X,pady=(4,6))

        hr=tk.Frame(outer,bg=C["bg"]); hr.pack(fill=tk.X)
        tk.Label(hr,text="⏰  Shift Configuration",font=FL,bg=C["bg"],fg=C["accent"]).pack(side=tk.LEFT,pady=(0,4))

        # Save button — prominent, always visible
        self._save_btn=tk.Button(hr,text="💾 Save",font=FS,bg=C["green"],fg=C["bg"],
                  relief=tk.FLAT,bd=0,cursor="hand2",padx=8,pady=2,
                  command=self._save_config)
        self._save_btn.pack(side=tk.RIGHT,padx=(4,0))
        tk.Button(hr,text="+ Add",font=FS,bg=C["purple"],fg=C["white"],
                  relief=tk.FLAT,bd=0,cursor="hand2",padx=8,pady=2,
                  command=self._add_shift).pack(side=tk.RIGHT)

        # Save status label
        self._save_lbl=tk.Label(outer,text="",font=FS,bg=C["bg"],fg=C["green"])
        self._save_lbl.pack(anchor=tk.W)

        # Column header
        hrow=tk.Frame(outer,bg=C["surface"]); hrow.pack(fill=tk.X,pady=(2,2))
        for txt,w in [("Shift Name",10),("Start",5),("End",5),("Reg Hrs",6),(" ",2)]:
            tk.Label(hrow,text=txt,font=("Segoe UI",8,"bold"),bg=C["surface"],
                     fg=C["muted"],width=w).pack(side=tk.LEFT,padx=3,pady=3)

        self._shift_container=tk.Frame(outer,bg=C["bg"]); self._shift_container.pack(fill=tk.X)

        # Load saved shifts (or defaults)
        for sh in self._config.get("shifts", DEFAULT_CONFIG["shifts"]):
            self._add_shift(sh["name"], sh["start"], sh["end"], sh["reg"])

    def _save_config(self):
        """Persist current shifts + settings to config.json."""
        cfg = {
            "shifts": [
                {"name":sv["name"].get().strip(),
                 "start":sv["start"].get().strip(),
                 "end":sv["end"].get().strip(),
                 "reg":sv["reg"].get().strip()}
                for sv in self._shifts
            ],
            "sun_reg_hrs": self._vars.get("sun_reg_hrs", tk.StringVar(value="6")).get(),
            "late_deduct": bool(self._vars.get("late_deduct", tk.BooleanVar(value=True)).get()),
            "hourly_rate": self._vars.get("hourly_rate", tk.StringVar(value="200")).get(),
            "emp_overrides": {eid: {"shift": v["shift"], "skip_first": v["skip_first"]} for eid, v in self._get_emp_overrides().items()},
        }
        if save_config(cfg):
            self._config = cfg
            self._save_lbl.configure(text="✓ Shifts saved", fg=C["green"])
            self.after(2500, lambda: self._save_lbl.configure(text=""))
            if hasattr(self, "_override_save_lbl"):
                self._override_save_lbl.configure(text="✓ Overrides saved", fg=C["green"])
                self.after(2500, lambda: self._override_save_lbl.configure(text=""))
        else:
            self._save_lbl.configure(text="⚠ Save failed", fg=C["red"])
            self.after(2500, lambda: self._save_lbl.configure(text=""))

    def _add_shift(self,name="",start="",end="",reg="8"):
        row=tk.Frame(self._shift_container,bg=C["card"],
                     highlightthickness=1,highlightbackground=C["border"])
        row.pack(fill=tk.X,pady=2)
        vn=tk.StringVar(value=name); vs=tk.StringVar(value=start)
        ve=tk.StringVar(value=end);  vr=tk.StringVar(value=reg)
        def ent(var,w):
            e=tk.Entry(row,textvariable=var,font=FM,width=w,
                       bg=C["surface"],fg=C["text"],insertbackground=C["accent"],relief=tk.FLAT,bd=0)
            e.pack(side=tk.LEFT,padx=2,pady=5,ipady=4)
        ent(vn,10); ent(vs,5); ent(ve,5); ent(vr,5)
        sv={"frame":row,"name":vn,"start":vs,"end":ve,"reg":vr}
        self._shifts.append(sv)
        def remove():
            self._shifts.remove(sv)
            row.destroy()
            self._save_config()   # auto-save when a shift is deleted
        tk.Button(row,text="✕",font=("Segoe UI",9),bg=C["card"],fg=C["red"],
                  relief=tk.FLAT,bd=0,cursor="hand2",command=remove).pack(side=tk.LEFT,padx=3)

    def _get_shifts(self):
        out=[]
        for sv in self._shifts:
            try:
                nm=sv["name"].get().strip() or "Shift"
                sh,sm=[int(x) for x in sv["start"].get().strip().split(":")]
                eh,em=[int(x) for x in sv["end"].get().strip().split(":")]
                rg=float(sv["reg"].get().strip())
                out.append({"name":nm,"start_h":sh,"start_m":sm,"end_h":eh,"end_m":em,"reg_hours":rg})
            except: pass
        return out

    # ── General settings ──────────────────────────────────
    def _build_general(self,parent):
        card=self._card(parent,"⚙  General Settings")
        tk.Label(card,text="Sunday Regular Hours (OT kicks in after)",
                 font=FS,bg=C["card"],fg=C["muted"]).pack(anchor=tk.W,pady=(8,2))
        v=tk.StringVar(value=self._config.get("sun_reg_hrs","6")); self._vars["sun_reg_hrs"]=v
        e=tk.Entry(card,textvariable=v,font=FB,bg=C["surface"],fg=C["text"],
                   insertbackground=C["accent"],relief=tk.FLAT,bd=0)
        e.pack(fill=tk.X,ipady=7,padx=2)
        sep=tk.Frame(card,bg=C["border"],height=1); sep.pack(fill=tk.X,padx=2)
        e.bind("<FocusIn>", lambda ev,s=sep:s.configure(bg=C["accent"]))
        e.bind("<FocusOut>",lambda ev,s=sep:s.configure(bg=C["border"]))
        tk.Label(card,text="Mon–Fri: OT after 9hrs (excl. lunch)",
                 font=FS,bg=C["card"],fg=C["muted"],wraplength=270,justify=tk.LEFT
                 ).pack(anchor=tk.W,pady=(6,2))
        tk.Label(card,text="Saturday: OT after 6hrs (5hrs work + 1hr lunch)",
                 font=FS,bg=C["card"],fg=C["orange"],wraplength=270,justify=tk.LEFT
                 ).pack(anchor=tk.W,pady=(0,2))
        tk.Label(card,text="Sunday: OT only if any Mon–Sat work that week (otherwise regular)",
                 font=FS,bg=C["card"],fg=C["yellow"],wraplength=270,justify=tk.LEFT
                 ).pack(anchor=tk.W,pady=(0,4))
        ld=self._config.get("late_deduct",True)
        self._vars["late_deduct"]=tk.BooleanVar(value=ld)
        tk.Checkbutton(card,text="Deduct late arrival minutes from OT",
                       variable=self._vars["late_deduct"],bg=C["card"],fg=C["text"],
                       selectcolor=C["surface"],activebackground=C["card"],
                       font=FS,anchor=tk.W).pack(fill=tk.X,padx=2,pady=(8,4))

    # ── File upload ───────────────────────────────────────
    def _build_upload(self,parent):
        card=self._card(parent,"📂  Load Biometric File")
        drop=tk.Frame(card,bg=C["surface"],highlightthickness=2,highlightbackground=C["border"])
        drop.pack(fill=tk.X,pady=8)
        tk.Label(drop,text="📄",font=("Segoe UI",20),bg=C["surface"],fg=C["muted"]).pack(pady=(8,2))
        tk.Label(drop,text="Click to browse file",font=FH,bg=C["surface"],fg=C["text"]).pack()
        tk.Label(drop,text="Excel (.xlsx .xls) or PDF",font=FS,bg=C["surface"],fg=C["muted"]).pack(pady=(2,8))
        for w in [drop]+drop.winfo_children():
            w.bind("<Button-1>",lambda e:self._browse())
        drop.bind("<Enter>",lambda e:drop.configure(highlightbackground=C["accent"]))
        drop.bind("<Leave>",lambda e:drop.configure(highlightbackground=C["border"]))
        self._file_lbl=tk.Label(card,text="No file selected",font=FM,
                                 bg=C["card"],fg=C["muted"],wraplength=275,justify=tk.LEFT)
        self._file_lbl.pack(anchor=tk.W,pady=4)
        self._status_lbl=tk.Label(card,text="",font=FS,bg=C["card"],
                                   fg=C["green"],wraplength=275,justify=tk.LEFT)
        self._status_lbl.pack(anchor=tk.W,pady=2)

    # ── Employee shift overrides ─────────────────────────
    def _build_emp_overrides(self, parent):
        card = self._card(parent, "👤  Employee Shift Overrides")
        tk.Label(card, text="Assign a specific shift to an employee by Staff ID. Useful for night-shift staff whose punches look like day shifts.",
                 font=FS, bg=C["card"], fg=C["muted"], wraplength=270, justify=tk.LEFT).pack(anchor=tk.W, pady=(0,6))

        # Column headers
        hrow = tk.Frame(card, bg=C["surface"]); hrow.pack(fill=tk.X, pady=(0,2))
        for txt, w in [("Staff ID", 10), ("Shift Name", 10), (" ", 2)]:
            tk.Label(hrow, text=txt, font=("Segoe UI",8,"bold"), bg=C["surface"],
                     fg=C["muted"], width=w).pack(side=tk.LEFT, padx=3, pady=3)

        self._override_container = tk.Frame(card, bg=C["card"])
        self._override_container.pack(fill=tk.X)
        self._overrides = []  # list of {"frame","id","shift"}

        # Load saved overrides
        for eid, val in self._config.get("emp_overrides", {}).items():
            if isinstance(val, dict):
                self._add_override(eid, val.get("shift",""), val.get("skip_first", False))
            else:
                self._add_override(eid, val, False)

        bf2 = tk.Frame(card, bg=C["card"]); bf2.pack(anchor=tk.W, pady=(6,0))
        tk.Button(bf2, text="+ Add Override", font=FS, bg=C["purple"], fg=C["white"],
                  relief=tk.FLAT, bd=0, cursor="hand2", padx=8, pady=2,
                  command=self._add_override).pack(side=tk.LEFT, padx=(0,6))
        tk.Button(bf2, text="💾 Save Overrides", font=FS, bg=C["green"], fg=C["bg"],
                  relief=tk.FLAT, bd=0, cursor="hand2", padx=8, pady=2,
                  command=self._save_config).pack(side=tk.LEFT)
        self._override_save_lbl = tk.Label(card, text="", font=FS, bg=C["card"], fg=C["green"])
        self._override_save_lbl.pack(anchor=tk.W)

    def _add_override(self, eid="", shift_name="", skip_first=False):
        row = tk.Frame(self._override_container, bg=C["card"],
                       highlightthickness=1, highlightbackground=C["border"])
        row.pack(fill=tk.X, pady=2)
        vid = tk.StringVar(value=eid)
        vsh = tk.StringVar(value=shift_name)
        vsk = tk.BooleanVar(value=skip_first)

        # Staff ID entry
        tk.Entry(row, textvariable=vid, font=FM, width=10,
                 bg=C["surface"], fg=C["text"], insertbackground=C["accent"],
                 relief=tk.FLAT, bd=0).pack(side=tk.LEFT, padx=2, pady=5, ipady=4)

        # Shift dropdown
        shift_names = [sv["name"].get().strip() for sv in self._shifts if sv["name"].get().strip()]
        if not shift_names:
            shift_names = ["(no shifts defined)"]
        if shift_name and shift_name not in shift_names:
            shift_names.insert(0, shift_name)
        cb = ttk.Combobox(row, textvariable=vsh, values=shift_names, font=FM,
                          width=14, state="readonly")
        cb.pack(side=tk.LEFT, padx=2, pady=5)
        if vsh.get() in shift_names:
            cb.set(vsh.get())
        elif shift_names:
            cb.set(shift_names[0])

        # Skip first punch checkbox
        tk.Checkbutton(row, text="Skip 1st", variable=vsk,
                       bg=C["card"], fg=C["yellow"], selectcolor=C["surface"],
                       activebackground=C["card"], font=FS).pack(side=tk.LEFT, padx=4)

        ov = {"frame": row, "id": vid, "shift": vsh, "skip_first": vsk}
        self._overrides.append(ov)
        def remove():
            self._overrides.remove(ov)
            row.destroy()
            self._save_config()
        tk.Button(row, text="✕", font=("Segoe UI",9), bg=C["card"], fg=C["red"],
                  relief=tk.FLAT, bd=0, cursor="hand2", command=remove).pack(side=tk.LEFT, padx=3)

    def _get_emp_overrides(self):
        out = {}
        for ov in getattr(self, "_overrides", []):
            try:
                eid = ov["id"].get().strip()
                sname = ov["shift"].get().strip()
                skip = bool(ov.get("skip_first", tk.BooleanVar()).get())
                if eid and sname:
                    out[eid] = {"shift": sname, "skip_first": skip}
            except Exception:
                pass
        return out

    # ── Results panel ─────────────────────────────────────
    def _build_results(self,parent):
        # Rate bar — hourly rate only, no multiplier
        rc=tk.Frame(parent,bg=C["card"],highlightthickness=1,highlightbackground=C["border"])
        rc.pack(fill=tk.X,pady=(0,10))
        tk.Label(rc,text="💰  OT Pay — enter hourly rate then click Apply",
                 font=FL,bg=C["card"],fg=C["accent"]).pack(anchor=tk.W,padx=14,pady=(8,4))
        rr=tk.Frame(rc,bg=C["card"]); rr.pack(fill=tk.X,padx=14,pady=(0,10))
        f=tk.Frame(rr,bg=C["card"]); f.pack(side=tk.LEFT,padx=(0,18))
        tk.Label(f,text="Hourly Rate (KES)",font=FS,bg=C["card"],fg=C["muted"]).pack(anchor=tk.W)
        var=tk.StringVar(value=self._config.get("hourly_rate","200")); self._vars["hourly_rate"]=var
        e=tk.Entry(f,textvariable=var,font=FB,width=12,
                   bg=C["surface"],fg=C["text"],insertbackground=C["accent"],relief=tk.FLAT,bd=0)
        e.pack(fill=tk.X,ipady=6)
        sep=tk.Frame(f,bg=C["border"],height=1); sep.pack(fill=tk.X)
        e.bind("<FocusIn>", lambda ev,s=sep:s.configure(bg=C["accent"]))
        e.bind("<FocusOut>",lambda ev,s=sep:s.configure(bg=C["border"]))
        tk.Button(rr,text="▶  Apply Rate",font=FL,bg=C["purple"],fg=C["white"],
                  activebackground="#5b21b6",relief=tk.FLAT,bd=0,cursor="hand2",
                  command=lambda:self._apply_rate()).pack(side=tk.LEFT,ipady=8,ipadx=14,pady=(14,0))

        # Summary cards
        cf=tk.Frame(parent,bg=C["bg"]); cf.pack(fill=tk.X,pady=(0,8))
        self._sv={}
        for title,key,color in [("Employees","emp",C["accent"]),("Total OT Hrs","ot",C["orange"]),
                                  ("Total OT Pay","pay",C["purple"]),("Highest OT","top",C["yellow"])]:
            f=tk.Frame(cf,bg=C["card"],highlightthickness=2,highlightbackground=color)
            f.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=4)
            tk.Label(f,text=title,font=FS,bg=C["card"],fg=C["muted"]).pack(anchor=tk.W,padx=12,pady=(8,2))
            v=tk.StringVar(value="—"); self._sv[key]=v
            tk.Label(f,textvariable=v,font=("Segoe UI",16,"bold"),bg=C["card"],fg=color).pack(anchor=tk.W,padx=12,pady=(0,8))

        # Search bar
        sf=tk.Frame(parent,bg=C["surface"],highlightthickness=1,highlightbackground=C["border"])
        sf.pack(fill=tk.X,pady=(0,6))
        tk.Label(sf,text="⌕",font=("Segoe UI",13),bg=C["surface"],fg=C["muted"]).pack(side=tk.LEFT,padx=10)
        self._search=tk.StringVar()
        self._search.trace_add("write",lambda *_:self._filter() if self._tree else None)
        se=tk.Entry(sf,textvariable=self._search,font=FB,bg=C["surface"],
                    fg=C["text"],insertbackground=C["accent"],relief=tk.FLAT,bd=0)
        se.pack(fill=tk.X,ipady=8,padx=(0,10))
        ph="Search by name or staff ID..."; se.insert(0,ph)
        se.bind("<FocusIn>", lambda e:se.delete(0,tk.END) if se.get()==ph else None)
        se.bind("<FocusOut>",lambda e:se.insert(0,ph) if not se.get() else None)

        # Main table — IMPORTANT: tree parent = tf frame, not this parent
        tf=tk.Frame(parent,bg=C["bg"]); tf.pack(fill=tk.BOTH,expand=True)
        self._tree,_,_=_make_tree(tf,[
            ("id","Staff ID",100,tk.W),("name","Full Name",175,tk.W),
            ("days","Days",48,tk.CENTER),("regular","Reg Hrs",80,tk.CENTER),
            ("wot","Mon–Fri OT",82,tk.CENTER),("satok","Sat OT",72,tk.CENTER),
            ("sot","Sun OT",72,tk.CENTER),("tot","Total OT",82,tk.CENTER),
            ("late","Late(m)",68,tk.CENTER),("pay","OT Pay (KES)",115,tk.E),
        ])
        for tag,fg in [("high",C["orange"]),("mid",C["yellow"]),("low",C["green"]),("none",C["muted"])]:
            self._tree.tag_configure(tag,foreground=fg)
        self._tree.tag_configure("alt",background="#1a2436")
        self._tree.bind("<Double-1>",self._show_breakdown)

        tk.Label(parent,text="💡 Double-click a row to see full daily breakdown",
                 font=FS,bg=C["bg"],fg=C["muted"]).pack(anchor=tk.W,pady=4)

    def _card(self,parent,title):
        outer=tk.Frame(parent,bg=C["bg"],pady=4); outer.pack(fill=tk.X,pady=4)
        tk.Label(outer,text=title,font=FL,bg=C["bg"],fg=C["accent"]).pack(anchor=tk.W,pady=(0,4))
        inner=tk.Frame(outer,bg=C["card"],highlightthickness=1,
                       highlightbackground=C["border"],padx=14,pady=10)
        inner.pack(fill=tk.X); return inner

    # ── Theme toggle ─────────────────────────────────────
    def _toggle_theme(self):
        global C
        self._dark_mode = not self._dark_mode
        C.update(DARK if self._dark_mode else LIGHT)
        self._theme_btn.configure(
            text="☀ Light" if self._dark_mode else "🌙 Dark",
            bg=C["border"], fg=C["text"])
        self._apply_theme(self)

    def _apply_theme(self, widget):
        cls = widget.winfo_class()
        try:
            if cls in ("Frame","Labelframe"):
                widget.configure(bg=C["bg"] if widget.master and widget.master.winfo_class()=="Tk" else C["card"])
        except: pass
        try:
            if cls=="Label":
                widget.configure(bg=widget.master.cget("bg") if widget.master else C["bg"])
        except: pass
        for child in widget.winfo_children():
            self._apply_theme(child)
        # Refresh styles
        _setup_style()
        if self._tree:
            self._tree.configure(style="OT.Treeview")
        if self._results:
            self._populate(self._results)

    # ── Charts panel ─────────────────────────────────────
    def _build_charts(self, parent):
        self._chart_frame = parent
        tk.Label(parent, text="Run a calculation to see charts",
                 font=FH, bg=C["bg"], fg=C["muted"]).pack(expand=True)

    def _refresh_charts(self):
        if not self._results: return
        for w in self._chart_frame.winfo_children():
            w.destroy()

        cf = tk.Frame(self._chart_frame, bg=C["bg"])
        cf.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ── Top OT earners (horizontal bars) ──────────────
        tk.Label(cf, text="🏆 Top 10 OT Earners", font=FH,
                 bg=C["bg"], fg=C["accent"]).pack(anchor=tk.W, pady=(0,4))
        top10 = sorted(self._results, key=lambda r: r["total_ot"], reverse=True)[:10]
        if top10:
            max_ot = max(r["total_ot"] for r in top10) or 1
            bar_frame = tk.Frame(cf, bg=C["bg"]); bar_frame.pack(fill=tk.X, pady=(0,16))
            for i, r in enumerate(top10):
                row = tk.Frame(bar_frame, bg=C["bg"]); row.pack(fill=tk.X, pady=2)
                name = r["name"][:22] if len(r["name"])>22 else r["name"]
                tk.Label(row, text=name, font=FS, bg=C["bg"], fg=C["text"],
                         width=24, anchor=tk.W).pack(side=tk.LEFT)
                bar_bg = tk.Frame(row, bg=C["surface"], height=18)
                bar_bg.pack(side=tk.LEFT, fill=tk.X, expand=True)
                bar_bg.pack_propagate(False)
                pct = r["total_ot"] / max_ot
                fill_color = C["orange"] if i == 0 else C["accent"]
                tk.Frame(bar_bg, bg=fill_color, height=18).place(
                    relx=0, rely=0, relwidth=pct, relheight=1)
                h = int(r["total_ot"]); m = round((r["total_ot"]-h)*60)
                tk.Label(row, text=f"{h}h{m:02d}m", font=FS, bg=C["bg"],
                         fg=C["yellow"], width=8).pack(side=tk.LEFT)

        # ── OT by day of week (vertical bars) ─────────────
        tk.Label(cf, text="📅 OT Hours by Day of Week", font=FH,
                 bg=C["bg"], fg=C["accent"]).pack(anchor=tk.W, pady=(8,4))
        day_ot = defaultdict(float)
        for r in self._results:
            for d in r.get("breakdown", []):
                if not d.get("note","").startswith("⚠"):
                    day_ot[d.get("day","")] += d.get("ot", 0)
        days_order = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
        day_vals   = [day_ot.get(d, 0) for d in days_order]
        day_clrs   = [C["blue"]]*5 + [C["green"], C["purple"]]
        max_day    = max(day_vals) or 1
        BAR_MAX    = 80
        dow_frame  = tk.Frame(cf, bg=C["bg"]); dow_frame.pack(fill=tk.X, pady=(0,16))
        for day, val, color in zip(days_order, day_vals, day_clrs):
            col = tk.Frame(dow_frame, bg=C["bg"])
            col.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=3)
            bar_h = max(int((val / max_day) * BAR_MAX), 3) if val > 0 else 2
            tk.Frame(col, bg=C["bg"], height=BAR_MAX - bar_h).pack(fill=tk.X)
            tk.Frame(col, bg=color, height=bar_h).pack(fill=tk.X)
            h=int(val); m=round((val-h)*60)
            tk.Label(col, text=f"{h}h{m:02d}m" if val>0 else "—",
                     font=("Segoe UI",7), bg=C["bg"], fg=color).pack()
            # Day label on contrasting card chip
            chip = tk.Frame(col, bg=C["surface"])
            chip.pack(fill=tk.X)
            tk.Label(chip, text=day, font=("Segoe UI",8,"bold"),
                     bg=C["surface"], fg=C["text"]).pack(pady=2)

        # ── OT type breakdown (horizontal bars) ───────────
        tk.Label(cf, text="📊 OT Type Breakdown", font=FH,
                 bg=C["bg"], fg=C["accent"]).pack(anchor=tk.W, pady=(8,4))
        wot   = sum(r["weekday_ot"]         for r in self._results)
        satok = sum(r.get("saturday_ot", 0) for r in self._results)
        sunot = sum(r["sunday_ot"]           for r in self._results)
        total = wot + satok + sunot or 1
        pie_frame = tk.Frame(cf, bg=C["bg"]); pie_frame.pack(fill=tk.X)
        for label, val, color in [("Mon–Fri OT", wot, C["orange"]),
                                   ("Saturday OT", satok, C["green"]),
                                   ("Sunday OT",  sunot, C["purple"])]:
            row = tk.Frame(pie_frame, bg=C["bg"]); row.pack(fill=tk.X, pady=3)
            tk.Label(row, text="█", font=("Segoe UI",14), bg=C["bg"],
                     fg=color).pack(side=tk.LEFT)
            tk.Label(row, text=f"{label}:", font=FS, bg=C["bg"],
                     fg=C["text"], width=14, anchor=tk.W).pack(side=tk.LEFT)
            h=int(val); m=round((val-h)*60)
            tk.Label(row, text=f"{h}hr {m}m", font=FS, bg=C["bg"],
                     fg=color, width=11, anchor=tk.W).pack(side=tk.LEFT)
            pct = val / total * 100
            bar_bg = tk.Frame(row, bg=C["surface"], height=14)
            bar_bg.pack(side=tk.LEFT, fill=tk.X, expand=True)
            bar_bg.pack_propagate(False)
            tk.Frame(bar_bg, bg=color, height=14).place(
                relx=0, rely=0, relwidth=pct/100, relheight=1)
            tk.Label(row, text=f"{pct:.1f}%", font=FS, bg=C["bg"],
                     fg=C["muted"], width=6).pack(side=tk.LEFT)

    # ── Logic ─────────────────────────────────────────────
    def _browse(self):
        path=filedialog.askopenfilename(title="Select Biometric File",
            filetypes=[("Supported","*.xlsx *.xls *.csv *.pdf"),
                       ("Excel","*.xlsx *.xls"),("PDF","*.pdf"),("All","*.*")])
        if path:
            self._filepath=path
            self._file_lbl.configure(text=f"📄 {os.path.basename(path)}",fg=C["text"])
            self._status_lbl.configure(text="File loaded — click Calculate",fg=C["accent"])
            self._calc_btn.configure(state=tk.NORMAL)
            self._records=[]; self._results=[]

    def _run_calc(self):
        if not self._filepath:
            messagebox.showwarning("No File","Select a biometric file first."); return
        shifts=self._get_shifts()
        if not shifts:
            messagebox.showwarning("No Shifts","Add at least one shift."); return

        # Validate Sunday Regular Hours (kept for compatibility, but no longer used)
        sun_reg_hrs_str = self._vars["sun_reg_hrs"].get().strip()
        if not sun_reg_hrs_str:
            sun_reg_hrs_str = "6"  # fallback to default
        try:
            sun_reg_hrs = float(sun_reg_hrs_str)
        except ValueError:
            self._status_lbl.configure(text="Invalid Sunday Hours", fg=C["red"])
            messagebox.showerror("Invalid Input", "Sunday Regular Hours must be a number.")
            return

        self._calc_btn.configure(state=tk.DISABLED, text="Processing…")
        self._status_lbl.configure(text="⏳ Parsing file…", fg=C["yellow"])
        self.update()

        def worker():
            try:
                ext = os.path.splitext(self._filepath)[1].lower()
                recs = parse_pdf(self._filepath) if ext == ".pdf" else parse_excel(self._filepath)
                if not recs:
                    self.after(0, lambda: self._on_err("No valid records found.\n\nCheck file has Staff ID, Name and Date/Time columns."))
                    return
                late_d = self._vars["late_deduct"].get()
                emp_overrides = self._get_emp_overrides()
                res = calculate_overtime(recs, shifts, sun_reg_hrs, late_d, emp_overrides)
                self.after(0, lambda r=recs, rs=res: self._on_ok(r, rs))
            except Exception as ex:
                msg = f"Error:\n{ex}\n\n{traceback.format_exc()}"
                self.after(0, lambda m=msg: self._on_err(m))

        threading.Thread(target=worker, daemon=True).start()

    def _on_ok(self,recs,results):
        self._records=recs; self._results=results
        tot_ot=sum(r["total_ot"] for r in results)
        top=results[0]["total_ot"] if results else 0
        self._sv["emp"].set(str(len(results)))
        self._sv["ot"].set(_fmt_hm(tot_ot) + " hrs")
        self._sv["pay"].set("— set rate →")
        self._sv["top"].set(_fmt_hm(top) + " hrs")
        self._populate(results)
        self._status_lbl.configure(text=f"✓ {len(recs)} records → {len(results)} employees",fg=C["green"])
        self._calc_btn.configure(state=tk.NORMAL,text="⚡  CALCULATE OVERTIME")
        self._export_btn.configure(state=tk.NORMAL)
        self._refresh_charts()

    def _on_err(self,msg):
        self._status_lbl.configure(text="⚠ Error",fg=C["red"])
        self._calc_btn.configure(state=tk.NORMAL,text="⚡  CALCULATE OVERTIME")
        messagebox.showerror("Error",msg)

    def _apply_rate(self,hrly_override=None):
        if not self._results:
            messagebox.showwarning("No Data","Calculate overtime first."); return None
        # Validate hourly rate
        rate_str = self._vars["hourly_rate"].get().strip()
        if not rate_str:
            rate_str = "200"  # fallback
        try:
            hrly = hrly_override if hrly_override is not None else float(rate_str)
        except ValueError:
            messagebox.showerror("Invalid", "Hourly Rate must be a number.")
            return None
        for r in self._results:
            r["ot_pay"]=round(r["total_ot"]*hrly,2)
        self._sv["pay"].set(f"KES {sum(r['ot_pay'] for r in self._results):,.0f}")
        self._populate(self._results)
        self._status_lbl.configure(text=f"✓ KES {hrly:,.0f}/hr applied",fg=C["green"])
        # Persist the rate so it's remembered next launch
        self._vars["hourly_rate"].set(str(hrly))
        self._save_config()
        return hrly

    def _populate(self,data):
        if not self._tree: return
        for item in self._tree.get_children(): self._tree.delete(item)
        for i,r in enumerate(data):
            tag="none"
            if r["total_ot"]>20: tag="high"
            elif r["total_ot"]>10: tag="mid"
            elif r["total_ot"]>0: tag="low"
            tags=(tag,)+(("alt",) if i%2 else ())
            late=(lambda m: f"{int(m)//60}h {int(m)%60}m" if m>=60 else f"{int(m)}m")(r.get('total_late_min',0)) if r.get("total_late_min",0)>0 else "—"
            pay=f"KES {r['ot_pay']:,.2f}" if r["ot_pay"]>0 else "—"
            self._tree.insert("",tk.END,iid=f"emp_{r['id']}",tags=tags,values=(
                r["id"],r["name"],r["days_worked"],
                f"{r['regular_hours']:.2f}",_fmt_hm(r['weekday_ot']),
                _fmt_hm(r.get('saturday_ot',0)),_fmt_hm(r['sunday_ot']),
                _fmt_hm(r['total_ot']),late,pay))

    def _filter(self):
        if not self._tree or not self._results: return
        q=self._search.get().strip().lower()
        ph="search by name or staff id..."
        if q==ph or not q: self._populate(self._results); return
        self._populate([r for r in self._results if q in r["name"].lower() or q in r["id"].lower()])

    def _sort(self,col):
        self._sort_rev=not self._sort_rev if self._sort_col==col else True
        self._sort_col=col
        key={"id":lambda r:r["id"],"name":lambda r:r["name"],"days":lambda r:r["days_worked"],
             "regular":lambda r:r["regular_hours"],"wot":lambda r:r["weekday_ot"],
             "sot":lambda r:r["sunday_ot"],"tot":lambda r:r["total_ot"],
             "late":lambda r:r.get("total_late_min",0),"pay":lambda r:r["ot_pay"]}[col]
        self._results.sort(key=key,reverse=self._sort_rev)
        self._populate(self._results)

    # ── Breakdown window ──────────────────────────────────
    def _show_breakdown(self,event):
        item=self._tree.identify_row(event.y)
        if not item: return
        emp=next((r for r in self._results if r["id"]==item.replace("emp_","",1)),None)
        if not emp: return

        win=tk.Toplevel(self)
        win.title(f"Breakdown — {emp['name']}")
        win.geometry("1050x620")
        win.configure(bg=C["bg"])
        # No grab_set() → window can be minimised, moved behind main window freely

        # ── Header ──
        hf=tk.Frame(win,bg=C["surface"]); hf.pack(fill=tk.X)
        tk.Label(hf,text=f"  {emp['name']}  ({emp['id']})",
                 font=FH,bg=C["surface"],fg=C["accent"]).pack(side=tk.LEFT,pady=10,padx=6)

        # ── Inline rate bar ──
        rf=tk.Frame(win,bg=C["card"],highlightthickness=1,highlightbackground=C["border"])
        rf.pack(fill=tk.X)
        tk.Label(rf,text="Hourly Rate (KES):",font=FL,bg=C["card"],fg=C["muted"]).pack(side=tk.LEFT,padx=(14,6),pady=8)
        v_hrly=tk.StringVar(value=self._vars["hourly_rate"].get())
        pay_var=tk.StringVar(value=f"KES {emp['ot_pay']:,.2f}" if emp["ot_pay"]>0 else "—")
        tk.Entry(rf,textvariable=v_hrly,font=FB,width=10,bg=C["surface"],fg=C["text"],
                 insertbackground=C["accent"],relief=tk.FLAT,bd=0).pack(side=tk.LEFT,pady=8,ipady=4)
        tk.Label(rf,text="=",font=FH,bg=C["card"],fg=C["muted"]).pack(side=tk.LEFT,padx=8,pady=8)
        tk.Label(rf,textvariable=pay_var,font=("Segoe UI",13,"bold"),bg=C["card"],fg=C["yellow"]).pack(side=tk.LEFT,padx=4,pady=8)

        sum_vars={}

        def apply_local():
            try: hrly=float(v_hrly.get())
            except: messagebox.showerror("Invalid","Enter a valid rate.",parent=win); return
            self._vars["hourly_rate"].set(str(hrly))
            emp["ot_pay"]=round(emp["total_ot"]*hrly,2)
            pay_var.set(f"KES {emp['ot_pay']:,.2f}")
            self._apply_rate(hrly)
            sum_vars["OT Pay"].set(f"KES {emp['ot_pay']:,.2f}")

        tk.Button(rf,text="▶ Apply",font=FL,bg=C["purple"],fg=C["white"],
                  activebackground="#5b21b6",relief=tk.FLAT,bd=0,cursor="hand2",
                  command=apply_local,padx=10).pack(side=tk.LEFT,pady=8,ipady=4,padx=(8,14))

        # ── Summary strip ──
        sf=tk.Frame(win,bg=C["card"]); sf.pack(fill=tk.X)
        summary_items=[
            ("Days",        str(emp["days_worked"]),                    C["text"]),
            ("Reg Hrs",     f"{emp['regular_hours']:.2f}",              C["text"]),
            ("Mon–Fri OT",  _fmt_hm(emp['weekday_ot']),               C["orange"]),
            ("Sat OT",      _fmt_hm(emp.get('saturday_ot',0)),        C["blue"]),
            ("Sun OT",      _fmt_hm(emp['sunday_ot']),                C["purple"]),
            ("Total OT",    _fmt_hm(emp['total_ot']),                 C["accent"]),
            ("Total Late",  (lambda m: f"{int(m)//60}h {int(m)%60}m" if m>=60 else f"{int(m)}m")(emp.get('total_late_min',0)),    C["red"]),
            ("OT Pay",      f"KES {emp['ot_pay']:,.2f}" if emp["ot_pay"]>0 else "—", C["yellow"]),
        ]
        for lbl,val,color in summary_items:
            f=tk.Frame(sf,bg=C["card"]); f.pack(side=tk.LEFT,padx=10,pady=8)
            tk.Label(f,text=lbl,font=FS,bg=C["card"],fg=C["muted"]).pack()
            sv=tk.StringVar(value=val); sum_vars[lbl]=sv
            tk.Label(f,textvariable=sv,font=("Segoe UI",11,"bold"),bg=C["card"],fg=color).pack()

        # ── Breakdown table ──
        # CRITICAL: tree parent must be tf (the grid container), NOT win
        tf=tk.Frame(win,bg=C["bg"]); tf.pack(fill=tk.BOTH,expand=True,padx=10,pady=(6,0))

        tree,_,_=_make_tree(tf,[
            ("date","Shift Date", 100,tk.CENTER),
            ("day", "Day",         44,tk.CENTER),
            ("shft","Shift",       88,tk.CENTER),
            ("si",  "Sched In",    68,tk.CENTER),
            ("so",  "Sched Out",   80,tk.CENTER),
            ("cin", "Check-In",   122,tk.CENTER),
            ("cout","Check-Out",  140,tk.CENTER),
            ("wrk", "Eff. Hours",  76,tk.CENTER),   # effective hours (early clipped)
            ("early","Early(m)",   60,tk.CENTER),
            ("late","Late(m)",     60,tk.CENTER),
            ("ot",  "OT Hrs",      65,tk.CENTER),
            ("note","Note",       200,tk.W),
        ])
        tree.tag_configure("sun",  background="#1e1530",foreground="#c4b5fd")
        tree.tag_configure("sat",  background="#1a2010",foreground="#86efac")
        tree.tag_configure("night",background="#0f1a2a",foreground="#93c5fd")
        tree.tag_configure("late", foreground=C["red"])
        tree.tag_configure("ot",   foreground=C["orange"])
        tree.tag_configure("warn", foreground=C["yellow"])
        tree.tag_configure("alt2", background="#181e2e")

        for i,d in enumerate(emp["breakdown"]):
            tags=set()
            is_night="night" in d.get("shift","").lower()
            if d["is_sunday"]:          tags.add("sun")
            elif d.get("is_saturday"):  tags.add("sat")
            elif is_night:              tags.add("night")
            elif i%2==0:                tags.add("alt2")
            if d.get("late_min",0)>0:   tags.add("late")
            elif d.get("ot",0)>0:       tags.add("ot")
            if str(d.get("note","")).startswith("⚠"): tags.add("warn")

            # Night shift: show MM-DD HH:MM:SS so the date crossing is obvious
            cin_full  = d.get("check_in_full","")
            cout_full = d.get("check_out_full","")
            if is_night and cin_full:
                cin_disp  = cin_full[5:]   # "MM-DD HH:MM:SS"
                cout_disp = cout_full[5:] if cout_full and cout_full != "—" else "—"
            else:
                cin_disp  = d["check_in"]
                cout_disp = d["check_out"]

            tree.insert("",tk.END,tags=tuple(tags),values=(
                d["date"],
                d["day"]+(" ☀" if d["is_sunday"] else ""),
                d.get("shift","—"),
                d.get("sched_in","—"),
                d.get("sched_out","—"),
                cin_disp,
                cout_disp,
                _fmt_hm(d['worked']) if d["worked"]>0 else "—",
                f"{int(d.get('early_min',0))}m" if d.get("early_min",0)>0 else "—",
                (lambda m: f"{int(m)//60}h {int(m)%60}m" if m>=60 else f"{int(m)}m")(d.get('late_min',0)) if d.get("late_min",0)>0 else "—",
                _fmt_hm(d['ot']) if d.get("ot",0)>0 else "—",
                d.get("note",""),
            ))

        lf=tk.Frame(win,bg=C["bg"]); lf.pack(fill=tk.X,padx=10,pady=(2,8))
        for t,col in [("🌙 Night","#93c5fd"),("🟩 Saturday","#86efac"),
                      ("☀ Sunday","#c4b5fd"),("🔴 Late",C["red"]),("🔥 OT",C["orange"])]:
            tk.Label(lf,text=t,font=FS,bg=C["bg"],fg=col).pack(side=tk.LEFT,padx=10)

    # ── Export ────────────────────────────────────────────
    def _export(self):
        if not self._results:
            messagebox.showwarning("No Data","Calculate first."); return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],initialfile="Overtime_Summary.xlsx")
        if path:
            try:
                export_to_excel(self._results,path)
                messagebox.showinfo("Exported",f"Saved:\n{path}")
            except Exception as ex:
                messagebox.showerror("Export Error",str(ex))

# ══════════════════════════════════════════════════════════
if __name__=="__main__":
    app=OvertimeApp()
    app.mainloop()