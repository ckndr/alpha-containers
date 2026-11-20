"""
update_plans.py
Tubex -- Daily Plan Tracker & Forecast Updater (PET + Tubes Hook)
-----------------------------------------------------------------
Reads Production_Log and MRP from the active Tubex_<Mon><YY>.xlsx workbook:
  1. Updates PET_Plan in-place (keeps Baseline Planned Qty static by default;
     populates Actual Good, Wastage, Variance, Status for dates <= max_prod_date;
     dynamically simulates Forecast Finish following the revised sequence:
     8001 (Nov 18-21) -> Sun Nov 22 CO -> 8008 -> 8007 -> Nov 28 CO -> 8011 ->
     8005 -> Dec 7 CO -> 8015 -> Dec 9 CO -> 8006 -> Dec 15 CO -> 8020 ->
     Dec 19 CO -> remaining 8001).
  2. Supports --replan to intentionally re-cascade PET_Plan forward schedule
     from max_prod_date + 1 using live MRP remaining balances.
  3. Hook for Tubes_Plan (updates Actuals/Status/Forecast when sheet exists,
     excluding '(Varnish)' passes per PLANNING_RULES.md; skips cleanly otherwise).
  4. Generates/updates standalone shareable d:\\Alpha\\<MonthName>_Plan.xlsx
     (e.g., November_Plan.xlsx), falling back to <MonthName>_Plan_LATEST.xlsx
     if locked by Excel. Never modifies historical August_Plan.xlsx.
"""

import os
import sys
import re
import argparse
import calendar
from datetime import datetime, date, timedelta

import warnings
warnings.filterwarnings("ignore", message=".*Data Validation.*")
warnings.filterwarnings("ignore", message=".*extension.*")

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from alpha_checks import get_active_tubex_file, check_not_locked, atomic_save

DAILY_PET_CAPACITY = 10000
PAKISTAN_HOLIDAYS = {
    date(2026, 12, 25),  # Quaid-e-Azam Day / Christmas
    date(2027, 2, 5),    # Kashmir Day
    date(2027, 3, 23),   # Pakistan Day
}

MONTH_ABBR_TO_FULL = {
    'jan': 'January', 'feb': 'February', 'mar': 'March', 'apr': 'April',
    'may': 'May', 'jun': 'June', 'jul': 'July', 'aug': 'August',
    'sep': 'September', 'oct': 'October', 'nov': 'November', 'dec': 'December',
}

# Mould/colour group for changeover transitions
PET_MOULD_GROUP = {
    8001: '150_TRP',
    8008: '200_SAMSOL',
    8007: '200_SAMSOL',
    8006: '200_SAMSOL',
    8011: '120_SAMSOL',
    8005: '120_SAMSOL',
    8015: '130_WHITE',
    8020: '150_CAPSULE',
}

PET_SKU_META = {
    8001: ('150 ml', 'Alpha Labs PVT LTD', 'Alpha 150ml TRP'),
    8008: ('200 ml', 'Samsol', 'Samsol Compact Black 200ml'),
    8007: ('200 ml', 'Samsol', 'Samsol White 200ml'),
    8011: ('120 ml', 'Samsol', 'Samsol Compact Black 120ml'),
    8005: ('120 ml', 'Samsol', 'Samsol Yellow 120ml'),
    8015: ('130 ml', 'Mablay Beauty PVT LTD.', 'Mablay 130ml White'),
    8006: ('200 ml', 'Samsol', 'Samsol Yellow 200ml'),
    8020: ('150 ml', 'Alpha Labs PVT LTD', 'Alpha 150ml Capsule'),
}

# Revised campaign sequence after initial Alpha TRP tranche
POST_TRANCHE1_ORDER = [8008, 8007, 8011, 8005, 8015, 8006, 8020]


def _to_date(val):
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, str):
        for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d-%b-%Y', '%d-%b-%y'):
            try:
                return datetime.strptime(val.strip(), fmt).date()
            except ValueError:
                pass
    return None


def _eval_num(val):
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(round(val))
    s = str(val).strip().lstrip('=').replace(',', '')
    if re.match(r'^[0-9+\-*/().\s]+$', s):
        try:
            return int(round(float(eval(s, {'__builtins__': None}, {}))))
        except Exception:
            return 0
    return 0


def detect_plan_month_name(excel_path):
    base = os.path.basename(excel_path)
    m = re.search(r'Tubex_([A-Za-z]{3})\d{2}', base, re.IGNORECASE)
    if m:
        abbr = m.group(1).lower()
        if abbr in MONTH_ABBR_TO_FULL:
            return MONTH_ABBR_TO_FULL[abbr]
    return calendar.month_name[datetime.now().month]


def find_pet_plan_total_row(ws_pp):
    """Locates the TOTAL row in PET_Plan left table (e.g. Row 80)."""
    for r in range(3, ws_pp.max_row + 1):
        if str(ws_pp.cell(r, 5).value or '').strip().upper() == 'TOTAL':
            return r
    return 80


def read_pet_production_log(wb):
    """Returns (by_date, by_pid_total, last_date_by_pid, max_prod_date)."""
    ws_pl = wb['Production_Log']
    by_date = {}
    by_pid_total = {}
    last_date_by_pid = {}
    max_prod_date = None

    for r in range(3, ws_pl.max_row + 1):
        dt = _to_date(ws_pl.cell(r, 1).value)
        mach = str(ws_pl.cell(r, 2).value or '').strip().upper()
        if not dt or not (mach.startswith('PF') or mach.startswith('PET')):
            continue

        pid_raw = ws_pl.cell(r, 6).value
        try:
            pid = int(pid_raw) if pid_raw is not None else None
        except (ValueError, TypeError):
            pid = None

        good = _eval_num(ws_pl.cell(r, 8).value)
        waste = _eval_num(ws_pl.cell(r, 9).value)

        if max_prod_date is None or dt > max_prod_date:
            max_prod_date = dt

        entry = by_date.setdefault(dt, {'good': 0, 'waste': 0, 'pids': set()})
        entry['good'] += good
        entry['waste'] += waste
        if pid is not None and (good > 0 or waste > 0):
            entry['pids'].add(pid)

        if pid is not None:
            by_pid_total[pid] = by_pid_total.get(pid, 0) + good
            if good > 0:
                if pid not in last_date_by_pid or dt > last_date_by_pid[pid]:
                    last_date_by_pid[pid] = dt

    return by_date, by_pid_total, last_date_by_pid, max_prod_date


def read_mrp_pet_required(wb):
    """Reads PET Required Qty by PID directly from MRP sheet."""
    ws_mrp = wb['MRP']
    pet_req = {}
    in_pet = False
    for r in range(1, ws_mrp.max_row + 1):
        c1 = str(ws_mrp.cell(r, 1).value or '').strip().upper()
        c5 = str(ws_mrp.cell(r, 5).value or '').strip().upper()
        if 'PET REQUIRED' in c1:
            in_pet = True
            continue
        if in_pet:
            if c5 == 'TOTAL:' or 'PET MATERIAL REQUIREMENT' in c1:
                break
            pid_raw = ws_mrp.cell(r, 4).value
            if pid_raw is not None:
                try:
                    pid = int(pid_raw)
                    pet_req[pid] = _eval_num(ws_mrp.cell(r, 6).value)
                except (ValueError, TypeError):
                    pass
    return pet_req


def simulate_pet_forecast(ws_pp, total_row, rem_by_pid, by_date, last_date_by_pid, max_prod_date):
    """
    Simulates forward from sim_start following the revised sequence:
      8001 (Tranche 1 through Nov 21) -> Sun Nov 22 CO -> 8008 -> 8007 ->
      Nov 28 CO -> 8011 -> 8005 -> Dec 7 CO -> 8015 -> Dec 9 CO -> 8006 ->
      Dec 15 CO -> 8020 -> Dec 19 CO -> remaining 8001.
    Respects scheduled Sunday changeovers (e.g., Nov 22), skips WEEKLY OFF
    Sundays and Pakistan national holidays, and honours half-shift purge
    capacities encoded in the baseline schedule.
    """
    plan_start = date(2026, 11, 18)
    sim_start = max((max_prod_date + timedelta(days=1)) if max_prod_date else plan_start, plan_start)

    plan_holidays = set(PAKISTAN_HOLIDAYS)
    plan_co_dates = set()
    daily_cap_by_date = {}
    first_co_date = None
    planned_8001_before_co = 0

    for r in range(3, total_row):
        dt = _to_date(ws_pp.cell(r, 2).value)
        cust = str(ws_pp.cell(r, 4).value or '').strip().upper()
        pid_val = ws_pp.cell(r, 6).value
        pqty = _eval_num(ws_pp.cell(r, 7).value)
        if not dt:
            continue
        if cust == 'PUBLIC HOLIDAY':
            plan_holidays.add(dt)
        elif cust == 'CHANGEOVER':
            plan_co_dates.add(dt)
            if first_co_date is None:
                first_co_date = dt
        elif cust != 'WEEKLY OFF' and pqty > 0:
            daily_cap_by_date[(dt, pid_val)] = pqty
            if pid_val == 8001 and (first_co_date is None or dt < first_co_date):
                planned_8001_before_co += pqty

    prod_8001_in_plan = sum(
        info['good'] for dt, info in by_date.items()
        if dt >= plan_start and (not first_co_date or dt < first_co_date) and 8001 in info['pids']
    )

    rem = dict(rem_by_pid)
    forecast_finish = {}
    for pid, rqty in rem.items():
        if rqty <= 0:
            forecast_finish[pid] = last_date_by_pid.get(pid, max_prod_date or plan_start)

    tranche1_target = max(0, planned_8001_before_co - prod_8001_in_plan)
    tranche1_qty = min(rem.get(8001, 0), tranche1_target) if (not first_co_date or sim_start <= first_co_date) else 0
    tranche2_qty = max(0, rem.get(8001, 0) - tranche1_qty)

    queue = []
    if tranche1_qty > 0:
        queue.append([8001, tranche1_qty])
    for pid in POST_TRANCHE1_ORDER:
        if rem.get(pid, 0) > 0:
            queue.append([pid, rem[pid]])
    if tranche2_qty > 0:
        queue.append([8001, tranche2_qty])

    prev_group = None
    if max_prod_date and max_prod_date in by_date and by_date[max_prod_date]['pids']:
        last_pid = sorted(by_date[max_prod_date]['pids'])[-1]
        prev_group = PET_MOULD_GROUP.get(last_pid)
    elif tranche1_qty > 0:
        prev_group = PET_MOULD_GROUP.get(8001)

    cur_dt = sim_start
    q_idx = 0
    need_co = False
    if queue and prev_group and PET_MOULD_GROUP.get(queue[0][0]) != prev_group:
        need_co = True

    safety = 0
    while q_idx < len(queue) and safety < 730:
        safety += 1
        # Allow explicit changeover even on Sunday (e.g. Nov 22 Sun CO)
        if cur_dt in plan_holidays:
            cur_dt += timedelta(days=1)
            continue

        if cur_dt.weekday() == 6 and cur_dt not in plan_co_dates:
            cur_dt += timedelta(days=1)
            continue

        if need_co:
            need_co = False
            cur_dt += timedelta(days=1)
            continue

        cur_pid, cur_rem = queue[q_idx]
        cur_group = PET_MOULD_GROUP.get(cur_pid)
        if cur_dt in plan_co_dates and prev_group != cur_group:
            cur_dt += timedelta(days=1)
            continue

        # Honour half-shift purge capacity on matching planned day (e.g. Dec 2 = 5,000), else 10,000
        cap = daily_cap_by_date.get((cur_dt, cur_pid), DAILY_PET_CAPACITY)
        run_qty = min(cap, cur_rem)
        queue[q_idx][1] -= run_qty
        prev_group = cur_group

        if queue[q_idx][1] <= 0:
            forecast_finish[cur_pid] = cur_dt
            q_idx += 1
            if q_idx < len(queue):
                next_pid = queue[q_idx][0]
                if PET_MOULD_GROUP.get(next_pid) != cur_group:
                    need_co = True

        cur_dt += timedelta(days=1)

    return forecast_finish, queue


def update_pet_plan_sheet(wb, replan=False):
    if 'PET_Plan' not in wb.sheetnames:
        print('  [WARN] PET_Plan sheet not found in workbook -- skipped')
        return None

    ws_pp = wb['PET_Plan']
    total_row = find_pet_plan_total_row(ws_pp)
    by_date, by_pid_total, last_date_by_pid, max_prod_date = read_pet_production_log(wb)
    pet_req = read_mrp_pet_required(wb)

    rem_by_pid = {
        pid: max(0, pet_req.get(pid, 0) - by_pid_total.get(pid, 0))
        for pid in PET_SKU_META
    }

    if replan:
        plan_start = date(2026, 11, 18)
        sim_start = max((max_prod_date + timedelta(days=1)) if max_prod_date else plan_start, plan_start)
        plan_holidays = set(PAKISTAN_HOLIDAYS)
        plan_co_dates = set()
        for r in range(3, total_row):
            dt = _to_date(ws_pp.cell(r, 2).value)
            cust = str(ws_pp.cell(r, 4).value or '').strip().upper()
            if dt and cust == 'PUBLIC HOLIDAY':
                plan_holidays.add(dt)
            elif dt and cust == 'CHANGEOVER':
                plan_co_dates.add(dt)

        _, queue = simulate_pet_forecast(ws_pp, total_row, rem_by_pid, by_date, last_date_by_pid, max_prod_date)
        q_idx = 0
        need_co = False

        for r in range(3, total_row):
            dt = _to_date(ws_pp.cell(r, 2).value)
            if not dt or dt < sim_start:
                continue
            if dt in plan_holidays:
                ws_pp.cell(r, 3, '-')
                ws_pp.cell(r, 4, 'PUBLIC HOLIDAY')
                ws_pp.cell(r, 5, 'Public Holiday')
                ws_pp.cell(r, 6, '-')
                ws_pp.cell(r, 7, 0)
                ws_pp.cell(r, 11, 'Off/Changeover')
                continue
            if dt.weekday() == 6 and dt not in plan_co_dates:
                ws_pp.cell(r, 3, '-')
                ws_pp.cell(r, 4, 'WEEKLY OFF')
                ws_pp.cell(r, 5, 'Sunday - Weekly Off')
                ws_pp.cell(r, 6, '-')
                ws_pp.cell(r, 7, 0)
                ws_pp.cell(r, 11, 'Off/Changeover')
                continue
            if q_idx >= len(queue):
                ws_pp.cell(r, 3, '-')
                ws_pp.cell(r, 4, 'IDLE / BUFFER')
                ws_pp.cell(r, 5, 'Plan Complete')
                ws_pp.cell(r, 6, '-')
                ws_pp.cell(r, 7, 0)
                ws_pp.cell(r, 11, 'Off/Changeover')
                continue
            if need_co:
                nxt_pid = queue[q_idx][0]
                nxt_ml = PET_SKU_META[nxt_pid][0]
                ws_pp.cell(r, 3, '-')
                ws_pp.cell(r, 4, 'CHANGEOVER')
                ws_pp.cell(r, 5, f'Mould Swap -> {PET_SKU_META[nxt_pid][2]} ({nxt_ml})')
                ws_pp.cell(r, 6, '-')
                ws_pp.cell(r, 7, 0)
                ws_pp.cell(r, 11, 'Off/Changeover')
                need_co = False
                continue

            cur_pid, cur_rem = queue[q_idx]
            ml, cust, pname = PET_SKU_META[cur_pid]
            qty = min(DAILY_PET_CAPACITY, cur_rem)
            ws_pp.cell(r, 3, ml)
            ws_pp.cell(r, 4, cust)
            ws_pp.cell(r, 5, pname)
            ws_pp.cell(r, 6, cur_pid)
            ws_pp.cell(r, 7, qty)
            ws_pp.cell(r, 11, 'Scheduled')
            cur_group = PET_MOULD_GROUP[cur_pid]
            queue[q_idx][1] -= qty
            if queue[q_idx][1] <= 0:
                q_idx += 1
                if q_idx < len(queue) and PET_MOULD_GROUP[queue[q_idx][0]] != cur_group:
                    need_co = True

    # Populate Actuals (H, I) and Status (K) for dates <= max_prod_date
    for r in range(3, total_row):
        dt = _to_date(ws_pp.cell(r, 2).value)
        if not dt:
            continue
        cust = str(ws_pp.cell(r, 4).value or '').strip().upper()
        is_non_prod = cust in ('WEEKLY OFF', 'CHANGEOVER', 'PUBLIC HOLIDAY', 'IDLE / BUFFER')
        planned_pid = ws_pp.cell(r, 6).value
        try:
            planned_pid_int = int(planned_pid) if planned_pid not in (None, '-', '') else None
        except (ValueError, TypeError):
            planned_pid_int = None
        planned_qty = _eval_num(ws_pp.cell(r, 7).value)

        if max_prod_date and dt <= max_prod_date:
            day_info = by_date.get(dt, {'good': 0, 'waste': 0, 'pids': set()})
            act_good = day_info['good']
            act_waste = day_info['waste']
            act_pids = day_info['pids']
            ws_pp.cell(r, 8, act_good)
            ws_pp.cell(r, 9, act_waste)

            if is_non_prod:
                if act_good > 0:
                    pids_txt = ', '.join(str(p) for p in sorted(act_pids)) or 'PET'
                    status = f'Unplanned: {pids_txt}'
                else:
                    status = 'Off/Changeover'
            else:
                if act_good > 0 and planned_pid_int is not None and planned_pid_int not in act_pids:
                    pids_txt = ', '.join(str(p) for p in sorted(act_pids))
                    status = f'Unplanned: {pids_txt}'
                else:
                    diff = act_good - planned_qty
                    if diff == 0:
                        status = 'Done'
                    elif diff > 0:
                        status = f'Ahead (+{diff:,})'
                    else:
                        status = f'Behind ({diff:,})'
            ws_pp.cell(r, 11, status)
        else:
            ws_pp.cell(r, 8, None)
            ws_pp.cell(r, 9, None)
            ws_pp.cell(r, 11, 'Off/Changeover' if is_non_prod else 'Scheduled')

    # Update Forecast Finish (Col U = col 21, rows 10..17)
    forecast_finish, _ = simulate_pet_forecast(ws_pp, total_row, rem_by_pid, by_date, last_date_by_pid, max_prod_date)
    for r in range(10, 18):
        pid_val = ws_pp.cell(r, 13).value
        try:
            pid_int = int(pid_val)
        except (ValueError, TypeError):
            continue
        fin_dt = forecast_finish.get(pid_int)
        if fin_dt:
            ws_pp.cell(r, 21, datetime(fin_dt.year, fin_dt.month, fin_dt.day))

    print(f"  [OK] PET_Plan updated (max PF date={max_prod_date}, total PET MTD={sum(by_pid_total.values()):,}, remaining={sum(rem_by_pid.values()):,})")
    return {
        'by_date': by_date,
        'by_pid_total': by_pid_total,
        'pet_req': pet_req,
        'rem_by_pid': rem_by_pid,
        'forecast_finish': forecast_finish,
        'max_prod_date': max_prod_date,
        'total_row': total_row,
    }


def update_tubes_plan_hook(wb):
    """
    Hook for Tubes_Plan when added to Tubex_<Mon><YY>.xlsx.
    Excludes '(Varnish)' passes per PLANNING_RULES.md.
    """
    if 'Tubes_Plan' not in wb.sheetnames:
        print('  [INFO] Tubes_Plan not present yet -- skipped')
        return False

    ws_tp = wb['Tubes_Plan']
    ws_pl = wb['Production_Log']
    by_date = {}
    max_tube_dt = None

    for r in range(3, ws_pl.max_row + 1):
        dt = _to_date(ws_pl.cell(r, 1).value)
        mach = str(ws_pl.cell(r, 2).value or '').strip().upper()
        pname = str(ws_pl.cell(r, 4).value or '').strip().upper()
        if not dt or not (mach.startswith('PRINT') or mach.startswith('PLINE')):
            continue
        if '(VARNISH)' in pname:
            continue
        pid_raw = ws_pl.cell(r, 6).value
        try:
            pid = int(pid_raw) if pid_raw is not None else None
        except (ValueError, TypeError):
            pid = None
        good = _eval_num(ws_pl.cell(r, 8).value)
        waste = _eval_num(ws_pl.cell(r, 9).value)
        if max_tube_dt is None or dt > max_tube_dt:
            max_tube_dt = dt
        d = by_date.setdefault(dt, {'good': 0, 'waste': 0, 'pids': set()})
        d['good'] += good
        d['waste'] += waste
        if pid is not None and good > 0:
            d['pids'].add(pid)

    for r in range(3, ws_tp.max_row + 1):
        dt = _to_date(ws_tp.cell(r, 2).value)
        if not dt:
            continue
        if max_tube_dt and dt <= max_tube_dt:
            info = by_date.get(dt, {'good': 0, 'waste': 0, 'pids': set()})
            ws_tp.cell(r, 8, info['good'])
            ws_tp.cell(r, 9, info['waste'])
            planned_qty = _eval_num(ws_tp.cell(r, 7).value)
            diff = info['good'] - planned_qty
            if diff == 0:
                ws_tp.cell(r, 11, 'Done')
            elif diff > 0:
                ws_tp.cell(r, 11, f'Ahead (+{diff:,})')
            else:
                ws_tp.cell(r, 11, f'Behind ({diff:,})')
    print(f'  [OK] Tubes_Plan updated (max Print date={max_tube_dt})')
    return True


def export_standalone_plan(wb, month_name, pet_stats):
    """
    Exports standalone shareable d:\\Alpha\\<MonthName>_Plan.xlsx with:
      - '<MonthName> Plan PET'
      - '<MonthName> PET Weekly'
    Never touches August_Plan.xlsx. Uses <MonthName>_Plan_LATEST.xlsx if locked.
    """
    if month_name.lower() == 'august':
        print('  [SKIP] Historical August_Plan.xlsx is protected and never modified.')
        return None

    target_path = os.path.join(DIR, f'{month_name}_Plan.xlsx')
    fallback_path = os.path.join(DIR, f'{month_name}_Plan_LATEST.xlsx')
    lock_file = os.path.join(DIR, f'~${month_name}_Plan.xlsx')

    is_locked = os.path.exists(lock_file)
    if not is_locked and os.path.exists(target_path):
        try:
            with open(target_path, 'r+b'):
                pass
        except PermissionError:
            is_locked = True

    out_path = fallback_path if is_locked else target_path
    if is_locked:
        print(f'  [WARN] {os.path.basename(target_path)} is locked by Excel -- writing fallback {os.path.basename(out_path)}')

    if 'PET_Plan' not in wb.sheetnames:
        return None
    ws_src = wb['PET_Plan']
    total_row = pet_stats['total_row'] if pet_stats else find_pet_plan_total_row(ws_src)
    last_data_row = total_row - 1

    out_wb = openpyxl.Workbook()
    ws_daily = out_wb.active
    ws_daily.title = f'{month_name} Plan PET'
    ws_weekly = out_wb.create_sheet(title=f'{month_name} PET Weekly')

    thin = Side(style='thin', color='BFBFBF')
    med = Side(style='medium', color='808080')
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_total = Border(left=thin, right=thin, top=med, bottom=med)

    fill_hdr = PatternFill('solid', fgColor='D9D9D9')
    fill_off = PatternFill('solid', fgColor='F2F2F2')
    fill_white = PatternFill('solid', fgColor='FFFFFF')
    fill_co = PatternFill('solid', fgColor='FFF2CC')
    fill_total = PatternFill('solid', fgColor='FFF2CC')

    font_title = Font(name='Segoe UI', size=11, bold=True, color='1A1A2E')
    font_hdr = Font(name='Segoe UI', size=9, bold=True, color='1A1A2E')
    font_body = Font(name='Segoe UI', size=9, bold=False, color='1A1A2E')
    font_bold = Font(name='Segoe UI', size=9, bold=True, color='1A1A2E')
    font_muted = Font(name='Segoe UI', size=9, italic=True, color='595959')

    align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_l = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_r = Alignment(horizontal='right', vertical='center', wrap_text=True)

    ws_daily.merge_cells('A1:J1')
    ws_daily['A1'] = 'PET DAILY PRODUCTION PLAN (NOV 2026 - FEB 2027)'
    for c in range(1, 11):
        cell = ws_daily.cell(1, c)
        cell.font = font_title
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_l if c == 1 else align_c
    ws_daily.row_dimensions[1].height = 24

    daily_headers = ['Week', 'Date', 'ML', 'Customer', 'Product Name',
                     'Planned Qty', 'Actual Good', 'Wastage', 'Variance', 'Status']
    for c, h in enumerate(daily_headers, 1):
        cell = ws_daily.cell(2, c, h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_c
    ws_daily.row_dimensions[2].height = 22

    for r in range(3, total_row):
        wk = ws_src.cell(r, 1).value
        dt = ws_src.cell(r, 2).value
        ml = ws_src.cell(r, 3).value
        cust = ws_src.cell(r, 4).value
        pname = ws_src.cell(r, 5).value
        pqty = ws_src.cell(r, 7).value
        agood = ws_src.cell(r, 8).value
        waste = ws_src.cell(r, 9).value
        status = ws_src.cell(r, 11).value

        ws_daily.cell(r, 1, wk)
        ws_daily.cell(r, 2, dt)
        ws_daily.cell(r, 3, ml)
        ws_daily.cell(r, 4, cust)
        ws_daily.cell(r, 5, pname)
        ws_daily.cell(r, 6, pqty)
        ws_daily.cell(r, 7, agood)
        ws_daily.cell(r, 8, waste)
        ws_daily.cell(r, 9, f'=IF(G{r}="","",G{r}-F{r})')
        ws_daily.cell(r, 10, status)

        cust_u = str(cust or '').upper()
        is_co = cust_u == 'CHANGEOVER'
        is_off = cust_u in ('WEEKLY OFF', 'PUBLIC HOLIDAY', 'IDLE / BUFFER')
        r_fill = fill_co if is_co else (fill_off if is_off else fill_white)
        r_font = font_muted if (is_co or is_off) else font_body

        for c in range(1, 11):
            cell = ws_daily.cell(r, c)
            cell.fill = r_fill
            cell.font = font_bold if (c == 6 and not (is_co or is_off)) else r_font
            cell.border = border_cell
            if c == 2:
                cell.number_format = 'ddd, dd-mmm-yy'
                cell.alignment = align_c
            elif c in (1, 3, 10):
                cell.alignment = align_c
            elif c in (4, 5):
                cell.alignment = align_l
            elif c in (6, 7, 8, 9):
                cell.number_format = '#,##0'
                cell.alignment = align_r
        ws_daily.row_dimensions[r].height = 18

    ws_daily.cell(total_row, 5, 'TOTAL')
    ws_daily.cell(total_row, 6, f'=SUM(F3:F{last_data_row})')
    ws_daily.cell(total_row, 7, f'=SUM(G3:G{last_data_row})')
    ws_daily.cell(total_row, 8, f'=SUM(H3:H{last_data_row})')
    ws_daily.cell(total_row, 9, f'=G{total_row}-F{total_row}')
    for c in range(1, 11):
        cell = ws_daily.cell(total_row, c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = border_total
        if c in (6, 7, 8, 9):
            cell.number_format = '#,##0'
            cell.alignment = align_r
        else:
            cell.alignment = align_r if c == 5 else align_c

    # Right-side PET PLAN Order Progress & Forecast Finish (L8:T18)
    ws_daily.merge_cells('L8:T8')
    ws_daily['L8'] = 'PET PLAN — ORDER PROGRESS & FORECAST FINISH'
    for c in range(12, 21):
        cell = ws_daily.cell(8, c)
        cell.font = font_title
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_l if c == 12 else align_c

    r1_hdrs = ['PID', 'ML', 'Client', 'Product', 'MRP Required', 'MTD Produced',
               'Remaining', 'Baseline Finish', 'Forecast Finish']
    for idx, h in enumerate(r1_hdrs):
        c = 12 + idx
        cell = ws_daily.cell(9, c, h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_c

    pet_req = pet_stats['pet_req'] if pet_stats else {}
    by_pid = pet_stats['by_pid_total'] if pet_stats else {}
    for rr in range(10, 18):
        pid_int = int(ws_src.cell(rr, 13).value)
        ml = ws_src.cell(rr, 14).value
        client = ws_src.cell(rr, 15).value
        pname = ws_src.cell(rr, 16).value
        req_q = pet_req.get(pid_int, 0)
        prod_q = by_pid.get(pid_int, 0)
        base_dt = ws_src.cell(rr, 20).value
        fcst_dt = ws_src.cell(rr, 21).value

        ws_daily.cell(rr, 12, pid_int)
        ws_daily.cell(rr, 13, ml)
        ws_daily.cell(rr, 14, client)
        ws_daily.cell(rr, 15, pname)
        ws_daily.cell(rr, 16, req_q)
        ws_daily.cell(rr, 17, prod_q)
        ws_daily.cell(rr, 18, f'=MAX(0,P{rr}-Q{rr})')
        ws_daily.cell(rr, 19, base_dt)
        ws_daily.cell(rr, 20, fcst_dt)

        r_fill = fill_off if rr % 2 == 0 else fill_white
        for c in range(12, 21):
            cell = ws_daily.cell(rr, c)
            cell.font = font_bold if c in (12, 18, 20) else font_body
            cell.fill = fill_co if c == 20 else r_fill
            cell.border = border_cell
            if c in (12, 13):
                cell.alignment = align_c
            elif c in (14, 15):
                cell.alignment = align_l
            elif c in (16, 17, 18):
                cell.number_format = '#,##0'
                cell.alignment = align_r
            elif c in (19, 20):
                cell.number_format = 'dd-mmm-yy'
                cell.alignment = align_c

    ws_daily.cell(18, 15, 'TOTAL')
    ws_daily.cell(18, 16, '=SUM(P10:P17)')
    ws_daily.cell(18, 17, '=SUM(Q10:Q17)')
    ws_daily.cell(18, 18, '=SUM(R10:R17)')
    ws_daily.cell(18, 19, '=MAX(S10:S17)')
    ws_daily.cell(18, 20, '=MAX(T10:T17)')
    for c in range(12, 21):
        cell = ws_daily.cell(18, c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = border_total
        if c in (16, 17, 18):
            cell.number_format = '#,##0'
            cell.alignment = align_r
        elif c in (19, 20):
            cell.number_format = 'dd-mmm-yy'
            cell.alignment = align_c
        else:
            cell.alignment = align_r if c == 15 else align_c

    # Reconciliation mini-box L3:O6
    ws_daily.merge_cells('L3:O3')
    ws_daily['L3'] = 'MRP RECONCILIATION (PRE-PLAN + SCHEDULE)'
    for c in range(12, 16):
        cell = ws_daily.cell(3, c)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
    pre_plan_mtd = sum(by_pid.values())
    for rr, lbl, val in [
        (4, 'Total MRP Required (8 SKUs)', '=P18'),
        (5, 'Pre-Plan MTD Produced (Nov 15–17)', pre_plan_mtd),
        (6, 'Plan Remaining Balance (From Nov 18)', '=O4-O5'),
    ]:
        ws_daily.merge_cells(start_row=rr, start_column=12, end_row=rr, end_column=14)
        ws_daily.cell(rr, 12, lbl)
        ws_daily.cell(rr, 15, val)
        for c in range(12, 16):
            cell = ws_daily.cell(rr, c)
            cell.font = font_bold if rr == 6 else font_body
            cell.fill = fill_total if rr == 6 else fill_white
            cell.border = border_cell
            cell.alignment = align_l if c < 15 else align_r
            if c == 15:
                cell.number_format = '#,##0'

    daily_widths = {
        1: 9, 2: 16, 3: 9, 4: 26, 5: 38, 6: 13, 7: 13, 8: 11, 9: 12, 10: 18,
        11: 3, 12: 8, 13: 9, 14: 30, 15: 22, 16: 14, 17: 14, 18: 13, 19: 15, 20: 15,
    }
    for c, w in daily_widths.items():
        ws_daily.column_dimensions[get_column_letter(c)].width = w
    ws_daily.freeze_panes = 'A3'

    # Weekly Sheet: A1:H17
    ws_weekly.merge_cells('A1:H1')
    ws_weekly['A1'] = 'PET WEEKLY PLAN & TRACKER'
    for c in range(1, 9):
        cell = ws_weekly.cell(1, c)
        cell.font = font_title
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_l if c == 1 else align_c
    ws_weekly.row_dimensions[1].height = 24

    wk_headers = ['Month', 'Week', 'Dates', 'Working Days', 'Products', 'Planned Qty', 'Actual Good', 'Variance']
    for c, h in enumerate(wk_headers, 1):
        cell = ws_weekly.cell(2, c, h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_c
    ws_weekly.row_dimensions[2].height = 22

    weekly_products = [
        'Alpha 150ml TRP',
        'Samsol Compact Black 200ml, Samsol White 200ml',
        'Samsol Compact Black 120ml',
        'Samsol Compact Black 120ml, Samsol Yellow 120ml',
        'Mablay 130ml White, Samsol Yellow 200ml',
        'Samsol Yellow 200ml, Alpha 150ml Capsule, Alpha 150ml TRP',
        'Alpha 150ml TRP',
        'Alpha 150ml TRP',
        'Alpha 150ml TRP',
        'Alpha 150ml TRP',
        'Alpha 150ml TRP',
        'Alpha 150ml TRP',
        'Alpha 150ml TRP',
        'Alpha 150ml TRP',
    ]

    daily_sheet_ref = f"'{ws_daily.title}'"
    for idx in range(14):
        src_r = 21 + idx
        dst_r = 3 + idx
        m_dt = ws_src.cell(src_r, 13).value
        wk = ws_src.cell(src_r, 14).value
        d_lbl = ws_src.cell(src_r, 15).value

        ws_weekly.cell(dst_r, 1, m_dt)
        ws_weekly.cell(dst_r, 2, wk)
        ws_weekly.cell(dst_r, 3, d_lbl)
        ws_weekly.cell(dst_r, 4, f'=COUNTIFS({daily_sheet_ref}!$A$3:$A${last_data_row},B{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},">="&A{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},"<="&EOMONTH(A{dst_r},0),{daily_sheet_ref}!$D$3:$D${last_data_row},"<>WEEKLY OFF",{daily_sheet_ref}!$D$3:$D${last_data_row},"<>PUBLIC HOLIDAY")')
        ws_weekly.cell(dst_r, 5, weekly_products[idx])
        ws_weekly.cell(dst_r, 6, f'=SUMIFS({daily_sheet_ref}!$F$3:$F${last_data_row},{daily_sheet_ref}!$A$3:$A${last_data_row},B{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},">="&A{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},"<="&EOMONTH(A{dst_r},0))')
        ws_weekly.cell(dst_r, 7, f'=SUMIFS({daily_sheet_ref}!$G$3:$G${last_data_row},{daily_sheet_ref}!$A$3:$A${last_data_row},B{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},">="&A{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},"<="&EOMONTH(A{dst_r},0))')
        ws_weekly.cell(dst_r, 8, f'=IF(COUNTIFS({daily_sheet_ref}!$A$3:$A${last_data_row},B{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},">="&A{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},"<="&EOMONTH(A{dst_r},0),{daily_sheet_ref}!$G$3:$G${last_data_row},">="&0)=0,"",G{dst_r}-F{dst_r})')

        r_fill = fill_off if dst_r % 2 == 0 else fill_white
        for c in range(1, 9):
            cell = ws_weekly.cell(dst_r, c)
            cell.font = font_body
            cell.fill = r_fill
            cell.border = border_cell
            if c == 1:
                cell.number_format = 'mmm-yy'
                cell.alignment = align_c
            elif c in (2, 3):
                cell.alignment = align_c
            elif c == 5:
                cell.alignment = align_l
            else:
                cell.number_format = '#,##0'
                cell.alignment = align_r

    ws_weekly.cell(17, 1, 'TOTAL')
    ws_weekly.cell(17, 4, '=SUM(D3:D16)')
    ws_weekly.cell(17, 6, '=SUM(F3:F16)')
    ws_weekly.cell(17, 7, '=SUM(G3:G16)')
    ws_weekly.cell(17, 8, '=IF(COUNT(H3:H16)=0,"",SUM(H3:H16))')
    for c in range(1, 9):
        cell = ws_weekly.cell(17, c)
        cell.font = font_bold
        cell.fill = fill_total
        cell.border = border_total
        if c in (4, 6, 7, 8):
            cell.number_format = '#,##0'
            cell.alignment = align_r
        else:
            cell.alignment = align_c

    for c, w in {1: 11, 2: 10, 3: 18, 4: 14, 5: 56, 6: 14, 7: 14, 8: 13}.items():
        ws_weekly.column_dimensions[get_column_letter(c)].width = w

    atomic_save(out_wb, out_path)
    out_wb.close()
    print(f'  [OK] Exported standalone plan: {os.path.basename(out_path)}')
    return out_path


def main(argv=None):
    parser = argparse.ArgumentParser(description='Update Tubex PET & Tubes plans')
    parser.add_argument('--replan', action='store_true',
                        help='Re-cascade PET_Plan forward schedule from max_prod_date + 1')
    args = parser.parse_args(argv)

    excel_path = os.environ.get('TUBEX_FILE') or get_active_tubex_file(DIR)
    if not excel_path:
        raise FileNotFoundError(f'No Tubex*.xlsx found in {DIR}')

    print(f'Update Plans: {os.path.basename(excel_path)} (replan={args.replan})')
    check_not_locked(excel_path)

    wb = openpyxl.load_workbook(excel_path, data_only=False)
    pet_stats = update_pet_plan_sheet(wb, replan=args.replan)
    update_tubes_plan_hook(wb)
    atomic_save(wb, excel_path)

    month_name = detect_plan_month_name(excel_path)
    export_standalone_plan(wb, month_name, pet_stats)
    wb.close()


if __name__ == '__main__':
    main()
