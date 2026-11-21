"""
update_plans.py
Tubex -- Daily Plan Tracker & Forecast Updater (PET + Tubes Hook)
-----------------------------------------------------------------
Reads Production_Log, MRP, Inventory, and BOM from active Tubex_<Mon><YY>.xlsx:
  1. Updates PET_Plan in-place (keeps Baseline Planned Qty static by default;
     populates Actual Good, Wastage, Variance, Status for dates <= max_prod_date;
     updates the 4-row MRP Reconciliation box, Order Progress & Forecast Finish,
     Weekly Summary, and the 12-column Live Material Readiness & Cumulative
     Depletion table across all 9 scheduled batches).
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
PLAN_START_DATE = date(2026, 11, 18)
PRE_PLAN_CUTOFF = date(2026, 11, 17)

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

# Order Progress display sequence (matches Boss's revised production sequence)
ORDER_PROGRESS_ROWS = [
    (8001, '150 ml', 'Alpha Labs PVT LTD', 'Alpha 150ml TRP', date(2027, 2, 2)),
    (8008, '200 ml', 'Samsol', 'Samsol Black 200ml', date(2026, 11, 25)),
    (8007, '200 ml', 'Samsol', 'Samsol White 200ml', date(2026, 11, 27)),
    (8011, '120 ml', 'Samsol', 'Samsol Black 120ml', date(2026, 12, 1)),
    (8005, '120 ml', 'Samsol', 'Samsol Yellow 120ml', date(2026, 12, 5)),
    (8015, '130 ml', 'Mablay Beauty PVT LTD.', 'Mablay 130ml White', date(2026, 12, 8)),
    (8006, '200 ml', 'Samsol', 'Samsol Yellow 200ml', date(2026, 12, 14)),
    (8020, '150 ml', 'Alpha Labs PVT LTD', 'Alpha 150ml Capsule', date(2026, 12, 18)),
]

WEEKLY_ROWS_META = [
    (date(2026, 11, 1), 'Week 3', '18 Nov – 21 Nov', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2026, 11, 1), 'Week 4', '22 Nov – 28 Nov', 'Samsol Compact Black 200ml, Samsol White 200ml', 'Samsol Blk 200ml, Samsol Wht 200ml'),
    (date(2026, 11, 1), 'Week 5', '29 Nov – 30 Nov', 'Samsol Compact Black 120ml', 'Samsol Blk 120ml'),
    (date(2026, 12, 1), 'Week 1', '01 Dec – 07 Dec', 'Samsol Compact Black 120ml, Samsol Yellow 120ml', 'Samsol Blk 120ml, Samsol Ylw 120ml'),
    (date(2026, 12, 1), 'Week 2', '08 Dec – 14 Dec', 'Mablay 130ml White, Samsol Yellow 200ml', 'Mablay Wht 130ml, Samsol Ylw 200ml'),
    (date(2026, 12, 1), 'Week 3', '15 Dec – 21 Dec', 'Samsol Yellow 200ml, Alpha 150ml Capsule, Alpha 150ml TRP', 'Samsol Ylw 200ml, Alpha Cap, Alpha TRP'),
    (date(2026, 12, 1), 'Week 4', '22 Dec – 28 Dec', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2026, 12, 1), 'Week 5', '29 Dec – 31 Dec', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2027, 1, 1),  'Week 1', '01 Jan – 07 Jan', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2027, 1, 1),  'Week 2', '08 Jan – 14 Jan', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2027, 1, 1),  'Week 3', '15 Jan – 21 Jan', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2027, 1, 1),  'Week 4', '22 Jan – 28 Jan', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2027, 1, 1),  'Week 5', '29 Jan – 31 Jan', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
    (date(2027, 2, 1),  'Week 1', '01 Feb – 02 Feb', 'Alpha 150ml TRP', 'Alpha 150ml TRP'),
]

# Scheduled 9-batch sequence for Material Readiness & Cumulative Depletion
SCHEDULED_BATCHES = [
    (1, 'Nov 18–21',    'Alpha Labs PVT LTD',     8001, 'Alpha 150ml TRP (T1)',       'T1'),
    (2, 'Nov 23–25',    'Samsol',                 8008, 'Samsol Compact Black 200ml', 'FULL'),
    (3, 'Nov 26–27',    'Samsol',                 8007, 'Samsol White 200ml',         'FULL'),
    (4, 'Nov 30–Dec 1', 'Samsol',                 8011, 'Samsol Compact Black 120ml', 'FULL'),
    (5, 'Dec 2–5',      'Samsol',                 8005, 'Samsol Yellow 120ml',        'FULL'),
    (6, 'Dec 8',        'Mablay Beauty PVT LTD.', 8015, 'Mablay 130ml White',         'FULL'),
    (7, 'Dec 10–14',    'Samsol',                 8006, 'Samsol Yellow 200ml',        'FULL'),
    (8, 'Dec 16–18',    'Alpha Labs PVT LTD',     8020, 'Alpha 150ml Capsule',        'FULL'),
    (9, 'Dec 21–Feb 2', 'Alpha Labs PVT LTD',     8001, 'Alpha 150ml TRP (T2)',       'T2'),
]

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


def _eval_float(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().lstrip('=').replace(',', '')
    if re.match(r'^[0-9+\-*/().\s]+$', s):
        try:
            return float(eval(s, {'__builtins__': None}, {}))
        except Exception:
            return 0.0
    return 0.0


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
    """
    Returns dict with:
      by_date, by_pid_total, last_date_by_pid, max_prod_date,
      pre_plan_mtd (dt <= 2026-11-17), plan_period_mtd (dt >= 2026-11-18),
      post_plan_by_pid
    """
    ws_pl = wb['Production_Log']
    by_date = {}
    by_pid_total = {}
    post_plan_by_pid = {}
    last_date_by_pid = {}
    max_prod_date = None
    pre_plan_mtd = 0
    plan_period_mtd = 0

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

        if dt <= PRE_PLAN_CUTOFF:
            pre_plan_mtd += good
        else:
            plan_period_mtd += good
            if pid is not None:
                post_plan_by_pid[pid] = post_plan_by_pid.get(pid, 0) + good

        if pid is not None:
            by_pid_total[pid] = by_pid_total.get(pid, 0) + good
            if good > 0:
                if pid not in last_date_by_pid or dt > last_date_by_pid[pid]:
                    last_date_by_pid[pid] = dt

    return {
        'by_date': by_date,
        'by_pid_total': by_pid_total,
        'post_plan_by_pid': post_plan_by_pid,
        'last_date_by_pid': last_date_by_pid,
        'max_prod_date': max_prod_date,
        'pre_plan_mtd': pre_plan_mtd,
        'plan_period_mtd': plan_period_mtd,
    }


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


def read_inventory_stock(wb):
    """Reads live Store + WIP stock by Item ID from Inventory sheet."""
    ws_inv = wb['Inventory']
    stock = {}
    for r in range(3, ws_inv.max_row + 1):
        iid_raw = ws_inv.cell(r, 1).value
        if iid_raw is None:
            continue
        try:
            iid = int(iid_raw)
        except (ValueError, TypeError):
            continue
        opening = _eval_float(ws_inv.cell(r, 5).value)
        received = _eval_float(ws_inv.cell(r, 6).value)
        issued = _eval_float(ws_inv.cell(r, 7).value)
        bal_raw = ws_inv.cell(r, 8).value
        # Col H may be formula '=E+F-G' without cached value if edited via openpyxl
        if isinstance(bal_raw, (int, float)):
            bal = float(bal_raw)
        else:
            bal = opening + received - issued
        wip = _eval_float(ws_inv.cell(r, 9).value)
        stock[iid] = max(0.0, bal) + max(0.0, wip)
    return stock


def read_pet_bom_rates(wb):
    """
    Reads gross BOM rates per 1,000 units: Per 1000 Units * (1 + Scrap %)
    Returns dict {pid: {item_id: gross_per_1k}}.
    """
    ws_bom = wb['BOM']
    bom = {}
    for r in range(2, ws_bom.max_row + 1):
        pid_raw = ws_bom.cell(r, 1).value
        iid_raw = ws_bom.cell(r, 7).value
        if pid_raw is None or iid_raw is None:
            continue
        try:
            pid = int(pid_raw)
            iid = int(iid_raw)
        except (ValueError, TypeError):
            continue
        if pid not in PET_SKU_META:
            continue
        per1k = _eval_float(ws_bom.cell(r, 10).value)
        scrap_raw = ws_bom.cell(r, 12).value
        scrap = float(scrap_raw) if isinstance(scrap_raw, (int, float)) else 0.15
        bom.setdefault(pid, {})[iid] = per1k * (1.0 + scrap)
    return bom


def compute_material_readiness(wb, ws_pp, total_row, pl_info, pet_req, rem_by_pid):
    """
    Computes live SKU-by-SKU 6-column Executive PET Material Status & cumulative
    depletion across all 9 scheduled batches using live Inventory & BOM @ 15% scrap.
    Returns (rows_out, summary_dict).
    """
    inv = read_inventory_stock(wb)
    bom = read_pet_bom_rates(wb)

    # Determine Tranche 1 planned qty (8001 rows before first changeover)
    t1_planned = 0
    for r in range(3, total_row):
        cust = str(ws_pp.cell(r, 4).value or '').strip().upper()
        if cust == 'CHANGEOVER':
            break
        if ws_pp.cell(r, 6).value == 8001:
            t1_planned += _eval_num(ws_pp.cell(r, 7).value)
    if t1_planned == 0:
        t1_planned = 40000

    t1_produced = pl_info['post_plan_by_pid'].get(8001, 0)
    t1_rem = max(0, t1_planned - t1_produced)
    t2_rem = max(0, rem_by_pid.get(8001, 0) - t1_rem)

    init_stock = {
        2680: inv.get(2680, 0.0),  # PET Resin A-84
        7412: inv.get(7412, 0.0),  # Black MB
        5816: inv.get(5816, 0.0),  # White MB
        4002: inv.get(4002, 0.0),  # Yellow MB
        5977: inv.get(5977, 0.0),  # Master Carton
        2267: inv.get(2267, 0.0),  # Poly Bag 20x30
        3006: inv.get(3006, 0.0),  # Poly Bag 24x36
    }
    run = dict(init_stock)
    tot_req = {k: 0.0 for k in init_stock}

    mb_meta = {
        7412: 'Black',
        5816: 'White',
        4002: 'Yellow',
    }

    rows_out = []
    for seq, run_dates, client_lbl, pid, label, kind in SCHEDULED_BATCHES:
        if kind == 'T1':
            rem_qty = t1_rem
        elif kind == 'T2':
            rem_qty = t2_rem
        else:
            rem_qty = rem_by_pid.get(pid, 0)

        rates = bom.get(pid, {})
        k = rem_qty / 1000.0

        # 1. Resin (2680)
        resin_rate = rates.get(2680, 0.0)
        resin_req = k * resin_rate
        tot_req[2680] += resin_req
        resin_bef = run[2680]
        if rem_qty == 0:
            resin_txt = 'Completed'
            resin_ok = True
        elif resin_bef + 1e-6 >= resin_req:
            run[2680] = resin_bef - resin_req
            resin_txt = f'OK ({round(resin_req):,} kg req)'
            resin_ok = True
        elif resin_bef > 0.05:
            deficit = resin_req - resin_bef
            can_make = (resin_bef * 1000.0 / resin_rate) if resin_rate > 0 else 0.0
            run[2680] = 0.0
            resin_txt = f'Short {round(deficit):,} kg (has {round(can_make / 1000.0, 1):.1f}k)'
            resin_ok = False
        else:
            run[2680] = 0.0
            resin_txt = f'Short {round(resin_req):,} kg'
            resin_ok = False

        # 2. Masterbatch (7412 Black, 5816 White, 4002 Yellow)
        mb_id = next((mid for mid in (7412, 5816, 4002) if mid in rates), None)
        if mb_id is None:
            mb_txt = '— (Clear)'
            mb_ok = True
        elif rem_qty == 0:
            mb_txt = 'Completed'
            mb_ok = True
        else:
            mb_name = mb_meta[mb_id]
            mb_rate = rates[mb_id]
            mb_req = k * mb_rate
            tot_req[mb_id] += mb_req
            mb_bef = run[mb_id]
            if mb_bef + 1e-6 >= mb_req:
                run[mb_id] = mb_bef - mb_req
                mb_txt = f'OK ({round(mb_req):,} kg {mb_name})'
                mb_ok = True
            elif mb_bef > 0.05:
                deficit = mb_req - mb_bef
                can_make = (mb_bef * 1000.0 / mb_rate) if mb_rate > 0 else 0.0
                run[mb_id] = 0.0
                mb_txt = f'Short {round(deficit):,} kg {mb_name} (has {round(can_make / 1000.0, 1):.1f}k)'
                mb_ok = False
            else:
                run[mb_id] = 0.0
                mb_txt = f'Short {round(mb_req):,} kg {mb_name}'
                mb_ok = False

        # 3. Packing (Carton 5977, Bag 24x36 3006, Bag 20x30 2267)
        if rem_qty == 0:
            pack_txt = 'Completed'
            pack_ok = True
        else:
            pack_items = [
                (5977, 'Cartons', 'Cartons'),
                (3006, 'Bags 24x36', 'Bags'),
                (2267, 'Bags 20x30', 'Bags'),
            ]
            short_parts = []
            ok_parts = []
            has_ctn = 5977 in rates
            for iid, full_lbl, short_lbl in pack_items:
                if iid not in rates:
                    continue
                req_p = k * rates[iid]
                tot_req[iid] += req_p
                bef_p = run[iid]
                if bef_p + 1e-6 >= req_p:
                    run[iid] = bef_p - req_p
                    ok_parts.append(f'{round(req_p):,} {full_lbl}')
                else:
                    def_p = req_p - bef_p
                    run[iid] = 0.0
                    lbl = short_lbl if has_ctn else full_lbl
                    short_parts.append(f'{round(def_p):,} {lbl}')

            if not short_parts:
                pack_txt = f'OK ({", ".join(ok_parts)})'
                pack_ok = True
            else:
                pack_txt = f'Short {", ".join(short_parts)}'
                pack_ok = False

        rows_out.append({
            'seq': seq,
            'dates': run_dates,
            'client': client_lbl,
            'product': label,
            'rem_qty': rem_qty,
            'resin_txt': resin_txt,
            'resin_ok': resin_ok,
            'mb_txt': mb_txt,
            'mb_ok': mb_ok,
            'pack_txt': pack_txt,
            'pack_ok': pack_ok,
            'is_done': (rem_qty == 0),
        })

    net_resin_def = round(max(0.0, tot_req[2680] - init_stock[2680]))
    net_blk_def = round(max(0.0, tot_req[7412] - init_stock[7412]))
    net_wht_def = round(max(0.0, tot_req[5816] - init_stock[5816]))
    net_yel_def = round(max(0.0, tot_req[4002] - init_stock[4002]))
    net_ctn_def = round(max(0.0, tot_req[5977] - init_stock[5977]))
    net_bag24_def = round(max(0.0, tot_req[3006] - init_stock[3006]))
    net_bag20_def = round(max(0.0, tot_req[2267] - init_stock[2267]))

    summary = {
        'init_resin_kg': round(init_stock[2680]),
        'total_rem_qty': sum(r['rem_qty'] for r in rows_out),
        'total_resin_req': tot_req[2680],
        'total_resin_txt': f'Short {net_resin_def:,} kg Resin',
        'total_mb_txt': f'Short {net_blk_def:,} Blk, {net_wht_def:,} Wht, {net_yel_def:,} Yel kg',
        'total_pack_txt': f'Short {net_ctn_def:,} Ctn, {net_bag24_def:,} Bag 24x36, {net_bag20_def:,} Bag 20x30',
    }
    return rows_out, summary


def simulate_pet_forecast(ws_pp, total_row, rem_by_pid, by_date, last_date_by_pid, max_prod_date):
    """
    Simulates forward from sim_start following the revised sequence:
      8001 (Tranche 1 through Nov 21) -> Sun Nov 22 CO -> 8008 -> 8007 ->
      Nov 28 CO -> 8011 -> 8005 -> Dec 7 CO -> 8015 -> Dec 9 CO -> 8006 ->
      Dec 15 CO -> 8020 -> Dec 19 CO -> remaining 8001.
    """
    sim_start = max((max_prod_date + timedelta(days=1)) if max_prod_date else PLAN_START_DATE, PLAN_START_DATE)

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
        if dt >= PLAN_START_DATE and (not first_co_date or dt < first_co_date) and 8001 in info['pids']
    )

    rem = dict(rem_by_pid)
    forecast_finish = {}
    for pid, rqty in rem.items():
        if rqty <= 0:
            forecast_finish[pid] = last_date_by_pid.get(pid, max_prod_date or PLAN_START_DATE)

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


def _clear_right_side_area(ws, start_col, end_col, max_r):
    """Unmerges, clears, and removes extra columns beyond Order Progress."""
    for rng in list(ws.merged_cells.ranges):
        if rng.max_col >= start_col and rng.max_row <= max_r:
            ws.unmerge_cells(str(rng))
    for r in range(1, max_r + 1):
        for c in range(start_col, max(end_col, ws.max_column) + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()


def _write_material_readiness_block(ws, start_row, start_col, readiness_rows, summary):
    """
    Writes the unified 9-column-grid Material Readiness table (Col 8:9 merged for Packing)
    using clean #D9D9D9 header / #F2F2F2 / #FFFFFF / #FFF2CC standard Excel style,
    Calibri 10pt font, thin black borders, and uniform 17.5pt row heights.
    Returns last row written.
    """
    thin_blk = Side(style='thin', color='000000')
    med_blk = Side(style='medium', color='000000')
    dbl_blk = Side(style='double', color='000000')
    b_cell = Border(left=thin_blk, right=thin_blk, top=thin_blk, bottom=thin_blk)
    b_tot = Border(left=thin_blk, right=thin_blk, top=med_blk, bottom=dbl_blk)

    fill_hdr = PatternFill('solid', fgColor='D9D9D9')
    fill_alt = PatternFill('solid', fgColor='F2F2F2')
    fill_wht = PatternFill('solid', fgColor='FFFFFF')
    fill_hi = PatternFill('solid', fgColor='FFF2CC')

    f_title = Font(name='Calibri', size=10, bold=True, color='000000')
    f_hdr = Font(name='Calibri', size=10, bold=True, color='000000')
    f_body = Font(name='Calibri', size=9.5, bold=False, color='000000')
    f_bold = Font(name='Calibri', size=9.5, bold=True, color='000000')

    a_c = Alignment(horizontal='center', vertical='center', wrap_text=False)
    a_l = Alignment(horizontal='left', vertical='center', wrap_text=False)
    a_r = Alignment(horizontal='right', vertical='center', wrap_text=False)

    end_col = start_col + 8  # 9-col grid (idx 0..8)
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=end_col)
    ws.cell(start_row, start_col, 'PET MATERIAL STATUS (BY PRODUCTION SEQUENCE)')
    for c in range(start_col, end_col + 1):
        cell = ws.cell(start_row, c)
        cell.font = f_title
        cell.fill = fill_hdr
        cell.border = b_cell
        cell.alignment = a_l if c == start_col else a_c

    hdr_row = start_row + 1
    headers = [
        (0, 'Seq'),
        (1, 'Run Dates'),
        (2, 'Client'),
        (3, 'Product'),
        (4, 'Rem. Qty'),
        (5, f'Resin ({summary["init_resin_kg"]:,} kg Stock)'),
        (6, 'Masterbatch'),
        (7, 'Packing (Bags / Cartons)'),
    ]
    ws.merge_cells(start_row=hdr_row, start_column=start_col + 7, end_row=hdr_row, end_column=start_col + 8)
    for ci, h in headers:
        ws.cell(hdr_row, start_col + ci, h)
    for c in range(start_col, end_col + 1):
        cell = ws.cell(hdr_row, c)
        cell.font = f_hdr
        cell.fill = fill_hdr
        cell.border = b_cell
        cell.alignment = a_c

    first_data = hdr_row + 1
    for idx, item in enumerate(readiness_rows):
        rr = first_data + idx
        ws.merge_cells(start_row=rr, start_column=start_col + 7, end_row=rr, end_column=start_col + 8)
        col_vals = {
            0: item['seq'],
            1: item['dates'],
            2: item['client'],
            3: item['product'],
            4: item['rem_qty'],
            5: item['resin_txt'],
            6: item['mb_txt'],
            7: item['pack_txt'],
        }
        r_fill = fill_alt if item['is_done'] else (fill_wht if idx % 2 == 0 else fill_alt)

        for ci, val in col_vals.items():
            ws.cell(rr, start_col + ci, val)

        for ci in range(9):
            c = start_col + ci
            cell = ws.cell(rr, c)
            cell.font = f_bold if ci in (0, 4) else f_body
            if ci == 5 and not item['resin_ok'] and 'has ' in item['resin_txt']:
                cell.fill = fill_hi
            elif ci == 6 and not item['mb_ok'] and 'has ' in item['mb_txt']:
                cell.fill = fill_hi
            else:
                cell.fill = r_fill
            cell.border = b_cell
            if ci in (0, 1):
                cell.alignment = a_c
            elif ci == 4:
                cell.number_format = '#,##0'
                cell.alignment = a_r
            else:
                cell.alignment = a_l

    last_data = first_data + len(readiness_rows) - 1
    tot_row = last_data + 1
    col_rem_let = get_column_letter(start_col + 4)

    ws.merge_cells(start_row=tot_row, start_column=start_col, end_row=tot_row, end_column=start_col + 2)
    ws.merge_cells(start_row=tot_row, start_column=start_col + 7, end_row=tot_row, end_column=start_col + 8)
    ws.cell(tot_row, start_col, 'TOTAL SHORTAGE')
    ws.cell(tot_row, start_col + 3, 'All Remaining Orders')
    ws.cell(tot_row, start_col + 4, f'=SUM({col_rem_let}{first_data}:{col_rem_let}{last_data})')
    ws.cell(tot_row, start_col + 5, summary['total_resin_txt'])
    ws.cell(tot_row, start_col + 6, summary['total_mb_txt'])
    ws.cell(tot_row, start_col + 7, summary['total_pack_txt'])

    for ci in range(9):
        c = start_col + ci
        cell = ws.cell(tot_row, c)
        cell.font = f_bold
        cell.fill = fill_hdr
        cell.border = b_tot
        if ci in (0, 1, 2):
            cell.alignment = a_c
        elif ci == 4:
            cell.number_format = '#,##0'
            cell.alignment = a_r
        else:
            cell.alignment = a_l

    return tot_row


def update_pet_plan_sheet(wb, replan=False):
    if 'PET_Plan' not in wb.sheetnames:
        print('  [WARN] PET_Plan sheet not found in workbook -- skipped')
        return None

    ws_pp = wb['PET_Plan']
    total_row = find_pet_plan_total_row(ws_pp)
    last_data_row = total_row - 1
    pl_info = read_pet_production_log(wb)
    by_date = pl_info['by_date']
    by_pid_total = pl_info['by_pid_total']
    last_date_by_pid = pl_info['last_date_by_pid']
    max_prod_date = pl_info['max_prod_date']
    pet_req = read_mrp_pet_required(wb)

    rem_by_pid = {
        pid: max(0, pet_req.get(pid, 0) - by_pid_total.get(pid, 0))
        for pid in PET_SKU_META
    }

    if replan:
        sim_start = max((max_prod_date + timedelta(days=1)) if max_prod_date else PLAN_START_DATE, PLAN_START_DATE)
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

    # Forecast Finish simulation
    forecast_finish, _ = simulate_pet_forecast(ws_pp, total_row, rem_by_pid, by_date, last_date_by_pid, max_prod_date)

    # Clear & rebuild Right Side cleanly (M..X) so 4-row Reconciliation + shifted tables + Readiness never collide
    _clear_right_side_area(ws_pp, 13, 24, min(60, ws_pp.max_row))

    thin = Side(style='thin', color='BFBFBF')
    med = Side(style='medium', color='808080')
    b_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    b_tot = Border(left=thin, right=thin, top=med, bottom=med)

    fill_hdr = PatternFill('solid', fgColor='D9D9D9')
    fill_off = PatternFill('solid', fgColor='F2F2F2')
    fill_wht = PatternFill('solid', fgColor='FFFFFF')
    fill_co = PatternFill('solid', fgColor='FFF2CC')

    f_title = Font(name='Segoe UI', size=10, bold=True, color='1A1A2E')
    f_hdr = Font(name='Segoe UI', size=9, bold=True, color='1A1A2E')
    f_body = Font(name='Segoe UI', size=9, bold=False, color='1A1A2E')
    f_bold = Font(name='Segoe UI', size=9, bold=True, color='1A1A2E')

    a_c = Alignment(horizontal='center', vertical='center', wrap_text=False)
    a_l = Alignment(horizontal='left', vertical='center', wrap_text=False)
    a_r = Alignment(horizontal='right', vertical='center', wrap_text=False)

    # 1. MRP Reconciliation Box (M3:Q7 — 4 metric rows, Col 1:5 merged title, Col 1:4 labels, Col 5 values)
    ws_pp.merge_cells('M3:Q3')
    ws_pp['M3'] = 'PET MRP & PLAN RECONCILIATION'
    for c in range(13, 18):
        cell = ws_pp.cell(3, c)
        cell.font = f_hdr
        cell.fill = fill_hdr
        cell.border = b_cell
        cell.alignment = a_l if c == 13 else a_c

    recon_rows = [
        (4, 'Total MRP Required (8 SKUs)', '=Q19'),
        (5, 'Pre-Plan MTD Produced (Nov 15–17)', f'=SUMIFS(Production_Log!$H$3:$H$8961,Production_Log!$B$3:$B$8961,"PF Machine",Production_Log!$A$3:$A$8961,"<"&$B$3)'),
        (6, 'Plan Period Produced (From Nov 18)', f'=SUM(H3:H{last_data_row})'),
        (7, 'Current Remaining Balance', '=Q4-Q5-Q6'),
    ]
    for rr, lbl, formula in recon_rows:
        ws_pp.merge_cells(start_row=rr, start_column=13, end_row=rr, end_column=16)
        ws_pp.cell(rr, 13, lbl)
        ws_pp.cell(rr, 17, formula)
        for c in range(13, 18):
            cell = ws_pp.cell(rr, c)
            cell.font = f_bold if rr == 7 else f_body
            cell.fill = fill_co if rr == 7 else fill_wht
            cell.border = b_tot if rr == 7 else b_cell
            cell.alignment = a_l if c < 17 else a_r
            if c == 17:
                cell.number_format = '#,##0'

    # 2. Order Progress & Forecast Finish (M9:U19)
    ws_pp.merge_cells('M9:U9')
    ws_pp['M9'] = 'PET ORDER PROGRESS & FORECAST FINISH'
    for c in range(13, 22):
        cell = ws_pp.cell(9, c)
        cell.font = f_title
        cell.fill = fill_hdr
        cell.border = b_cell
        cell.alignment = a_l if c == 13 else a_c

    r1_hdrs = ['PID', 'ML', 'Client', 'Product', 'MRP Required', 'MTD Produced',
               'Remaining', 'Baseline Finish', 'Forecast Finish']
    for idx, h in enumerate(r1_hdrs):
        c = 13 + idx
        cell = ws_pp.cell(10, c, h)
        cell.font = f_hdr
        cell.fill = fill_hdr
        cell.border = b_cell
        cell.alignment = a_c

    for idx, (pid, ml, client, pname, base_fin) in enumerate(ORDER_PROGRESS_ROWS):
        rr = 11 + idx
        fin_dt = forecast_finish.get(pid, base_fin)
        ws_pp.cell(rr, 13, pid)
        ws_pp.cell(rr, 14, ml)
        ws_pp.cell(rr, 15, client)
        ws_pp.cell(rr, 16, pname)
        ws_pp.cell(rr, 17, f'=IFERROR(INDEX(MRP!$F$96:$F$120,MATCH(M{rr},MRP!$D$96:$D$120,0)),0)')
        ws_pp.cell(rr, 18, f'=SUMIFS(Production_Log!$H$3:$H$8961,Production_Log!$F$3:$F$8961,M{rr})')
        ws_pp.cell(rr, 19, f'=MAX(0,Q{rr}-R{rr})')
        ws_pp.cell(rr, 20, datetime(base_fin.year, base_fin.month, base_fin.day))
        ws_pp.cell(rr, 21, datetime(fin_dt.year, fin_dt.month, fin_dt.day))

        r_fill = fill_off if rr % 2 == 0 else fill_wht
        for c in range(13, 22):
            cell = ws_pp.cell(rr, c)
            cell.font = f_bold if c in (13, 19, 21) else f_body
            cell.fill = fill_co if c == 21 else r_fill
            cell.border = b_cell
            if c in (13, 14):
                cell.alignment = a_c
            elif c in (15, 16):
                cell.alignment = a_l
            elif c in (17, 18, 19):
                cell.number_format = '#,##0'
                cell.alignment = a_r
            elif c in (20, 21):
                cell.number_format = 'dd-mmm-yy'
                cell.alignment = a_c

    ws_pp.cell(19, 16, 'TOTAL')
    ws_pp.cell(19, 17, '=SUM(Q11:Q18)')
    ws_pp.cell(19, 18, '=SUM(R11:R18)')
    ws_pp.cell(19, 19, '=SUM(S11:S18)')
    ws_pp.cell(19, 20, '=MAX(T11:T18)')
    ws_pp.cell(19, 21, '=MAX(U11:U18)')
    for c in range(13, 22):
        cell = ws_pp.cell(19, c)
        cell.font = f_bold
        cell.fill = fill_co
        cell.border = b_tot
        if c in (17, 18, 19):
            cell.number_format = '#,##0'
            cell.alignment = a_r
        elif c in (20, 21):
            cell.number_format = 'dd-mmm-yy'
            cell.alignment = a_c
        else:
            cell.alignment = a_r if c == 16 else a_c

    # 3. Weekly Summary (M21:U36, with T:U merged for Products)
    r2_hdrs = [
        (13, 'Month'), (14, 'Week'), (15, 'Dates'), (16, 'Work Days'),
        (17, 'Planned Qty'), (18, 'Actual Good'), (19, 'Variance'), (20, 'Products'),
    ]
    ws_pp.merge_cells('T21:U21')
    for c_idx, h in r2_hdrs:
        ws_pp.cell(21, c_idx, h)
    for c in range(13, 22):
        cell = ws_pp.cell(21, c)
        cell.font = f_hdr
        cell.fill = fill_hdr
        cell.border = b_cell
        cell.alignment = a_c

    for idx, (m_dt, wk, d_lbl, _prods_full, prods_short) in enumerate(WEEKLY_ROWS_META):
        rr = 22 + idx
        ws_pp.merge_cells(start_row=rr, start_column=20, end_row=rr, end_column=21)
        ws_pp.cell(rr, 13, datetime(m_dt.year, m_dt.month, m_dt.day))
        ws_pp.cell(rr, 14, wk)
        ws_pp.cell(rr, 15, d_lbl)
        ws_pp.cell(rr, 16, f'=COUNTIFS($A$3:$A${last_data_row},N{rr},$B$3:$B${last_data_row},">="&M{rr},$B$3:$B${last_data_row},"<="&EOMONTH(M{rr},0),$D$3:$D${last_data_row},"<>WEEKLY OFF",$D$3:$D${last_data_row},"<>PUBLIC HOLIDAY")')
        ws_pp.cell(rr, 17, f'=SUMIFS($G$3:$G${last_data_row},$A$3:$A${last_data_row},N{rr},$B$3:$B${last_data_row},">="&M{rr},$B$3:$B${last_data_row},"<="&EOMONTH(M{rr},0))')
        ws_pp.cell(rr, 18, f'=SUMIFS($H$3:$H${last_data_row},$A$3:$A${last_data_row},N{rr},$B$3:$B${last_data_row},">="&M{rr},$B$3:$B${last_data_row},"<="&EOMONTH(M{rr},0))')
        ws_pp.cell(rr, 19, f'=IF(COUNTIFS($A$3:$A${last_data_row},N{rr},$B$3:$B${last_data_row},">="&M{rr},$B$3:$B${last_data_row},"<="&EOMONTH(M{rr},0),$H$3:$H${last_data_row},">="&0)=0,"",R{rr}-Q{rr})')
        ws_pp.cell(rr, 20, prods_short)

        r_fill = fill_off if rr % 2 == 0 else fill_wht
        for c in range(13, 22):
            cell = ws_pp.cell(rr, c)
            cell.font = f_body
            cell.fill = r_fill
            cell.border = b_cell
            if c == 13:
                cell.number_format = 'mmm-yy'
                cell.alignment = a_c
            elif c in (14, 15):
                cell.alignment = a_c
            elif c >= 20:
                cell.alignment = a_l
            else:
                cell.number_format = '#,##0'
                cell.alignment = a_r

    ws_pp.merge_cells('T36:U36')
    ws_pp.cell(36, 15, 'TOTAL')
    ws_pp.cell(36, 16, '=SUM(P22:P35)')
    ws_pp.cell(36, 17, '=SUM(Q22:Q35)')
    ws_pp.cell(36, 18, '=SUM(R22:R35)')
    ws_pp.cell(36, 19, '=IF(COUNT(S22:S35)=0,"",SUM(S22:S35))')
    for c in range(13, 22):
        cell = ws_pp.cell(36, c)
        cell.font = f_bold
        cell.fill = fill_co
        cell.border = b_tot
        if c in (16, 17, 18, 19):
            cell.number_format = '#,##0'
            cell.alignment = a_r
        else:
            cell.alignment = a_r if c == 15 else a_c

    # 4. Live Material Status (Unified 9-col grid: M38:U49, Col T:U merged for Packing)
    readiness_rows, mat_summary = compute_material_readiness(wb, ws_pp, total_row, pl_info, pet_req, rem_by_pid)
    _write_material_readiness_block(ws_pp, 38, 13, readiness_rows, mat_summary)

    # Remove any leftover columns beyond Col U (21) and rows beyond TOTAL row (80)
    if ws_pp.max_column > 21:
        ws_pp.delete_cols(22, ws_pp.max_column - 21)
    if ws_pp.max_row > total_row:
        ws_pp.delete_rows(total_row + 1, ws_pp.max_row - total_row)
    for col_let in ['V', 'W', 'X', 'Y', 'Z']:
        ws_pp.column_dimensions.pop(col_let, None)

    # Exact unified column widths across A..U and uniform row heights R1..R80
    pet_plan_widths = {
        1: 9.5, 2: 18.5, 3: 9.0, 4: 24.0, 5: 50.0, 6: 8.0,
        7: 13.5, 8: 13.5, 9: 13.5, 10: 13.5, 11: 18.0,
        12: 3.0,
        13: 9.0, 14: 16.0, 15: 24.0, 16: 31.0, 17: 15.5,
        18: 26.5, 19: 33.0, 20: 17.5, 21: 29.5,
    }
    for c, w in pet_plan_widths.items():
        ws_pp.column_dimensions[get_column_letter(c)].width = w

    ws_pp.row_dimensions[1].height = 26.0
    ws_pp.row_dimensions[2].height = 20.0
    for rr in range(3, total_row + 1):
        ws_pp.row_dimensions[rr].height = 17.5

    print(
        f"  [OK] PET_Plan updated (max PF date={max_prod_date}, "
        f"pre-plan={pl_info['pre_plan_mtd']:,}, plan-period={pl_info['plan_period_mtd']:,}, "
        f"remaining={sum(rem_by_pid.values()):,}, resin_req={mat_summary['total_resin_req']:,.2f} kg)"
    )
    return {
        'by_date': by_date,
        'by_pid_total': by_pid_total,
        'pre_plan_mtd': pl_info['pre_plan_mtd'],
        'plan_period_mtd': pl_info['plan_period_mtd'],
        'pet_req': pet_req,
        'rem_by_pid': rem_by_pid,
        'forecast_finish': forecast_finish,
        'max_prod_date': max_prod_date,
        'total_row': total_row,
        'readiness_rows': readiness_rows,
        'mat_summary': mat_summary,
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

    font_title = Font(name='Segoe UI', size=10, bold=True, color='1A1A2E')
    font_hdr = Font(name='Segoe UI', size=9, bold=True, color='1A1A2E')
    font_body = Font(name='Segoe UI', size=9, bold=False, color='1A1A2E')
    font_bold = Font(name='Segoe UI', size=9, bold=True, color='1A1A2E')
    font_muted = Font(name='Segoe UI', size=9, italic=True, color='595959')

    align_c = Alignment(horizontal='center', vertical='center', wrap_text=False)
    align_l = Alignment(horizontal='left', vertical='center', wrap_text=False)
    align_r = Alignment(horizontal='right', vertical='center', wrap_text=False)

    ws_daily.merge_cells('A1:J1')
    ws_daily['A1'] = 'PET DAILY PRODUCTION PLAN (NOV 2026 - FEB 2027)'
    for c in range(1, 11):
        cell = ws_daily.cell(1, c)
        cell.font = font_title
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_l if c == 1 else align_c
    ws_daily.row_dimensions[1].height = 26.0

    daily_headers = ['Week', 'Date', 'ML', 'Customer', 'Product Name',
                     'Planned Qty', 'Actual Good', 'Wastage', 'Variance', 'Status']
    for c, h in enumerate(daily_headers, 1):
        cell = ws_daily.cell(2, c, h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_c
    ws_daily.row_dimensions[2].height = 20.0

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
        ws_daily.row_dimensions[r].height = 17.5

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

    # 1. Reconciliation Box (L3:P7 — Col 1:5 title, Col 1:4 labels, Col 5 values)
    ws_daily.merge_cells('L3:P3')
    ws_daily['L3'] = 'PET MRP & PLAN RECONCILIATION'
    for c in range(12, 17):
        cell = ws_daily.cell(3, c)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_l if c == 12 else align_c

    pre_plan_mtd = pet_stats['pre_plan_mtd'] if pet_stats else 0
    for rr, lbl, val in [
        (4, 'Total MRP Required (8 SKUs)', '=P19'),
        (5, 'Pre-Plan MTD Produced (Nov 15–17)', pre_plan_mtd),
        (6, 'Plan Period Produced (From Nov 18)', f'=SUM(G3:G{last_data_row})'),
        (7, 'Current Remaining Balance', '=P4-P5-P6'),
    ]:
        ws_daily.merge_cells(start_row=rr, start_column=12, end_row=rr, end_column=15)
        ws_daily.cell(rr, 12, lbl)
        ws_daily.cell(rr, 16, val)
        for c in range(12, 17):
            cell = ws_daily.cell(rr, c)
            cell.font = font_bold if rr == 7 else font_body
            cell.fill = fill_total if rr == 7 else fill_white
            cell.border = border_total if rr == 7 else border_cell
            cell.alignment = align_l if c < 16 else align_r
            if c == 16:
                cell.number_format = '#,##0'

    # 2. Right-side PET PLAN Order Progress & Forecast Finish (L9:T19)
    ws_daily.merge_cells('L9:T9')
    ws_daily['L9'] = 'PET PLAN — ORDER PROGRESS & FORECAST FINISH'
    for c in range(12, 21):
        cell = ws_daily.cell(9, c)
        cell.font = font_title
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_l if c == 12 else align_c

    r1_hdrs = ['PID', 'ML', 'Client', 'Product', 'MRP Required', 'MTD Produced',
               'Remaining', 'Baseline Finish', 'Forecast Finish']
    for idx, h in enumerate(r1_hdrs):
        c = 12 + idx
        cell = ws_daily.cell(10, c, h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_c

    pet_req = pet_stats['pet_req'] if pet_stats else {}
    by_pid = pet_stats['by_pid_total'] if pet_stats else {}
    for idx_ord, (pid_int, _ml, client_short, _pname, _base) in enumerate(ORDER_PROGRESS_ROWS):
        rr = 11 + idx_ord
        ml = ws_src.cell(rr, 14).value
        pname = ws_src.cell(rr, 16).value
        req_q = pet_req.get(pid_int, 0)
        prod_q = by_pid.get(pid_int, 0)
        base_dt = ws_src.cell(rr, 20).value
        fcst_dt = ws_src.cell(rr, 21).value

        ws_daily.cell(rr, 12, pid_int)
        ws_daily.cell(rr, 13, ml)
        ws_daily.cell(rr, 14, client_short)
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

    ws_daily.cell(19, 15, 'TOTAL')
    ws_daily.cell(19, 16, '=SUM(P11:P18)')
    ws_daily.cell(19, 17, '=SUM(Q11:Q18)')
    ws_daily.cell(19, 18, '=SUM(R11:R18)')
    ws_daily.cell(19, 19, '=MAX(S11:S18)')
    ws_daily.cell(19, 20, '=MAX(T11:T18)')
    for c in range(12, 21):
        cell = ws_daily.cell(19, c)
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

    # 3. Live Material Status (Unified 9-col grid: L21:T32, Col S:T merged for Packing)
    if pet_stats and 'readiness_rows' in pet_stats:
        _write_material_readiness_block(
            ws_daily, 21, 12,
            pet_stats['readiness_rows'],
            pet_stats['mat_summary'],
        )

    # Uniform 17.5pt height on all R3..R80 rows
    for rr in range(3, total_row + 1):
        ws_daily.row_dimensions[rr].height = 17.5

    daily_widths = {
        1: 9.5, 2: 18.5, 3: 9.0, 4: 24.0, 5: 50.0,
        6: 13.5, 7: 13.5, 8: 13.5, 9: 13.5, 10: 18.0,
        11: 3.0,
        12: 9.0, 13: 16.0, 14: 24.0, 15: 31.0, 16: 15.5,
        17: 26.5, 18: 33.0, 19: 17.5, 20: 29.5,
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
    ws_weekly.row_dimensions[1].height = 26.0

    wk_headers = ['Month', 'Week', 'Dates', 'Working Days', 'Products', 'Planned Qty', 'Actual Good', 'Variance']
    for c, h in enumerate(wk_headers, 1):
        cell = ws_weekly.cell(2, c, h)
        cell.font = font_hdr
        cell.fill = fill_hdr
        cell.border = border_cell
        cell.alignment = align_c
    ws_weekly.row_dimensions[2].height = 20.0

    daily_sheet_ref = f"'{ws_daily.title}'"
    for idx, (m_dt, wk, d_lbl, prods_full, _prods_short) in enumerate(WEEKLY_ROWS_META):
        dst_r = 3 + idx
        ws_weekly.cell(dst_r, 1, datetime(m_dt.year, m_dt.month, m_dt.day))
        ws_weekly.cell(dst_r, 2, wk)
        ws_weekly.cell(dst_r, 3, d_lbl)
        ws_weekly.cell(dst_r, 4, f'=COUNTIFS({daily_sheet_ref}!$A$3:$A${last_data_row},B{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},">="&A{dst_r},{daily_sheet_ref}!$B$3:$B${last_data_row},"<="&EOMONTH(A{dst_r},0),{daily_sheet_ref}!$D$3:$D${last_data_row},"<>WEEKLY OFF",{daily_sheet_ref}!$D$3:$D${last_data_row},"<>PUBLIC HOLIDAY")')
        ws_weekly.cell(dst_r, 5, prods_full)
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
        ws_weekly.row_dimensions[dst_r].height = 17.5

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
    ws_weekly.row_dimensions[17].height = 17.5

    for c, w in {1: 10.0, 2: 10.0, 3: 18.5, 4: 14.5, 5: 60.0, 6: 14.5, 7: 14.5, 8: 14.5}.items():
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
