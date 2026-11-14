"""
Tubex - Pre-Run Safety Checks
=========================================
Shared helper functions used by all updater scripts.

WHAT THIS MODULE DOES (3 things):
  1. FRESHNESS CHECK — Warns if an ERP file (dispatch.xls, inventory.xls, etc.)
     hasn't been modified recently. Catches the "forgot to download fresh export" mistake.
  2. EXCEL LOCK GUARD — Checks if Tubex*.xlsx is open in Excel.
     If it is, the script stops immediately to prevent silent data corruption.
  3. MISMATCH LOGGING — Saves unmatched product names to Logs/mismatches.log
     so you can review them later without scrolling through console output.

HOW TO USE:
  from alpha_checks import check_freshness, check_not_locked, log_mismatches

  # At the start of your script, after finding files:
  check_freshness(xls_path, max_hours=26)       # warns if file is old
  check_not_locked(excel_path)                    # stops if Excel has it open

  # At the end, after collecting unmatched items:
  log_mismatches("dispatch", unmapped_list)       # saves to Logs/mismatches.log
"""

import os
import re
import glob
import time
from datetime import datetime


# ── Logs folder path (one level up from Scripts/) ────────────────────────────
LOGS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Logs")


def check_freshness(filepath, max_hours=26, label=None):
    """
    Check if a file was modified within the last `max_hours` hours.
    Prints a warning if the file is stale. Does NOT stop the script —
    stale data is still useful (e.g. yesterday's dispatch).

    Args:
        filepath: Full path to the file to check.
        max_hours: How many hours old is "too old". Default 26h
                   (allows for running at different times of day).
        label: Display name for the file. Defaults to basename.

    Returns:
        True if fresh, False if stale.
    """
    if not os.path.exists(filepath):
        name = label or os.path.basename(filepath)
        print("  !! ERROR: %s not found at: %s" % (name, filepath))
        return False

    name = label or os.path.basename(filepath)
    mod_time = os.path.getmtime(filepath)
    age_hours = (time.time() - mod_time) / 3600
    mod_str = datetime.fromtimestamp(mod_time).strftime("%d-%b-%Y %H:%M")

    if age_hours > max_hours:
        print("")
        print("  !! WARNING: %s is %.0f hours old (last modified: %s)" % (name, age_hours, mod_str))
        print("  !! Did you download a fresh export from ERP?")
        print("  !! The script will continue, but data may be outdated.")
        print("")
        return False
    else:
        print("  %s last modified: %s (%.1fh ago) -- OK" % (name, mod_str, age_hours))
        return True


def check_not_locked(filepath, label=None):
    """
    Check if a file is currently open/locked by another process (e.g. Excel).
    If it IS locked, print an error and EXIT the script immediately.
    This prevents openpyxl from silently corrupting data.

    Args:
        filepath: Full path to the .xlsx file.
        label: Display name. Defaults to basename.

    Returns:
        True if file is available (not locked).
        Calls sys.exit(1) if file is locked.
    """
    if not os.path.exists(filepath):
        return True  # file-not-found is handled elsewhere

    name = label or os.path.basename(filepath)

    try:
        # Try to open the file exclusively for writing.
        # If Excel has it open, this will raise PermissionError.
        with open(filepath, 'r+b'):
            pass
        return True
    except PermissionError:
        import sys
        print("")
        print("  " + "=" * 55)
        print("  !! ERROR: %s is open in Excel!" % name)
        print("  !!")
        print("  !! Close Excel first, then run this script again.")
        print("  !! Writing to a locked file causes data corruption.")
        print("  " + "=" * 55)
        print("")
        sys.exit(1)
    except Exception:
        # Other errors (e.g. file doesn't exist) — let the caller handle
        return True


def log_mismatches(source_name, unmapped_items, folder=None):
    """
    Append unmatched product entries to Logs/mismatches.log.
    Each entry includes a timestamp and the source script name.

    Args:
        source_name: Which script found the mismatch (e.g. "dispatch", "production").
        unmapped_items: List of tuples. Each tuple should have at least
                       (product_name, ...) — extra fields are included as-is.
        folder: Override for logs folder path.
    """
    if not unmapped_items:
        return

    log_dir = folder or LOGS_DIR
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_path = os.path.join(log_dir, "mismatches.log")
    timestamp = datetime.now().strftime("%d-%b-%Y %H:%M")

    with open(log_path, "a", encoding="utf-8") as f:
        f.write("\n--- %s | %s ---\n" % (timestamp, source_name))
        for item in unmapped_items:
            if isinstance(item, (tuple, list)):
                f.write("  " + " | ".join(str(x) for x in item) + "\n")
            else:
                f.write("  " + str(item) + "\n")

    print("  Mismatches logged to: Logs/mismatches.log (%d item(s))" % len(unmapped_items))


def replace_copy_export(folder, target_name):
    """
    Look for a '- copy' or '- copy (N)' variant of target_name in folder.
    If found, replace the target with the latest copy (fresh ERP export)
    and delete all other copies.
    Returns True if replacement happened, False otherwise.
    """
    import re
    target_path = os.path.join(folder, target_name)
    stem, ext = os.path.splitext(target_name)
    
    # Matches target_name - copy, target_name - copy (2), etc. case-insensitively
    pattern = re.compile(rf"^{re.escape(stem)} - copy(?: \(\d+\))?{re.escape(ext)}$", re.IGNORECASE)
    
    matches = [
        os.path.join(folder, name)
        for name in os.listdir(folder)
        if pattern.match(name)
    ]

    if not matches:
        return False

    # Find the most recently modified copy file
    latest_copy_path = max(matches, key=os.path.getmtime)
    
    # Verify candidate copy is valid and non-empty (>512 bytes) before replacing (Rule R1-20)
    try:
        if os.path.getsize(latest_copy_path) < 512:
            print("  !! WARNING: Downloaded copy %s is incomplete/empty (<512 bytes). Skipping replacement." % os.path.basename(latest_copy_path))
            return False
    except Exception as e:
        print("  !! WARNING: Could not inspect %s: %s" % (os.path.basename(latest_copy_path), e))
        return False

    # Overwrite the target file with the latest copy
    try:
        os.replace(latest_copy_path, target_path)
        print("  Fresh export found: %s -> %s" % (
            os.path.basename(latest_copy_path),
            os.path.basename(target_path),
        ))
    except Exception as e:
        print("  Error: Could not replace %s with latest copy %s: %s" % (
            os.path.basename(target_path),
            os.path.basename(latest_copy_path),
            e
        ))
        return False

    # Remove all other copies to keep the folder clean
    for match_path in matches:
        if match_path != latest_copy_path:
            try:
                os.remove(match_path)
                print("  Removed older copy file: %s" % os.path.basename(match_path))
            except Exception as e:
                print("  Warning: could not remove older copy file %s: %s" % (
                    os.path.basename(match_path),
                    e
                ))
                
    return True


_MONTHS = {m: i for i, m in enumerate(
    ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"], 1)}

def _tubex_sort_key(path):
    name = os.path.basename(path).lower()
    m = re.fullmatch(r"tubex_([a-z]{3,9})(\d{2})(?:_(\d+))?\.xlsx", name)
    if m and m.group(1)[:3] in _MONTHS:
        return (2, 2000 + int(m.group(2)), _MONTHS[m.group(1)[:3]], int(m.group(3) or 0))
    m = re.fullmatch(r"tubex_v(\d+)_(\d+)\.xlsx", name)
    if m:
        return (1, int(m.group(1)), int(m.group(2)), 0)
    return None   # anything else (temp files, "- Copy", .dry_run.tmp) is ignored

def get_active_tubex_file(folder):
    override = os.environ.get("TUBEX_FILE")
    if override and os.path.exists(override):
        return override
    ranked = []
    for f in glob.glob(os.path.join(folder, "Tubex*.xlsx")):
        key = _tubex_sort_key(f)
        if key is not None:
            ranked.append((key, f))
    return max(ranked)[1] if ranked else None


_FULL_MONTHS = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
}

LEGACY_MONTHS = [
    ("November 2025", None, 11, 2025),
    ("December 2025", None, 12, 2025),
    ("January 2026", None, 1, 2026),
    ("February 2026", None, 2, 2026),
    ("March 2026", None, 3, 2026),
    ("April 2026", None, 4, 2026),
    ("May 2026", None, 5, 2026),
    ("June 2026", None, 6, 2026),
]


def get_month_registry(folder=None):
    """
    Returns a list of (label, path, month_number, year) sorted oldest to newest.
    Built from:
      - constant LEGACY_MONTHS (Nov 2025 to Jun 2026)
      - every Tubex_<Mon><YY>.xlsx in 'Tubex Records'
      - active workbook in project root (from get_active_tubex_file)
    """
    if folder is None:
        folder = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    records_dir = os.path.join(folder, "Tubex Records")
    months_by_key = {}

    for label, path, mnum, yr in LEGACY_MONTHS:
        months_by_key[(yr, mnum)] = (label, path, mnum, yr)

    if os.path.exists(records_dir):
        for f in glob.glob(os.path.join(records_dir, "Tubex*.xlsx")):
            key = _tubex_sort_key(f)
            if key is not None and key[0] == 2:
                yr = key[1]
                mnum = key[2]
                label = f"{_FULL_MONTHS[mnum]} {yr}"
                existing = months_by_key.get((yr, mnum))
                if existing is None or existing[1] is None or key > _tubex_sort_key(existing[1]):
                    months_by_key[(yr, mnum)] = (label, f, mnum, yr)

    active = get_active_tubex_file(folder)
    if active:
        key = _tubex_sort_key(active)
        if key is not None and key[0] == 2:
            yr = key[1]
            mnum = key[2]
            label = f"{_FULL_MONTHS[mnum]} {yr}"
            months_by_key[(yr, mnum)] = (label, active, mnum, yr)

    return [months_by_key[k] for k in sorted(months_by_key.keys())]


def cleanup_stale_lockfiles(folder):
    """
    Remove orphaned Excel lockfiles (~$*.xlsx) that are not currently held by open Excel sessions (Rule R4-07).
    """
    import glob
    lockfiles = glob.glob(os.path.join(folder, "~$*.xlsx"))
    cleaned = 0
    for lf in lockfiles:
        try:
            os.remove(lf)
            cleaned += 1
            print("  Cleaned stale lockfile: %s" % os.path.basename(lf))
        except Exception:
            pass
    return cleaned


def atomic_save(wb, target_path):
    """
    Safely save an openpyxl workbook using an atomic write.
    Writes to a temporary file first, then replaces the target.
    Prevents corrupting/truncating the workbook if interrupted (e.g. Ctrl+C).
    """
    temp_path = target_path + ".tmp"
    try:
        wb.save(temp_path)
        os.replace(temp_path, target_path)
    except Exception:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        raise


def get_inactive_inventory_items():
    """
    Returns a set of integer item IDs configured as inactive in custom_aliases.json.
    """
    import json
    cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "custom_aliases.json")
    if os.path.exists(cfg_path):
        try:
            with open(cfg_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            items = data.get("inactive_inventory_items", [])
            return {int(x) for x in items if str(x).isdigit() or isinstance(x, (int, float))}
        except Exception:
            pass
    return set()


# ── Safe Row Insertion and Deletion Engine (Rule R4-07) ─────────────────────

def _shift_row_bounds(min_row, max_row, idx, amount=1, is_delete=False):
    """
    Calculate shifted row bounds [min_row, max_row] when inserting or deleting
    `amount` rows at 1-based index `idx`.

    Returns:
        (new_min, new_max) tuple or None if the range is completely deleted.
    """
    if not is_delete:
        # INSERT amount rows at idx
        if max_row < idx:
            return (min_row, max_row)
        elif min_row >= idx:
            return (min_row + amount, max_row + amount)
        else:  # min_row < idx <= max_row: insertion inside range expands it
            return (min_row, max_row + amount)
    else:
        # DELETE amount rows starting at idx (inclusive: idx .. idx + amount - 1)
        del_start = idx
        del_end = idx + amount - 1

        if max_row < del_start:
            return (min_row, max_row)
        elif min_row > del_end:
            return (min_row - amount, max_row - amount)
        elif min_row >= del_start and max_row <= del_end:
            # Entirely within deleted region
            return None
        elif min_row < del_start and max_row > del_end:
            # Deleted region is strictly inside the range
            new_min = min_row
            new_max = max_row - amount
            return (new_min, new_max) if new_min <= new_max else None
        elif min_row < del_start and del_start <= max_row <= del_end:
            # End of range falls inside deleted region
            new_min = min_row
            new_max = del_start - 1
            return (new_min, new_max) if new_min <= new_max else None
        elif del_start <= min_row <= del_end and max_row > del_end:
            # Start of range falls inside deleted region
            new_min = del_start
            new_max = max_row - amount
            return (new_min, new_max) if new_min <= new_max else None

        return None


def _shift_range_list(ranges, idx, amount=1, is_delete=False):
    """
    Shift an iterable of CellRange objects (or coordinate strings) by `amount` rows at `idx`.
    Returns a list of shifted CellRange objects, omitting any that were completely deleted.
    """
    from openpyxl.worksheet.cell_range import CellRange
    shifted = []
    for r in ranges:
        if isinstance(r, str):
            cr = CellRange(r)
        else:
            cr = r
        bounds = _shift_row_bounds(cr.min_row, cr.max_row, idx, amount, is_delete)
        if bounds is not None:
            new_min_r, new_max_r = bounds
            if new_min_r <= new_max_r:
                shifted.append(CellRange(min_col=cr.min_col, min_row=new_min_r,
                                         max_col=cr.max_col, max_row=new_max_r))
    return shifted


def _apply_safe_row_shift(ws, idx, amount=1, is_delete=False):
    """
    Internal engine for safe_insert_rows and safe_delete_rows.
    Preserves and shifts merged cells, row_dimensions (height, hidden, styles),
    auto_filter, data_validations, conditional_formatting, and freeze_panes.
    """
    from collections import OrderedDict
    from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
    from openpyxl.formatting.formatting import ConditionalFormatting
    from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

    if amount <= 0:
        return

    # 1. Snapshot & unmerge any ws.merged_cells.ranges with max_row >= idx BEFORE row operation
    merged_to_shift = []
    for cr in list(ws.merged_cells.ranges):
        if cr.max_row >= idx:
            merged_to_shift.append((cr.min_col, cr.min_row, cr.max_col, cr.max_row))
            ws.unmerge_cells(range_string=cr.coord)

    # 2. Call openpyxl's native row operation
    if not is_delete:
        ws.insert_rows(idx, amount)
    else:
        ws.delete_rows(idx, amount)

    # 1b. Re-merge shifted valid ranges (skipping collapsed or 1x1 ranges)
    for min_col, min_row, max_col, max_row in merged_to_shift:
        bounds = _shift_row_bounds(min_row, max_row, idx, amount, is_delete)
        if bounds is not None:
            new_min_row, new_max_row = bounds
            if new_min_row <= new_max_row:
                # Skip 1x1 cells
                if not (min_col == max_col and new_min_row == new_max_row):
                    ws.merge_cells(start_row=new_min_row, start_column=min_col,
                                   end_row=new_max_row, end_column=max_col)

    # 3. Shift ws.row_dimensions keys cleanly
    # (insert: >= idx by +amount; delete: > idx + amount - 1 by -amount; drop deleted; clear newly inserted)
    new_row_dims = {}
    del_start = idx
    del_end = idx + amount - 1
    for r, dim in list(ws.row_dimensions.items()):
        if not is_delete:
            if r >= idx:
                new_r = r + amount
                dim.index = new_r
                new_row_dims[new_r] = dim
            else:
                new_row_dims[r] = dim
        else:
            if r < del_start:
                new_row_dims[r] = dim
            elif r > del_end:
                new_r = r - amount
                dim.index = new_r
                new_row_dims[new_r] = dim
            # keys inside [del_start, del_end] are omitted (deleted)
    ws.row_dimensions.clear()
    ws.row_dimensions.update(new_row_dims)

    # 4. Shift ws.auto_filter.ref using _shift_row_bounds, preserving <filterColumn> criteria
    if ws.auto_filter and ws.auto_filter.ref:
        cr = CellRange(ws.auto_filter.ref)
        bounds = _shift_row_bounds(cr.min_row, cr.max_row, idx, amount, is_delete)
        if bounds is not None:
            new_min_r, new_max_r = bounds
            new_cr = CellRange(min_col=cr.min_col, min_row=new_min_r,
                               max_col=cr.max_col, max_row=new_max_r)
            ws.auto_filter.ref = new_cr.coord
        else:
            ws.auto_filter.ref = None

    # 5. Shift ws.data_validations.dataValidation sqref ranges (drop DV if all ranges deleted)
    if hasattr(ws, 'data_validations') and ws.data_validations and hasattr(ws.data_validations, 'dataValidation'):
        to_remove = []
        for dv in list(ws.data_validations.dataValidation):
            if dv.sqref:
                new_ranges = _shift_range_list(dv.sqref.ranges, idx, amount, is_delete)
                if new_ranges:
                    dv.sqref = MultiCellRange(new_ranges)
                else:
                    to_remove.append(dv)
        for dv in to_remove:
            ws.data_validations.dataValidation.remove(dv)

    # 6. Shift ws.conditional_formatting._cf_rules safely by rebuilding OrderedDict
    # (NEVER mutate cf.sqref in-place while it's a dict key in _cf_rules!)
    if hasattr(ws, 'conditional_formatting') and ws.conditional_formatting and hasattr(ws.conditional_formatting, '_cf_rules'):
        new_cf_rules = OrderedDict()
        for cf, rules in ws.conditional_formatting._cf_rules.items():
            if cf.sqref:
                new_ranges = _shift_range_list(cf.sqref.ranges, idx, amount, is_delete)
                if new_ranges:
                    new_cf = ConditionalFormatting(
                        sqref=MultiCellRange(new_ranges),
                        pivot=getattr(cf, 'pivot', None),
                        extLst=getattr(cf, 'extLst', None)
                    )
                    new_cf_rules[new_cf] = rules
            else:
                new_cf_rules[cf] = rules
        ws.conditional_formatting._cf_rules = new_cf_rules

    # 7. Shift ws.freeze_panes if set and row is at/below idx
    if ws.freeze_panes and isinstance(ws.freeze_panes, str):
        row, col = coordinate_to_tuple(ws.freeze_panes)
        if not is_delete:
            if row >= idx:
                new_row = row + amount
                ws.freeze_panes = f"{get_column_letter(col)}{new_row}"
        else:
            if row > del_end:
                new_row = row - amount
                ws.freeze_panes = f"{get_column_letter(col)}{new_row}"
            elif del_start <= row <= del_end:
                new_row = max(1, idx)
                ws.freeze_panes = f"{get_column_letter(col)}{new_row}"


def safe_insert_rows(ws, idx, amount=1):
    """
    Safely insert `amount` rows at 1-based index `idx` in worksheet `ws`.
    Maintains merged cells, row_dimensions (heights, hidden), auto_filter,
    data_validations, conditional_formatting, and freeze_panes without corruption.
    """
    _apply_safe_row_shift(ws, idx, amount=amount, is_delete=False)


def safe_delete_rows(ws, idx, amount=1):
    """
    Safely delete `amount` rows starting at 1-based index `idx` in worksheet `ws`.
    Maintains merged cells, row_dimensions (heights, hidden), auto_filter,
    data_validations, conditional_formatting, and freeze_panes without corruption.
    """
    _apply_safe_row_shift(ws, idx, amount=amount, is_delete=True)





