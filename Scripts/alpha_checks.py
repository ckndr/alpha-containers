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
        return True  # file-not-found is handled elsewhere

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
