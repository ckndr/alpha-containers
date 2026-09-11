"""
Tubex — Daily Update Master Script
───────────────────────────────────
Complete replacement for Run_All_Updates.bat. One script handles:
  1. Pre-run backup of Excel workbook
  2. Detect fresh ERP exports (inventory, dispatch, dispatch_pet)
  3. Find Imran's production report from Downloads
  4. Apply WIP message from Mehmood
  5. Run full pipeline (production → inventory → dispatch → sort → HTML)
  6. Cross-check results against Imran's source data
  7. Take dashboard screenshot for WhatsApp
  8. Git push to GitHub

Usage:
  python daily.py              (interactive — walks you through everything)
  python daily.py --skip-wip   (skip WIP prompt)
  python daily.py --skip-prod  (skip Production file search)
  python daily.py --skip-git   (skip Git push)

Author: Sikander
"""

import os
import sys
import re
import glob
import shutil
import time
import subprocess
import logging
from datetime import datetime

# ── PATH SETUP ──────────────────────────────────────────────────────────────
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
ALPHA_DIR   = os.path.dirname(SCRIPTS_DIR)

# Guard: If invoked from OneDrive mirror, redirect to master repository at D:\Alpha
if "onedrive" in ALPHA_DIR.lower() and os.path.exists(r"D:\Alpha"):
    print("[NOTICE] Detected execution from OneDrive mirror folder.")
    print("[NOTICE] Switching workspace root to master repository at D:\\Alpha...")
    ALPHA_DIR = r"D:\Alpha"
    SCRIPTS_DIR = os.path.join(ALPHA_DIR, "Scripts")
    if SCRIPTS_DIR not in sys.path:
        sys.path.insert(0, SCRIPTS_DIR)

LOGS_DIR    = os.path.join(ALPHA_DIR, "Logs")
REPORTS_ARCHIVE_DIR = os.path.join(LOGS_DIR, "Reports_Archive")

# ── CONFIGURABLE ────────────────────────────────────────────────────────────
DOWNLOADS_DIR = r"C:\Users\HP\Downloads"

# Production file pattern: "Production report June-2026.xlsx" etc.
PROD_FILE_PATTERN = "Production report "   # files starting with this
PROD_TARGET_NAME  = "Production.xlsx"      # what we rename/copy it to

# How many backup files to keep
MAX_BACKUPS = 15

# ── COLORS (Windows terminal) ──────────────────────────────────────────────
if sys.platform == 'win32':
    os.system('')  # enables ANSI escape codes

GREEN  = '\033[92m'
YELLOW = '\033[93m'
RED    = '\033[91m'
CYAN   = '\033[96m'
BOLD   = '\033[1m'
DIM    = '\033[2m'
RESET  = '\033[0m'

TOTAL_STEPS = 8

def safe_print(msg=""):
    try:
        print(msg)
    except UnicodeEncodeError:
        clean = msg.replace('✓', '[OK]').replace('─', '-').replace('⚠', '[WARN]').replace('✗', '[FAIL]')
        clean = clean.replace('╔', '+').replace('═', '-').replace('╗', '+').replace('║', '|').replace('╚', '+').replace('╝', '+')
        try:
            print(clean)
        except UnicodeEncodeError:
            enc = getattr(sys.stdout, 'encoding', 'ascii') or 'ascii'
            print(clean.encode(enc, errors='replace').decode(enc))

def ok(msg):
    safe_print(f"    {GREEN}✓{RESET} {msg}")

def warn(msg):
    safe_print(f"    {YELLOW}⚠{RESET} {msg}")

def fail(msg):
    safe_print(f"    {RED}✗{RESET} {msg}")

def timed_input(prompt, timeout=2.0):
    print(prompt, end='', flush=True)
    if not sys.stdin or not hasattr(sys.stdin, 'isatty') or not sys.stdin.isatty():
        print()
        return None

    if sys.platform != 'win32':
        try:
            import select
            rlist, _, _ = select.select([sys.stdin], [], [], timeout)
            if rlist:
                return sys.stdin.readline().strip()
            else:
                print()
                return None
        except Exception:
            return input().strip()

    import msvcrt
    start_time = time.time()
    input_str = ""

    while True:
        elapsed = time.time() - start_time
        if not input_str and elapsed >= timeout:
            print()
            return None
        if input_str and elapsed >= (timeout + 10.0):
            print()
            return input_str.strip()

        if msvcrt.kbhit():
            ch = msvcrt.getwch()
            if ch in ('\r', '\n'):
                print()
                return input_str.strip()
            elif ch == '\b':  # Backspace
                if len(input_str) > 0:
                    input_str = input_str[:-1]
                    sys.stdout.write('\b \b')
                    sys.stdout.flush()
            elif ch in ('\x00', '\xe0'):  # Special keys
                if msvcrt.kbhit():
                    msvcrt.getwch()
            elif ord(ch) >= 32:  # Printable
                input_str += ch
                sys.stdout.write(ch)
                sys.stdout.flush()
        time.sleep(0.05)

def header(step, title):
    print(f"\n  {CYAN}{BOLD}[{step}/{TOTAL_STEPS}]{RESET} {BOLD}{title}{RESET}")


def banner():
    msg = f"""
  {CYAN}╔══════════════════════════════════════════════════╗
  ║  {BOLD}Tubex — Daily Update{RESET}{CYAN}                              ║
  ║  {DIM}{datetime.now().strftime('%A, %d-%b-%Y %H:%M')}{RESET}{CYAN}                     ║
  ╚══════════════════════════════════════════════════╝{RESET}
"""
    try:
        print(msg)
    except UnicodeEncodeError:
        ascii_msg = f"""
  +--------------------------------------------------+
     Tubex - Daily Update
     {datetime.now().strftime('%A, %d-%b-%Y %H:%M')}
  +--------------------------------------------------+
"""
        print(ascii_msg)


# ── LOGGING SETUP ──────────────────────────────────────────────────────────
def setup_logging():
    """Mirror all print output to a timestamped log file."""
    os.makedirs(LOGS_DIR, exist_ok=True)
    log_name = f"update_{datetime.now().strftime('%Y%m%d_%H%M')}.log"
    log_path = os.path.join(LOGS_DIR, log_name)

    # Custom stream that writes to both console and file
    class TeeStream:
        def __init__(self, original, log_file):
            self.original = original
            self.log_file = log_file
        def write(self, data):
            try:
                self.original.write(data)
            except UnicodeEncodeError:
                # Fallback: replace common box-drawing / checkmark symbols for console output
                clean_data = data.replace('✓', '[OK]').replace('⚠', '[WARN]').replace('✗', '[FAIL]')
                clean_data = clean_data.replace('─', '-').replace('╔', '+').replace('═', '-').replace('╗', '+')
                clean_data = clean_data.replace('║', '|').replace('╚', '+').replace('╝', '+')
                try:
                    self.original.write(clean_data)
                except UnicodeEncodeError:
                    enc = getattr(self.original, 'encoding', 'ascii') or 'ascii'
                    self.original.write(clean_data.encode(enc, errors='replace').decode(enc))
            self.log_file.write(data)
            self.log_file.flush()
        def flush(self):
            self.original.flush()
            self.log_file.flush()

    log_file = open(log_path, 'w', encoding='utf-8')
    sys.stdout = TeeStream(sys.__stdout__, log_file)
    sys.stderr = TeeStream(sys.__stderr__, log_file)
    return log_path


def check_excel_running():
    """Pre-flight check: verify if Microsoft Excel is running, prompt operator to close or terminate."""
    try:
        res = subprocess.run(["tasklist", "/fi", "imagename eq excel.exe", "/fo", "csv", "/nh"],
                             capture_output=True, text=True)
        out = res.stdout.lower()
        if "excel.exe" in out and "no tasks" not in out:
            warn("Microsoft Excel is currently running!")
            print(f"    {YELLOW}Excel locks workbook files and will cause save failures or corruption.{RESET}")
            if sys.stdin and hasattr(sys.stdin, 'isatty') and sys.stdin.isatty():
                try:
                    ans = input(f"    Close Excel and press Enter, or type 'k' to terminate Excel process: ").strip().lower()
                    if ans == 'k':
                        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], capture_output=True)
                        ok("Excel process terminated.")
                    else:
                        res2 = subprocess.run(["tasklist", "/fi", "imagename eq excel.exe", "/fo", "csv", "/nh"],
                                              capture_output=True, text=True)
                        out2 = res2.stdout.lower()
                        if "excel.exe" in out2 and "no tasks" not in out2:
                            warn("Excel is still running. Proceeding with caution...")
                        else:
                            ok("Excel closed.")
                except (EOFError, KeyboardInterrupt):
                    pass
            else:
                warn("Non-interactive mode: proceeding with caution while Excel runs.")
    except Exception:
        pass


# ═══════════════════════════════════════════════════════════════════════════
# STEP 1: PRE-RUN BACKUP
# ═══════════════════════════════════════════════════════════════════════════
def step_backup():
    header(1, "Pre-run backup & workspace cleanup...")
    os.makedirs(LOGS_DIR, exist_ok=True)

    # Clean orphaned Excel lockfiles (Rule R4-07)
    try:
        from alpha_checks import cleanup_stale_lockfiles
        cleanup_stale_lockfiles(ALPHA_DIR)
    except Exception:
        pass

    timestamp_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_files = sorted(glob.glob(os.path.join(ALPHA_DIR, "Tubex*.xlsx")))

    if not excel_files:
        warn("No Tubex*.xlsx found — nothing to back up")
        return

    import zipfile

    for src in excel_files:
        name = os.path.basename(src)
        
        # Verify file integrity before backing up or operating on it
        if not zipfile.is_zipfile(src):
            warn(f"Corrupted or truncated workbook detected: {name} ({os.path.getsize(src)} bytes)!")
            try:
                subprocess.run(["git", "-C", ALPHA_DIR, "checkout", "HEAD", "--", name], check=True, capture_output=True)
                if zipfile.is_zipfile(src):
                    ok(f"Auto-recovered {name} from git repository (clean state) ✓")
                else:
                    fail(f"Could not recover {name} from git.")
            except Exception as ge:
                warn(f"Git auto-recovery failed: {ge}")

        if zipfile.is_zipfile(src):
            dst = os.path.join(LOGS_DIR, f"backup_daily_{timestamp_str}_{name}")
            shutil.copy2(src, dst)
            ok(f"Backed up: {name}")

    # Clean old backups — keep only the last MAX_BACKUPS of daily automated backups
    backups = sorted(glob.glob(os.path.join(LOGS_DIR, "backup_daily_*.xlsx")),
                     key=os.path.getmtime, reverse=True)
    for old in backups[MAX_BACKUPS:]:
        try:
            os.remove(old)
            print(f"    {DIM}Cleaned old daily backup: {os.path.basename(old)}{RESET}")
        except Exception:
            pass

    # Snapshot raw ERP exports to Logs/ERP_Archives/YYYYMMDD_HHMMSS/
    erp_archive_root = os.path.join(LOGS_DIR, "ERP_Archives")
    os.makedirs(erp_archive_root, exist_ok=True)
    erp_snapshot_dir = os.path.join(erp_archive_root, timestamp_str)

    erp_raw_files = ['inventory.xls', 'dispatch.xls', 'dispatch_pet.xls', PROD_TARGET_NAME]
    archived_count = 0
    for ef in erp_raw_files:
        ef_path = os.path.join(ALPHA_DIR, ef)
        if os.path.exists(ef_path) and os.path.getsize(ef_path) > 0:
            os.makedirs(erp_snapshot_dir, exist_ok=True)
            shutil.copy2(ef_path, os.path.join(erp_snapshot_dir, ef))
            archived_count += 1

    if archived_count > 0:
        ok(f"Archived {archived_count} raw ERP/Production file(s) → Logs/ERP_Archives/{timestamp_str}/")

    # Prune old ERP snapshot folders — keep last MAX_BACKUPS
    existing_dirs = sorted(
        [os.path.join(erp_archive_root, d) for d in os.listdir(erp_archive_root)
         if os.path.isdir(os.path.join(erp_archive_root, d))],
        key=os.path.getmtime,
        reverse=True
    )
    for old_dir in existing_dirs[MAX_BACKUPS:]:
        try:
            shutil.rmtree(old_dir)
            print(f"    {DIM}Cleaned old ERP snapshot: {os.path.basename(old_dir)}{RESET}")
        except Exception:
            pass


# ═══════════════════════════════════════════════════════════════════════════
# STEP 2: CHECK ERP EXPORTS
# ═══════════════════════════════════════════════════════════════════════════
def step_check_erp():
    header(2, "Checking ERP exports...")
    warnings_list = []

    erp_files = {
        'inventory.xls':    'Inventory',
        'dispatch.xls':     'Dispatch (Tube)',
        'dispatch_pet.xls': 'Dispatch (PET)',
    }

    all_ok = True
    for filename, label in erp_files.items():
        target = os.path.join(ALPHA_DIR, filename)

        # Check for "- copy" variant first (Sikander's ERP download workflow)
        import re
        stem, ext = os.path.splitext(filename)
        pattern = re.compile(rf"^{re.escape(stem)} - copy(?: \(\d+\))?{re.escape(ext)}$", re.IGNORECASE)
        copy_matches = [
            os.path.join(ALPHA_DIR, f) for f in os.listdir(ALPHA_DIR)
            if pattern.match(f)
        ]

        if copy_matches:
            latest_copy = max(copy_matches, key=os.path.getmtime)
            age_min = (time.time() - os.path.getmtime(latest_copy)) / 60
            try:
                shutil.copy2(latest_copy, target)
                for old_copy in copy_matches:
                    try:
                        os.remove(old_copy)
                    except Exception:
                        pass
                ok(f"{label}: auto-replaced with fresh copy ({os.path.basename(latest_copy)}, {int(age_min)} min ago)")
            except Exception as ce:
                warn(f"{label}: could not auto-replace: {ce}")

        if os.path.exists(target):
            size_b = os.path.getsize(target)
            if size_b < 1024:
                fail(f"{label}: {filename} is empty or corrupted ({size_b} bytes)")
                warnings_list.append(f"{label}: file ({filename}) is empty/corrupt ({size_b} bytes)")
                all_ok = False
                continue

            age_h = (time.time() - os.path.getmtime(target)) / 3600
            if age_h < 26:
                ok(f"{label}: {filename} ({age_h:.1f}h old)")
            else:
                warn(f"{label}: {filename} is {age_h:.0f}h old (weekend / advisory notice)")
        else:
            fail(f"{label}: {filename} NOT FOUND")
            warnings_list.append(f"{label}: file ({filename}) not found")
            all_ok = False

    if not all_ok:
        print(f"\n    {DIM}Copy fresh ERP exports to {ALPHA_DIR}")
        print(f"    as 'filename - copy.xls' — pipeline auto-replaces.{RESET}")

    return warnings_list


# ═══════════════════════════════════════════════════════════════════════════
# STEP 3: FIND PRODUCTION FILE
# ═══════════════════════════════════════════════════════════════════════════
def step_find_production(skip=False):
    header(3, "Finding & archiving reports (Production & Pending Orders)...")

    os.makedirs(REPORTS_ARCHIVE_DIR, exist_ok=True)

    # 1. Archive ANY pending order files currently in Downloads into Logs/Reports_Archive
    if os.path.isdir(DOWNLOADS_DIR):
        pending_archived = 0
        for f in os.listdir(DOWNLOADS_DIR):
            if f.upper().startswith("PENDING ORDER ") and (f.endswith('.xlsx') or f.endswith('.xls')) and '~$' not in f:
                p_src = os.path.join(DOWNLOADS_DIR, f)
                p_dst = os.path.join(REPORTS_ARCHIVE_DIR, f)
                try:
                    if os.path.exists(p_dst):
                        try: os.remove(p_dst)
                        except Exception: pass
                    shutil.move(p_src, p_dst)
                    pending_archived += 1
                except Exception as pe:
                    warn(f"Could not move pending order file {f}: {pe}")
        if pending_archived > 0:
            ok(f"Archived {pending_archived} Pending Order file(s) from Downloads → Logs/Reports_Archive/")

    if skip:
        warn("Skipped Production report search (--skip-prod)")
        return True

    target = os.path.join(ALPHA_DIR, PROD_TARGET_NAME)

    # 2. Search Downloads for fresh "Production report *.xlsx"
    candidates = []
    if os.path.isdir(DOWNLOADS_DIR):
        for f in os.listdir(DOWNLOADS_DIR):
            if (f.lower().startswith(PROD_FILE_PATTERN.lower())
                    and f.lower().endswith('.xlsx')
                    and '~$' not in f):
                full = os.path.join(DOWNLOADS_DIR, f)
                age_h = (time.time() - os.path.getmtime(full)) / 3600
                candidates.append((full, f, age_h))

    # Sort by modification time — most recent first
    candidates.sort(key=lambda x: x[2])

    if candidates:
        best_path, best_name, best_age = candidates[0]

        if best_age < 1:
            age_str = f"{int(best_age * 60)} min ago"
        elif best_age < 24:
            age_str = f"{best_age:.1f}h ago"
        else:
            age_str = f"{best_age / 24:.1f} days ago"

        print(f"    Found: {CYAN}{best_name}{RESET}  ({age_str})")

        if len(candidates) > 1:
            print(f"    {DIM}({len(candidates)} total files found in Downloads — using most recent){RESET}")

        def is_file_locked(filepath):
            if not os.path.exists(filepath):
                return False
            try:
                with open(filepath, 'r+b'):
                    pass
                return False
            except (IOError, PermissionError):
                return True

        if is_file_locked(target):
            fail(f"Existing {PROD_TARGET_NAME} is open/locked in Microsoft Excel!")
            if sys.stdin and hasattr(sys.stdin, 'isatty') and sys.stdin.isatty():
                try:
                    input(f"    {YELLOW}Please close {PROD_TARGET_NAME} in Excel and press Enter: {RESET}")
                except (EOFError, KeyboardInterrupt):
                    pass

        # Check if we already have the exact same file in Alpha
        should_copy = True
        if os.path.exists(target):
            if os.path.getsize(target) == os.path.getsize(best_path) and abs(os.path.getmtime(target) - os.path.getmtime(best_path)) < 2:
                should_copy = False
                ok(f"Existing {PROD_TARGET_NAME} is already up-to-date")

        if should_copy:
            try:
                shutil.copy2(best_path, target)
                ok(f"Copied → {PROD_TARGET_NAME}")
            except PermissionError:
                fail(f"Could not overwrite {PROD_TARGET_NAME} — file is locked by another program")
                return False

        # Move source file to project archive folder (Logs/Reports_Archive/)
        dest_archive = os.path.join(REPORTS_ARCHIVE_DIR, best_name)
        try:
            if os.path.abspath(best_path) != os.path.abspath(dest_archive):
                if os.path.exists(dest_archive):
                    try:
                        os.remove(dest_archive)
                    except Exception:
                        pass
                shutil.move(best_path, dest_archive)
                ok(f"Archived source report → Logs/Reports_Archive/{best_name}")
        except Exception as me:
            warn(f"Could not move source report to archive: {me}")

        # Also clean any other older production reports lingering in Downloads
        for full_p, f_name, _ in candidates[1:]:
            try:
                d_p = os.path.join(REPORTS_ARCHIVE_DIR, f_name)
                if os.path.exists(d_p):
                    try: os.remove(d_p)
                    except Exception: pass
                shutil.move(full_p, d_p)
            except Exception:
                pass

        return True

    # If not found directly in Downloads root, search project archive folder
    archive_candidates = []
    if os.path.isdir(REPORTS_ARCHIVE_DIR):
        for f in os.listdir(REPORTS_ARCHIVE_DIR):
            if (f.lower().startswith(PROD_FILE_PATTERN.lower())
                    and f.lower().endswith('.xlsx')
                    and '~$' not in f):
                full = os.path.join(REPORTS_ARCHIVE_DIR, f)
                age_h = (time.time() - os.path.getmtime(full)) / 3600
                archive_candidates.append((full, f, age_h))
    if archive_candidates:
        archive_candidates.sort(key=lambda x: x[2])
        best_arch_path, best_arch_name, best_arch_age = archive_candidates[0]
        if not os.path.exists(target) or os.path.getmtime(best_arch_path) > os.path.getmtime(target):
            try:
                shutil.copy2(best_arch_path, target)
                ok(f"Copied from archive ({best_arch_name}) → {PROD_TARGET_NAME}")
                return True
            except Exception:
                pass

    # No fresh candidates found in Downloads root or archive
    if os.path.exists(target):
        age_h = (time.time() - os.path.getmtime(target)) / 3600
        if age_h < 26:
            ok(f"Using existing {PROD_TARGET_NAME} ({age_h:.1f}h old)")
            return True
        else:
            warn(f"Existing {PROD_TARGET_NAME} is {age_h:.0f}h old — no fresh file found in Downloads or archive")
    else:
        fail(f"No 'Production report *.xlsx' found in {DOWNLOADS_DIR} or {REPORTS_ARCHIVE_DIR}")
        fail(f"No existing {PROD_TARGET_NAME} in {ALPHA_DIR}")

    # Manual fallback
    manual = input(f"    {DIM}Drag+drop file here, or Enter to skip: {RESET}").strip().strip('"')
    if manual and os.path.exists(manual):
        shutil.copy2(manual, target)
        ok(f"Copied → {PROD_TARGET_NAME}")
        return True

    warn("Continuing without fresh Production data")
    return True


def get_clipboard_text():
    """Retrieve plain text from Windows clipboard via win32clipboard, tkinter, or powershell."""
    try:
        import win32clipboard
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                return data or ""
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass

    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        data = root.clipboard_get()
        root.destroy()
        return data or ""
    except Exception:
        pass

    try:
        res = subprocess.run(["powershell", "-NoProfile", "-Command", "Get-Clipboard"],
                             capture_output=True, text=True, timeout=2)
        if res.returncode == 0 and res.stdout:
            return res.stdout.strip()
    except Exception:
        pass

    return ""


def is_wip_pattern(text):
    """Check if clipboard text matches a Mehmood-style WIP WhatsApp message."""
    if not text or len(text) > 250:
        return False
    t = text.lower().strip()
    if ('mm' in t or 'kg' in t or '#' in t) and any(c.isdigit() for c in t):
        try:
            sys.path.insert(0, SCRIPTS_DIR)
            from update_wip import parse_wip_message
            res = parse_wip_message(t)
            valid_dias = {12.5, 13.5, 16.0, 16, 19.0, 19, 20.5, 25.0, 25, 30.0, 30, 32.0, 32, 35.0, 35}
            if res and any(d in valid_dias for d in res.keys()):
                return True
        except Exception:
            pass
    return False


def apply_wip(msg):
    """Apply WIP weights to column I of Inventory sheet in active Tubex workbook."""
    if not msg:
        return False
    sys.path.insert(0, SCRIPTS_DIR)
    try:
        from update_wip import parse_wip_message, build_slug_map, find_excel, pick_row
        from update_wip import INV_WIP_COL
        from openpyxl import load_workbook
    except ImportError as e:
        fail(f"Could not import WIP module: {e}")
        return False

    wip_data = parse_wip_message(msg)
    if not wip_data:
        fail("Could not parse WIP message. Expected: #19mm 10kg #30mm 125kg")
        return False

    excel_path, _ = find_excel()
    if not excel_path:
        fail("No Tubex*.xlsx found for WIP update")
        return False

    wb = load_workbook(excel_path)
    ws = wb['Inventory']
    slug_map = build_slug_map(ws)

    # Clear existing WIP
    for dia, rows in slug_map.items():
        for r, issued, name in rows:
            if ws.cell(r, INV_WIP_COL).value is not None:
                ws.cell(r, INV_WIP_COL).value = None

    # Write new
    written = []
    for dia, kg in sorted(wip_data.items()):
        if dia in slug_map:
            row, name = pick_row(slug_map[dia])
            ws.cell(row, INV_WIP_COL).value = kg
            written.append(f"{dia}mm→{kg}kg")

    wb.save(excel_path)
    wb.close()
    if written:
        ok(f"WIP applied: {', '.join(written)}")
        return True
    else:
        warn("No matching slug rows found in Inventory sheet")
        return False


def step_wip(skip=False, wip_text=None):
    header(4, "WIP Ingestion (Mehmood's message)...")

    if skip:
        warn("Skipped (--skip-wip)")
        return None

    if wip_text:
        msg = wip_text.strip()
        print(f"    Using CLI WIP input: {msg}")
        return msg

    clip = get_clipboard_text().strip()
    if is_wip_pattern(clip):
        ok(f"Detected WIP in clipboard: {CYAN}{clip}{RESET}")
        return clip

    msg = timed_input(f"    Paste WIP message (5s timeout, or Enter to skip):\n    {CYAN}>{RESET} ", timeout=5.0)
    if not msg:
        warn("No WIP message — skipping")
        return None

    return msg.strip()


# ═══════════════════════════════════════════════════════════════════════════
# STEP 5: RUN PIPELINE
# ═══════════════════════════════════════════════════════════════════════════
def step_pipeline(wip_msg=None):
    header(5, "Running update pipeline...")

    # Check Excel not open
    excel_files = sorted(glob.glob(os.path.join(ALPHA_DIR, "Tubex*.xlsx")))
    if excel_files:
        try:
            with open(excel_files[-1], 'r+b'):
                pass
        except PermissionError:
            fail(f"{os.path.basename(excel_files[-1])} is OPEN in Excel!")
            fail("Close Excel and run again.")
            return False

    # Clear old mismatches log
    mismatch_log = os.path.join(LOGS_DIR, "mismatches.log")
    if os.path.exists(mismatch_log):
        os.remove(mismatch_log)

    # Pipeline order is CRITICAL:
    # 1. Production first (populates Production_Log that sort_dashboard reads)
    # 2. Inventory (refreshes ERP stock & issuances)
    # 2b. WIP (applied immediately after inventory using fresh issuance data)
    # 3. Dispatch (can add dispatch to inactive products)
    # 4. Sort Dashboard AFTER 1-3 (rearranges based on fresh production+dispatch)
    # 5. HTML last (reads everything)
    scripts = [
        ("update_production.py",  "Production Log + FG Stock"),
        ("update_inventory.py",   "Inventory"),
        ("update_dispatch.py",    "Dispatch"),
        ("sort_dashboard.py",     "Sort Dashboard"),
        ("build_archives.py",     "Build Archives"),
        ("update_html.py",        "HTML Dashboard"),
    ]

    failures = []
    for script_name, label in scripts:
        path = os.path.join(SCRIPTS_DIR, script_name)
        if not os.path.exists(path):
            fail(f"{script_name} NOT FOUND")
            failures.append(label)
            continue

        print(f"\n    {DIM}── {label} ──{RESET}")
        result = subprocess.run(
            [sys.executable, path],
            cwd=SCRIPTS_DIR
        )

        if result.returncode != 0:
            fail(f"{label} FAILED (exit code {result.returncode})")
            failures.append(label)
            # R4-01: Interactive prompt to ask operator whether to continue or stop
            if sys.stdin and sys.stdin.isatty():
                try:
                    ans = input(f"\n    [?] {label} failed (exit code {result.returncode}). Do you want to continue anyway? (y/N): ").strip().lower()
                    if ans not in ('y', 'yes'):
                        print(f"    Pipeline stopped by user after {label} failure.")
                        return False
                except (EOFError, KeyboardInterrupt):
                    return False
            else:
                print(f"    Non-interactive execution: stopping pipeline after {label} failure.")
                return False
        else:
            ok(label)
            if script_name == "update_inventory.py" and wip_msg:
                print(f"\n    {DIM}── Applying WIP weights to Inventory (fresh issuance) ──{RESET}")
                apply_wip(wip_msg)

    # Show mismatches if any
    if os.path.exists(mismatch_log):
        try:
            print(f"\n    {YELLOW}{'─'*50}")
        except UnicodeEncodeError:
            print(f"\n    {YELLOW}{'-'*50}")
        print(f"    {BOLD}MISMATCHES DETECTED:{RESET}")
        with open(mismatch_log, 'r', encoding='utf-8', errors='replace') as f:
            for line in f:
                print(f"    {line.rstrip()}")
        try:
            print(f"    {YELLOW}{'─'*50}{RESET}")
        except UnicodeEncodeError:
            print(f"    {YELLOW}{'-'*50}{RESET}")

    if failures:
        fail(f"Pipeline had {len(failures)} failure(s): {', '.join(failures)}")
        return False

    ok(f"{BOLD}Pipeline completed successfully{RESET}")
    return True


# ═══════════════════════════════════════════════════════════════════════════
# STEP 6: CROSS-CHECK
# ═══════════════════════════════════════════════════════════════════════════
def step_crosscheck():
    header(6, "Cross-checking with Imran's data...")

    critical_errors = []
    pending_warnings = []
    prod_path = os.path.join(ALPHA_DIR, PROD_TARGET_NAME)
    if not os.path.exists(prod_path):
        warn("Production.xlsx not found — skipping machine-level and summary checks")
        critical_errors.append("Production.xlsx not found - skipped cross-checks")
        return critical_errors, pending_warnings

    try:
        import pandas as pd
        from openpyxl import load_workbook
    except ImportError:
        warn("pandas or openpyxl not installed — skipping cross-check")
        critical_errors.append("pandas or openpyxl not installed - skipped cross-checks")
        return critical_errors, pending_warnings

    # --- Part A: Machine-level production totals comparison ---
    try:
        # Detect header row dynamically by scanning the first 10 rows
        raw = pd.read_excel(prod_path, sheet_name='Production Day wise', header=None, nrows=10)
        best_row = 0
        best_score = 0
        KNOWN_KEYWORDS = {'date', 'machines', 'machine', 'customer', 'product name', 'good production'}
        for idx, r_vals in raw.iterrows():
            vals_clean = [str(v).strip().lower() for v in r_vals if str(v).strip() not in ('', 'nan')]
            score = sum(1 for v in vals_clean if any(k in v for k in KNOWN_KEYWORDS))
            if score > best_score:
                best_score = score
                best_row = idx
        
        # Read Imran's raw data using the detected header row
        df_imran = pd.read_excel(prod_path, sheet_name='Production Day wise', header=best_row)

        # Find the good-production and machine columns (flexible matching)
        good_col = None
        for col in df_imran.columns:
            c = str(col).lower().strip()
            if 'good' in c and ('qty' in c or 'prod' in c or 'quantity' in c):
                good_col = col
                break
        if good_col is None:
            for col in df_imran.columns:
                if 'good' in str(col).lower():
                    good_col = col
                    break

        machine_col = None
        for col in df_imran.columns:
            if 'machine' in str(col).lower():
                machine_col = col
                break

        if not good_col or not machine_col:
            warn("Could not identify Machine/Good columns in Imran's file")
            critical_errors.append("Could not identify Machine/Good columns in Production.xlsx")
        else:
            df_imran[good_col] = pd.to_numeric(df_imran[good_col], errors='coerce').fillna(0)
            imran_totals = df_imran.groupby(machine_col)[good_col].sum()

            # Read our Production_Log
            excel_files = sorted(glob.glob(os.path.join(ALPHA_DIR, "Tubex*.xlsx")))
            if not excel_files:
                warn("No Tubex*.xlsx — skipping machine totals check")
                critical_errors.append("No Tubex*.xlsx found for machine totals check")
            else:
                df_dash = pd.read_excel(excel_files[-1], sheet_name='Production_Log', header=1)

                good_col_d = None
                for col in df_dash.columns:
                    if 'good' in str(col).lower():
                        good_col_d = col
                        break
                machine_col_d = None
                for col in df_dash.columns:
                    if 'machine' in str(col).lower():
                        machine_col_d = col
                        break

                if not good_col_d or not machine_col_d:
                    warn("Could not identify columns in Production_Log")
                    critical_errors.append("Could not identify columns in Production_Log")
                else:
                    df_dash[good_col_d] = pd.to_numeric(df_dash[good_col_d], errors='coerce').fillna(0)
                    dash_totals = df_dash.groupby(machine_col_d)[good_col_d].sum()

                    # Compare machine by machine
                    all_machines = sorted(set(imran_totals.index) | set(dash_totals.index))
                    mismatches = []
                    print()
                    for m in all_machines:
                        iv = int(imran_totals.get(m, 0))
                        dv = int(dash_totals.get(m, 0))
                        if iv == dv:
                            ok(f"{str(m):14s} Imran={iv:>8,}  Dashboard={dv:>8,}")
                        else:
                            diff = dv - iv
                            fail(f"{str(m):14s} Imran={iv:>8,}  Dashboard={dv:>8,}  ({diff:+,})")
                            mismatches.append(m)
                            critical_errors.append(f"Machine total mismatch for {m}: Imran={iv:,}, Dashboard={dv:,} (diff={diff:+,})")

                    it = int(imran_totals.sum())
                    dt = int(dash_totals.sum())
                    safe_print(f"    {'─'*52}")

                    if it == dt:
                        ok(f"{'TOTAL':14s} Imran={it:>8,}  Dashboard={dt:>8,}")
                    else:
                        fail(f"{'TOTAL':14s} Imran={it:>8,}  Dashboard={dt:>8,}  ({dt-it:+,})")
                        critical_errors.append(f"Grand Total production mismatch: Imran={it:,}, Dashboard={dt:,} (diff={dt-it:+,})")

                    if mismatches:
                        warn(f"Mismatches in: {', '.join(str(m) for m in mismatches)}")
                    else:
                        ok("All machines match!")
    except Exception as e:
        warn(f"Machine totals cross-check error: {e}")
        critical_errors.append(f"Machine totals cross-check error: {e}")

    # --- Part B: Summary Sheet comparison ---
    try:
        wb_imran = load_workbook(prod_path, data_only=True)
        summary_sheet_name = None
        candidate_summary_sheets = []
        for name in wb_imran.sheetnames:
            n_low = name.lower().strip()
            if "summary" in n_low and "downtime" not in n_low:
                parsed_dt = None
                date_match = re.search(r'(\d{2}[-_/]\d{2}[-_/]\d{4})', name)
                if date_match:
                    raw_dt = date_match.group(1).replace('/', '-').replace('_', '-')
                    try:
                        parsed_dt = datetime.strptime(raw_dt, "%d-%m-%Y")
                    except ValueError:
                        pass
                candidate_summary_sheets.append((name, parsed_dt))

        if candidate_summary_sheets:
            with_dates = [c for c in candidate_summary_sheets if c[1] is not None]
            if with_dates:
                with_dates.sort(key=lambda x: x[1], reverse=True)
                summary_sheet_name = with_dates[0][0]
            else:
                summary_sheet_name = candidate_summary_sheets[0][0]
        
        if not summary_sheet_name:
            warn("No Summary sheet found in Production.xlsx")
            critical_errors.append("No Summary sheet found in Production.xlsx")
        else:
            ws_sum = wb_imran[summary_sheet_name]
            excel_files = sorted(glob.glob(os.path.join(ALPHA_DIR, "Tubex*.xlsx")))
            if not excel_files:
                warn("No Tubex*.xlsx found — skipping summary check")
            else:
                wb_dash = load_workbook(excel_files[-1], data_only=True)
                ws_dash = wb_dash['Tubex_Dashboard']
                
                def _to_int(v):
                    if v is None: return 0
                    try: return int(float(str(v).replace(',', '').strip()))
                    except Exception: return 0

                from openpyxl.utils import get_column_letter

                # Dynamic rule-based metric finder in Imran's summary sheet (Rule R1-17)
                def find_imran_metric(ws, metric_key, fallback_cell):
                    rules = {
                        'tube_prod_today': {
                            'any_type': ['print', 'tube'],
                            'any_time': ['today', 'day', 'daily'],
                            'exclude': ['target', 'plan', 'pending', 'average', 'avg', 'require', 'compliance', '%', 'disp', 'dispatch']
                        },
                        'tube_prod_mtd': {
                            'any_type': ['print', 'tube'],
                            'any_time': ['month', 'monthly', 'mtd', 'total'],
                            'exclude': ['target', 'plan', 'pending', 'average', 'avg', 'require', 'compliance', '%', 'today', 'day', 'disp', 'dispatch']
                        },
                        'pet_prod_today': {
                            'any_type': ['pet'],
                            'any_time': ['today', 'day', 'daily'],
                            'exclude': ['target', 'plan', 'pending', 'average', 'avg', 'require', 'compliance', '%', 'disp', 'dispatch']
                        },
                        'pet_prod_mtd': {
                            'any_type': ['pet'],
                            'any_time': ['month', 'monthly', 'mtd', 'total'],
                            'exclude': ['target', 'plan', 'pending', 'average', 'avg', 'require', 'compliance', '%', 'today', 'day', 'disp', 'dispatch']
                        },
                        'tube_disp_mtd': {
                            'any_type': ['disp', 'dispatch'],
                            'any_time': ['month', 'monthly', 'mtd', 'total'],
                            'exclude': ['pet', 'target', 'plan', 'pending', 'average', 'avg', 'require', 'compliance', '%', 'today', 'day']
                        },
                        'pet_disp_mtd': {
                            'any_type': ['pet'],
                            'any_time': ['disp', 'dispatch'],
                            'any_extra': ['month', 'monthly', 'mtd', 'total'],
                            'exclude': ['target', 'plan', 'pending', 'average', 'avg', 'require', 'compliance', '%', 'today', 'day']
                        },
                    }
                    cfg = rules.get(metric_key, {})
                    for r in range(1, ws.max_row + 1):
                        for c in range(1, min(ws.max_column + 1, 10)):
                            val = ws.cell(r, c).value
                            if not val or not isinstance(val, str):
                                continue
                            txt = val.strip().lower()
                            if any(ex in txt for ex in cfg.get('exclude', [])):
                                continue
                            if not any(t in txt for t in cfg.get('any_type', [])):
                                continue
                            if not any(t in txt for t in cfg.get('any_time', [])):
                                continue
                            if 'any_extra' in cfg and not any(t in txt for t in cfg['any_extra']):
                                continue
                            
                            # Found matching label row! Find numeric value across subsequent columns
                            for val_col in range(c + 1, ws.max_column + 1):
                                num_val = ws.cell(r, val_col).value
                                if num_val is not None and not isinstance(num_val, str):
                                    col_let = get_column_letter(val_col)
                                    return f"{col_let}{r}", _to_int(num_val)
                                elif isinstance(num_val, str):
                                    clean_str = num_val.replace(',', '').strip()
                                    try:
                                        int_v = int(float(clean_str))
                                        col_let = get_column_letter(val_col)
                                        return f"{col_let}{r}", int_v
                                    except ValueError:
                                        pass
                    # Fallback if no dynamic row matched
                    return fallback_cell, _to_int(ws[fallback_cell].value)

                # Dynamically locate KPI card rows on Tubex_Dashboard (in top header section before row 10)
                tube_dash_row = 6
                pet_dash_row = 8
                for r_d in range(1, 10):
                    c_b = str(ws_dash.cell(r_d, 2).value or '').strip().upper()
                    c_d = str(ws_dash.cell(r_d, 4).value or '').strip().upper()
                    if 'TUBE' in c_b and ('MTD' in c_d or 'TUBE' in c_d):
                        tube_dash_row = r_d + 1
                    elif 'PET' in c_b and ('MTD' in c_d or 'PET' in c_d):
                        pet_dash_row = r_d + 1

                # Define checks: (Label, Metric Key, Fallback Imran Cell, Dash Cell)
                checks = [
                    ("Printing Production (Today)", "tube_prod_today", "B14", f"B{tube_dash_row}"),
                    ("Printing Production (MTD)",   "tube_prod_mtd",   "B15", f"D{tube_dash_row}"),
                    ("PET Production (Today)",      "pet_prod_today",  "B3",  f"B{pet_dash_row}"),
                    ("PET Production (MTD)",        "pet_prod_mtd",    "B4",  f"D{pet_dash_row}"),
                    ("Tube Dispatch (MTD)",         "tube_disp_mtd",   "B22", f"J{tube_dash_row}"),
                    ("PET Dispatch (MTD)",          "pet_disp_mtd",    "B11", f"J{pet_dash_row}"),
                ]

                safe_print(f"\n    {DIM}── Summary Sheet ({summary_sheet_name}) vs Dashboard ──{RESET}")
                summary_mismatches = 0
                for label, metric_key, fallback_cell, dash_cell in checks:
                    imran_cell, imran_val = find_imran_metric(ws_sum, metric_key, fallback_cell)
                    dash_val = _to_int(ws_dash[dash_cell].value)
                    
                    if imran_val == dash_val:
                        ok(f"{label:28s} Imran ({imran_cell})={imran_val:>8,}  Dashboard ({dash_cell})={dash_val:>8,}")
                    else:
                        diff = dash_val - imran_val
                        fail(f"{label:28s} Imran ({imran_cell})={imran_val:>8,}  Dashboard ({dash_cell})={dash_val:>8,}  ({diff:+,})")
                        critical_errors.append(f"Summary mismatch - {label}: Imran={imran_val:,}, Dashboard={dash_val:,} (diff={diff:+,})")
                        summary_mismatches += 1
                
                if summary_mismatches == 0:
                    ok("All Summary sheet KPIs match Dashboard!")
                else:
                    warn(f"{summary_mismatches} Summary sheet KPI mismatch(es) found!")
                
                wb_dash.close()
        wb_imran.close()
    except Exception as e:
        warn(f"Summary sheet cross-check error: {e}")
        critical_errors.append(f"Summary sheet cross-check error: {e}")

    # --- Part C: Pending Tube Orders comparison (Line-by-Line & Grand Total - Advisory) ---
    try:
        safe_print(f"\n    {DIM}── Pending Tube Orders: MRP vs PENDING ORDER file (Line-by-Line) ──{RESET}")
        pending_files = []
        for folder in [REPORTS_ARCHIVE_DIR, DOWNLOADS_DIR]:
            if os.path.exists(folder):
                for f in os.listdir(folder):
                    if f.upper().startswith("PENDING ORDER ") and (f.endswith(".xlsx") or f.endswith(".xls")) and '~$' not in f:
                        path = os.path.join(folder, f)
                        pending_files.append((path, os.path.getmtime(path)))

        if not pending_files:
            warn("No PENDING ORDER file found in Logs/Reports_Archive or Downloads — skipping comparison")
            pending_warnings.append("No PENDING ORDER file found for floor comparison")
        else:
            pending_files.sort(key=lambda x: x[1], reverse=True)
            most_recent_pending = pending_files[0][0]
            pending_basename = os.path.basename(most_recent_pending)

            excel_files = sorted(glob.glob(os.path.join(ALPHA_DIR, "Tubex*.xlsx")))
            if not excel_files:
                warn("No Tubex*.xlsx found — skipping Pending Order check")
                pending_warnings.append("No Tubex*.xlsx found for Pending Order check")
            else:
                active_tb = excel_files[-1]
                wb_mrp = load_workbook(active_tb, data_only=True)
                if 'MRP' not in wb_mrp.sheetnames:
                    warn("No MRP sheet found in active Tubex workbook")
                    pending_warnings.append("No MRP sheet found in active Tubex workbook")
                else:
                    ws_mrp = wb_mrp['MRP']
                    mrp_rows = []
                    mrp_total = None
                    for r in range(3, 40):
                        val_e = str(ws_mrp.cell(r, 5).value or '').strip().upper()
                        if val_e == 'TOTAL:':
                            mrp_total = ws_mrp.cell(r, 8).value
                            break
                        dia = ws_mrp.cell(r, 1).value
                        cust = str(ws_mrp.cell(r, 2).value or '').strip()
                        pname = str(ws_mrp.cell(r, 3).value or '').strip()
                        pid = ws_mrp.cell(r, 4).value
                        job = ws_mrp.cell(r, 5).value
                        req = ws_mrp.cell(r, 6).value or 0
                        prod = ws_mrp.cell(r, 7).value or 0
                        bal = ws_mrp.cell(r, 8).value or 0
                        remarks = str(ws_mrp.cell(r, 9).value or '').strip()
                        if pname:
                            try:
                                req_int = int(float(str(req).replace(',', '').strip()))
                                prod_int = int(float(str(prod).replace(',', '').strip()))
                                bal_int = int(float(str(bal).replace(',', '').strip()))
                            except Exception:
                                req_int, prod_int, bal_int = 0, 0, 0
                            mrp_rows.append({
                                'row': r, 'dia': dia, 'customer': cust, 'product': pname,
                                'pid': pid, 'job': job, 'req': req_int, 'prod': prod_int,
                                'bal': bal_int, 'remarks': remarks
                            })
                    wb_mrp.close()

                    if mrp_total is not None:
                        try:
                            mrp_total = int(float(str(mrp_total).replace(',', '').strip()))
                        except Exception:
                            mrp_total = None

                    # Open Imran's pending workbook
                    wb_p = load_workbook(most_recent_pending, data_only=True)
                    today_str = datetime.now().strftime("%d-%m-%Y")
                    target_sheet = None
                    if today_str in wb_p.sheetnames:
                        target_sheet = today_str
                    else:
                        date_sheets = []
                        for s in wb_p.sheetnames:
                            try:
                                dt = datetime.strptime(s.strip(), "%d-%m-%Y")
                                date_sheets.append((s, dt))
                            except ValueError:
                                continue
                        if date_sheets:
                            date_sheets.sort(key=lambda x: x[1], reverse=True)
                            target_sheet = date_sheets[0][0]

                    if not target_sheet:
                        warn(f"No date sheet (DD-MM-YYYY) found in {pending_basename} — skipping comparison")
                        pending_warnings.append(f"No date sheet (DD-MM-YYYY) found in {pending_basename}")
                        wb_p.close()
                    else:
                        ws_p = wb_p[target_sheet]
                        imran_rows = []
                        pending_total = None

                        for r in range(4, ws_p.max_row + 1):
                            c1 = ws_p.cell(r, 1).value
                            c4 = ws_p.cell(r, 4).value
                            c5 = ws_p.cell(r, 5).value
                            c6 = ws_p.cell(r, 6).value
                            c7 = ws_p.cell(r, 7).value
                            c8 = ws_p.cell(r, 8).value

                            s1 = str(c1 or '').strip().upper()
                            s4 = str(c4 or '').strip().upper()

                            if 'GRAND TOTAL' in s1 or 'GRAND TOTAL' in s4:
                                pending_total = c8 or ws_p.cell(r, 2).value or ws_p.cell(r, 7).value
                                continue

                            if any(k in s1 for k in ('TOTAL', 'JOB NO')):
                                continue
                            if any(k in s4 for k in ('TOTAL', 'PRODUCT NAME')):
                                continue
                            if not c4 and not c6 and not c8:
                                continue
                            if c4:
                                try:
                                    req_int = int(float(str(c6 or 0).replace(',', '').strip()))
                                    prod_int = int(float(str(c7 or 0).replace(',', '').strip()))
                                    bal_int = int(float(str(c8 or 0).replace(',', '').strip()))
                                except Exception:
                                    req_int, prod_int, bal_int = 0, 0, 0
                                imran_rows.append({
                                    'row': r, 'job': c1, 'product': str(c4).strip(), 'dia': c5,
                                    'req': req_int, 'prod': prod_int, 'bal': bal_int
                                })

                        if pending_total is not None:
                            try:
                                pending_total = int(float(str(pending_total).replace(',', '').strip()))
                            except Exception:
                                pending_total = None

                        wb_p.close()

                        def _tokens(s):
                            return set(re.findall(r'[a-zA-Z0-9]+', str(s).lower()))

                        matched_mrp_indices = set()
                        line_mismatches = 0

                        for ir in imran_rows:
                            best_idx = None
                            best_score = -1
                            ir_job_str = str(ir['job']).strip() if ir['job'] is not None else ''
                            ir_toks = _tokens(ir['product'])

                            for idx, mr in enumerate(mrp_rows):
                                if idx in matched_mrp_indices:
                                    continue
                                score = 0
                                mr_job_str = str(mr['job']).strip() if mr['job'] is not None else ''
                                if ir_job_str and ir_job_str in mr_job_str:
                                    score += 100
                                try:
                                    if ir['dia'] is not None and mr['dia'] is not None and float(ir['dia']) == float(mr['dia']):
                                        score += 20
                                except Exception:
                                    pass
                                mr_toks = _tokens(mr['product'] + ' ' + mr['customer'])
                                common = ir_toks & mr_toks
                                score += len(common) * 10
                                if score > best_score and score >= 20:
                                    best_score = score
                                    best_idx = idx

                            if best_idx is not None:
                                matched_mrp_indices.add(best_idx)
                                mr = mrp_rows[best_idx]
                                req_ok = (ir['req'] == mr['req'])
                                prod_ok = (ir['prod'] == mr['prod'])
                                bal_ok = (ir['bal'] == mr['bal'])
                                job_disp = f"Job {ir['job']}" if ir['job'] else "No Job #"
                                if req_ok and prod_ok and bal_ok:
                                    ok(f"Dia {ir['dia']}mm | {ir['product']} ({job_disp}): Req={ir['req']:,}, Prod={ir['prod']:,}, Bal={ir['bal']:,}")
                                else:
                                    diff_req = ir['req'] - mr['req']
                                    diff_prod = ir['prod'] - mr['prod']
                                    diff_bal = ir['bal'] - mr['bal']
                                    warn(f"Dia {ir['dia']}mm | Imran '{ir['product']}' vs MRP '{mr['product']}':")
                                    if not req_ok: print(f"         Req:  Imran={ir['req']:,} vs MRP={mr['req']:,} ({diff_req:+,})")
                                    if not prod_ok: print(f"         Prod: Imran={ir['prod']:,} vs MRP={mr['prod']:,} ({diff_prod:+,})")
                                    if not bal_ok: print(f"         Bal:  Imran={ir['bal']:,} vs MRP={mr['bal']:,} ({diff_bal:+,})")
                                    pending_warnings.append(f"Pending Order SKU difference for '{ir['product']}': Req diff={diff_req:+,}, Prod diff={diff_prod:+,}, Bal diff={diff_bal:+,}")
                                    line_mismatches += 1
                            else:
                                warn(f"Unmatched Imran floor line: Dia {ir['dia']}mm | '{ir['product']}' | Bal: {ir['bal']:,}")
                                pending_warnings.append(f"Unmatched Imran floor line in {pending_basename}: '{ir['product']}' (Bal: {ir['bal']:,})")
                                line_mismatches += 1

                        # Print unscheduled MRP orders
                        unscheduled_mrp = [mr for idx, mr in enumerate(mrp_rows) if idx not in matched_mrp_indices]
                        if unscheduled_mrp:
                            safe_print(f"    {CYAN}[INFO]{RESET} Booked in MRP / Awaiting shop-floor schedule:")
                            for mr in unscheduled_mrp:
                                rem = f" ({mr['remarks']})" if mr['remarks'] else ""
                                safe_print(f"      • Dia {mr['dia']}mm | {mr['customer']} - {mr['product']}: Req {mr['req']:,}, Bal {mr['bal']:,}{rem}")

                        unscheduled_mrp_bal = sum(mr['bal'] for mr in unscheduled_mrp)

                        if pending_total is not None and mrp_total is not None:
                            accounted_diff = pending_total - (mrp_total - unscheduled_mrp_bal)
                            if accounted_diff == 0:
                                ok(f"Scheduled Floor Orders: Imran={pending_total:,} == MRP Scheduled Total={mrp_total - unscheduled_mrp_bal:,}")
                                safe_print(f"    {CYAN}ℹ{RESET} Total MRP Order Book: {mrp_total:,} ({pending_total:,} floor + {unscheduled_mrp_bal:,} unscheduled Vince)")
                            else:
                                warn(f"Pending Orders difference: Imran={pending_total:,} vs MRP Scheduled={mrp_total - unscheduled_mrp_bal:,} (diff={accounted_diff:+,})")
                                print(f"    {DIM}(Advisory notice: e.g. Imran pending sheet missing completed batches such as Vince 17k or timing lag — non-blocking){RESET}")
                                pending_warnings.append(f"Pending Orders floor difference: Imran={pending_total:,} vs MRP Scheduled={mrp_total - unscheduled_mrp_bal:,} (diff={accounted_diff:+,})")
                        elif pending_total is not None:
                            ok(f"Imran Floor Total: {pending_total:,}")
    except Exception as e:
        warn(f"Pending Tube Orders cross-check error: {e}")
        pending_warnings.append(f"Pending Tube Orders cross-check error: {e}")

    # --- Part D: Check for unassigned PIDs (PID=0) in Production_Log ---
    try:
        excel_files = sorted(glob.glob(os.path.join(ALPHA_DIR, "Tubex*.xlsx")))
        if excel_files:
            wb_check = load_workbook(excel_files[-1], data_only=True)
            if 'Production_Log' in wb_check.sheetnames:
                ws_plog = wb_check['Production_Log']
                pid_zeros = []
                for r in range(3, ws_plog.max_row + 1):
                    val_pid = ws_plog.cell(r, 6).value
                    val_name = ws_plog.cell(r, 4).value
                    if val_pid == 0 or str(val_pid).strip() == '0':
                        pid_zeros.append(f"Row {r}: {val_name}")
                if pid_zeros:
                    fail(f"Found {len(pid_zeros)} unassigned PID=0 entry/entries in Production_Log: {', '.join(pid_zeros[:3])}")
                    critical_errors.append(f"Unassigned PID=0 in Production_Log: {len(pid_zeros)} entries ({', '.join(pid_zeros[:3])})")
                else:
                    ok("All Production_Log entries have valid assigned PIDs (0 unassigned PID=0)")
            wb_check.close()
    except Exception as e:
        warn(f"PID check error: {e}")
        critical_errors.append(f"PID check error: {e}")

    return critical_errors, pending_warnings


# ═══════════════════════════════════════════════════════════════════════════
# STEP 7: SCREENSHOT
# ═══════════════════════════════════════════════════════════════════════════
def step_screenshot():
    header(7, "Dashboard screenshot...")

    html_path = os.path.join(ALPHA_DIR, "Tubex.html")
    if not os.path.exists(html_path):
        warn("Tubex.html not found")
        return

    date_str = datetime.now().strftime('%Y%m%d')
    ss_path = os.path.join(LOGS_DIR, f"dashboard_{date_str}.png")

    # Try Playwright
    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page(viewport={'width': 430, 'height': 932})
            page.goto(f"file:///{html_path.replace(os.sep, '/')}")
            page.wait_for_timeout(2000)
            try:
                page.click('text=Dashboard', timeout=3000)
                page.wait_for_timeout(1000)
            except Exception:
                pass
            page.screenshot(path=ss_path, full_page=False)
            browser.close()

        ok(f"Saved: Logs/dashboard_{date_str}.png")
        headless_mode = '--headless' in sys.argv or '--no-open' in sys.argv or not (sys.stdin and hasattr(sys.stdin, 'isatty') and sys.stdin.isatty())
        if sys.platform == 'win32' and not headless_mode:
            os.startfile(ss_path)
            print(f"    {DIM}Image opened — share to WhatsApp{RESET}")
        return
    except ImportError:
        pass
    except Exception as e:
        warn(f"Playwright error: {e}")

    # Fallback — just open in browser
    warn("Playwright not installed — opening in browser instead")
    print(f"    {DIM}To enable: pip install playwright && playwright install chromium{RESET}")
    headless_mode = '--headless' in sys.argv or '--no-open' in sys.argv or not (sys.stdin and hasattr(sys.stdin, 'isatty') and sys.stdin.isatty())
    if sys.platform == 'win32' and not headless_mode:
        os.startfile(html_path)


# ═══════════════════════════════════════════════════════════════════════════
# STEP 8: GIT PUSH
# ═══════════════════════════════════════════════════════════════════════════
def step_git_push(skip=False):
    header(8, "Pushing to GitHub...")

    if skip:
        warn("Skipped (--skip-git)")
        return

    # Check git is available
    try:
        subprocess.run(["git", "--version"], capture_output=True, check=True)
    except (FileNotFoundError, subprocess.CalledProcessError):
        warn("Git not installed or not in PATH — skipping push")
        return

    # Check it's a git repo
    result = subprocess.run(
        ["git", "-C", ALPHA_DIR, "rev-parse", "--git-dir"],
        capture_output=True
    )
    if result.returncode != 0:
        warn("Not a git repo — skipping push")
        return

    # Stage all changes
    subprocess.run(["git", "-C", ALPHA_DIR, "add", "-A"], capture_output=True)

    # Check if there's anything to commit
    result = subprocess.run(
        ["git", "-C", ALPHA_DIR, "diff", "--quiet", "--cached"],
        capture_output=True
    )
    if result.returncode == 0:
        ok("No changes to push")
        return

    # Commit and push
    msg = f"Daily update {datetime.now().strftime('%d-%b-%Y %H:%M')}"
    subprocess.run(
        ["git", "-C", ALPHA_DIR, "commit", "-m", msg],
        capture_output=True
    )
    result = subprocess.run(
        ["git", "-C", ALPHA_DIR, "push", "origin", "main"],
        capture_output=True
    )

    if result.returncode == 0:
        ok("Pushed to GitHub ✓")
    else:
        fail("Git push failed — check internet/credentials")
        stderr = result.stderr.decode(errors='replace').strip()
        if stderr:
            print(f"    {DIM}{stderr}{RESET}")


def read_mismatches_log(log_path):
    import json
    import re
    from datetime import datetime
    
    inventory_warnings = []
    mapping_warnings = []
    
    state_file = os.path.join(LOGS_DIR, "previous_missing_items.json")
    state_data = {}
    if os.path.exists(state_file):
        try:
            with open(state_file, 'r', encoding='utf-8') as f:
                state_data = json.load(f)
        except Exception:
            pass

    today_str = datetime.now().strftime('%Y-%m-%d')
    last_run_date = state_data.get("last_run_date", "")
    
    if last_run_date != today_str:
        prev_missing = set(state_data.get("missing_today", []))
        state_data["missing_yesterday"] = list(prev_missing)
        state_data["last_run_date"] = today_str
    else:
        prev_missing = set(state_data.get("missing_yesterday", []))

    current_missing = set()

    # Load required item IDs from active MRP sheet (Rule R1-16 / AUDIT_NOTES.md)
    mrp_required_items = set()
    try:
        from alpha_checks import get_active_tubex_file
        active_tb = get_active_tubex_file(ALPHA_DIR)
        if active_tb and os.path.exists(active_tb):
            wb_mrp = load_workbook(active_tb, data_only=True)
            if 'MRP' in wb_mrp.sheetnames:
                ws_m = wb_mrp['MRP']
                for r_m in range(7, ws_m.max_row + 1):
                    item_id_val = ws_m.cell(row=r_m, column=1).value
                    req_qty_val = ws_m.cell(row=r_m, column=5).value
                    if item_id_val is not None and req_qty_val is not None:
                        try:
                            iid = int(float(str(item_id_val).strip()))
                            rqty = float(str(req_qty_val).replace(',', '').strip())
                            if rqty > 0:
                                mrp_required_items.add(str(iid))
                        except Exception:
                            pass
            wb_mrp.close()
    except Exception:
        pass

    if os.path.exists(log_path):
        with open(log_path, 'r', encoding='utf-8') as f:
            for line in f:
                l = line.strip()
                if not l or l.startswith('---') or l.startswith('==='):
                    continue
                if 'missing from inventory.xls' in l:
                    clean = l.replace('WARNING:', '').strip()
                    lower_clean = clean.lower()
                    
                    # Only highlight Slugs or Resin in daily summary
                    is_slug_or_resin = 'slug' in lower_clean or 'resin' in lower_clean
                    if not is_slug_or_resin:
                        continue
                        
                    # Extract ID
                    m = re.search(r'Item ID\s+(\d+)', clean)
                    item_id = m.group(1) if m else None
                    if item_id:
                        current_missing.add(item_id)
                        tag = "[PERSISTENT]" if (item_id in prev_missing) else "[NEW]"
                        inventory_warnings.append(f"{tag} {clean}")
                else:
                    mapping_warnings.append(l)

    state_data["missing_today"] = list(current_missing)

    try:
        with open(state_file, 'w', encoding='utf-8') as f:
            json.dump(state_data, f)
    except Exception:
        pass

    return inventory_warnings, mapping_warnings


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════
def main():
    banner()

    skip_wip  = '--skip-wip'  in sys.argv
    skip_prod = '--skip-prod' in sys.argv
    skip_git  = '--skip-git'  in sys.argv

    cli_wip = None
    for idx, arg in enumerate(sys.argv):
        if arg == '--wip' and idx + 1 < len(sys.argv):
            wip_parts = []
            for next_arg in sys.argv[idx + 1:]:
                if next_arg.startswith('--'):
                    break
                wip_parts.append(next_arg)
            cli_wip = " ".join(wip_parts)
            break
        elif arg.startswith('--wip='):
            cli_wip = arg.split('=', 1)[1]
            break

    # Start logging
    log_path = setup_logging()

    start = time.time()
    
    all_errors = []

    check_excel_running()      # Pre-flight check: warn or terminate if Excel is open
    step_backup()              # 1. Backup Excel
    
    erp_warnings = step_check_erp()  # 2. Check ERP exports
    all_errors.extend(erp_warnings)
    
    step_find_production(      # 3. Find Production file
        skip=skip_prod)
    wip_msg = step_wip(skip=skip_wip, wip_text=cli_wip)    # 4. WIP ingestion (clipboard / prompt / CLI)
    
    success = step_pipeline(wip_msg=wip_msg)  # 5. Run pipeline (applies WIP after update_inventory)
    if not success:
        all_errors.append("Pipeline execution had failures (one or more scripts failed)")

    # Read mismatches.log (populated during step 5)
    mismatch_log = os.path.join(LOGS_DIR, "mismatches.log")
    inv_warns, map_warns = read_mismatches_log(mismatch_log)

    critical_errors = []
    pending_warnings = []
    if success:
        critical_errors, pending_warnings = step_crosscheck()  # 6. Cross-check
        step_screenshot()                                      # 7. Screenshot
        # Deployment gating: ONLY block Git push if critical core discrepancies exist (Machine/KPI/PID)
        if critical_errors:
            fail(f"CRITICAL: Core cross-check found {len(critical_errors)} discrepancy/discrepancies (Machine/KPI/PID). Gating deployment — skipping Git push to protect production integrity.")
        else:
            if pending_warnings:
                safe_print(f"    {CYAN}ℹ{RESET} {len(pending_warnings)} pending order floor notice(s) detected — non-blocking, proceeding with Git push.")
            step_git_push(skip=skip_git)                       # 8. Git push
    else:
        fail("CRITICAL: Core pipeline experienced failure. Skipping Git push to protect production integrity.")

    elapsed = time.time() - start

    # ── Unified Error Summary ──
    has_blocking = bool(all_errors or inv_warns or map_warns or critical_errors)
    has_issues = bool(has_blocking or pending_warnings)
    error_summary_path = os.path.join(LOGS_DIR, "error_summary.txt")
    
    with open(error_summary_path, 'w', encoding='utf-8') as f_sum:
        def print_both(msg=""):
            print(msg)
            # Remove ANSI colors from file output
            clean_msg = msg
            for color in [GREEN, YELLOW, RED, CYAN, BOLD, DIM, RESET]:
                clean_msg = clean_msg.replace(color, "")
            f_sum.write(clean_msg + "\n")

        if has_issues:
            print_both()
            print_both(f"  {RED if has_blocking else YELLOW}{BOLD}╔══════════════════════════════════════════════════════════╗{RESET}")
            print_both(f"  {RED if has_blocking else YELLOW}{BOLD}║  ⚠ DAILY WORKFLOW ERROR & ADVISORY SUMMARY               ║{RESET}")
            print_both(f"  {RED if has_blocking else YELLOW}{BOLD}╚══════════════════════════════════════════════════════════╝{RESET}")
            print_both()
            
            if all_errors:
                print_both(f"  {YELLOW}{BOLD}[SYSTEM / FILE CHECK ISSUES]{RESET}")
                for err in all_errors:
                    print_both(f"    • {err}")
                print_both()
                
            if inv_warns:
                print_both(f"  {RED}{BOLD}[INVENTORY: CRITICAL ITEMS (SLUG/RESIN) MISSING FROM ERP]{RESET}")
                print_both(f"  {DIM}  (These rows are highlighted in RED in Excel and zeroed out){RESET}")
                for err in inv_warns:
                    print_both(f"    • {err}")
                print_both()
                
            if critical_errors:
                print_both(f"  {RED}{BOLD}[CORE PRODUCTION CROSS-CHECK MISMATCHES (BLOCKING)]{RESET}")
                for err in critical_errors:
                    print_both(f"    • {err}")
                print_both()

            if pending_warnings:
                print_both(f"  {CYAN}{BOLD}[PENDING ORDERS: FLOOR ADVISORY (NON-BLOCKING)]{RESET}")
                print_both(f"  {DIM}  (Shop-floor pending differences e.g. unrecorded prior batches, scheduling delays - does not halt push){RESET}")
                for w in pending_warnings:
                    print_both(f"    • {w}")
                print_both()
                
            if map_warns:
                print_both(f"  {YELLOW}{BOLD}[MAPPING MISMATCHES]{RESET}")
                print_both(f"  {DIM}  (Unmapped products found during production/dispatch/FG processing){RESET}")
                for err in map_warns:
                    print_both(f"    • {err}")
                print_both()
        else:
            print_both()
            print_both(f"  {GREEN}{BOLD}✓ ALL CHECKS PASSED: No errors, missing items, or mismatches detected!{RESET}")

    # ── Final Summary ──
    print(f"\n  {'='*52}")
    if not has_issues:
        print(f"  {GREEN}{BOLD}  ALL DONE{RESET} in {elapsed:.0f} seconds")
    elif not has_blocking:
        print(f"  {GREEN}{BOLD}  ALL CORE STEPS DONE{RESET} in {elapsed:.0f} seconds (Git updated, floor advisories logged)")
    else:
        print(f"  {YELLOW}{BOLD}  COMPLETED WITH BLOCKING ISSUES{RESET} — see summary above (Git push gated)")
    print(f"  {DIM}  Log: {os.path.basename(log_path)}{RESET}")
    print(f"  {DIM}  Error Summary: {os.path.basename(error_summary_path)}{RESET}")
    print(f"  {'='*52}\n")
    if sys.stdin and sys.stdin.isatty():
        try:
            input("  Press Enter to close...")
        except (EOFError, KeyboardInterrupt):
            pass


if __name__ == '__main__':
    main()
