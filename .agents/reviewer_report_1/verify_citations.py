import os
import re

def check_citation(label, filepath, start_line, end_line, expected_snippets):
    full_path = os.path.join('d:\\Alpha', filepath)
    if not os.path.exists(full_path):
        return False, f"File not found: {filepath}", ""
    
    with open(full_path, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.readlines()
    
    total_lines = len(lines)
    if start_line > total_lines:
        return False, f"Start line {start_line} exceeds total lines {total_lines}", ""
    
    actual_chunk = "".join(lines[max(0, start_line - 1):min(total_lines, end_line)])
    
    missing = []
    for snippet in expected_snippets:
        # Normalize whitespace
        snip_clean = " ".join(snippet.split())
        chunk_clean = " ".join(actual_chunk.split())
        if snip_clean not in chunk_clean:
            # Also search entire file to see where snippet actually is if line number shifted
            file_clean = " ".join("".join(lines).split())
            if snip_clean in file_clean:
                # Find approx line number
                matching_lines = [i+1 for i, l in enumerate(lines) if any(word in l for word in snippet.split()[:3])]
                missing.append(f"Snippet '{snippet[:40]}...' found elsewhere in file around lines {matching_lines[:5]} but not in range [{start_line}, {end_line}]")
            else:
                missing.append(f"Snippet '{snippet[:40]}...' NOT found anywhere in {filepath}")
    
    if missing:
        return False, "; ".join(missing), actual_chunk
    return True, f"Lines {start_line}-{end_line} verified OK", actual_chunk

citations = [
    # R1-01
    ("R1-01 (Part 1)", "Scripts/update_production.py", 598, 641, ["is_varnish", "session_overrides", "sys.stdin.isatty()", "Enter PID for"]),
    ("R1-01 (Part 2)", "Scripts/update_production.py", 1076, 1085, ["interactive", "pid"]),
    
    # R1-02
    ("R1-02 (Guardrail)", "Scripts/update_inventory.py", 219, 223, ["len(xls_items) < 5", "WARNING: ERP inventory export has only"]),
    ("R1-02 (Zeroing)", "Scripts/update_inventory.py", 277, 297, ["Not active in ERP", "0.0", "Font("]),
    
    # R1-03
    ("R1-03", "Scripts/sort_dashboard.py", 389, 394, ["orders_val = data['orders']", "re.sub(r'(?<![!$\\w])([FD])(\\d+)\\b'"]),
    
    # R1-04
    ("R1-04 (Python)", "Scripts/sort_dashboard.py", 130, 138, ["is_print = mach_up.startswith('PRINT') or mach_up.startswith('PLINE')"]),
    ("R1-04 (Formula)", "Scripts/sort_dashboard.py", 320, 326, ["LEFT(Production_Log!$B$3:$B", "Print", "PLINE"]),
    ("R1-04 (Line 596)", "Scripts/sort_dashboard.py", 590, 605, ["PLINE", "Print"]),
    
    # R1-05
    ("R1-05", "Scripts/sort_dashboard.py", 315, 335, ["pl_max_row = max(ws_pl.max_row, 1000)"]),
    
    # R1-06
    ("R1-06", "Scripts/update_dispatch.py", 175, 265, ["xldate_as_datetime", "dayfirst=True", "serial"]),
    
    # R1-07
    ("R1-07", "Scripts/update_dispatch.py", 189, 205, ["col_disp_idx = 7", "disp", "qty"]),
    
    # R1-08
    ("R1-08", "Scripts/update_production.py", 785, 799, ["FG Stock In hand", "raw_head", "header_row"]),
    
    # R1-09
    ("R1-09", "Scripts/update_production.py", 515, 546, ["dayfirst=True", "2020 <= d.year <= 2035"]),
    
    # R1-10
    ("R1-10", "Scripts/update_inventory.py", 97, 139, ["col_id = 0", "col_name = 1", "col_opening = 3"]),
    
    # R1-11
    ("R1-11", "Scripts/update_inventory.py", 192, 200, ["Slugs & Raw Materials Inventory", "date_range"]),
    
    # R1-12
    ("R1-12", "Scripts/update_production.py", 921, 931, ["max_c = max(ws.max_column or 8, 12)", "cell.value = None"]),
    
    # R1-13
    ("R1-13", "Scripts/update_html.py", 226, 241, ["cat_pid_type", "ml", "8000 <= pid_k < 9000", "TUBE", "PET"]),
    
    # R1-14
    ("R1-14 (daily.py)", "Scripts/daily.py", 441, 455, ["update_production.py", "update_inventory.py", "update_dispatch.py", "sort_dashboard.py", "build_archives.py", "update_html.py"]),
    
    # R1-15
    ("R1-15", "Scripts/daily.py", 170, 185, ["encoding='utf-8'"]),
    
    # R1-16
    ("R1-16", "Scripts/daily.py", 968, 1030, ["req_qty > 0", "PERSISTENT", "NEW"]),
    
    # R1-17
    ("R1-17", "Scripts/daily.py", 657, 705, ["imran_labels", "Summary", "print", "pet"]),
    
    # R1-18
    ("R1-18", "Scripts/alpha_checks.py", 45, 55, ["not os.path.exists(filepath)", "return False"]),
    
    # R1-19
    ("R1-19", "Scripts/alpha_checks.py", 34, 68, ["check_freshness"]),
    
    # R1-20
    ("R1-20", "Scripts/alpha_checks.py", 144, 206, ["os.path.getsize(latest_copy_path) >= 512", "os.replace"]),
    
    # R1-21
    ("R1-21", "Scripts/customer_normalization.py", 77, 95, ["len(raw_upper) >= 4", "issubset"]),
    
    # R1-22
    ("R1-22 (alpha_checks)", "Scripts/alpha_checks.py", 209, 220, ["get_active_tubex_file", "sorted(excels)[-1]"]),
    
    # R4-01
    ("R4-01", "Scripts/daily.py", 464, 480, ["result.returncode != 0", "sys.stdin.isatty()", "Do you want to continue anyway?"]),
    
    # R4-02
    ("R4-02", "Scripts/daily.py", 1065, 1080, ["if success:", "step_onedrive_backup()", "step_git_push", "CRITICAL: Core pipeline experienced failure"]),
    
    # R4-03 (update_html)
    ("R4-03 (update_html)", "Scripts/update_html.py", 40, 75, ["DispatchEx", "try:", "finally:", "wb.Close", "excel.Quit()"]),
    # R4-03 (build_archives)
    ("R4-03 (build_archives)", "Scripts/build_archives.py", 108, 185, ["DispatchEx", "try:", "finally:", "excel.Quit()"]),
    
    # R4-04
    ("R4-04", "Scripts/daily.py", 968, 1030, ["req_qty > 0", "PERSISTENT"]),
    
    # R4-05
    ("R4-05 (Push.bat)", "Scripts/Push.bat", 10, 20, ["OneDrive\\Alpha"]),
    ("R4-05 (daily.py)", "Scripts/daily.py", 860, 875, ["OneDrive", "Alpha"]),
    
    # R4-06
    ("R4-06 (daily.py)", "Scripts/daily.py", 865, 880, ["/E", "/COPY:DAT"]),
    ("R4-06 (Push.bat)", "Scripts/Push.bat", 30, 45, ["/E", "/COPY:DAT"]),
    
    # R4-07
    ("R4-07 (alpha_checks)", "Scripts/alpha_checks.py", 222, 238, ["cleanup_stale_lockfiles", "~$*"]),
    ("R4-07 (daily.py)", "Scripts/daily.py", 185, 195, ["cleanup_stale_lockfiles"]),
    ("R4-07 (Push.bat)", "Scripts/Push.bat", 30, 45, ["/XF \"~$*\""]),
    
    # R4-08
    ("R4-08 (PIPELINE.md)", "PIPELINE.md", 20, 45, ["update_production.py", "update_inventory.py", "update_dispatch.py", "sort_dashboard.py", "build_archives.py", "update_html.py"]),
    ("R4-08 (DAILY_WORKFLOW.md)", "DAILY_WORKFLOW.md", 70, 95, ["update_production.py", "update_inventory.py", "update_dispatch.py", "sort_dashboard.py", "build_archives.py", "update_html.py"]),
]

results = []
passed_count = 0
for label, fpath, sline, eline, snips in citations:
    ok, msg, chunk = check_citation(label, fpath, sline, eline, snips)
    if ok:
        passed_count += 1
        print(f"[PASS] {label:25s} in {fpath}:{sline}-{eline}")
    else:
        print(f"[FAIL] {label:25s} in {fpath}:{sline}-{eline} -> {msg}")
    results.append((label, fpath, sline, eline, ok, msg))

print(f"\nSummary: {passed_count}/{len(citations)} checks passed.")
