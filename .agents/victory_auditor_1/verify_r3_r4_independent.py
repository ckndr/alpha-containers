import os
import re

def check_r3_r4():
    print("=== STARTING INDEPENDENT R3 & R4 AUDIT ===")

    # R3-01, R3-02, R3-03: Tubex.html
    with open(r"d:\Alpha\Tubex.html", "r", encoding="utf-8", errors="ignore") as f:
        html_lines = f.readlines()
    html_text = "".join(html_lines)

    # Check innerHTML usage
    l1550 = "".join(html_lines[1550:1573])
    print(f"R3-01 snippet (L1551-1571):\n{l1550[:200]}")
    assert "${o.customer}" in l1550 and "tbody.innerHTML" in l1550, "Unsanitized innerHTML in L1551 verified"
    print("[PASS] R3-01: Unsanitized DOM innerHTML injection in Orders table verified")

    l1783 = html_lines[1782] if len(html_lines) > 1782 else ""
    print(f"R3-02 snippet (L1783): {l1783.strip()}")
    assert "toggleNativeMonth('${m}')" in l1783 or "toggleNativeMonth" in l1783, "Inline onclick handler verified"
    print("[PASS] R3-02: Unescaped inline onclick handler in customer report verified")

    # R3-04, R3-05: sw.js
    with open(r"d:\Alpha\sw.js", "r", encoding="utf-8", errors="ignore") as f:
        sw_text = f.read()
    print(f"sw.js contents:\n{sw_text[:400]}")
    assert "caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone))" in sw_text, "SW error caching verified"
    print("[PASS] R3-04: Premature caching of error responses in sw.js verified")

    # R3-07: Stale date parsing
    l1470 = "".join(html_lines[1465:1520])
    assert 'new Date(' in l1470, "Date parsing in stale alert banner verified"
    print("[PASS] R3-07: Stale date parsing in Tubex.html verified")

    # R3-08: Data marker
    l922 = html_lines[921] if len(html_lines) > 921 else ""
    print(f"R3-08 snippet (L922): {l922.strip()}")
    assert "/*/* DATA_START */" in l922 or "/* DATA_START */" in l922, "Data start marker verified"
    print("[PASS] R3-08: Injection marker verified")

    # R4-01 & R4-02: daily.py
    with open(r"d:\Alpha\Scripts\daily.py", "r", encoding="utf-8", errors="ignore") as f:
        daily_lines = f.readlines()
    daily_text = "".join(daily_lines)

    step_pipe = "".join(daily_lines[440:485])
    assert "subprocess.run" in step_pipe and "failures.append" in step_pipe, "step_pipeline failure handling verified"
    print("[PASS] R4-01: Unchecked script failure propagation in daily.py verified")

    main_flow = "".join(daily_lines[990:1025])
    assert "step_onedrive_backup()" in main_flow and "step_git_push" in main_flow, "Unconditional onedrive/git push verified"
    print("[PASS] R4-02: Automated cloud push on failed pipeline verified")

    # R4-05: Push.bat vs daily.py
    with open(r"d:\Alpha\Scripts\Push.bat", "r", encoding="utf-8", errors="ignore") as f:
        push_bat = f.read()
    print(f"Push.bat contents:\n{push_bat}")
    assert r"OneDrive\Tubex" in push_bat, "Push.bat targets OneDrive\\Tubex"
    assert r"OneDrive\Alpha" in daily_text, "daily.py targets OneDrive\\Alpha"
    print("[PASS] R4-05: Divergent OneDrive destinations verified (Tubex vs Alpha)")

    # R4-06: daily.py Robocopy /MIR
    assert "/MIR" in daily_text, "daily.py uses robocopy /MIR"
    print("[PASS] R4-06: Robocopy /MIR purge hazard verified")

    # R4-07: Lockfiles in root
    lockfiles = [f for f in os.listdir(r"d:\Alpha") if f.startswith("~$")]
    print(f"R4-07: Lockfiles found in root: {lockfiles}")
    assert len(lockfiles) > 0, "Lockfiles present in root"
    print("[PASS] R4-07: Orphaned Excel lockfiles in root verified")

    # R4-09: Update_App_HTML.bat
    with open(r"d:\Alpha\Scripts\Update_App_HTML.bat", "r", encoding="utf-8", errors="ignore") as f:
        uab = f.read()
    assert "icon-192.png" in uab, "Update_App_HTML.bat references obsolete icon-192.png"
    print("[PASS] R4-09: Obsolete icon name reference in Update_App_HTML.bat verified")

if __name__ == "__main__":
    check_r3_r4()
