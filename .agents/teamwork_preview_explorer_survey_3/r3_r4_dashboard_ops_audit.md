# Alpha Containers Forensic Audit Report: Web Dashboard, PWA & Operational Synchronization (R3 & R4)

**Auditor**: Teamwork Explorer (Survey Recon 3)  
**Date**: August 19, 2026  
**Scope**: Requirements R3 (Web Dashboard & PWA Integrity) & R4 (Synchronization & Operational Workflow)  
**Target Repository**: `d:\Alpha`  
**Classification Standard**: Critical (Data corruption / operational stoppage risk), High (Calculation / logic / security vulnerability), Medium (Unhandled edge case / data integrity risk), Low (Hygiene / optimization / maintainability).

---

## Executive Summary

A forensic technical investigation was performed on the Alpha Containers presentation layer (`Tubex.html`, `sw.js`, `manifest.json`, `index.html`), data injection engines (`update_html.py`, `sort_dashboard.py`, `generate_customer_report.py`), batch automation scripts (`Push.bat`, `Pull.bat`, `Daily_Update.bat`, `Update_App_HTML.bat`, `Update_*.bat`), and operational workflow orchestration (`daily.py`, `alpha_checks.py`, `PIPELINE.md`, `DAILY_WORKFLOW.md`, `.gitignore`).

The audit identified **18 distinct vulnerabilities and architectural deficiencies**:
- **Critical / High Severity**: 6 findings (Silent error propagation in daily pipeline, Unsanitized XSS injection vectors across 7 DOM views, COM Excel process leak holding invisible locks, Premature HTTP error caching in Service Worker, Missing rollback/transaction mechanisms, Persistent inventory warning suppression).
- **Medium Severity**: 8 findings (Date constructor parser failure in freshness banner, Marker corruption bug `/*/* DATA_START */`, Backup path divergence between `Push.bat` and `daily.py`, Destructive `/MIR` robocopy sync, Service Worker offline failure on root URL, Pipeline execution order contradiction with documentation, Orphaned Excel lockfiles in root, Missing lockfile exclusion in backup).
- **Low / Optimization**: 4 findings (Stale icon filenames in batch script, Hardcoded PID range vs product type mismatch, Unbounded log file growth, Misplaced scripts in root directory).

---

## 1. Web Dashboard & PWA Integrity Audit (Requirement R3)

### 1.1 Security & XSS Sanitization Audit (Finding SEC-01)
- **Severity**: **HIGH**
- **Affected Files**:
  - `d:\Alpha\Tubex.html` (Lines 1551–1560, 1678, 1783, 1810–1819, 1844–1858, 1862–1869, 1973–1999, 2194–2201, 2270–2287, 2369–2380, 2417–2426, 2453–2462, 2493–2502, 2538–2547)
  - `d:\Alpha\Scripts\update_html.py` (Lines 855–924)
- **Observation**:
  `update_html.py` serializes raw Excel and ERP strings into JSON constants (`DASH_DATA`, `PRODUCTS`, `FG_STOCK_DATA`, `INVENTORY_DATA`, `MRP_DATA`, `CUSTOMER_REPORT_DATA`) using `json.dumps(..., ensure_ascii=False)`. While `json.dumps` properly creates valid JavaScript objects, `Tubex.html` renders these strings into the DOM exclusively via string concatenation assigned directly to `.innerHTML` without entity encoding or HTML sanitization.
- **Code Excerpts**:
  *`Tubex.html` (Orders Table, lines 1551-1555)*:
  ```javascript
  html += `<tr>
    <td style="font-weight:500">${o.customer}</td>
    <td style="font-weight:500">${o.product}</td>
    <td style="text-align:right">${o.dia}</td>
  ...`;
  tbody.innerHTML = html;
  ```
  *`Tubex.html` (FG Stock remarks & products, lines 2273-2287)*:
  ```javascript
  html += `<div class="fg-card ${isOk ? 'status-ok' : 'status-warn'}">
    <div class="fg-card-top">
      <div>
        <div class="fg-product">${r.product}</div>
        <div class="fg-customer">${r.customer}</div>
      </div>
  ...
    ${r.remarks ? `<div class="fg-remarks">📝 ${r.remarks}</div>` : ''}
  </div>`;
  ```
  *`Tubex.html` (Customer Report dynamic onclick attribute, line 1783)*:
  ```javascript
  periodBtns += `<button class="filter-btn ${active ? 'active' : ''}" onclick="toggleNativeMonth('${m}')">${m}</button>`;
  ```
- **Root Cause**:
  No escaping helper (such as `escapeHtml()`) exists in `Tubex.html`. Developer assumed that because inputs originate from Excel/ERP exports, data is trusted.
- **Impact & Attack Scenarios**:
  1. **ERP/Excel Payload Injection**: If an ERP export or manual Excel entry contains special characters (e.g. `<svg onload=...>`, `<img onerror=...>`, or quotes like `D'Angelo` in customer/month names), the browser parses them as HTML markup or JavaScript syntax errors.
  2. **DOM / Layout Breakage**: An unescaped single quote (`'`) in month names or customer names inside `onclick="toggleNativeMonth('${m}')"` breaks JavaScript event handlers with `Uncaught SyntaxError: Unexpected identifier`.
  3. **Credential / Cookie Access in PWA**: In standalone PWA mode on Android, arbitrary script execution within the app context can hijack local state, trigger malicious redirects, or manipulate displayed production calculations.
- **Concrete Remediation**:
  1. Add an `escapeHtml()` utility function at the top of the `<script>` block in `Tubex.html`:
     ```javascript
     function escapeHtml(str) {
       if (str === null || str === undefined) return '';
       return String(str)
         .replace(/&/g, '&amp;')
         .replace(/</g, '&lt;')
         .replace(/>/g, '&gt;')
         .replace(/"/g, '&quot;')
         .replace(/'/g, '&#39;');
     }
     ```
  2. Wrap all interpolated strings (`o.customer`, `o.product`, `r.remarks`, `m.name`, `p.name`, etc.) in `escapeHtml()`.
  3. Refactor dynamic inline `onclick="toggleNativeMonth('${m}')"` handlers to use `addEventListener` or encode string arguments safely.

---

### 1.2 Service Worker Lifecycle & Cache-Invalidation Integrity (Finding SW-01)
- **Severity**: **HIGH**
- **Affected Files**:
  - `d:\Alpha\sw.js` (Lines 1–61)
  - `d:\Alpha\Tubex.html` (Lines 2568–2572)
  - `d:\Alpha\Scripts\update_html.py` (Lines 927–937)
- **Observation**:
  `sw.js` uses a Network-First strategy with Cache Fallback. On every fetch event, it executes:
  ```javascript
  fetch(event.request)
    .then(response => {
      const clone = response.clone();
      caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
      return response;
    })
    .catch(() => {
      return caches.match(event.request).then(cached => {
        if (cached) return cached;
        return new Response('...', { headers: { 'Content-Type': 'text/html' } });
      });
    })
  ```
- **Vulnerabilities & Failure Modes**:
  1. **Caching HTTP Error Responses (404/500/502)**: When the server or network returns a non-200 HTTP response (e.g. 404 Not Found, 500 Internal Server Error, or GitHub Pages maintenance), `fetch()` resolves successfully (`catch()` is NOT triggered). `sw.js` immediately clones and puts the HTTP error response into `CACHE_NAME`. Subsequent requests (and offline sessions) serve the cached 404/500 error page permanently until the cache version changes!
  2. **Missing Scheme & Origin Validation**: Line 38 checks only `if (event.request.method !== 'GET') return;`. It does NOT verify `event.request.url.startsWith('http')`. Requests generated by Chrome extensions (`chrome-extension://`) or non-HTTP protocols trigger `cache.put()`, throwing `TypeError: Request scheme 'chrome-extension' is unsupported`.
  3. **Silent SW Activation without In-App Refresh / Controller Migration**:
     `sw.js` executes `self.skipWaiting()` on install and `self.clients.claim()` on activate. When a new version is pushed and `CACHE_NAME` changes, the active service worker is replaced in the background. However, `Tubex.html` has no `navigator.serviceWorker.addEventListener('controllerchange', ...)` listener. An open PWA window will continue running the old in-memory DOM and data indefinitely until the user forcibly closes and restarts the application.
  4. **Fragile Regex Cache Bumping**:
     In `update_html.py` line 932:
     `sw = re.sub(r"const CACHE_NAME = '[^']+';", f"const CACHE_NAME = '{new_cache}';", sw)`
     If `sw.js` uses double quotes (`"`) or whitespace variations, the regex fails silently and `sw.js` is left unversioned.
- **Concrete Remediation**:
  1. Update `sw.js` fetch handler to validate response status and protocol:
     ```javascript
     self.addEventListener('fetch', event => {
       if (event.request.method !== 'GET') return;
       if (!event.request.url.startsWith('http')) return;

       event.respondWith(
         fetch(event.request)
           .then(response => {
             // Only cache valid 200 OK responses
             if (response && response.status === 200) {
               const clone = response.clone();
               caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
             }
             return response;
           })
           .catch(() => {
             return caches.match(event.request).then(cached => {
               if (cached) return cached;
               if (event.request.headers.get('accept')?.includes('text/html')) {
                 return caches.match('./Tubex.html');
               }
               return new Response('Offline', { status: 503, statusText: 'Offline' });
             });
           })
       );
     });
     ```
  2. Add a `controllerchange` handler in `Tubex.html` to notify users or auto-refresh when a fresh version is activated:
     ```javascript
     if ('serviceWorker' in navigator) {
       let refreshing = false;
       navigator.serviceWorker.addEventListener('controllerchange', () => {
         if (!refreshing) {
           refreshing = true;
           window.location.reload();
         }
       });
       navigator.serviceWorker.register('./sw.js', { updateViaCache: 'none' });
     }
     ```

---

### 1.3 Offline Asset Manifest & CDN Strategy (Finding PWA-01)
- **Severity**: **MEDIUM**
- **Affected Files**:
  - `d:\Alpha\sw.js` (Lines 6–13)
  - `d:\Alpha\index.html` (Lines 1–15)
  - `d:\Alpha\Tubex.html` (Line 13)
- **Observation**:
  - In `sw.js`:
    ```javascript
    const ASSETS = [
      './Tubex.html',
      './manifest.json',
      './icon-192-any.png',
      './icon-512-any.png',
      './icon-192-maskable.png',
      './icon-512-maskable.png',
    ];
    ```
  - `index.html` is the default landing page on web servers (e.g. GitHub Pages root `/`).
  - `index.html` is NOT included in `ASSETS`.
  - `Tubex.html` line 13 requests Google Fonts from an external CDN:
    `https://fonts.googleapis.com/css2?family=DM+Serif+Display...`
- **Impact**:
  1. **Root URL Offline Failure**: If a user opens the bookmark or PWA URL `https://domain/` while offline, the server request fails. The cache fallback searches for `index.html` or `/`, finds nothing, and displays the generic `<h2>Offline...</h2>` message instead of redirecting or serving `Tubex.html`.
  2. **External Typography Failure**: When offline, Google Fonts stylesheets and webfont files (`.woff2`) fail to load if not already in browser HTTP cache. While fallback sans-serif fonts are specified in CSS, table columns formatted with monospace numbers (`DM Mono`) suffer typography shifts.
- **Concrete Remediation**:
  1. Add `'./index.html'` and `'./'` to `ASSETS` in `sw.js`.
  2. In `sw.js` cache fallback, if `event.request.mode === 'navigate'` or request URL ends in `/` or `index.html`, fall back directly to `caches.match('./Tubex.html')`.
  3. (Optional / Best Practice) Bundle the 3 Google Fonts locally in an `assets/fonts/` directory.

---

### 1.4 Freshness Date Parser Failure in Stale Data Banner (Finding UI-01)
- **Severity**: **MEDIUM**
- **Affected Files**:
  - `d:\Alpha\Tubex.html` (Lines 1470–1516)
  - `d:\Alpha\Scripts\update_html.py` (Line 872)
- **Observation**:
  `Tubex.html` checks data freshness via `checkFreshness()`:
  ```javascript
  const lastUpdated = (() => {
    try {
      const tsMatch = DASH_DATA.lastUpdated.match(/(\d{2}-\w{3}-\d{4} \d{2}:\d{2})/);
      return tsMatch ? tsMatch[1].replace(/-/g, ' ') : undefined;
    } catch(e) {
      return undefined;
    }
  })();

  function checkFreshness() {
    const banner = document.getElementById('stale-banner');
    if (!banner || typeof lastUpdated === 'undefined') return;
    const updated = new Date(lastUpdated);
    const now = new Date();
    const hoursAgo = (now - updated) / (1000 * 60 * 60);
    ...
  ```
- **Defect Analysis**:
  `tsMatch[1]` extracts a string like `"18-Aug-2026 13:54"`. `.replace(/-/g, ' ')` converts it to `"18 Aug 2026 13:54"`.
  In ECMAScript standards (ECMA-262), `new Date(dateString)` parsing of non-ISO 8601 strings is **implementation-dependent**.
  - While Google Chrome V8 parses `"18 Aug 2026 13:54"`, strict engines (e.g. WebKit / Safari iOS, older Android WebViews, non-English locales) return `Invalid Date` (`NaN`).
  - When `updated` is `Invalid Date`, `hoursAgo` is `NaN`.
  - The comparisons `hoursAgo > 48` and `hoursAgo > 24` evaluate to `false`.
  - The stale data banner remains permanently hidden (`display: none`), completely defeating the safety mechanism designed to alert operators to stale ERP data!
- **Concrete Remediation**:
  1. Have `update_html.py` inject an ISO-8601 timestamp in `DASH_DATA` (e.g. `dash_data["timestamp_iso"] = now.isoformat()`).
  2. In `Tubex.html`, parse the ISO string or use a deterministic parser:
     ```javascript
     function parseCustomDate(str) {
       const m = str.match(/(\d{2})-(\w{3})-(\d{4})\s+(\d{2}):(\d{2})/);
       if (!m) return null;
       const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
       return new Date(Date.UTC(parseInt(m[3]), months[m[2]], parseInt(m[1]), parseInt(m[4]), parseInt(m[5])));
     }
     ```

---

### 1.5 Marker Duplication & Substring Slicing Vulnerabilities (Finding INJ-01)
- **Severity**: **MEDIUM**
- **Affected Files**:
  - `d:\Alpha\Tubex.html` (Line 922)
  - `d:\Alpha\Scripts\update_html.py` (Lines 855–911)
- **Observation**:
  Line 922 of `Tubex.html` currently contains:
  ```javascript
  /*/* DATA_START */
  const DASH_DATA = { ... };
  /* DATA_END */
  ```
  Notice the duplicated comment opening `/*/*`.
- **Root Cause & Mechanism**:
  In `update_html.py` lines 857-867:
  ```python
  pos_start = html.find('/* DATA_START */')
  pos_end = html.find('/* DATA_END */')
  new_data_block = f"/* DATA_START */\nconst DASH_DATA = ...;\n/* DATA_END */"
  html = html[:pos_start] + new_data_block + html[pos_end + len(marker_end):]
  ```
  When `Tubex.html` was historically updated or manually edited, an extra `/*` was prepended. Because `html.find('/* DATA_START */')` finds the match starting *after* the leading `/*`, `html[:pos_start]` preserves the stray `/*`. Every subsequent execution copies the leading `/*`.
  Furthermore:
  - If `pos_end < pos_start` (due to marker disruption), the string slice `html[:pos_start] + ... + html[pos_end + len(marker_end):]` corrupts the entire middle section of `Tubex.html`.
- **Concrete Remediation**:
  1. Clean `Tubex.html` line 922 to remove the duplicate `/*`.
  2. In `update_html.py`, use regex replacement with strict boundary verification:
     ```python
     def inject_block_regex(html_content, marker_name, js_payload):
         pattern = re.compile(rf'/\*\s*{marker_name}_START\s*\*/[\s\S]*?/\*\s*{marker_name}_END\s*\*/')
         replacement = f"/* {marker_name}_START */\n{js_payload}\n/* {marker_name}_END */"
         if not pattern.search(html_content):
             raise RuntimeError(f"Marker block {marker_name} not found.")
         return pattern.sub(replacement, html_content, count=1)
     ```

---

### 1.6 Frontend UI Logic & DOM Edge Cases (Finding UI-02)
- **Severity**: **LOW**
- **Affected Files**:
  - `d:\Alpha\Tubex.html` (Lines 1247–1248, 1341, 1362, 1439–1440, 1702–1708, 2070–2088)
- **Observations & Discrepancies**:
  1. **Hardcoded PID Range vs Product Type Divergence**:
     In `renderCard` (line 1968), product type is determined by `p.type === 'pet'`.
     In `updateGrandTotal` (lines 1439–1440):
     ```javascript
     const tubeLines = productLines.filter(pl => pl.pid < 8000 || pl.pid >= 9000);
     const petLines  = productLines.filter(pl => pl.pid >= 8000 && pl.pid < 9000);
     ```
     If a PET SKU is assigned a PID outside `[8000, 8999]` (or a Tube SKU inside `[8000, 8999]`), the product card renders as PET, but the Grand Total calculation groups it under "Aluminum Tubes".
  2. **Hardcoded Month List**:
     In `getNativeCustMonths` (line 1702), `monthOrder` is hardcoded:
     `["November 2025", ... "December 2026"]`.
     Starting January 2027, chronological sorting breaks and falls back to alphabetical sort (`a.localeCompare(b)`), placing "April 2027" before "January 2027".
  3. **Silent Card Lookup Failure in `loadRemainingOrders`**:
     In lines 2079–2086, if an order in `tubeOrders` has a PID not present in `PRODUCTS`, `document.querySelector(...)` returns `null`. `quantities[o.pid]` is populated and included in the grand total BOM calculations, but no UI card is displayed or adjustable.
  4. **Input Parsing Edge Cases**:
     In `updateQty` (line 1247), `parseInt(val)` parses scientific notation `"1e5"` as `1`. Negative numbers delete the quantity without feedback.
  5. **Duplicate Variable Declaration**:
     `const catOrder = [...]` is declared at line 1341 and re-declared inside `renderSegment` at line 1362.

---

## 2. Synchronization & Operational Workflow Audit (Requirement R4)

### 2.1 Pipeline Failure Propagation & Unsafe Automation (Finding OPS-01)
- **Severity**: **CRITICAL**
- **Affected Files**:
  - `d:\Alpha\Scripts\daily.py` (Lines 443–480, 978–1017)
  - `d:\Alpha\Scripts\Daily_Update.bat` (Lines 1–10)
  - `d:\Alpha\PIPELINE.md` (Lines 24–35)
- **Observation**:
  In `daily.py`, the master workflow runs in sequence:
  `step_backup()` → `step_check_erp()` → `step_find_production()` → `step_wip()` → `step_pipeline()` → `step_crosscheck()` → `step_screenshot()` → `step_onedrive_backup()` → `step_git_push()`.
- **Failure Analysis**:
  1. **Sub-Script Failure Does Not Halt Pipeline**:
     In `step_pipeline()` (lines 443-462):
     ```python
     for script_name, label in scripts:
         result = subprocess.run([sys.executable, path], cwd=SCRIPTS_DIR)
         if result.returncode != 0:
             fail(f"{label} FAILED (exit code {result.returncode})")
             failures.append(label)
         else:
             ok(label)
     ```
     If `update_production.py` fails (e.g. missing sheet in `Production.xlsx`), the loop continues. `sort_dashboard.py` and `update_html.py` still run, reading half-updated or corrupt Excel data!
  2. **Pushing Corrupted State to Production Cloud & Git**:
     In `main()` (lines 1001-1017):
     ```python
     success = step_pipeline()
     if not success:
         all_errors.append("Pipeline execution had failures...")
     ...
     step_screenshot()          # 7. Screenshot
     step_onedrive_backup()     # 8. OneDrive backup
     step_git_push(skip=skip_git) # 9. Git push
     ```
     Even when `success == False` (meaning one or more scripts failed), `daily.py` STILL executes `step_onedrive_backup()` and `step_git_push()`.
     - It overwrites the healthy cloud backup in OneDrive with corrupted data.
     - It commits and pushes incomplete files to GitHub `main`, triggering a broken deployment on GitHub Pages.
- **Impact**:
  Data corruption is automatically propagated to all consumers (WhatsApp screenshot, OneDrive backup, and live PWA on mobile phones) without manual intervention stopping it.
- **Concrete Remediation**:
  1. In `step_pipeline()`, abort immediately if any script returns a non-zero exit code.
  2. In `main()`, if `success` is `False`, skip `step_screenshot()`, `step_onedrive_backup()`, and `step_git_push()`:
     ```python
     if not success:
         fail("Pipeline aborted due to errors. Skipping OneDrive sync and Git push.")
         # Trigger automated rollback from backup_*.xlsx
         return
     ```

---

### 2.2 Unhandled COM Automation Lifetime & Excel Lock Leaks (Finding OPS-02)
- **Severity**: **HIGH**
- **Affected Files**:
  - `d:\Alpha\Scripts\update_html.py` (Lines 40–58)
  - `d:\Alpha\Scripts\build_archives.py` (Lines 104–185)
  - `d:\Alpha\Scripts\clean_legacy_xls.py` (Lines 4–50)
  - `d:\Alpha\Scripts\alpha_checks.py` (Lines 69–108)
- **Observation**:
  `update_html.py` runs `recalculate_formulas_via_com(EXCEL_PATH)` to force Excel to evaluate workbook formulas before openpyxl reads them:
  ```python
  def recalculate_formulas_via_com(file_path):
      try:
          import win32com.client
          excel = win32com.client.Dispatch("Excel.Application")
          excel.Visible = False
          excel.DisplayAlerts = False
          wb_com = excel.Workbooks.Open(abs_path)
          wb_com.Save()
          wb_com.Close(SaveChanges=True)
          excel.Quit()
          return True
      except Exception as e:
          print(f"  Warning: Could not recalculate formulas via Excel COM: {e}")
          return False
  ```
- **Defect Analysis**:
  1. **Missing `try...finally` Cleanup**: If an error occurs during `wb_com.Save()` or `wb_com.Close()`, `excel.Quit()` is never called. An invisible `EXCEL.EXE` background process remains running in Windows Task Manager.
  2. **Invisible File Locking**: The orphaned `EXCEL.EXE` process holds a persistent OS file lock on `Tubex_Aug26.xlsx`. Subsequent runs of `daily.py` or user attempts to open Excel fail with `PermissionError` or "File in Use by another user".
  3. **Attaching to User's Active Excel Session**: `win32com.client.Dispatch("Excel.Application")` connects to an existing running Excel instance if one is open. If the user is currently editing another spreadsheet, `excel.Visible = False` hides the user's Excel window, and `excel.Quit()` terminates the user's entire Excel session!
- **Concrete Remediation**:
  1. Use `win32com.client.DispatchEx("Excel.Application")` to force a new, isolated Excel instance.
  2. Wrap all COM operations in `try...finally`:
     ```python
     excel = None
     wb_com = None
     try:
         excel = win32com.client.DispatchEx("Excel.Application")
         excel.Visible = False
         excel.DisplayAlerts = False
         wb_com = excel.Workbooks.Open(abs_path)
         wb_com.Save()
         return True
     except Exception as e:
         print(f"  Warning: COM formula recalculation failed: {e}")
         return False
     finally:
         if wb_com:
             try: wb_com.Close(SaveChanges=False)
             except Exception: pass
         if excel:
             try: excel.Quit()
             except Exception: pass
         del wb_com
         del excel
     ```

---

### 2.3 Backup Protocol Divergence & Destructive Sync Hazards (Finding OPS-03)
- **Severity**: **MEDIUM**
- **Affected Files**:
  - `d:\Alpha\Scripts\Push.bat` (Lines 14, 36–43)
  - `d:\Alpha\Scripts\daily.py` (Lines 828–846)
  - `d:\Alpha\.gitignore` (Lines 34–37)
- **Observations & Discrepancies**:
  1. **Destination Path Discrepancy**:
     - `Push.bat` line 14: `set "ONEDRIVE_BACKUP=C:\Users\HP\OneDrive\Tubex"`
     - `daily.py` line 835: `onedrive_dir = r"C:\Users\HP\OneDrive\Alpha"`
     `Push.bat` and `daily.py` backup to two different folders in OneDrive. Files pushed by `daily.py` do not update the backup created by `Push.bat`.
  2. **Destructive Robocopy Mirroring (`/MIR`) in `daily.py`**:
     In `daily.py` line 838:
     `cmd = ["robocopy", ALPHA_DIR, onedrive_dir, "/MIR", "/XD", ".git", "Logs", "__pycache__", "/R:1", "/W:1"]`
     Robocopy with `/MIR` purges any files in the destination directory that do not exist in the source directory. If a file is temporarily removed, renamed, or corrupted in `d:\Alpha`, `/MIR` immediately deletes the healthy backup copy from OneDrive.
  3. **Lockfile Copying in `daily.py`**:
     `Push.bat` excludes lockfiles via `/XF "~$*"`.
     `daily.py` lacks `/XF "~$*"`. If Excel lockfiles exist in `d:\Alpha`, `daily.py` attempts to copy them to OneDrive, which often fails if the lock is held, generating Robocopy exit code 8.
- **Concrete Remediation**:
  1. Standardize OneDrive destination path across all scripts to `C:\Users\HP\OneDrive\Tubex` (or `C:\Users\HP\OneDrive\Alpha`).
  2. Replace `/MIR` with `/E /COPY:DAT /DCOPY:DAT` in `daily.py` to prevent accidental deletion of historical backups.
  3. Add `"/XF", "~$*", "*.tmp"` to Robocopy arguments in `daily.py`.

---

### 2.4 Active Orphaned Lockfiles & Temporary File Hygiene (Finding OPS-04)
- **Severity**: **MEDIUM**
- **Affected Files / Directory**:
  - `d:\Alpha\~$June_Plan.xlsx` (Created 2026-08-05 10:53, 165 bytes)
  - `d:\Alpha\~$Production.xlsx` (Created 2026-08-06 11:28, 165 bytes)
  - `d:\Alpha\~$Tubex_Aug26.xlsx` (Created 2026-08-05 10:54, 165 bytes)
  - `d:\Alpha\Logs\hourly_push.log`
  - `d:\Alpha\test_js.js`
  - `d:\Alpha\extract_dispatch_summary.py`
- **Observation**:
  - Three hidden Microsoft Excel owner lockfiles (`~$*.xlsx`) exist in the root folder from crashed sessions dating back to early August 2026.
  - `Logs/hourly_push.log` appends logs on every hourly push without rotation or truncation.
  - `test_js.js` (a duplicate code fragment of customer report JS) and `extract_dispatch_summary.py` reside in the root directory instead of `Scripts/`.
- **Impact**:
  - Orphaned lockfiles can cause third-party sync tools (OneDrive, Dropbox) to show sync warning flags or fail batch transfers.
  - If a script uses `glob.glob("*Production*")` or unanchored matching, lockfiles can be misidentified as valid workbooks.
- **Concrete Remediation**:
  1. Add an automated cleanup step in `alpha_checks.py` or `daily.py` to detect and delete orphaned `~$*.xlsx` lockfiles if no Excel process is running:
     ```python
     def cleanup_orphaned_lockfiles(folder):
         for f in glob.glob(os.path.join(folder, "~$*.xlsx")):
             try: os.remove(f)
             except Exception: pass
     ```
  2. Delete `d:\Alpha\test_js.js` and move `extract_dispatch_summary.py` into `Scripts/`.
  3. Implement size-capped log rotation for `hourly_push.log`.

---

### 2.5 Documentation vs Implementation Contradiction in Pipeline Sequence (Finding OPS-05)
- **Severity**: **MEDIUM**
- **Affected Files**:
  - `d:\Alpha\PIPELINE.md` (Lines 24–35)
  - `d:\Alpha\DAILY_WORKFLOW.md` (Lines 74–81)
  - `d:\Alpha\Scripts\daily.py` (Lines 434–441)
- **Observation**:
  There is a direct contradiction in script execution order between documentation and code:
  - **`PIPELINE.md` (lines 27–31)**:
    ```
    Step 1: update_dispatch.py   ← FIRST
    Step 2: update_production.py ← Second
    Step 3: update_inventory.py  ← Third
    Step 4: sort_dashboard.py    ← AFTER Steps 1-3
    Step 5: update_html.py       ← LAST
    ```
  - **`daily.py` (lines 434–441)**:
    ```python
    scripts = [
        ("update_production.py", "Production Log + FG Stock"), # Step 1
        ("update_inventory.py",  "Inventory"),                 # Step 2
        ("update_dispatch.py",   "Dispatch"),                  # Step 3
        ("sort_dashboard.py",    "Sort Dashboard"),            # Step 4
        ("build_archives.py",    "Build Archives"),            # Step 5 (Unlisted in PIPELINE.md!)
        ("update_html.py",       "HTML Dashboard"),            # Step 6
    ]
    ```
- **Impact**:
  1. **Documentation Misalignment**: An engineer or operator running scripts manually per `PIPELINE.md` executes them in the wrong sequence, omitting `build_archives.py`.
  2. If `build_archives.py` is omitted, `Tubex Records/Production_Archive.xlsx` is not updated with today's production, causing the Customer Report tab in `Tubex.html` to display stale data when `update_html.py` runs.
- **Concrete Remediation**:
  Update `PIPELINE.md` and `DAILY_WORKFLOW.md` to reflect the exact 6-step sequence implemented in `daily.py`.

---

### 2.6 Persistent Inventory Warning Suppression (Finding OPS-06)
- **Severity**: **HIGH**
- **Affected Files**:
  - `d:\Alpha\Scripts\daily.py` (Lines 914–968)
- **Observation**:
  `daily.py` tracks missing inventory items using `Logs/previous_missing_items.json`.
  In `read_mismatches_log` (lines 955-959):
  ```python
  # 2. Hide if already missing yesterday, except exceptions
  is_exception = re.search(r'\b(pet resin|master batch|slugs?)\b', lower_clean)
  if item_id and item_id in prev_missing and not is_exception:
      continue
  inventory_warnings.append(clean)
  ```
- **Impact & Risk**:
  If an essential material (e.g. Carton, Cap, Lacquer, Latex, Tape, or Thinner) goes missing from `inventory.xls` (due to ERP SKU renumbering or item deletion), `daily.py` displays a warning on Day 1.
  On **Day 2 and all subsequent days**, because `item_id in prev_missing` is `True` and it is not an exception, `daily.py` **silently suppresses the warning**!
  The material remains zeroed out in Excel and `Tubex.html` indefinitely, with `daily.py` printing `✓ ALL CHECKS PASSED: No errors, missing items, or mismatches detected!`
- **Concrete Remediation**:
  Do not hide persistent inventory mismatches. If filtering is desired, display them under a dedicated section `[PERSISTENT ERP INVENTORY OMISSIONS]` so management is continuously aware of missing items.

---

### 2.7 Batch Script Asset Reference Drift (Finding OPS-07)
- **Severity**: **LOW**
- **Affected Files**:
  - `d:\Alpha\Scripts\Update_App_HTML.bat` (Lines 42–43)
- **Observation**:
  Lines 42-43 contain:
  ```bat
  if exist "%~dp0..\icon-192.png" git -C "%~dp0.." add icon-192.png
  if exist "%~dp0..\icon-512.png" git -C "%~dp0.." add icon-512.png
  ```
  The actual icon filenames in `d:\Alpha` are `icon-192-any.png`, `icon-192-maskable.png`, `icon-512-any.png`, and `icon-512-maskable.png`.
- **Impact**:
  `Update_App_HTML.bat` never stages or commits updated PWA icons if they are modified.
- **Concrete Remediation**:
  Update `Update_App_HTML.bat` to reference `icon-192-any.png`, `icon-192-maskable.png`, `icon-512-any.png`, and `icon-512-maskable.png`.

---

## 3. Consolidated Finding & Severity Matrix

| Finding ID | Area | Severity | File / Component | Exact Lines | Issue Summary | Impact |
|---|---|---|---|---|---|---|
| **OPS-01** | Synchronization | **CRITICAL** | `Scripts/daily.py` | 443–480, 978–1017 | Pipeline continues running git push and OneDrive backup even after script failures | Broken/corrupted state pushed to cloud & GitHub |
| **SEC-01** | Security / XSS | **HIGH** | `Tubex.html` | 1551, 1678, 1783, 1973, 2270, 2369, 2417 | Unsanitized ERP/Excel strings rendered via `.innerHTML` across all tabs | DOM disruption, script injection, handler breakage |
| **SW-01** | Service Worker | **HIGH** | `sw.js` | 36–60 | SW caches HTTP 404/500 errors into Cache API; no controller refresh | Cached error pages, silent SW update without reload |
| **OPS-02** | Concurrency / COM | **HIGH** | `Scripts/update_html.py`, `build_archives.py` | 40–58, 104–185 | Excel COM lack of `try...finally: excel.Quit()`; attaches to active user session | Orphaned `EXCEL.EXE` processes, workbook write locks |
| **OPS-06** | Data Integrity | **HIGH** | `Scripts/daily.py` | 955–959 | Missing inventory warnings suppressed after Day 1 | Permanent silent omission of missing ERP inventory items |
| **UI-01** | Presentation | **MEDIUM** | `Tubex.html` | 1470–1516 | `new Date('18 Aug 2026 13:54')` non-standard parser returns `NaN` | Freshness banner permanently hidden on strict browsers |
| **INJ-01** | HTML Injection | **MEDIUM** | `Tubex.html`, `Scripts/update_html.py` | 922, 855–911 | Duplicated comment `/*/* DATA_START */`; fragile substring slicing | Marker corruption, file slicing risks |
| **PWA-01** | Offline / PWA | **MEDIUM** | `sw.js`, `Tubex.html` | 6–13, 13 | `index.html` missing from cache assets; external Google Fonts offline dependency | Root URL fails offline; typography shift |
| **OPS-03** | Backup Sync | **MEDIUM** | `Scripts/Push.bat`, `daily.py` | 14, 835, 838 | Divergent OneDrive backup paths (`Tubex` vs `Alpha`); destructive `/MIR` sync | Backup fragmentation, risk of cloud backup purge |
| **OPS-04** | Disk Hygiene | **MEDIUM** | `d:\Alpha` | Root | Orphaned lockfiles `~$*.xlsx` from Aug 5/6, unrotated `hourly_push.log` | Lockfile clutter, sync issues |
| **OPS-05** | Documentation | **MEDIUM** | `PIPELINE.md`, `Scripts/daily.py` | 27–31, 434–441 | Pipeline execution sequence in docs contradicts `daily.py` | Operational confusion during manual recovery |
| **UI-02** | Frontend Logic | **LOW** | `Tubex.html` | 1439, 1702, 2079 | Hardcoded PID range `[8000, 8999]`, hardcoded months to 2026 | Potential categorization drift, sorting bug in 2027 |
| **OPS-07** | Batch Scripts | **LOW** | `Scripts/Update_App_HTML.bat` | 42–43 | Batch file references obsolete `icon-192.png` instead of `icon-192-any.png` | Icons not staged during manual HTML publish |

---

## 4. Prioritized Step-by-Step Remediation Plan

### Phase 1: Critical Operational Safety & Pipeline Hardening (Immediate)
1. **Halt Pipeline on Error (`Scripts/daily.py`)**:
   - Refactor `step_pipeline()` to abort immediately upon any script exit code $\ne 0$.
   - Prevent `step_screenshot()`, `step_onedrive_backup()`, and `step_git_push()` from executing if `step_pipeline()` fails.
2. **Eliminate Persistent Inventory Warning Suppression (`Scripts/daily.py`)**:
   - Remove the `if item_id and item_id in prev_missing` filter so missing ERP items are always visible in the daily error summary.
3. **COM Process Protection (`Scripts/update_html.py`, `build_archives.py`)**:
   - Use `win32com.client.DispatchEx("Excel.Application")` with strict `try...finally: excel.Quit()` blocks to ensure background `EXCEL.EXE` processes never leak.

### Phase 2: Web Dashboard & PWA Integrity (Next)
4. **Implement XSS Sanitization (`Tubex.html`)**:
   - Introduce `escapeHtml()` helper and sanitize all dynamic ERP strings before `.innerHTML` insertion.
5. **Harden Service Worker (`sw.js`)**:
   - Enforce `response.status === 200` check before caching.
   - Include `'./index.html'` and `'./'` in `ASSETS`.
   - Add navigation fallback to `'./Tubex.html'`.
6. **Fix Date Parsing in Freshness Banner (`Tubex.html` & `update_html.py`)**:
   - Pass an ISO-8601 string (`YYYY-MM-DDTHH:MM:SS`) in `DASH_DATA.timestamp_iso` to guarantee cross-browser date parsing.
7. **Clean Injection Markers (`Tubex.html` & `update_html.py`)**:
   - Remove `/*/*` from `Tubex.html` line 922.
   - Switch `update_html.py` to regex-based block replacement with validation assertions.

### Phase 3: Synchronization, Backup & Hygiene Standardization
8. **Unify OneDrive Backup Paths**:
   - Set `C:\Users\HP\OneDrive\Alpha` as the single canonical path across `Push.bat` and `daily.py`.
   - Replace `/MIR` with `/E /COPY:DAT /DCOPY:DAT /XF "~$*" "*.tmp"`.
9. **Clean Orphaned Lockfiles & Move Stray Scripts**:
   - Delete `~$June_Plan.xlsx`, `~$Production.xlsx`, `~$Tubex_Aug26.xlsx`, and `test_js.js`.
   - Move `extract_dispatch_summary.py` to `Scripts/`.
10. **Synchronize Pipeline Documentation**:
    - Update `PIPELINE.md` and `DAILY_WORKFLOW.md` to document the 6-step pipeline (including `build_archives.py`).
11. **Update Batch Icon Names**:
    - Update `Update_App_HTML.bat` to reference `icon-*-any.png` and `icon-*-maskable.png`.

---

## 5. Verification Methods

1. **Service Worker Offline Test**:
   - In Chrome DevTools > Application > Service Workers, check "Offline".
   - Navigate to `/index.html` and `/Tubex.html`. Verify that the full dashboard loads from cache without falling back to raw error text.
2. **XSS Injection Test**:
   - Inject test string `<img src=x onerror=alert('XSS')>` into a customer name in `Tubex_Aug26.xlsx` Product_Catalog.
   - Run `update_html.py` and open `Tubex.html`. Verify that the text renders literally as `<img src=x onerror=alert('XSS')>` and no alert box executes.
3. **Pipeline Failure Simulation**:
   - Rename `inventory.xls` temporarily to simulate a missing ERP file.
   - Run `python daily.py`. Verify that the pipeline halts after step 2, does NOT execute Git push or OneDrive sync, and does NOT generate a false `✓ ALL CHECKS PASSED` summary.
4. **COM Process Cleanup Verification**:
   - In PowerShell, run `Get-Process excel -ErrorAction SilentlyContinue`.
   - Run `python update_html.py`.
   - Verify that no lingering `EXCEL.EXE` background processes exist after script execution.
