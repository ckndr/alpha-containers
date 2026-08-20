# Handoff Report: Survey Explorer 3 (R3 & R4 Audit)

## 1. Observation
- **Dashboard & PWA (`d:\Alpha\Tubex.html`, `d:\Alpha\sw.js`, `d:\Alpha\manifest.json`)**:
  - `Tubex.html`: Dynamic data rendering across all 7 views uses direct string interpolation into `.innerHTML` (e.g. lines 1551–1560, 1678, 1783, 1973–1999, 2194–2201, 2270–2287, 2369–2380, 2417–2426, 2453–2462, 2493–2502, 2538–2547). Zero HTML escaping is performed on ERP strings.
  - `Tubex.html` line 922: Contains duplicated block comment marker `/*/* DATA_START */`.
  - `Tubex.html` line 1483: `const updated = new Date(lastUpdated)` parses non-standard date format `"18 Aug 2026 13:54"`, causing `NaN` in strict ECMAScript engines and permanently hiding the stale data warning banner.
  - `sw.js` lines 36–60: Intercepts all GET requests and caches responses without checking `response.status === 200`. HTTP 404/500 errors are permanently stored in `CACHE_NAME`. Missing scheme check throws on `chrome-extension://`. `index.html` is omitted from `ASSETS`.
- **Pipeline & Operational Workflow (`d:\Alpha\Scripts\daily.py`, `d:\Alpha\Scripts\update_html.py`, `d:\Alpha\Scripts\Push.bat`, `d:\Alpha\Scripts\Pull.bat`)**:
  - `daily.py` lines 443–480, 978–1017: If sub-scripts fail during `step_pipeline()`, the pipeline continues and executes `step_onedrive_backup()` and `step_git_push()`, pushing broken/corrupted states to OneDrive and GitHub Pages.
  - `daily.py` lines 955–959: Missing ERP inventory items are suppressed from daily warnings after Day 1 (`item_id in prev_missing`), hiding long-term material omissions.
  - `update_html.py` lines 40–58: `recalculate_formulas_via_com` lacks `try...finally: excel.Quit()`. If COM fails or throws, invisible `EXCEL.EXE` background processes remain active holding write locks on master workbooks.
  - Backup path divergence: `Push.bat` uses `C:\Users\HP\OneDrive\Tubex`, while `daily.py` uses `C:\Users\HP\OneDrive\Alpha` with destructive `/MIR` mode.
  - Orphaned lockfiles: `~$June_Plan.xlsx`, `~$Production.xlsx`, `~$Tubex_Aug26.xlsx` exist in root directory from early August.
  - Documentation divergence: `PIPELINE.md` lines 27–31 specifies a 5-step sequence starting with dispatch, while `daily.py` lines 434–441 executes a 6-step sequence starting with production and including `build_archives.py`.

## 2. Logic Chain
1. *From observations in `daily.py` (lines 443-480 & 1001-1017)*: Because sub-script return codes do not raise fatal exceptions, and because `main()` calls `step_git_push()` and `step_onedrive_backup()` unconditionally regardless of `success`, a failure in `update_production.py` or `update_inventory.py` will still result in `Tubex.html` being generated from partial data and pushed to GitHub main, which breaks production dashboards on mobile devices.
2. *From observations in `Tubex.html` (lines 1551, 1678, 1783, 1973, 2270)*: Because all dynamic table rows, product names, customer names, and remarks are concatenated into HTML template literals and assigned directly to `.innerHTML`, any unescaped quotes or angle brackets from ERP strings will disrupt the DOM or execute injected scripts.
3. *From observations in `sw.js` (lines 41-47)*: Because `fetch(event.request)` resolves successfully for HTTP 404/500 error responses, the cloned error response is placed in the cache. When the operator goes offline, `caches.match()` returns the cached 404 error page.
4. *From observations in `update_html.py` (lines 40-58) and `alpha_checks.py` (lines 69-108)*: Because COM Dispatch is called without `DispatchEx` or `try...finally: excel.Quit()`, any unhandled error inside COM leaves an orphaned background Excel instance holding a file lock. When `alpha_checks.check_not_locked()` runs subsequently, it triggers `sys.exit(1)`, stalling the daily pipeline.

## 3. Caveats
- The investigation focused strictly on R3 (Web Dashboard, PWA, Injection, Frontend) and R4 (Operational Workflow, Synchronization, Batch scripts, Failure recovery).
- In-depth Excel cell-by-cell formula logic and Aerosol workbook commissioning models are scoped under R2 and audited by Explorer 2.
- Ingestion parsing logic inside `update_production.py`, `update_inventory.py`, and `update_dispatch.py` are scoped under R1 and audited by Explorer 1.

## 4. Conclusion
The Web Dashboard and Operational Synchronization layers have critical architectural vulnerabilities that permit silent error propagation, persistent file locks, and XSS/DOM disruption. The full forensic report has been compiled at `d:\Alpha\.agents\teamwork_preview_explorer_survey_3\r3_r4_dashboard_ops_audit.md`.

Immediate remediation priorities:
1. Halt `daily.py` pipeline immediately on any script error; skip git push and OneDrive sync on failure.
2. Add `escapeHtml()` sanitization across all `.innerHTML` renderings in `Tubex.html`.
3. Wrap Excel COM automation in `DispatchEx` and `try...finally: excel.Quit()` in `update_html.py`.
4. Fix `sw.js` to only cache `response.status === 200` and include `index.html` in cache assets.
5. Standardize OneDrive backup paths between `Push.bat` and `daily.py`, removing `/MIR` destructive mirroring.
6. Pass ISO 8601 timestamps in `DASH_DATA` to fix the freshness banner in `Tubex.html`.

## 5. Verification Method
1. **Pipeline Error Halting**: Run `python daily.py` with a simulated invalid `inventory.xls`. Verify that the script terminates at step 2 and does NOT execute `git push` or `robocopy`.
2. **XSS & DOM Sanitization**: Inject `<b>Test</b>` into a product name in `Tubex_Aug26.xlsx`, run `update_html.py`, and inspect `Tubex.html` in browser to confirm literal rendering.
3. **Service Worker Offline Cache**: In Chrome DevTools Application > Service Workers, toggle Offline mode and load `/index.html` and `/Tubex.html`.
4. **COM Process Cleanup**: Inspect Task Manager / PowerShell `Get-Process excel` before and after running `update_html.py` to confirm zero leaked `EXCEL.EXE` processes.
