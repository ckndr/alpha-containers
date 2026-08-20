# Requirement R1: Data Pipeline & Script Reliability Forensic Audit Report

**Author**: Survey Explorer 1 (Alpha Containers End-to-End Audit)  
**Date**: 2026-08-19  
**Status**: Completed  
**Scope**: Python Automation Scripts (`Scripts/`), ERP Data Ingestion Logic (`Production.xlsx`, `inventory.xls`, `dispatch.xls`, `dispatch_pet.xls`), Safety Assertions (`alpha_checks.py`), and Pipeline Architecture (`PIPELINE.md`).

---

## 1. Executive Summary

A comprehensive, forensic audit was performed across all 32 Python automation and maintenance scripts in `Scripts/` and the root workspace, alongside all ERP ingestion routines and workbook interfaces.

The audit uncovered 22 vulnerabilities across data pipeline integrity, calculation reliability, and error handling:
1. **Silent Production Dropping on Unmapped Aliases**: When new products or naming variants appear in Imran's production file, `update_production.py` writes them with `PID = None`. Downstream, `sort_dashboard.py` and `update_html.py` silently discard all rows without a PID, causing good production units to vanish from MTD KPIs and marking active lines as inactive without raising an error.
2. **Catastrophic Zeroing of Unmatched Inventory**: `update_inventory.py` automatically zeroes out Opening, Received, and Issued quantities for any item present in the master workbook but absent from `inventory.xls`. If an operator downloads a category-filtered or partial ERP export, all other items across all categories are instantly wiped to `0.0`.
3. **Formula & Regex Corruption in Dashboard Sorting**: `sort_dashboard.py` rewrites order formulas using regex `\b([FD])\d+\b`, which corrupts multi-cell range references (e.g., `MRP!$D$3:$F$50` becomes `MRP!$D$15:$F$15`), and injects `SUMPRODUCT` formulas hardcoded to row limit `$8963` that only match machine names starting with `"Print"` (ignoring `"PLINE"`, which Python includes).
4. **Dead Code & Date Filter Blind Spots in Dispatch Ingestion**: `update_dispatch.py` attempts dynamic date filtering to exclude "today's" dispatch, but because pandas reads numeric serial dates without header conversion, type checks fail and the filter is completely inoperative. Furthermore, filtering "today's" data creates an operational blind spot for same-day evening runs.
5. **Execution Order Discrepancy**: `PIPELINE.md` mandates running `update_dispatch.py` before `update_production.py`, whereas `daily.py` executes `update_production.py` first, then `update_inventory.py`, then `update_dispatch.py`.
6. **False-Safe Safety Checks**: `alpha_checks.py` returns `True` (fresh) if a file does not exist, checks freshness non-blockingly without halting execution, verifies lock status on only one file, and lacks schema, row count, referential integrity, and formula health assertions.

---

## 2. Master Vulnerability Matrix

| ID | Component / File | Severity | Vulnerability Category | Root Cause Summary |
|---|---|---|---|---|
| **V-01** | `update_production.py` (L612-616) / `sort_dashboard.py` (L130) | **CRITICAL** | Silent Data Loss | Products with missing aliases get `PID=None` and are silently dropped from MTD sums and dashboard. |
| **V-02** | `update_inventory.py` (L257-288) | **CRITICAL** | Data Corruption | Items absent from ERP export are zeroed out (0.0 opening/in/out); partial ERP exports destroy inventory history. |
| **V-03** | `sort_dashboard.py` (L388-392) | **HIGH** | Formula Corruption | Regex rewriting `\b([FD])\d+\b` corrupts range references across sheets (e.g. `MRP!D3:F50` -> `MRP!D15:F15`). |
| **V-04** | `sort_dashboard.py` (L319-324) | **HIGH** | Logic Divergence | Excel formula only matches `LEFT(...,5)="Print"`, while Python logic counts `PLINE` machines. |
| **V-05** | `sort_dashboard.py` (L319, L595, L661) | **MEDIUM** | Scalability / Truncation | SUMPRODUCT hardcoded to `$F$3:$F$8963`; rows beyond 8963 in Production_Log are ignored. |
| **V-06** | `update_dispatch.py` (L174-231) | **HIGH** | Ingestion Logic Bug | Dynamic date filter is dead code on serial dates and drops valid dispatches run on current day. |
| **V-07** | `update_dispatch.py` (L188-204) | **HIGH** | Column-Shift Hazard | Hardcoded `col0 = row[0]`, `col7 = row[7]` without header lookup; column shifts inject invalid quantities. |
| **V-08** | `update_production.py` (L743-751) | **HIGH** | Schema Assumption | Hardcoded `header=1` and positional rename in FG Stock sheet; extra rows or column additions break parsing. |
| **V-09** | `update_production.py` (L515-532) | **HIGH** | Date Parsing Ambiguity | `pd.Timestamp(date_raw)` without `dayfirst=True` defaults to MM/DD/YYYY, swapping days and months. |
| **V-10** | `update_inventory.py` (L98-105) | **HIGH** | Index Out of Bounds | Hardcoded fallback map assumes 11 columns, but actual ERP report has 8 columns. |
| **V-11** | `update_inventory.py` (L196) | **MEDIUM** | Broken Regex / UI Stale | `re.sub(r'\(.*?\)', ...)` fails on `Inventory!A1` because title has no parentheses and has corrupt encoding. |
| **V-12** | `update_production.py` (L869-877) | **MEDIUM** | Residual Formula Leak | `write_fg_stock` clears only columns 1-8, leaving orphan `=IFERROR(...)` formulas in column 9 (Col I). |
| **V-13** | `update_html.py` (L216-217) | **MEDIUM** | Hardcoded Partitioning | Tube vs PET MTD production strictly split by `PID < 8000` vs `PID >= 8000`. |
| **V-14** | `daily.py` (L434-441) vs `PIPELINE.md` (L27-31) | **MEDIUM** | Pipeline Architecture | Execution order mismatch between documentation and master orchestrator script. |
| **V-15** | `daily.py` (L470) | **MEDIUM** | Encoding Bug | Opens `mismatches.log` without `encoding='utf-8'`, causing `UnicodeDecodeError` on Windows cp1252. |
| **V-16** | `daily.py` (L945-960) | **HIGH** | Alert Suppression | Silently suppresses all "INK" warnings and recurring missing items from daily error summary. |
| **V-17** | `daily.py` (L638-646) | **HIGH** | Fragile Assertions | Machine and Summary cross-checks rely on hardcoded cell addresses (B14, B15, B3, B4, B22, B11). |
| **V-18** | `alpha_checks.py` (L50) | **HIGH** | Safety Blind Spot | `check_freshness` returns `True` for non-existent files. |
| **V-19** | `alpha_checks.py` (L34-67) | **HIGH** | Safety Bypass | `check_freshness` warns but never halts execution, allowing stale ERP data to overwrite master models. |
| **V-20** | `alpha_checks.py` (L142-195) | **MEDIUM** | File Corruption Risk | `replace_copy_export` replaces target without checking file size, magic bytes, or download completion. |
| **V-21** | `customer_normalization.py` (L80) | **MEDIUM** | False Match Hazard | Bi-directional substring match (`mc in raw or raw in mc`) can misclassify short customer names. |
| **V-22** | Across scripts (`find_files`) | **MEDIUM** | File Selection Conflict | Inconsistent sorting: `build_archives.py` uses `mtime`, while other scripts use alphabetical `sorted()[-1]`. |

---
## 3. Deep Forensic Investigation by Script & Module

### 3.1. `Scripts/update_production.py`

#### Finding F-01: Silent Rejection of Unmapped Production Rows
- **Exact Location**: `Scripts/update_production.py`, Lines 584-642, Lines 612-616
- **Code Snippet**:
  ```python
  catalog_name, pid = ALIASES.get((name_raw.lower().strip(), dia_raw), (None, None))
  if catalog_name is None:
      catalog_name = name_raw
      pid = None

  if pid is None:
      if not ("(varnish)" in catalog_name.lower() or "(varnish)" in name_raw.lower()):
          no_pid.append((catalog_name, dia_raw, cust_norm))
  ```
- **Root Cause**: When Imran introduces a new SKU, changes formatting (e.g. "150" instead of "150 ml"), or enters a typographical variation, `ALIASES` lookup yields `pid = None`. The row is written into `Production_Log` with an empty PID.
- **Downstream Consequence**: In `sort_dashboard.py` (Line 130) and `update_html.py` (Line 184), the parser executes `if not pid or not good_qty: continue`. As a result, the produced units are completely excluded from MTD production totals, MTD machine totals, and dashboard KPIs. The product is treated as having 0 production, and if it has no open order, it is sorted into the Inactive section.
- **Severity**: **CRITICAL**
- **Remediation**:
  1. Implement fuzzy matching with Levenshtein distance against `Product_Catalog` for unmatched names.
  2. If an alias cannot be resolved, halt pipeline execution or assign a fallback tracking ID (e.g., `PID_UNMAPPED_9999`) and emit a blocking error rather than a passive warning.
  3. Require `alpha_checks.py` to validate that every row in `Production.xlsx` maps to a non-null PID before modifying `Tubex_Aug26.xlsx`.

---

#### Finding F-02: Ambiguous Date Parsing in Production Data
- **Exact Location**: `Scripts/update_production.py`, Lines 515-532
- **Code Snippet**:
  ```python
  def parse_date(date_raw):
      ...
      try:
          ts = pd.Timestamp(date_raw)
          d  = ts.date()
          if not (2020 <= d.year <= 2035):
              return None
          return d
      except Exception:
          return None
  ```
- **Root Cause**: `pd.Timestamp(date_raw)` assumes ISO format or US format (MM/DD/YYYY) when parsing strings without explicit format specifications. In Pakistani manufacturing environments, dates are entered as DD/MM/YYYY (e.g. `06/08/2026` for August 6). `pd.Timestamp('06/08/2026')` parses as June 8, 2026.
- **Impact**: Production entries for the first 12 days of the month are assigned to incorrect months, causing MTD sums for the current month to be understated and corrupting historical records.
- **Severity**: **HIGH**
- **Remediation**:
  Use explicit date parsing: `pd.to_datetime(date_raw, dayfirst=True, format='mixed')` or try explicit formats `['%d-%m-%Y', '%d/%m/%Y', '%d-%b-%Y', '%Y-%m-%d']`.

---

#### Finding F-03: Positional Header and Column Assumptions in `read_fg_stock`
- **Exact Location**: `Scripts/update_production.py`, Lines 743-751
- **Code Snippet**:
  ```python
  df = pd.read_excel(prod_path, sheet_name='FG Stock In hand', header=1)
  df.columns = [
      'Sr', 'Date', 'Customer', 'Product', 'Diameter',
      'FG_Qty', 'Prod_Remarks', 'Dispatch_Remarks'
  ] + list(df.columns[8:])
  ```
- **Root Cause**: The script hardcodes `header=1` (skipping row 0) and directly overwrites column names with a hardcoded list of 8 fields.
- **Impact**: If Imran removes the top title banner (row 0), inserts a new column (e.g. Batch # or Location), or shifts columns, the script renames wrong columns. For instance, Remarks would become Qty, causing `int(float(r.get('FG_Qty', 0)))` to coerce to 0 and wiping all Finished Goods stock.
- **Severity**: **HIGH**
- **Remediation**:
  Dynamically scan the first 5 rows for header keywords (`'Customer'`, `'Product'`, `'Diameter'`, `'Stock in Hand'`, `'Remarks'`) and map columns by name rather than fixed position.

---

#### Finding F-04: Orphaned Formula Leakage in `write_fg_stock`
- **Exact Location**: `Scripts/update_production.py`, Lines 869-877
- **Code Snippet**:
  ```python
  max_r = ws.max_row
  for r in range(4, max_r + 1):
      for c in range(1, 9):
          cell = ws.cell(row=r, column=c)
          cell.value  = None
          cell.font   = _font()
          cell.fill   = _fill(None)
          cell.border = Border()
  ```
- **Root Cause**: `write_fg_stock` only clears columns 1 through 8 (`range(1, 9)`). In `Tubex_Aug26.xlsx`, Column 9 (Column I) contains cap lookup formulas: `=IFERROR(SUMPRODUCT((TableBOM[Product ID]=B21)*(TableBOM[Material Category]="CAP")*TableBOM[Item ID]),0)`.
- **Impact**: When the number of FG rows shrinks (e.g. from 29 rows to 20 rows), rows 21-29 have blank data in columns A-H, but retain active Excel formulas in Column I.
- **Severity**: **MEDIUM**
- **Remediation**:
  Expand cell clearing to `range(1, ws.max_column + 1)` or explicitly clear columns 1 to 12.

---

### 3.2. `Scripts/update_inventory.py`

#### Finding F-05: Destructive Zeroing of Inventory Items Absent from ERP Export
- **Exact Location**: `Scripts/update_inventory.py`, Lines 257-288
- **Code Snippet**:
  ```python
  # Find missing items (present in Inventory sheet but missing from new inventory.xls)
  missing_items = []
  missing_slugs = []
  for item_id, row in sorted(excel_ids.items()):
      if item_id not in xls_items:
          name_val = ws.cell(row=row, column=3).value
          cat_val = ws.cell(row=row, column=2).value
          missing_items.append((item_id, str(name_val or '').strip(), str(cat_val or '').strip()))

          # Zero out values to prevent phantom stock
          ws.cell(row=row, column=5).value = 0.0
          ws.cell(row=row, column=6).value = 0.0
          ws.cell(row=row, column=7).value = 0.0

          # Set font color to RED for visibility
          ...
          ws.cell(row=row, column=11).value = "Not active in ERP"
  ```
- **Root Cause**: The script assumes `inventory.xls` is an exhaustive export of every single item ever tracked in Tubex. If an item in Excel Col A is not in `inventory.xls`, it sets Opening (Col E), Inward (Col F), and Outward (Col G) to `0.0`.
- **Operational Hazard**: In ERP systems, consolidated reports are frequently filtered by warehouse, active status, or material group. If an operator exports only "Slugs & Resins" or if ERP suppresses zero-movement items, all 24+ packaging, cap, lacquer, carton, and ink inventories in `Tubex_Aug26.xlsx` are instantly wiped to `0.0`.
- **Severity**: **CRITICAL**
- **Remediation**:
  1. Never overwrite Opening balance to 0.0 unless explicitly verified against a confirmed full physical stocktake.
  2. Distinguish between "Zero Movement in Period" (Opening preserved, In=0, Out=0) and "Item Deletion".
  3. Require `alpha_checks.py` to verify that `inventory.xls` contains at least 80% of master catalog Item IDs before allowing the update script to write to disk.

---

#### Finding F-06: Out-of-Bounds Fallback Column Indices on 8-Column ERP Export
- **Exact Location**: `Scripts/update_inventory.py`, Lines 98-105
- **Code Snippet**:
  ```python
  # Default column indices (old format)
  col_id = 0
  col_name = 2
  col_opening = 6
  col_inward = 7
  col_out = 8
  col_balance = 9
  col_unit = 10
  ```
- **Root Cause**: The hardcoded fallback default assumes an 11-column legacy report format. The actual current ERP export (`inventory.xls`, sheet `'Admin_MIS_Stk_Rpt_IW_Con.rpt'`) contains exactly 8 columns:
  `['ID', 'Item Name', 'Make', 'Opening', 'In', 'Out', 'Balance', 'Unit']` (indices 0 to 7).
- **Impact**: If dynamic header detection fails to match `'id'` and `'item name'` (e.g. if ERP prints `'Item Code'` or localized headers), the script falls back to `col_out = 8`, `col_balance = 9`, `col_unit = 10`. On an 8-column row, `row[col_out]` triggers an unhandled `IndexError` or reads empty cells.
- **Severity**: **HIGH**
- **Remediation**:
  Update default fallback indices to match the 8-column standard: `col_id=0, col_name=1, col_opening=3, col_inward=4, col_out=5, col_balance=6, col_unit=7`, and raise an explicit `SchemaValidationError` if header detection fails.

---

#### Finding F-07: Ineffective Date Range Substitution & Corrupted Title String
- **Exact Location**: `Scripts/update_inventory.py`, Lines 193-197
- **Code Snippet**:
  ```python
  # Update date range in cell A1 if present
  if date_range:
      cell = ws.cell(row=1, column=1)
      if cell.value:
          cell.value = re.sub(r'\(.*?\)', '(' + date_range + ')', str(cell.value))
  ```
- **Root Cause**: In `Tubex_Aug26.xlsx`, cell `Inventory!A1` contains: `"Slugs Inventory  August 2026 MONTH TO DATE"`. It has no parentheses `(...)`.
- **Impact**: `re.sub(r'\(.*?\)', ...)` finds no match and does nothing. The inventory date range is never updated in the master workbook, leaving headers misleadingly static. Furthermore, the character `` is an unresolved encoding corruption.
- **Severity**: **MEDIUM**
- **Remediation**:
  Standardize `Inventory!A1` formatting and update using template formatting:
  `ws['A1'].value = f"Slugs & Raw Materials Inventory — ({date_range})"`.
### 3.3. `Scripts/update_dispatch.py`

#### Finding F-08: Ineffective Date Filter & Same-Day Data Loss Hazard
- **Exact Location**: `Scripts/update_dispatch.py`, Lines 174-231
- **Code Snippet**:
  ```python
  today = datetime.now()
  today_date = today.date()
  ...
  for val in row:
      if hasattr(val, 'date') and callable(getattr(val, 'date')):
          if val.date() == today_date:
              skip_row = True
              break
      elif isinstance(val, str):
          v_str = val.strip().lower()
          if v_str in today_strs:
              skip_row = True
              break
  ```
- **Root Cause**:
  1. `pd.read_excel(..., engine='xlrd', header=None)` reads ERP date cells as numeric floats (e.g. `46245.0`) when no date converters are attached. Neither `hasattr(val, 'date')` nor `isinstance(val, str)` matches float values, rendering the filter inoperative for serial dates.
  2. If the filter does operate (when dates are strings/datetimes), it unconditionally ignores all dispatches occurring on the system date (`today`). If an operator runs the daily pipeline at 6:00 PM to generate the evening management dashboard, all dispatches completed during that day are dropped.
- **Severity**: **HIGH**
- **Remediation**:
  Remove arbitrary system-date filtering. The ERP export date range is controlled by the user at export time in ERP. The script must parse all valid records in the file or allow explicit cutoff parameters (`--cutoff-date`).

---

#### Finding F-09: Unvalidated Hardcoded Column Indices in Dispatch Parsing
- **Exact Location**: `Scripts/update_dispatch.py`, Lines 188-235
- **Code Snippet**:
  ```python
  col0 = row[0]
  col7 = row[7]
  ...
  try:
      int(float(col0))
      if pd.isna(col7):
          continue
      ...
      disp_qty = float(col7)
  ```
- **Root Cause**: The parser hardcodes `col0` as record indicator / product name and `col7` as Dispatch Quantity without validating Row 5 headers (`['No.', 'Pk Id / Dly Id', 'Date', 'POF #', 'Client PO#', 'Dia', 'Ord. Qty', 'Disp. Qty', ...]`).
- **Impact**: If ERP adds a column (e.g. Vehicle # or Delivery Challan #) or re-orders fields, `row[7]` will read Ordered Quantity, Replaced Quantity, or Gate Pass strings, injecting corrupted dispatch numbers into `Tubex_Dashboard` Column K.
- **Severity**: **HIGH**
- **Remediation**:
  Locate header row dynamically, resolve index of `'Disp. Qty'` or `'Dispatched Qty'`, and extract values by resolved index.

---

### 3.4. `Scripts/sort_dashboard.py`

#### Finding F-10: Regex Formula Rewriting Corrupts Range Lookups
- **Exact Location**: `Scripts/sort_dashboard.py`, Lines 388-392
- **Code Snippet**:
  ```python
  orders_val = data['orders']
  if isinstance(orders_val, str) and orders_val.startswith('='):
      orders_val = re.sub(r'\b([FD])\d+\b', r'\g<1>' + str(r), orders_val)
  ws.cell(r, 7).value = orders_val
  ```
- **Root Cause**: The regex `\b([FD])\d+\b` matches any occurrence of letter F or D followed by digits. If an order cell in Column G contains an external sheet lookup such as `=VLOOKUP(F11, MRP!$D$3:$F$50, 3, FALSE)` or `=SUMIF(MRP!D3:D50, F11, MRP!F3:F50)`, the regex alters `F11` to `F{r}`, but it also alters `D3` to `D{r}`, `F50` to `F{r}`, and `D50` to `D{r}`!
- **Impact**: Destroys table lookup ranges in Excel formulas, replacing 2D matrix lookups with collapsed single-cell references and returning `#REF!` or incorrect order quantities.
- **Severity**: **HIGH**
- **Remediation**:
  Use openpyxl token parsing or target relative cell references only: `re.sub(r'(?<![!$\w])([FD])(\d+)\b', ...)`.

---

#### Finding F-11: Machine String Matching Discrepancy Between Python and Injected Excel Formula
- **Exact Location**: `Scripts/sort_dashboard.py`, Lines 133 vs Line 320
- **Code Comparison**:
  - Python Logic (Line 133):
    `is_print = mach_up.startswith('PRINT') or mach_up.startswith('PLINE')`
  - Injected Excel Formula (Line 320):
    `TUBE_H_TPL = '=SUMPRODUCT((Production_Log!$F$3:$F$8963=F{r})*(LEFT(Production_Log!$B$3:$B$8963,5)="Print")*(ISERROR(SEARCH("(Varnish)",Production_Log!$D$3:$D$8963)))*Production_Log!$H$3:$H$8963)'`
- **Root Cause**: In Python, machines named `"PLINE 1"` or `"PLINE 2"` are recognized as printing lines. However, the injected Excel formula checks only `LEFT(Production_Log!$B$3:$B$8963,5)="Print"`. In Excel, `LEFT("PLINE 1", 5)` is `"PLINE"`, which does not equal `"Print"`.
- **Impact**: Any production logged under machine name `"PLINE"` evaluates to `0` in Excel formulas on the Dashboard, creating a severe divergence between Python dashboard reports and Excel workbook values.
- **Severity**: **HIGH**
- **Remediation**:
  Align Excel formula with Python logic:
  `*( (LEFT(Production_Log!$B$3:$B$8963,5)="Print") + (LEFT(Production_Log!$B$3:$B$8963,5)="PLINE") )*`

---

#### Finding F-12: Hardcoded Production Log Row Bound `$8963`
- **Exact Location**: `Scripts/sort_dashboard.py`, Lines 320, 327, 595, 661
- **Code Snippet**:
  `Production_Log!$F$3:$F$8963`
- **Root Cause**: All injected SUMPRODUCT formulas hardcode row 8963 as the upper boundary.
- **Impact**: When `Production_Log` grows beyond 8963 rows (e.g. in multi-month archives or high-volume logging), any rows below row 8963 are completely ignored in production sums and downtime calculations.
- **Severity**: **MEDIUM**
- **Remediation**:
  Dynamically determine `max_pl_row = ws_pl.max_row` and format formulas with dynamic upper bound `{max_row}` (e.g. `$F$3:$F${max_pl_row}`).

---

### 3.5. `Scripts/update_html.py`

#### Finding F-13: Rigid PID Partitioning for Product Types
- **Exact Location**: `Scripts/update_html.py`, Lines 216-217, Line 424
- **Code Snippet**:
  ```python
  tube_mtd = sum(v for k, v in mtd_by_pid.items() if k < 8000)
  pet_mtd  = sum(v for k, v in mtd_by_pid.items() if k >= 8000)
  ```
- **Root Cause**: The script hardcodes business logic assuming `PID < 8000` is Tube and `PID >= 8000` is PET.
- **Impact**: If a new PET product is assigned an ERP PID in the 6000 or 7000 range, or if a Tube product is assigned an 8000+ PID, all KPI calculations (MTD Production, MTD Scrap, MTD Dispatch) silently misclassify and corrupt the executive summary.
- **Severity**: **MEDIUM**
- **Remediation**:
  Classify product type by querying `Product_Catalog` col B (`Type` = `TUBE` or `PET`) or diameter unit (`ml` vs `mm`), not by arithmetic PID magnitude.
### 3.6. `Scripts/daily.py`

#### Finding F-14: Silent Error & Alert Suppression in Daily Reporting
- **Exact Location**: `Scripts/daily.py`, Lines 945-960
- **Code Snippet**:
  ```python
  # 1. Ignore INKs completely in daily summary
  if re.search(r'\binks?\b', lower_clean):
      continue

  # 2. Hide if already missing yesterday, except exceptions
  is_exception = re.search(r'\b(pet resin|master batch|slugs?)\b', lower_clean)
  if item_id and item_id in prev_missing and not is_exception:
      continue
  ```
- **Root Cause**: `daily.py` explicitly ignores ink mismatches and suppresses warnings for any items that were already missing on the previous run (stored in `previous_missing_items.json`).
- **Impact**: If 15 carton, lacquer, and cap SKUs go missing from ERP and are zeroed out by `update_inventory.py`, the supervisor is notified on Day 1. On Day 2 onwards, `daily.py` reports `"✓ ALL CHECKS PASSED: No errors, missing items, or mismatches detected!"`, masking persistent inventory wipeouts.
- **Severity**: **HIGH**
- **Remediation**:
  Remove suppression logic. Always report the complete list of missing or zeroed items in `error_summary.txt` and highlight recurring vs newly missing items clearly.

---

#### Finding F-15: Default Encoding Crash Risk on Windows
- **Exact Location**: `Scripts/daily.py`, Line 470
- **Code Snippet**:
  ```python
  with open(mismatch_log, 'r') as f:
      for line in f:
          print(f"    {line.rstrip()}")
  ```
- **Root Cause**: `mismatches.log` is written in UTF-8 by `alpha_checks.py` (Line 131: `encoding="utf-8"`). In `daily.py` Line 470, it is opened without specifying `encoding='utf-8'`. On Windows, Python defaults to `locale.getpreferredencoding()` (`cp1252`).
- **Impact**: If product names or customer remarks contain non-ASCII characters, em-dashes, or special symbols, `daily.py` crashes with `UnicodeDecodeError` mid-pipeline.
- **Severity**: **MEDIUM**
- **Remediation**:
  Explicitly specify `encoding='utf-8'` on all file open operations across all scripts.

---

### 3.7. `Scripts/alpha_checks.py`

#### Finding F-16: Non-Existent File Freshness Check False Positive
- **Exact Location**: `Scripts/alpha_checks.py`, Lines 49-50
- **Code Snippet**:
  ```python
  def check_freshness(filepath, max_hours=26, label=None):
      if not os.path.exists(filepath):
          return True  # file-not-found is handled elsewhere
  ```
- **Root Cause**: If a file is missing entirely, `check_freshness` returns `True`.
- **Impact**: If an updater script calls `check_freshness` without separate `os.path.exists` validation, the check succeeds silently.
- **Severity**: **HIGH**
- **Remediation**:
  Return `False` and print an explicit missing-file warning if `not os.path.exists(filepath)`.

---

#### Finding F-17: Non-Blocking Safety Assertions
- **Exact Location**: `Scripts/alpha_checks.py`, Lines 34-67
- **Root Cause**: `check_freshness` only prints a warning to stdout and returns `False`. None of the caller scripts check the boolean return value or halt execution.
- **Impact**: If `inventory.xls` or `dispatch.xls` is 3 weeks old, the pipeline prints a 2-line warning and proceeds to overwrite current production models with stale data.
- **Severity**: **HIGH**
- **Remediation**:
  Provide a strict mode or command-line flag (`--strict`) that raises `StaleDataError` and halts execution if input files exceed freshness thresholds.

---

#### Finding F-18: Unchecked File Replacement in `replace_copy_export`
- **Exact Location**: `Scripts/alpha_checks.py`, Lines 142-195
- **Code Snippet**:
  ```python
  latest_copy_path = max(matches, key=os.path.getmtime)
  os.replace(latest_copy_path, target_path)
  ```
- **Root Cause**: `replace_copy_export` replaces the target file with the latest copy file without verifying:
  1. That the copy file has a non-zero size.
  2. That the file is not currently locked by another process (e.g. browser download in progress).
  3. That the file is a valid OLE2/ZIP Excel file (valid magic bytes `\xd0\xcf\x11\xe0` or `PK\x03\x04`).
- **Impact**: If an operator triggers an export and runs the script before the download finishes, `replace_copy_export` replaces the production database with a 0-byte or corrupt temporary file.
- **Severity**: **MEDIUM**
- **Remediation**:
  Verify file size > 1024 bytes, check magic header bytes, and test file readability with `open(..., 'rb')` before replacing.

---

### 3.8. `Scripts/build_archives.py`

#### Finding F-19: Sorting Strategy Conflict for Active Monthly Workbook
- **Exact Location**: `Scripts/build_archives.py`, Line 41 vs other scripts
- **Code Comparison**:
  - `build_archives.py` (Line 41):
    `active_files = sorted(glob.glob(r"d:\Alpha\Tubex_*.xlsx"), key=os.path.getmtime)`
  - `daily.py`, `update_production.py`, `sort_dashboard.py`:
    `sorted(glob.glob(...))[-1]` (alphabetical sort)
- **Root Cause**: `build_archives.py` picks the active month workbook by most recent file modification timestamp (`getmtime`), whereas all daily update scripts pick the active workbook alphabetically.
- **Impact**: If an older monthly workbook (e.g. `Tubex_July26.xlsx`) is opened, inspected, or touched, its `mtime` updates. `build_archives.py` will treat `Tubex_July26.xlsx` as the active August workbook, while `daily.py` writes to `Tubex_Aug26.xlsx`, causing archive desynchronization.
- **Severity**: **MEDIUM**
- **Remediation**:
  Standardize active workbook resolution using a single unified function in `alpha_checks.py` that verifies the current calendar month and file naming convention.

---

## 4. ERP Ingestion & Schema Vulnerabilities Analysis

### 4.1. `Production.xlsx` (Day wise & FG Stock In hand)
1. **Header Row Shifting**: In `Production Day wise`, Imran often inserts notes or shifts rows. `detect_header_row` checks `nrows=10`, but if more than 10 rows of notes are added, detection fails.
2. **Rejection Formula Assumption**: The formula `=IFERROR(I{r}/(H{r}+I{r}), "")` assumes Imran's Column H is Good and Column I is Rejection. If Imran swaps columns in `Production.xlsx`, total production is calculated erroneously.

### 4.2. `inventory.xls` (Item Wise Consolidated Report)
1. **OLE2 Header Warning**: `inventory.xls` outputs `WARNING *** OLE2 inconsistency: SSCS size is 0 but SSAT size is non-zero` when read by `xlrd`. While `xlrd` parses it, newer parsers or strict libraries fail to open the file.
2. **Category Row Mixing**: Category banners (e.g. `'CONSUMABLE ITEMS - ADMIN.'`) sit in the same column as numeric Item IDs. Ingestion relies on `int(float(val))` throwing a `ValueError` to skip headers. If a category or item description begins with a number (e.g. `'19mm Caps'`), it can be misparsed as an Item ID.

### 4.3. `dispatch.xls` & `dispatch_pet.xls` (Date Wise Dispatch Report)
1. **Merged / Blank Sub-rows**: Each dispatch transaction spans multiple rows (e.g. Row 9 has transaction details, Row 10 has secondary ID and date, Row 11 has `'POSTED' / 'ACTUAL'`). The parser uses `int(float(col0))` to detect primary rows, which skips multi-line delivery splits.
2. **Product Name Delimitation**: The parser relies on `pd.isna(row[1])` to identify a new product header banner. If ERP changes layout so `row[1]` contains data on header rows, product grouping breaks entirely.

---

## 5. PIPELINE.md vs Reality & Operational Blind Spots

| Aspect | `PIPELINE.md` Specification | Actual Implementation | Integrity Risk |
|---|---|---|---|
| **Execution Sequence** | Dispatch → Production → Inventory → Sort → HTML | Production → Inventory → Dispatch → Sort → Archives → HTML | Medium (Order dependencies) |
| **Archival Step** | Not documented in daily pipeline | Executed as Step 5 in `daily.py` via `build_archives.py` | Low (Hidden dependency) |
| **WIP Handling** | Manual script `update_wip.py` | Integrated into Step 4 of `daily.py` with 2s timeout | Low |
| **Data Validation** | Documents `check_freshness` and `check_not_locked` | Safety checks are non-blocking and lack schema validation | High (Silent corruption) |
| **Error Handling** | Claims exclusive lock stops corruption | Does not check lock on ERP inputs or secondary files | Medium |

---

## 6. Prioritized Remediation Roadmap

### Phase 1: Immediate Safety Hardening (Critical / High)
1. **Fix Inventory Wipeout Logic**: Update `update_inventory.py` to preserve historical Opening balances and require full catalog verification before zeroing missing items.
2. **Fix Alias Resolution & Dropped Production**: In `update_production.py`, enforce strict error raising when unmapped products occur instead of silently continuing with `PID=None`.
3. **Fix Date Parsing**: Replace `pd.Timestamp(date_raw)` with explicit `dayfirst=True` parsing across `update_production.py` and `update_dispatch.py`.
4. **Fix Dashboard Formula Regex**: Update `sort_dashboard.py` regex to prevent corruption of 2D lookup ranges (`MRP!D3:F50`).
5. **Fix Machine Formula Alignment**: Update `TUBE_H_TPL` in `sort_dashboard.py` to match `"Print"` and `"PLINE"`.

### Phase 2: Schema & Ingestion Robustness (High / Medium)
6. **Implement Dynamic Header & Column Lookups**: Replace all hardcoded positional indices (`col7`, `header=1`, `range(1, 9)`) in `update_dispatch.py` and `update_production.py` with dynamic column name mapping.
7. **Expand `alpha_checks.py`**:
   - Add `validate_schema(df, required_columns)`.
   - Add `validate_row_count(df, min_rows, max_rows)`.
   - Add `check_formula_integrity(workbook_path)`.
   - Ensure `check_freshness` returns `False` for missing files and supports strict blocking mode.
8. **Fix Daily Mismatch Alerting**: Remove suppression of "INK" items and recurring missing items in `daily.py`.

### Phase 3: Code Hygiene & Architecture Alignment (Medium / Low)
9. **Standardize File Selection**: Replace ad-hoc globbing across all scripts with a central `get_active_tubex_file()` utility in `alpha_checks.py`.
10. **Align Execution Order**: Harmonize `PIPELINE.md` and `daily.py` to ensure consistent execution sequences.
11. **Enforce UTF-8 Everywhere**: Add `encoding='utf-8'` to all file open calls and ensure all batch scripts set `chcp 65001`.
