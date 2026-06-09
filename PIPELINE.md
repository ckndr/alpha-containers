# Tubex Data Pipeline

## How Data Flows (Read This First)

```
YOUR INPUTS (manual downloads)          SCRIPTS (automated)              OUTPUT
================================       ======================          ========

Production.xlsx (from Imran)  ──┐
                                ├──► update_production.py ──► Production_Log sheet
                                │                          ──► FG Stock sheet
                                │
inventory.xls (from ERP)  ──────┤──► update_inventory.py  ──► Inventory sheet
                                │
dispatch.xls (from ERP)  ───────┤──► update_dispatch.py   ──► Dashboard col K
dispatch_pet.xls (from ERP) ────┘
                                │
            (all Excel data) ───┤──► sort_dashboard.py    ──► Dashboard row order
                                │
            (all Excel data) ───┴──► update_html.py       ──► Tubex.html ──► GitHub Pages
                                                           ──► sw.js (cache bump)
```

## Script Execution Order (MUST follow this sequence)

```
Step 1: update_dispatch.py      ← FIRST (writes fresh dispatch data to Dashboard)
Step 2: update_production.py    ← Second (populates Production_Log and FG Stock)
Step 3: update_inventory.py     ← Third (updates Inventory sheet)
Step 4: sort_dashboard.py       ← AFTER Steps 1-3 (needs fresh production + dispatch data)
Step 5: update_html.py          ← LAST (reads everything, generates final HTML)
```

> `Run_All_Updates.bat` runs these in the correct order automatically.

## Manual/Optional Scripts (NOT in the daily pipeline)

| Script | When to Use |
|---|---|
| `update_wip.py` | When Aurangzeb sends WIP message on WhatsApp |
| `update_mrp.py` | ⚠️ ONE-TIME migration only. Do NOT run again. |
| `update_from_images.py` | When you have screenshot images instead of Production.xlsx |
| `Excel_Scrap.py` | When you need to dump Excel data for AI context |

## What Each Script Reads and Writes

### update_production.py
```
READS:  Production.xlsx → "Production Day wise" sheet (production entries)
        Production.xlsx → "FG Stock In hand" sheet (finished goods)
WRITES: Tubex_v*.xlsx  → "Production_Log" sheet (wipe and rewrite all rows)
        Tubex_v*.xlsx  → "FG Stock" sheet (wipe and rewrite, latest date only)
```

### update_inventory.py
```
READS:  inventory.xls  → ERP Item Wise Consolidated report
WRITES: Tubex_v*.xlsx  → "Inventory" sheet cols E/F/G (Opening/Received/Issued)
```

### update_dispatch.py
```
READS:  dispatch.xls     → ERP Dispatch Report (TUBEX-ALUM)
        dispatch_pet.xls  → ERP Dispatch Report (TUBEX-PET)
        Tubex_v*.xlsx     → "Product_Catalog" sheet (name→PID mapping)
WRITES: Tubex_v*.xlsx     → "Tubex_Dashboard" col K (dispatch quantities)
```

### sort_dashboard.py
```
READS:  Tubex_v*.xlsx → "Tubex_Dashboard" (all product rows)
        Tubex_v*.xlsx → "Production_Log" (MTD production by PID)
        Tubex_v*.xlsx → "MRP" (order quantities)
WRITES: Tubex_v*.xlsx → "Tubex_Dashboard" (reorders rows: active→top, inactive→bottom)
```

### update_html.py
```
READS:  Tubex_v*.xlsx → ALL sheets (Dashboard, Production_Log, Inventory, MRP, BOM, Catalog, FG Stock)
WRITES: Tubex.html    → Injects JSON data constants between /* MARKER */ comments
        sw.js         → Updates CACHE_NAME timestamp
```

## File Dependencies (What breaks if you delete something)

| File | If Missing... |
|---|---|
| `Tubex_v*.xlsx` | **Everything breaks.** This is the central database. |
| `Production.xlsx` | Step 1 fails. Production_Log stays stale. |
| `inventory.xls` | Step 2 fails. Inventory numbers stay stale. |
| `dispatch.xls` | Step 3 fails. Dispatch column stays empty. |
| `dispatch_pet.xls` | Step 3 fails. PET dispatch stays empty. |
| `alpha_checks.py` | All scripts fail on import. Shared safety utilities. |
| `Tubex.html` | Step 5 fails. No web dashboard generated. |
| `sw.js` | PWA offline caching stops. Dashboard still works online. |

## Shared Module: alpha_checks.py

All scripts import from this module:
- `check_freshness(filepath)` — Warns if ERP file is stale (>26h old)
- `check_not_locked(filepath)` — Stops if Excel has the file open
- `log_mismatches(source, items)` — Saves unmatched products to Logs/mismatches.log
- `replace_copy_export(folder, name)` — Handles "- copy" file renaming from ERP downloads
