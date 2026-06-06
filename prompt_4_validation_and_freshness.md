# PROMPT: Implement Post-Write Validation + Freshness Improvements

## YOUR TASK

You will make 3 improvements to the Tubex repository at `d:\Alpha`. Each is precisely specified.

**Rules:**
- Do NOT modify any data-processing or business logic
- Only ADD validation code AFTER existing save operations
- Only MODIFY the freshness banner in Tubex.html — do not change any other HTML/CSS
- Test that `Run_All_Updates.bat` still works after your changes

---

## TASK 1: Add Post-Write Validation to All Updater Scripts

After each script saves to Excel, it should re-open the file (read-only) and verify the write actually worked. This catches silent openpyxl failures.

### 1A. `d:\Alpha\Scripts\update_production.py`

Find the `main()` function. After the line `print("  Saved:   " + os.path.basename(ac_path))` (around line ~930), add this validation block:

```python
    # ── Post-write validation ────────────────────────────────────────────
    print("\n  Validating writes...")
    try:
        wb_check = openpyxl.load_workbook(ac_path, read_only=True, data_only=True)
        
        # Check Production_Log
        ws_check = wb_check['Production_Log']
        pl_count = sum(1 for r in range(3, ws_check.max_row + 1) 
                       if ws_check.cell(r, 1).value is not None)
        if pl_count == len(source_rows):
            print(f"  ✓ Production_Log: {pl_count} rows written (matches source)")
        else:
            print(f"  !! Production_Log: expected {len(source_rows)} rows, found {pl_count}")
        
        # Check FG Stock
        if 'FG Stock' in wb_check.sheetnames and fg_rows:
            ws_fg = wb_check['FG Stock']
            fg_count = sum(1 for r in range(4, ws_fg.max_row + 1) 
                          if ws_fg.cell(r, 1).value is not None)
            if fg_count == len(fg_rows):
                print(f"  ✓ FG Stock: {fg_count} rows written (matches source)")
            else:
                print(f"  !! FG Stock: expected {len(fg_rows)} rows, found {fg_count}")
        
        wb_check.close()
    except Exception as e:
        print(f"  !! Validation error: {e}")
```

**Important**: You need to add `import openpyxl` at the top of the file if it's not already there. Check the existing imports — the script uses `openpyxl` inside functions already, so add it at module level near the other imports.

### 1B. `d:\Alpha\Scripts\update_inventory.py`

Find the `main()` function. After the line `print("  Saved: " + os.path.basename(excel_path))` (around line ~220), add:

```python
    # ── Post-write validation ────────────────────────────────────────────
    print("  Validating writes...")
    try:
        from openpyxl import load_workbook as _lw
        wb_check = _lw(excel_path, read_only=True, data_only=True)
        ws_check = wb_check['Inventory']
        # Spot-check: count rows with non-None values in col E (Opening)
        inv_count = sum(1 for r in range(2, ws_check.max_row + 1) 
                       if ws_check.cell(r, 5).value is not None)
        print(f"  ✓ Inventory: {inv_count} rows have Opening values")
        wb_check.close()
    except Exception as e:
        print(f"  !! Validation error: {e}")
```

### 1C. `d:\Alpha\Scripts\update_dispatch.py`

Find the `main()` function. After `print("  Saved:  " + os.path.basename(ac_path))` (around line ~415), add:

```python
    # ── Post-write validation ────────────────────────────────────────────
    print("  Validating writes...")
    try:
        wb_check = load_workbook(ac_path, read_only=True, data_only=True)
        ws_check = wb_check[DASHBOARD_SHEET]
        dispatch_cells = sum(1 for r in range(DASHBOARD_ROW_MIN, DASHBOARD_ROW_MAX + 1) 
                            if ws_check.cell(r, DASHBOARD_DISP_COL).value is not None
                            and isinstance(ws_check.cell(r, DASHBOARD_DISP_COL).value, (int, float)))
        if dispatch_cells == updated_count:
            print(f"  ✓ Dashboard col K: {dispatch_cells} dispatch values verified")
        else:
            print(f"  !! Dashboard col K: expected {updated_count} values, found {dispatch_cells}")
        wb_check.close()
    except Exception as e:
        print(f"  !! Validation error: {e}")
```

---

## TASK 2: Enhanced Freshness Banner in Tubex.html

**File:** `d:\Alpha\Tubex.html`

Find the existing freshness check code. Search for `stale-banner` or `lastUpdated` in the JavaScript. There's currently a simple >24h check.

Replace the existing freshness/stale banner logic with this enhanced version that:
- Shows exact age (e.g., "Updated 3 hours ago" vs "Updated 2 days ago")
- Uses yellow for 24–48h and red for >48h
- Shows green "Live" indicator when fresh

Find the `checkStaleness` or equivalent function and replace it with:

```javascript
function checkFreshness() {
  const banner = document.getElementById('stale-banner');
  if (!banner || typeof lastUpdated === 'undefined') return;
  
  const updated = new Date(lastUpdated);
  const now = new Date();
  const hoursAgo = (now - updated) / (1000 * 60 * 60);
  
  // Format the age string
  let ageStr;
  if (hoursAgo < 1) {
    const mins = Math.round(hoursAgo * 60);
    ageStr = mins + ' minute' + (mins !== 1 ? 's' : '') + ' ago';
  } else if (hoursAgo < 24) {
    const hrs = Math.round(hoursAgo);
    ageStr = hrs + ' hour' + (hrs !== 1 ? 's' : '') + ' ago';
  } else {
    const days = Math.round(hoursAgo / 24);
    ageStr = days + ' day' + (days !== 1 ? 's' : '') + ' ago';
  }
  
  if (hoursAgo > 48) {
    // RED — critically stale
    banner.style.display = 'block';
    banner.style.background = 'linear-gradient(135deg, #e74c3c, #c0392b)';
    banner.style.color = '#fff';
    banner.innerHTML = '⚠️ Data is ' + ageStr + ' — download fresh ERP exports and run the updater';
  } else if (hoursAgo > 24) {
    // YELLOW — stale
    banner.style.display = 'block';
    banner.style.background = 'linear-gradient(135deg, #f39c12, #e67e22)';
    banner.style.color = '#fff';
    banner.innerHTML = '⏳ Data updated ' + ageStr + ' — consider refreshing';
  } else {
    // FRESH — hide banner or show green
    banner.style.display = 'none';
  }
}
```

Then find where `checkStaleness()` (or whatever the old function name is) is called in the init block, and replace it with `checkFreshness()`.

Also find the **stale-banner CSS** and ensure it has these styles (update if different):
```css
#stale-banner {
  display: none;
  padding: 10px 16px;
  text-align: center;
  font-size: 13px;
  font-weight: 500;
  border-radius: 8px;
  margin: 8px 12px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}
```

If the freshness banner element doesn't exist in the HTML yet, add it right after the opening of the main content area (after the tab navigation bar):
```html
<div id="stale-banner"></div>
```

---

## TASK 3: Update .gitignore for New Files

**File:** `d:\Alpha\.gitignore`

Add these lines at the end of the file:

```gitignore

# ── Documentation (keep in repo) ──
# MONTH_CHECKLIST.md and PIPELINE.md are tracked intentionally
# (they are NOT in the ignore list)

# ── Pre-run backups (local only) ──
Logs/backup_*
```

This ensures the pre-run backups created by `Run_All_Updates.bat` are not committed to git (they contain business data).

---

## VERIFICATION

After all changes:

1. **Close Excel** (all Tubex files)
2. Run `Run_All_Updates.bat`
3. Check output — you should now see validation lines like:
   ```
   ✓ Production_Log: 245 rows written (matches source)
   ✓ FG Stock: 18 rows written (matches source)
   ✓ Inventory: 42 rows have Opening values
   ✓ Dashboard col K: 15 dispatch values verified
   ```
4. Open `Tubex.html` in browser
5. The freshness banner should show green/hidden (if data is fresh) or yellow/red (if old)
6. Run `git status` — `MONTH_CHECKLIST.md` and `PIPELINE.md` should appear as new files to commit. `Logs/backup_*` should NOT appear.

---

## IMPORTANT RULES

1. Validation code goes AFTER the existing save — never before
2. Use `read_only=True, data_only=True` when validating (don't lock the file)
3. Always wrap validation in try/except — a validation failure should print a warning, NOT crash the script
4. Do NOT change any existing data processing logic
5. The freshness banner enhancement should work with the existing `lastUpdated` variable that `update_html.py` already injects
