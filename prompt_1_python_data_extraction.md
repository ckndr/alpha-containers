# PROMPT 1 OF 3: Modify update_html.py — Extract & Inject New Data

## YOUR TASK

You must modify `d:\Alpha\Scripts\update_html.py` to extract 4 additional datasets from the Excel workbook and inject them as JavaScript constants into `Tubex.html`. The existing script already does this for `DASH_DATA`, `PRODUCTS`, and `BOM` — you are adding `INVENTORY_DATA`, `PRODUCTION_LOG_DATA`, `FG_STOCK_DATA`, and `MRP_DATA`.

**DO NOT** modify `Tubex.html` in this prompt — that is handled in Prompts 2 and 3.

---

## EXISTING ARCHITECTURE (READ CAREFULLY)

### File Locations
- Script: `d:\Alpha\Scripts\update_html.py`
- Excel: `d:\Alpha\Tubex_v10_29.xlsx` (glob pattern `Tubex_v*.xlsx`, latest is used)
- HTML: `d:\Alpha\Tubex.html`

### How the Injection Pattern Works
The script loads the HTML as a string, finds marker comment pairs, and replaces everything between them:

```python
# Existing markers:
#   /* DATA_START */ ... /* DATA_END */       → DASH_DATA JSON
#   /* CATALOG_START */ ... /* CATALOG_END */ → PRODUCTS + BOM arrays

marker_start = '/* DATA_START */'
marker_end   = '/* DATA_END */'
pos_start    = html.find(marker_start)
pos_end      = html.find(marker_end)
new_data_block = f"{marker_start}\nconst DASH_DATA = {json.dumps(dash_data)};\n{marker_end}"
html = html[:pos_start] + new_data_block + html[pos_end + len(marker_end):]
```

You must add **4 NEW marker pairs** in the same pattern. On the **first run**, auto-insert the markers into the HTML (just like the existing `CATALOG_START` self-patch logic at lines 434-447).

### Excel Workbook Sheets Available
The workbook `Tubex_v10_29.xlsx` is already loaded at line 46:
```python
wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
```
Currently accessed sheets:
- `ws_pl  = wb['Production_Log']`
- `ws_db  = wb['Tubex_Dashboard']`
- `ws_bom = wb['BOM']`
- `ws_cat = wb['Product_Catalog']`

You need to also access:
- `ws_inv = wb['Inventory']`
- `ws_fg  = wb['FG Stock']`
- `ws_mrp = wb['MRP']`

---

## WHAT TO ADD — 4 DATA EXTRACTORS

### 1. INVENTORY_DATA (from `Inventory` sheet)

**Sheet layout** (the Inventory sheet in Tubex workbook):
- Row 1: Title row (e.g., "Raw Material Inventory (01-06-2026 to 03-06-2026)")
- Row 2: Headers
- Row 3+: Data rows

| Column | Letter | Content |
|--------|--------|---------|
| A | Item ID | Integer (e.g., 2, 6, 19, 22...) |
| B | Category | Text (SLUG, BASE COAT, CAP, CARTON, ZINC POWDER, LACQUER, TAPE, LATEX, THINNER, INK...) |
| C | Item Name | Text |
| D | UOM | Text (kg, L, pcs) |
| E | Opening | Number (from ERP) |
| F | Received | Number (from ERP) |
| G | Issued | Number (from ERP) |
| H | Store Balance | Formula: `=E+F-G` (read as value with data_only=True) |
| I | Work In Process | Number or formula |

**Extraction logic:**
```python
inventory_data = []
ws_inv = wb['Inventory']
for r in range(2, ws_inv.max_row + 1):
    item_id = ws_inv.cell(row=r, column=1).value
    if item_id is None:
        continue
    try:
        item_id = int(item_id)
    except (TypeError, ValueError):
        continue
    
    cat      = str(ws_inv.cell(row=r, column=2).value or '').strip()
    name     = str(ws_inv.cell(row=r, column=3).value or '').strip()
    uom      = str(ws_inv.cell(row=r, column=4).value or '').strip()
    opening  = float(ws_inv.cell(row=r, column=5).value or 0)
    received = float(ws_inv.cell(row=r, column=6).value or 0)
    issued   = float(ws_inv.cell(row=r, column=7).value or 0)
    balance  = ws_inv.cell(row=r, column=8).value
    wip      = ws_inv.cell(row=r, column=9).value
    
    # balance and wip may be formulas — data_only=True gives calculated values
    balance = float(balance) if balance and isinstance(balance, (int, float)) else 0.0
    wip     = float(wip) if wip and isinstance(wip, (int, float)) else 0.0
    
    if not name:
        continue
    
    inventory_data.append({
        'id': item_id,
        'cat': cat,
        'name': name,
        'uom': uom,
        'opening': round(opening, 2),
        'received': round(received, 2),
        'issued': round(issued, 2),
        'balance': round(balance, 2),
        'wip': round(wip, 2),
    })
```

Also extract the title row date range:
```python
inv_title = str(ws_inv.cell(row=1, column=1).value or 'Inventory')
```

**JS output format:**
```javascript
const INVENTORY_DATA = {
  title: "Raw Material Inventory (01-06-2026 to 03-06-2026)",
  items: [
    {id:2, cat:"SLUG", name:"30X4.3", uom:"kg", opening:1200, received:500, issued:800, balance:900, wip:150},
    ...
  ]
};
```

---

### 2. PRODUCTION_LOG_DATA (from `Production_Log` sheet)

The script already iterates `ws_pl` (lines 95-149). You should extract ALL rows for the **current reporting month** (already determined as `cur_month` / `cur_year` at line 64-66).

**Sheet layout (Production_Log):**
| Col | Letter | Content |
|-----|--------|---------|
| A(1) | Date | datetime |
| B(2) | Machine | Text (Press1, Print1, PLine1, PF1...) |
| C(3) | Customer | Text |
| D(4) | Product | Text |
| E(5) | Dia | Number or text |
| F(6) | PID | Integer |
| G(7) | Target/Total | Number |
| H(8) | Good | Number |
| I(9) | Reject | Number |
| J(10) | Waste% | Formula (percentage) |

**Extraction logic:**
```python
prodlog_data = []
for row in ws_pl.iter_rows(min_row=3, values_only=True):
    row_date = row[0]
    machine  = row[1]
    if not row_date or not machine:
        continue
    if not isinstance(row_date, (date, datetime)):
        continue
    row_dt = row_date if isinstance(row_date, datetime) else datetime(row_date.year, row_date.month, row_date.day)
    if row_dt.month != cur_month or row_dt.year != cur_year:
        continue
    
    good    = int(row[7]) if row[7] and isinstance(row[7], (int, float)) else 0
    reject  = int(row[8]) if row[8] and isinstance(row[8], (int, float)) else 0
    total   = int(row[6]) if row[6] and isinstance(row[6], (int, float)) else 0
    
    prodlog_data.append({
        'date': row_dt.strftime('%d-%b'),
        'machine': str(machine).strip(),
        'customer': str(row[2] or '').strip(),
        'product': str(row[3] or '').strip(),
        'dia': str(row[4] or '').strip(),
        'pid': int(row[5]) if row[5] else None,
        'total': total,
        'good': good,
        'reject': reject,
    })
```

**JS output format:**
```javascript
const PRODUCTION_LOG_DATA = {
  month: "June 2026",
  rows: [
    {date:"01-Jun", machine:"Print1", customer:"Samsol", product:"TUBES", dia:"25", pid:3726, total:5000, good:4800, reject:200},
    ...
  ]
};
```

---

### 3. FG_STOCK_DATA (from `FG Stock` sheet)

**Sheet layout (FG Stock in Tubex workbook):**
- Row 1: Title "FG STOCK IN HAND — Last Updated: 03-Jun-2026"
- Row 2: (empty or sub-header)
- Row 3: Column headers — Sr, PID, Customer, Product, Dia, FG Qty, Status, Dispatch Remarks
- Row 4+: Data rows

| Col | Content |
|-----|---------|
| A(1) | Sr# |
| B(2) | PID |
| C(3) | Customer |
| D(4) | Product Name |
| E(5) | Dia |
| F(6) | FG Qty |
| G(7) | Status (OK, Not Ready, In Progress) |
| H(8) | Dispatch Remarks |

**Extraction logic:**
```python
fg_data = []
ws_fg = wb['FG Stock']
fg_title = str(ws_fg.cell(row=1, column=1).value or 'FG Stock')

for r in range(4, ws_fg.max_row + 1):
    pid_val = ws_fg.cell(row=r, column=2).value
    product = ws_fg.cell(row=r, column=4).value
    if not product:
        continue
    
    fg_qty = ws_fg.cell(row=r, column=6).value
    fg_qty = int(fg_qty) if fg_qty and isinstance(fg_qty, (int, float)) else 0
    
    fg_data.append({
        'sr': ws_fg.cell(row=r, column=1).value,
        'pid': int(pid_val) if pid_val else None,
        'customer': str(ws_fg.cell(row=r, column=3).value or '').strip(),
        'product': str(product).strip(),
        'dia': str(ws_fg.cell(row=r, column=5).value or '').strip(),
        'qty': fg_qty,
        'status': str(ws_fg.cell(row=r, column=7).value or '').strip(),
        'remarks': str(ws_fg.cell(row=r, column=8).value or '').strip(),
    })
```

**JS output format:**
```javascript
const FG_STOCK_DATA = {
  title: "FG STOCK IN HAND — Last Updated: 03-Jun-2026",
  rows: [
    {sr:1, pid:3726, customer:"Samsol International", product:"TUBES", dia:"25", qty:15000, status:"OK", remarks:"Ready for dispatch"},
    ...
  ]
};
```

---

### 4. MRP_DATA (from `MRP` sheet)

This is the most complex one. The MRP sheet has multiple sections:

**Section 1: Required Orders (rows 2–15)**
- Row 2: Headers (Dia, Customer, Product Name, Product ID, Job Order#, Required Qty, Produced, Remaining Balance)
- Rows 3–14: Active tube orders
- Row 15: Total row

**Section 2: Tubes Material Requirement (rows ~17–101)**
- Row 17: Headers (Item ID, Category, Item Name, UOM, Required Qty, Current Stock, Surplus/Deficit, Product Names, Status)
- Rows 18–101: Material items

**Section 3: PET Orders (rows ~104–112)**
**Section 4: PET Material Requirement (rows ~116–123)**
**Section 5: INK Table (rows ~127–158)**

**Extraction logic (simplified — read what's there):**
```python
ws_mrp = wb['MRP']

# Read orders (rows 3-14)
mrp_orders = []
for r in range(3, 15):
    pid = ws_mrp.cell(row=r, column=4).value
    if not pid:
        continue
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        continue
    
    required = ws_mrp.cell(row=r, column=6).value
    produced = ws_mrp.cell(row=r, column=7).value
    remaining = ws_mrp.cell(row=r, column=8).value
    
    mrp_orders.append({
        'dia': ws_mrp.cell(row=r, column=1).value,
        'customer': str(ws_mrp.cell(row=r, column=2).value or '').strip(),
        'product': str(ws_mrp.cell(row=r, column=3).value or '').strip(),
        'pid': pid,
        'jobOrder': str(ws_mrp.cell(row=r, column=5).value or '').strip(),
        'required': float(required) if required and isinstance(required, (int, float)) else 0,
        'produced': float(produced) if produced and isinstance(produced, (int, float)) else 0,
        'remaining': float(remaining) if remaining and isinstance(remaining, (int, float)) else 0,
        'remarks': str(ws_mrp.cell(row=r, column=9).value or '').strip(),
    })

# Read material requirements (rows 18-101 for tubes)
mrp_materials = []
for r in range(18, 102):
    item_id = ws_mrp.cell(row=r, column=1).value
    if item_id is None:
        continue
    try:
        item_id_int = int(item_id)
    except (TypeError, ValueError):
        continue
    
    cat      = str(ws_mrp.cell(row=r, column=2).value or '').strip()
    name     = str(ws_mrp.cell(row=r, column=3).value or '').strip()
    uom      = str(ws_mrp.cell(row=r, column=4).value or '').strip()
    req_qty  = ws_mrp.cell(row=r, column=5).value
    stock    = ws_mrp.cell(row=r, column=6).value
    surplus  = ws_mrp.cell(row=r, column=7).value
    products = str(ws_mrp.cell(row=r, column=8).value or '').strip()
    status   = str(ws_mrp.cell(row=r, column=9).value or '').strip()
    
    req_qty = round(float(req_qty), 2) if req_qty and isinstance(req_qty, (int, float)) else 0
    stock   = round(float(stock), 2) if stock and isinstance(stock, (int, float)) else 0
    surplus = round(float(surplus), 2) if surplus and isinstance(surplus, (int, float)) else 0
    
    if not name:
        continue
    
    mrp_materials.append({
        'id': item_id_int,
        'cat': cat,
        'name': name,
        'uom': uom,
        'required': req_qty,
        'stock': stock,
        'surplus': surplus,
        'products': products,
        'status': status,
        'section': 'tube',
    })

# Read PET material requirements (rows 116-123)
for r in range(116, 124):
    item_id = ws_mrp.cell(row=r, column=1).value
    if item_id is None:
        continue
    try:
        item_id_int = int(item_id)
    except (TypeError, ValueError):
        continue
    
    cat      = str(ws_mrp.cell(row=r, column=2).value or '').strip()
    name     = str(ws_mrp.cell(row=r, column=3).value or '').strip()
    uom      = str(ws_mrp.cell(row=r, column=4).value or '').strip()
    req_qty  = ws_mrp.cell(row=r, column=5).value
    stock    = ws_mrp.cell(row=r, column=6).value
    surplus  = ws_mrp.cell(row=r, column=7).value
    products = str(ws_mrp.cell(row=r, column=8).value or '').strip()
    status   = str(ws_mrp.cell(row=r, column=9).value or '').strip()
    
    req_qty = round(float(req_qty), 2) if req_qty and isinstance(req_qty, (int, float)) else 0
    stock   = round(float(stock), 2) if stock and isinstance(stock, (int, float)) else 0
    surplus = round(float(surplus), 2) if surplus and isinstance(surplus, (int, float)) else 0
    
    if not name:
        continue
    
    mrp_materials.append({
        'id': item_id_int,
        'cat': cat,
        'name': name,
        'uom': uom,
        'required': req_qty,
        'stock': stock,
        'surplus': surplus,
        'products': products,
        'status': status,
        'section': 'pet',
    })

# Read INK table (rows 127-158)
mrp_inks = []
for r in range(127, 159):
    item_id = ws_mrp.cell(row=r, column=1).value
    if item_id is None:
        continue
    try:
        item_id_int = int(item_id)
    except (TypeError, ValueError):
        continue
    
    name    = str(ws_mrp.cell(row=r, column=3).value or '').strip()
    uom     = str(ws_mrp.cell(row=r, column=4).value or '').strip()
    avg_use = ws_mrp.cell(row=r, column=5).value  # Avg Monthly Usage
    days    = ws_mrp.cell(row=r, column=6).value   # Days of Stock
    status  = str(ws_mrp.cell(row=r, column=7).value or '').strip()
    stock   = ws_mrp.cell(row=r, column=8).value   # Current Stock
    
    avg_use = round(float(avg_use), 2) if avg_use and isinstance(avg_use, (int, float)) else 0
    days    = round(float(days), 1) if days and isinstance(days, (int, float)) else 0
    stock   = round(float(stock), 2) if stock and isinstance(stock, (int, float)) else 0
    
    if not name:
        continue
    
    mrp_inks.append({
        'id': item_id_int,
        'name': name,
        'uom': uom,
        'avgUse': avg_use,
        'daysLeft': days,
        'status': status,
        'stock': stock,
    })
```

**JS output format:**
```javascript
const MRP_DATA = {
  title: "May 2026 Tube Required Orders",
  orders: [...],
  materials: [...],
  inks: [...]
};
```

---

## INJECTION INTO HTML

After you build the 4 data objects, inject them into the HTML using the same marker pattern. Add the self-patch logic (like lines 434-447) that inserts new markers on the first run.

The markers should be placed **after** the `/* CATALOG_END */` marker, inside the `<script>` tag, before `const CAT_ICON`.

**First-run self-patch** — if `/* INVENTORY_START */` is not in the HTML, add all 4 marker pairs:
```python
if '/* INVENTORY_START */' not in html:
    print("  First run: adding new data markers to HTML...")
    # Insert after CATALOG_END
    html = html.replace(
        '/* CATALOG_END */',
        '/* CATALOG_END */\n'
        '/* INVENTORY_START */\n/* INVENTORY_END */\n'
        '/* PRODLOG_START */\n/* PRODLOG_END */\n'
        '/* FGSTOCK_START */\n/* FGSTOCK_END */\n'
        '/* MRP_START */\n/* MRP_END */'
    )
```

Then inject each data block:
```python
def inject_block(html, start_marker, end_marker, js_content):
    ps = html.find(start_marker)
    pe = html.find(end_marker)
    if ps == -1 or pe == -1:
        raise RuntimeError(f"{start_marker} / {end_marker} markers not found.")
    return html[:ps] + f"{start_marker}\n{js_content}\n{end_marker}" + html[pe + len(end_marker):]

html = inject_block(html, '/* INVENTORY_START */', '/* INVENTORY_END */',
    f"const INVENTORY_DATA = {json.dumps({'title': inv_title, 'items': inventory_data}, ensure_ascii=False)};")
html = inject_block(html, '/* PRODLOG_START */', '/* PRODLOG_END */',
    f"const PRODUCTION_LOG_DATA = {json.dumps({'month': month_name, 'rows': prodlog_data}, ensure_ascii=False)};")
html = inject_block(html, '/* FGSTOCK_START */', '/* FGSTOCK_END */',
    f"const FG_STOCK_DATA = {json.dumps({'title': fg_title, 'rows': fg_data}, ensure_ascii=False)};")
html = inject_block(html, '/* MRP_START */', '/* MRP_END */',
    f"const MRP_DATA = {json.dumps({'title': mrp_title, 'orders': mrp_orders, 'materials': mrp_materials, 'inks': mrp_inks}, ensure_ascii=False)};")
```

Get `mrp_title` from: `mrp_title = str(ws_mrp.cell(row=1, column=1).value or 'Material Requirement Plan')`

---

## PRINT STATEMENTS

Add summary prints matching the existing style:
```python
print(f"\nInventory items loaded: {len(inventory_data)}")
print(f"Production Log rows (current month): {len(prodlog_data)}")
print(f"FG Stock rows: {len(fg_data)}")
print(f"MRP orders: {len(mrp_orders)}  |  Materials: {len(mrp_materials)}  |  Inks: {len(mrp_inks)}")
```

---

## IMPORTANT RULES

1. Do NOT modify the existing DASH_DATA, PRODUCTS, or BOM logic — it works perfectly.
2. Do NOT modify Tubex.html structure — only inject data between markers.
3. Use `data_only=True` workbook (already loaded as `wb`) for all formula cells.
4. Handle `None` values defensively — many cells may be empty.
5. Keep the existing code style: no external imports beyond what's already there.
6. The `MRP` sheet uses `data_only=True` which means formula results are cached values from the last time Excel saved. This is acceptable.
7. After your changes, running `python update_html.py` should print the new data summaries and update the HTML with the new JS constants.
