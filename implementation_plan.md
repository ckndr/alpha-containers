# Add Inventory, Production Log, FG Stock & Material Requirement to Tubex HTML

## Background

The Tubex app currently has **2 tabs**: Dashboard and Material Calculator. Data is stored in `Tubex_v10_29.xlsx` across multiple sheets. The Python script `update_html.py` reads Excel → builds JSON/JS constants → injects into `Tubex.html` between marker comments.

We need to add **4 new tabs** so the data guy (and Sikander) can view everything from their phone without opening Excel:

1. **📦 Inventory (Raw Material Stock)** — current stock levels from the `Inventory` sheet
2. **🏭 Production Log** — daily production entries from `Production_Log` sheet
3. **📋 FG Stock** — Finished Goods stock from `FG Stock` sheet (latest date only)
4. **⚖️ Material Requirement** — MRP view showing Required Qty vs Current Stock, Surplus/Deficit

## User Review Required

> [!IMPORTANT]
> **FG Stock date logic**: The data entry guy repeats the same FG stock for each new date. The plan filters to **only the latest date** — matching what `update_production.py` already does. Older dates are discarded.

> [!IMPORTANT]
> **MRP data cannot use formulas**: The MRP sheet uses SUMPRODUCT formulas referencing `TableBOM`, `TableInventory`, etc. `openpyxl` with `data_only=True` reads **cached calculated values** — these are only available if the workbook was last saved with Excel open (Ctrl+Shift+F9 recalculates). If the user runs `Run_All_Updates.bat` and then opens/saves in Excel, the values will be fresh. If not, they may be stale. The plan reads with `data_only=True` and shows a "last recalculated" timestamp.

> [!WARNING]
> **Inventory sheet needs `data_only=True`**: The Inventory sheet has formula columns (col H = Store Balance = Opening + Received - Issued, col I = Work In Process). These formulas must be read as values, so we must open the workbook with `data_only=True` in the HTML updater — the same workbook load used for Dashboard/BOM already does this.

## Open Questions

> [!IMPORTANT]
> **Tab order**: Proposed tab order is: Dashboard → Production Log → FG Stock → Inventory → Material Req → Material Calculator. Is this acceptable?

> [!IMPORTANT]
> **Production Log date range**: Should the Production Log tab show the **full current month** (matching the dashboard MTD logic), or the **entire log** (all months)? Plan assumes **current month only** to keep the HTML size reasonable.

## Proposed Changes

The work is split into **3 sequential prompts** for a weaker AI model. Each prompt is self-contained with all context needed.

---

### Component 1: Python — Data Extraction (update_html.py)

#### [MODIFY] [update_html.py](file:///d:/Alpha/Scripts/update_html.py)

Add 4 new data readers after the existing BOM/Products logic, and inject them into the HTML via new marker blocks:

**1a. Read Inventory Sheet** (cols A–I from `Inventory` sheet of `Tubex_v*.xlsx`):
- Item ID (col A), Category (col B), Item Name (col C), UOM (col D)
- Opening (col E), Received (col F), Issued (col G)
- Store Balance (col H, formula → read with `data_only=True`), Work In Process (col I)
- Filter: skip rows where Item ID is not a valid integer
- Output: `INVENTORY_DATA` JS array

**1b. Read Production Log** (from `Production_Log` sheet, already loaded as `ws_pl`):
- Re-use the existing row iteration but capture ALL fields for the current month
- Date, Machine, Customer, Product, Dia, PID, Target, Good, Reject, Waste%
- Also all downtime columns
- Output: `PRODUCTION_LOG_DATA` JS array

**1c. Read FG Stock** (from `FG Stock` sheet of `Tubex_v*.xlsx`):
- Sr, PID, Customer, Product, Dia, FG Qty, Status, Dispatch Remarks
- Filter to only rows with data (row ≥ 4)
- Title row A1 contains the date string
- Output: `FG_STOCK_DATA` JS object with `{date, rows}`

**1d. Read MRP Data** (from `MRP` sheet, `data_only=True`):
- Rows 3–14: Required Orders (Dia, Customer, Product, PID, Job#, Required Qty, Produced, Remaining)
- Rows 17–101: Tubes Material Requirement (Item ID, Category, Item Name, UOM, Required Qty, Current Stock, Surplus/Deficit, Product Names, Status)
- Rows ~106–111: PET orders section
- Rows ~116–123: PET Material Requirement
- Rows ~127–158: INK Table
- Output: `MRP_DATA` JS object

**Injection markers**: New marker pairs in the HTML:
```
/* INVENTORY_START */ ... /* INVENTORY_END */
/* PRODLOG_START */ ... /* PRODLOG_END */
/* FGSTOCK_START */ ... /* FGSTOCK_END */
/* MRP_START */ ... /* MRP_END */
```

---

### Component 2: HTML/CSS — New Tab Panels & Styles (Tubex.html)

#### [MODIFY] [Tubex.html](file:///d:/Alpha/Tubex.html)

**2a. Add 4 new tab buttons** in the `.tabs` div (line ~321):
```html
<button class="tab" onclick="switchTab('prodlog',this)">🏭 Production</button>
<button class="tab" onclick="switchTab('fgstock',this)">📋 FG Stock</button>
<button class="tab" onclick="switchTab('inventory',this)">📦 Inventory</button>
<button class="tab" onclick="switchTab('mrp',this)">⚖️ MRP</button>
```

**2b. Add 4 new tab panel `<div>`s** in the `.main` div, between the dashboard and calc panels:

Each panel follows the existing design system (navy/accent theme, DM Sans/Mono fonts, card-based layout):

- **Production Log panel**: Responsive table with date/machine/product/good/reject columns. Date filter buttons for quick switching. Mobile-first: horizontal scroll on narrow screens.
- **FG Stock panel**: Card-grid or table showing each product's FG qty, status badge (OK=green, other=orange), dispatch remarks. Date badge at top.
- **Inventory panel**: Searchable/filterable table grouped by category. Shows Item ID, Name, UOM, Balance, with color coding for low/zero stock.
- **MRP panel**: Split view — top section shows required orders with remaining balance. Bottom section shows material requirement with Required/Stock/Surplus columns and status badges (OK/LOW/SHORTAGE).

**2c. CSS additions**:
- Status badge styles: `.status-ok`, `.status-low`, `.status-shortage`
- FG card styles matching the existing card aesthetic
- MRP surplus/deficit color coding (green for surplus, red for deficit)
- Responsive breakpoints matching existing 768px / 380px pattern

---

### Component 3: JavaScript — Rendering Functions (Tubex.html)

#### [MODIFY] [Tubex.html](file:///d:/Alpha/Tubex.html)

**3a. `renderProductionLog()`**: 
- Builds table rows from `PRODUCTION_LOG_DATA`
- Groups by date with date separators
- Machine badges colored by type (Press=blue, Print=navy, PET=green)
- Shows downtime columns collapsed on mobile

**3b. `renderFGStock()`**:
- Renders cards/table from `FG_STOCK_DATA`
- Status badges: OK=green pill, other=orange pill
- Shows date header from FG_STOCK_DATA.date

**3c. `renderInventory()`**:
- Renders searchable table from `INVENTORY_DATA`
- Search by item name or ID
- Category filter buttons
- Row coloring: zero stock = red background, low stock = orange
- Sort by category then item name

**3d. `renderMRP()`**:
- Material Requirement tab with two sections:
  - **Required Orders**: table from MRP rows 3–14 (and PET equivalent)
  - **Material Status**: table from MRP material rows, showing:
    - Item Name, Required Qty, Current Stock (Store + WIP)
    - **Surplus / Deficit** column: `Current Stock - Required Qty`
    - Status badge: `SHORTAGE` (red), `LOW` (orange), `OK` (green), `Not needed` (grey)
- Filter: "Show shortages only" button
- This is the key feature — lets Sikander instantly see what raw materials need ordering

**3e. Update `renderDashboard()` init**: Call new render functions on page load.

---

## Verification Plan

### Automated Tests
1. Run `python Scripts/update_html.py` — verify no errors, check that new markers are injected
2. Open `Tubex.html` in Chrome DevTools mobile emulator (iPhone SE) — verify all 6 tabs render
3. Check each tab has data populated (not empty arrays)
4. Verify MRP surplus/deficit math: `surplus = currentStock - requiredQty`

### Manual Verification
1. Open on actual phone (Android Chrome) — verify touch scrolling, tab switching, search
2. Compare Inventory numbers in HTML vs Excel to confirm data accuracy
3. Compare FG Stock against Production.xlsx latest date rows
4. Run full `Run_All_Updates.bat` pipeline → verify HTML is correct after full refresh
