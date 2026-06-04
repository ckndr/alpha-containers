# PROMPT 2 OF 3: Add HTML Structure & CSS for New Tabs

## YOUR TASK

You must modify `d:\Alpha\Tubex.html` to add **4 new tab buttons** and **4 new tab panel divs** with all required CSS. After Prompt 1 has been run, the HTML will have new JS constants (`INVENTORY_DATA`, `PRODUCTION_LOG_DATA`, `FG_STOCK_DATA`, `MRP_DATA`) available. You are building the visual containers for them.

**DO NOT** modify the `<script>` section or any JavaScript. That is Prompt 3's job.
**DO NOT** modify the existing Dashboard or Material Calculator tabs.

---

## EXISTING DESIGN SYSTEM

The app uses this design language — you MUST match it exactly:

### CSS Variables (already defined in `:root`)
```css
--navy:#0d1f3c; --navy2:#1a2f52; --blue:#2355a0; --blue-light:#3a6bc7;
--accent:#e8a020; --accent2:#f5c842;
--red:#c0392b; --red-bg:#fff0ee;
--green:#1a7a4a; --green-bg:#edf7f1;
--orange:#d97706; --orange-bg:#fff7ed;
--bg:#f4f6fa; --card:#ffffff; --border:#dde3ef;
--text:#1a2035; --muted:#6b7b9a; --light:#eef1f8;
```

### Fonts
- Headings: `"DM Serif Display", serif`
- Body: `"DM Sans", sans-serif`
- Mono/numbers: `"DM Mono", monospace`

### Existing Patterns to Reuse
- Section headers: `<div class="sec-hdr"><h2>Title</h2><span class="sec-badge">info</span></div>`
- Tables: `<div class="tbl-wrap"><table class="orders">...</table></div>` — with existing `.orders` styling
- Cards: `.kpi-card` pattern with border-radius:12px, border:2px solid var(--border)
- Month badge: `<div class="month-badge">📅 <span>June 2026</span></div>`
- Mobile breakpoint: `@media(max-width:768px)` and `@media(max-width:380px)`

---

## STEP 1: Add New Tab Buttons

**Location**: Find the `.tabs` div (around line 321-324). Currently has 2 buttons:
```html
<div class="tabs">
    <button class="tab active" onclick="switchTab('dashboard',this)">📊 Dashboard</button>
    <button class="tab" onclick="switchTab('calc',this)">🧮 Material Calculator</button>
</div>
```

**Replace with 6 buttons** (add 4 new ones between Dashboard and Material Calculator):
```html
<div class="tabs">
    <button class="tab active" onclick="switchTab('dashboard',this)">📊 Dashboard</button>
    <button class="tab" onclick="switchTab('prodlog',this)">🏭 Production</button>
    <button class="tab" onclick="switchTab('fgstock',this)">📋 FG Stock</button>
    <button class="tab" onclick="switchTab('inventory',this)">📦 Inventory</button>
    <button class="tab" onclick="switchTab('mrp',this)">⚖️ MRP</button>
    <button class="tab" onclick="switchTab('calc',this)">🧮 Material Calc</button>
</div>
```

---

## STEP 2: Add New CSS Styles

**Location**: Add these styles BEFORE the closing `</style>` tag (before line 305). Add them after the existing `@media print` block (~line 263).

```css
/* ══ PRODUCTION LOG TAB ══ */
.prodlog-filters{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center}
.prodlog-filters .search-wrap{flex:1;min-width:200px}
.date-chip{padding:6px 12px;border:2px solid var(--border);border-radius:8px;background:var(--card);font-family:"DM Sans",sans-serif;font-size:12px;font-weight:500;color:var(--muted);cursor:pointer;transition:.2s;white-space:nowrap}
.date-chip:hover,.date-chip.active{border-color:var(--blue);color:var(--blue);background:#eef3ff}
.date-chip.active{background:var(--blue);color:#fff}
.machine-badge{display:inline-block;padding:2px 8px;border-radius:5px;font-size:10px;font-weight:600;letter-spacing:.3px}
.machine-badge.press{background:#e8f0fe;color:#2355a0}
.machine-badge.print{background:#1a2f52;color:#fff}
.machine-badge.pet{background:var(--green-bg);color:var(--green)}
.machine-badge.pline{background:#f3e8ff;color:#7c3aed}
.machine-badge.other{background:var(--light);color:var(--muted)}

/* ══ FG STOCK TAB ══ */
.fg-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:14px}
.fg-card{background:var(--card);border-radius:12px;border:2px solid var(--border);padding:16px;transition:.2s;box-shadow:0 2px 6px rgba(0,0,0,.04)}
.fg-card:hover{box-shadow:0 4px 16px rgba(13,31,60,.08);transform:translateY(-1px)}
.fg-card.status-ok{border-left:4px solid var(--green)}
.fg-card.status-warn{border-left:4px solid var(--orange)}
.fg-card-top{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px}
.fg-product{font-size:14px;font-weight:600;color:var(--text);line-height:1.3}
.fg-customer{font-size:11px;color:var(--muted);margin-top:2px}
.fg-qty{font-family:"DM Mono",monospace;font-size:22px;font-weight:600;color:var(--navy);text-align:right}
.fg-qty-label{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.8px;text-align:right}
.fg-meta{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:8px}
.fg-dia{font-family:"DM Mono",monospace;font-size:11px;color:var(--muted);background:var(--light);padding:2px 8px;border-radius:4px}
.fg-remarks{font-size:11px;color:var(--muted);font-style:italic;margin-top:6px}

/* ══ STATUS BADGES (shared) ══ */
.status-pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:10px;font-weight:600;letter-spacing:.5px;text-transform:uppercase}
.status-pill.ok{background:var(--green-bg);color:var(--green)}
.status-pill.low{background:var(--orange-bg);color:var(--orange)}
.status-pill.shortage{background:var(--red-bg);color:var(--red)}
.status-pill.not-needed{background:var(--light);color:var(--muted)}
.status-pill.not-ready{background:var(--orange-bg);color:var(--orange)}
.status-pill.in-progress{background:#e8f0fe;color:#2355a0}

/* ══ INVENTORY TAB ══ */
.inv-toolbar{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;align-items:center}
.inv-stat-row{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;margin-bottom:18px}
.inv-stat{background:var(--card);border-radius:10px;border:2px solid var(--border);padding:12px;text-align:center}
.inv-stat .val{font-family:"DM Mono",monospace;font-size:20px;font-weight:600;color:var(--navy)}
.inv-stat .lbl{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.8px;margin-top:2px}
table.inv-table{width:100%;border-collapse:collapse;background:var(--card)}
table.inv-table th{text-align:left;font-size:10px;font-weight:600;letter-spacing:.8px;text-transform:uppercase;color:var(--muted);padding:10px 10px;background:var(--light);white-space:nowrap;position:sticky;top:0;z-index:1}
table.inv-table th:nth-child(n+5){text-align:right}
table.inv-table td{padding:9px 10px;font-size:12px;border-bottom:1px solid var(--border);vertical-align:middle}
table.inv-table td:nth-child(n+5){text-align:right;font-family:"DM Mono",monospace}
table.inv-table tr:hover td{background:#f7f9ff}
table.inv-table tr.zero-stock td{background:#fff5f5}
table.inv-table tr.low-stock td{background:#fffbeb}
.cat-filter-row{display:flex;gap:6px;flex-wrap:wrap;align-items:center}

/* ══ MRP TAB ══ */
.mrp-section{margin-bottom:24px}
.mrp-section-title{font-family:"DM Serif Display",serif;font-size:16px;font-weight:400;color:var(--navy);margin-bottom:12px}
table.mrp-table{width:100%;border-collapse:collapse;background:var(--card)}
table.mrp-table th{text-align:left;font-size:10px;font-weight:600;letter-spacing:.8px;text-transform:uppercase;color:var(--muted);padding:10px 10px;background:var(--light);white-space:nowrap}
table.mrp-table th:nth-child(n+4){text-align:right}
table.mrp-table td{padding:9px 10px;font-size:12px;border-bottom:1px solid var(--border);vertical-align:middle}
table.mrp-table td:nth-child(n+4){text-align:right;font-family:"DM Mono",monospace}
table.mrp-table tr:hover td{background:#f7f9ff}
.surplus-val{color:var(--green);font-weight:600}
.deficit-val{color:var(--red);font-weight:600}
.mrp-filter-bar{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center}

/* ══ MOBILE OVERRIDES FOR NEW TABS ══ */
@media(max-width:768px){
  .fg-grid{grid-template-columns:1fr}
  .fg-qty{font-size:18px}
  .inv-toolbar{flex-direction:column}
  .inv-stat-row{grid-template-columns:repeat(2,1fr)}
  .prodlog-filters{flex-direction:column}
  table.inv-table th,table.inv-table td{padding:7px 6px;font-size:10.5px}
  table.mrp-table th,table.mrp-table td{padding:7px 6px;font-size:10.5px}
  .mrp-section-title{font-size:14px}
}
```

---

## STEP 3: Add Tab Panel HTML

**Location**: Add these 4 new panel `<div>`s inside `<div class="main">`, AFTER the Dashboard panel `</div>` (after line ~413) and BEFORE the Material Calculator panel (before line ~418).

### Panel 1: Production Log
```html
  <!-- ═══════════════════════════════════════════ -->
  <!-- PRODUCTION LOG TAB                          -->
  <!-- ═══════════════════════════════════════════ -->
  <div class="tab-panel" id="panel-prodlog">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:12px;margin-top:4px">
      <div class="month-badge">🏭 <span id="prodlog-month">—</span></div>
      <span class="sec-badge" id="prodlog-count">—</span>
    </div>
    
    <div class="prodlog-filters">
      <div class="search-wrap">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#666" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        <input class="search-input" type="text" id="prodlogSearch" placeholder="Search product, customer, machine…" oninput="filterProdLog()">
      </div>
      <div id="prodlog-date-chips" style="display:flex;gap:6px;flex-wrap:wrap"></div>
    </div>
    
    <div class="tbl-wrap">
      <table class="orders" id="prodlog-table">
        <thead><tr>
          <th>Date</th><th>Machine</th><th>Product</th>
          <th style="text-align:right">Good</th><th style="text-align:right">Reject</th>
          <th style="text-align:right">Total</th>
        </tr></thead>
        <tbody id="prodlog-body"></tbody>
      </table>
    </div>
  </div>
```

### Panel 2: FG Stock
```html
  <!-- ═══════════════════════════════════════════ -->
  <!-- FG STOCK TAB                                -->
  <!-- ═══════════════════════════════════════════ -->
  <div class="tab-panel" id="panel-fgstock">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:12px;margin-top:4px">
      <div class="month-badge">📋 <span id="fgstock-title">FG Stock</span></div>
      <span class="sec-badge" id="fgstock-count">—</span>
    </div>
    
    <div style="display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;align-items:center">
      <div class="search-wrap" style="flex:1;min-width:200px">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#666" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        <input class="search-input" type="text" id="fgSearch" placeholder="Search product or customer…" oninput="filterFGStock()">
      </div>
      <button class="filter-btn active" id="fg-all" onclick="setFGFilter('all',this)">All</button>
      <button class="filter-btn" id="fg-ok" onclick="setFGFilter('ok',this)">✅ OK</button>
      <button class="filter-btn" id="fg-warn" onclick="setFGFilter('warn',this)">⚠️ Not Ready</button>
    </div>
    
    <div class="fg-grid" id="fgstock-grid"></div>
  </div>
```

### Panel 3: Inventory
```html
  <!-- ═══════════════════════════════════════════ -->
  <!-- INVENTORY TAB                               -->
  <!-- ═══════════════════════════════════════════ -->
  <div class="tab-panel" id="panel-inventory">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:12px;margin-top:4px">
      <div class="month-badge">📦 <span id="inv-title">Inventory</span></div>
      <span class="sec-badge" id="inv-count">—</span>
    </div>
    
    <div class="inv-stat-row" id="inv-stats"></div>
    
    <div class="inv-toolbar">
      <div class="search-wrap" style="flex:1;min-width:200px">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#666" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
        <input class="search-input" type="text" id="invSearch" placeholder="Search item name or ID…" oninput="filterInventory()">
      </div>
      <div class="cat-filter-row" id="inv-cat-filters"></div>
    </div>
    
    <div class="tbl-wrap">
      <table class="inv-table" id="inv-table">
        <thead><tr>
          <th>ID</th><th>Category</th><th>Item Name</th><th>UOM</th>
          <th style="text-align:right">Opening</th><th style="text-align:right">Received</th>
          <th style="text-align:right">Issued</th><th style="text-align:right">Balance</th>
          <th style="text-align:right">WIP</th>
        </tr></thead>
        <tbody id="inv-body"></tbody>
      </table>
    </div>
  </div>
```

### Panel 4: Material Requirement (MRP)
```html
  <!-- ═══════════════════════════════════════════ -->
  <!-- MATERIAL REQUIREMENT TAB                    -->
  <!-- ═══════════════════════════════════════════ -->
  <div class="tab-panel" id="panel-mrp">
    <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:12px;margin-top:4px">
      <div class="month-badge">⚖️ <span id="mrp-title">Material Requirement</span></div>
    </div>

    <!-- Required Orders Section -->
    <div class="mrp-section">
      <div class="sec-hdr"><h2>Required Orders (Tubes)</h2><span class="sec-badge" id="mrp-orders-badge">—</span></div>
      <div class="tbl-wrap">
        <table class="orders" id="mrp-orders-table">
          <thead><tr>
            <th>Dia</th><th>Product</th><th>Customer</th>
            <th style="text-align:right">Required</th><th style="text-align:right">Produced</th>
            <th style="text-align:right">Remaining</th><th>Remarks</th>
          </tr></thead>
          <tbody id="mrp-orders-body"></tbody>
        </table>
      </div>
    </div>

    <!-- Material Status Section -->
    <div class="mrp-section">
      <div class="sec-hdr"><h2>Raw Material Status</h2><span class="sec-badge" id="mrp-mat-badge">—</span></div>
      
      <div class="mrp-filter-bar">
        <div class="search-wrap" style="flex:1;min-width:200px">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#666" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
          <input class="search-input" type="text" id="mrpSearch" placeholder="Search material…" oninput="filterMRP()">
        </div>
        <button class="filter-btn active" id="mrp-f-all" onclick="setMRPFilter('all',this)">All</button>
        <button class="filter-btn" id="mrp-f-shortage" onclick="setMRPFilter('SHORTAGE',this)">🔴 Shortage</button>
        <button class="filter-btn" id="mrp-f-low" onclick="setMRPFilter('LOW',this)">🟡 Low</button>
        <button class="filter-btn" id="mrp-f-ok" onclick="setMRPFilter('OK',this)">🟢 OK</button>
      </div>
      
      <div class="tbl-wrap">
        <table class="mrp-table" id="mrp-mat-table">
          <thead><tr>
            <th>ID</th><th>Category</th><th>Item Name</th>
            <th style="text-align:right">Required</th><th style="text-align:right">Stock</th>
            <th style="text-align:right">Surplus / Deficit</th><th>Status</th>
          </tr></thead>
          <tbody id="mrp-mat-body"></tbody>
        </table>
      </div>
    </div>

    <!-- INK Status Section -->
    <div class="mrp-section">
      <div class="sec-hdr"><h2>Ink Stock Status</h2><span class="sec-badge" id="mrp-ink-badge">—</span></div>
      <div class="tbl-wrap">
        <table class="mrp-table" id="mrp-ink-table">
          <thead><tr>
            <th>ID</th><th>Ink Name</th>
            <th style="text-align:right">Avg Monthly Use</th><th style="text-align:right">Current Stock</th>
            <th style="text-align:right">Days Left</th><th>Status</th>
          </tr></thead>
          <tbody id="mrp-ink-body"></tbody>
        </table>
      </div>
    </div>
  </div>
```

---

## IMPORTANT RULES

1. Preserve ALL existing HTML exactly as-is — only ADD new content.
2. The tab panel IDs must match the `switchTab()` parameter: `panel-prodlog`, `panel-fgstock`, `panel-inventory`, `panel-mrp`.
3. Keep `class="tab-panel"` (no `active` class) on all new panels — Dashboard stays as the default active.
4. CSS must use the existing CSS variables — do NOT introduce new colors.
5. All new panels must be inserted BETWEEN the Dashboard panel closing `</div>` and the Material Calculator panel opening comment.
6. The `.tabs` div will have 6 buttons — on mobile, the scrollable tab bar (existing CSS handles `overflow-x:auto`) will let users swipe.
