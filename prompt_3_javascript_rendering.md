# PROMPT 3 OF 3: Add JavaScript Render Functions for New Tabs

## YOUR TASK

You must add JavaScript rendering functions to `d:\Alpha\Tubex.html` that populate the 4 new tab panels created in Prompt 2. The data constants (`INVENTORY_DATA`, `PRODUCTION_LOG_DATA`, `FG_STOCK_DATA`, `MRP_DATA`) were injected by Prompt 1's Python changes.

**Location**: Add all new JavaScript **before** the `// ─── INIT ───` comment (currently around line ~1241), inside the existing `<script>` tag.

**Also**: Update the init block to call the new render functions.

---

## DATA FORMAT REFERENCE

These JS constants will be available (injected by update_html.py):

```javascript
// From Prompt 1:
const PRODUCTION_LOG_DATA = {
  month: "June 2026",
  rows: [
    {date:"01-Jun", machine:"Print1", customer:"Samsol", product:"TUBES", dia:"25", pid:3726, total:5000, good:4800, reject:200},
    ...
  ]
};

const FG_STOCK_DATA = {
  title: "FG STOCK IN HAND — Last Updated: 03-Jun-2026",
  rows: [
    {sr:1, pid:3726, customer:"Samsol International", product:"TUBES", dia:"25", qty:15000, status:"OK", remarks:"Ready for dispatch"},
    ...
  ]
};

const INVENTORY_DATA = {
  title: "Raw Material Inventory (01-06-2026 to 03-06-2026)",
  items: [
    {id:2, cat:"SLUG", name:"30X4.3", uom:"kg", opening:1200, received:500, issued:800, balance:900, wip:150},
    ...
  ]
};

const MRP_DATA = {
  title: "May 2026 Tube Required Orders",
  orders: [
    {dia:16, customer:"Brookes Pharma", product:"ECZEMUS OINTMENT", pid:6561, jobOrder:"6849", required:25000, produced:0, remaining:25000, remarks:"CAPS Not Approved"},
    ...
  ],
  materials: [
    {id:22, cat:"SLUG", name:"19X4.2", uom:"kg", required:120.5, stock:500, surplus:379.5, products:"PYODINE GEL 20GM", status:"OK", section:"tube"},
    ...
  ],
  inks: [
    {id:1234, name:"CYAN 2000", uom:"kg", avgUse:12.5, daysLeft:45, status:"OK", stock:18.7},
    ...
  ]
};
```

---

## FUNCTION 1: renderProductionLog()

```javascript
// ─── PRODUCTION LOG ───
let prodlogDateFilter = 'all';

function getMachineBadge(machine) {
  const m = machine.toUpperCase();
  let cls = 'other';
  if (m.startsWith('PRESS')) cls = 'press';
  else if (m.startsWith('PRINT')) cls = 'print';
  else if (m.startsWith('PLINE')) cls = 'pline';
  else if (m.startsWith('PF') || m.startsWith('PET')) cls = 'pet';
  return `<span class="machine-badge ${cls}">${machine}</span>`;
}

function renderProductionLog() {
  const data = (typeof PRODUCTION_LOG_DATA !== 'undefined') ? PRODUCTION_LOG_DATA : null;
  if (!data || !data.rows || data.rows.length === 0) {
    document.getElementById('prodlog-body').innerHTML = '<tr><td colspan="6" class="no-data"><div class="big">🏭</div>No production log data available.</td></tr>';
    return;
  }
  
  document.getElementById('prodlog-month').textContent = data.month;
  document.getElementById('prodlog-count').textContent = data.rows.length + ' entries';
  
  // Build unique date chips
  const dates = [...new Set(data.rows.map(r => r.date))];
  const chipDiv = document.getElementById('prodlog-date-chips');
  chipDiv.innerHTML = '<button class="date-chip active" onclick="setProdLogDate(\'all\',this)">All Dates</button>' +
    dates.map(d => `<button class="date-chip" onclick="setProdLogDate('${d}',this)">${d}</button>`).join('');
  
  filterProdLog();
}

function setProdLogDate(d, btn) {
  prodlogDateFilter = d;
  document.querySelectorAll('#prodlog-date-chips .date-chip').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  filterProdLog();
}

function filterProdLog() {
  const data = PRODUCTION_LOG_DATA;
  const q = (document.getElementById('prodlogSearch').value || '').toLowerCase();
  const tbody = document.getElementById('prodlog-body');
  
  let html = '';
  let totalGood = 0, totalReject = 0, totalTotal = 0;
  let visibleCount = 0;
  
  data.rows.forEach(r => {
    // Date filter
    if (prodlogDateFilter !== 'all' && r.date !== prodlogDateFilter) return;
    // Search filter
    if (q && !r.product.toLowerCase().includes(q) && !r.customer.toLowerCase().includes(q) && !r.machine.toLowerCase().includes(q)) return;
    
    visibleCount++;
    totalGood += r.good;
    totalReject += r.reject;
    totalTotal += r.total;
    
    html += `<tr>
      <td style="white-space:nowrap;font-family:'DM Mono',monospace;font-size:11px">${r.date}</td>
      <td>${getMachineBadge(r.machine)}</td>
      <td style="font-weight:500">${r.product}<div style="font-size:10px;color:var(--muted)">${r.customer}</div></td>
      <td style="font-weight:600;color:var(--navy)">${r.good > 0 ? r.good.toLocaleString() : '—'}</td>
      <td style="color:${r.reject > 0 ? 'var(--red)' : 'var(--muted)'}">${r.reject > 0 ? r.reject.toLocaleString() : '—'}</td>
      <td>${r.total > 0 ? r.total.toLocaleString() : '—'}</td>
    </tr>`;
  });
  
  // Total row
  if (visibleCount > 0) {
    html += `<tr class="total-row">
      <td colspan="3">TOTAL (${visibleCount} entries)</td>
      <td>${totalGood.toLocaleString()}</td>
      <td>${totalReject.toLocaleString()}</td>
      <td>${totalTotal.toLocaleString()}</td>
    </tr>`;
  } else {
    html = '<tr><td colspan="6" class="no-data">No entries match your filters.</td></tr>';
  }
  
  tbody.innerHTML = html;
}
```

---

## FUNCTION 2: renderFGStock()

```javascript
// ─── FG STOCK ───
let fgFilter = 'all';

function getStatusPill(status) {
  const s = (status || '').toLowerCase();
  if (s === 'ok') return '<span class="status-pill ok">OK</span>';
  if (s === 'not ready') return '<span class="status-pill not-ready">Not Ready</span>';
  if (s === 'in progress') return '<span class="status-pill in-progress">In Progress</span>';
  if (s.includes('short')) return '<span class="status-pill shortage">Shortage</span>';
  if (s) return `<span class="status-pill not-ready">${status}</span>`;
  return '<span class="status-pill not-needed">—</span>';
}

function renderFGStock() {
  const data = (typeof FG_STOCK_DATA !== 'undefined') ? FG_STOCK_DATA : null;
  if (!data || !data.rows || data.rows.length === 0) {
    document.getElementById('fgstock-grid').innerHTML = '<div class="no-data"><div class="big">📋</div>No FG Stock data available.</div>';
    return;
  }
  
  document.getElementById('fgstock-title').textContent = data.title || 'FG Stock';
  document.getElementById('fgstock-count').textContent = data.rows.length + ' products · ' + data.rows.reduce((s,r) => s + r.qty, 0).toLocaleString() + ' total pcs';
  
  filterFGStock();
}

function setFGFilter(f, btn) {
  fgFilter = f;
  document.querySelectorAll('#panel-fgstock .filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  filterFGStock();
}

function filterFGStock() {
  const data = FG_STOCK_DATA;
  const q = (document.getElementById('fgSearch').value || '').toLowerCase();
  const grid = document.getElementById('fgstock-grid');
  
  let html = '';
  let visible = 0;
  
  data.rows.forEach(r => {
    const statusLower = (r.status || '').toLowerCase();
    const isOk = statusLower === 'ok';
    
    // Filter
    if (fgFilter === 'ok' && !isOk) return;
    if (fgFilter === 'warn' && isOk) return;
    if (q && !r.product.toLowerCase().includes(q) && !r.customer.toLowerCase().includes(q)) return;
    
    visible++;
    html += `<div class="fg-card ${isOk ? 'status-ok' : 'status-warn'}">
      <div class="fg-card-top">
        <div>
          <div class="fg-product">${r.product}</div>
          <div class="fg-customer">${r.customer}</div>
        </div>
        <div>
          <div class="fg-qty">${r.qty.toLocaleString()}</div>
          <div class="fg-qty-label">pieces</div>
        </div>
      </div>
      <div class="fg-meta">
        <span class="fg-dia">${r.dia || '—'}</span>
        ${r.pid ? `<span class="pid-tag">PID ${r.pid}</span>` : ''}
        ${getStatusPill(r.status)}
      </div>
      ${r.remarks ? `<div class="fg-remarks">📝 ${r.remarks}</div>` : ''}
    </div>`;
  });
  
  if (visible === 0) {
    html = '<div class="no-data" style="grid-column:1/-1"><div class="big">📋</div>No products match your filters.</div>';
  }
  
  grid.innerHTML = html;
}
```

---

## FUNCTION 3: renderInventory()

```javascript
// ─── INVENTORY ───
let invCatFilter = 'all';

function renderInventory() {
  const data = (typeof INVENTORY_DATA !== 'undefined') ? INVENTORY_DATA : null;
  if (!data || !data.items || data.items.length === 0) {
    document.getElementById('inv-body').innerHTML = '<tr><td colspan="9" class="no-data"><div class="big">📦</div>No inventory data available.</td></tr>';
    return;
  }
  
  document.getElementById('inv-title').textContent = data.title || 'Inventory';
  document.getElementById('inv-count').textContent = data.items.length + ' items';
  
  // Category stats
  const cats = {};
  data.items.forEach(item => {
    if (!cats[item.cat]) cats[item.cat] = {count:0, zeroCount:0};
    cats[item.cat].count++;
    if (item.balance <= 0 && item.wip <= 0) cats[item.cat].zeroCount++;
  });
  
  // Stats cards
  const statsDiv = document.getElementById('inv-stats');
  const totalItems = data.items.length;
  const zeroStock = data.items.filter(i => i.balance <= 0 && i.wip <= 0).length;
  statsDiv.innerHTML = `
    <div class="inv-stat"><div class="val">${totalItems}</div><div class="lbl">Total Items</div></div>
    <div class="inv-stat"><div class="val" style="color:var(--green)">${totalItems - zeroStock}</div><div class="lbl">In Stock</div></div>
    <div class="inv-stat"><div class="val" style="color:var(--red)">${zeroStock}</div><div class="lbl">Zero Stock</div></div>
    <div class="inv-stat"><div class="val">${Object.keys(cats).length}</div><div class="lbl">Categories</div></div>
  `;
  
  // Category filter buttons
  const catDiv = document.getElementById('inv-cat-filters');
  catDiv.innerHTML = '<button class="filter-btn active" onclick="setInvCat(\'all\',this)">All</button>' +
    Object.keys(cats).sort().map(c => 
      `<button class="filter-btn" onclick="setInvCat('${c}',this)">${c} <span style="font-size:10px;opacity:.6">(${cats[c].count})</span></button>`
    ).join('');
  
  filterInventory();
}

function setInvCat(cat, btn) {
  invCatFilter = cat;
  document.querySelectorAll('#inv-cat-filters .filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  filterInventory();
}

function fmtNum(v, uom) {
  if (v === 0) return '0';
  if (uom === 'pcs') return Math.round(v).toLocaleString();
  if (Math.abs(v) >= 100) return Math.round(v).toLocaleString();
  return v.toFixed(2);
}

function filterInventory() {
  const data = INVENTORY_DATA;
  const q = (document.getElementById('invSearch').value || '').toLowerCase();
  const tbody = document.getElementById('inv-body');
  
  let html = '';
  let visible = 0;
  
  data.items.forEach(item => {
    if (invCatFilter !== 'all' && item.cat !== invCatFilter) return;
    if (q && !item.name.toLowerCase().includes(q) && !String(item.id).includes(q)) return;
    
    visible++;
    const totalStock = item.balance + item.wip;
    const rowClass = totalStock <= 0 ? 'zero-stock' : (totalStock < 10 && item.uom === 'kg') ? 'low-stock' : '';
    
    html += `<tr class="${rowClass}">
      <td style="font-family:'DM Mono',monospace;font-size:11px;color:var(--muted)">${item.id}</td>
      <td><span style="font-size:10px;font-weight:600;letter-spacing:.5px;text-transform:uppercase;color:var(--muted)">${item.cat}</span></td>
      <td style="font-weight:500">${item.name}</td>
      <td style="font-size:11px;color:var(--muted)">${item.uom}</td>
      <td>${fmtNum(item.opening, item.uom)}</td>
      <td style="color:${item.received > 0 ? 'var(--green)' : 'var(--muted)'}">${fmtNum(item.received, item.uom)}</td>
      <td style="color:${item.issued > 0 ? 'var(--red)' : 'var(--muted)'}">${fmtNum(item.issued, item.uom)}</td>
      <td style="font-weight:600;color:${item.balance <= 0 ? 'var(--red)' : 'var(--navy)'}">${fmtNum(item.balance, item.uom)}</td>
      <td style="color:${item.wip > 0 ? 'var(--blue)' : 'var(--muted)'}">${fmtNum(item.wip, item.uom)}</td>
    </tr>`;
  });
  
  if (visible === 0) {
    html = '<tr><td colspan="9" class="no-data">No items match your search.</td></tr>';
  }
  
  tbody.innerHTML = html;
}
```

---

## FUNCTION 4: renderMRP()

```javascript
// ─── MRP (MATERIAL REQUIREMENT) ───
let mrpStatusFilter = 'all';

function getMRPStatusPill(status) {
  const s = (status || '').toUpperCase();
  if (s === 'SHORTAGE') return '<span class="status-pill shortage">SHORTAGE</span>';
  if (s === 'LOW') return '<span class="status-pill low">LOW</span>';
  if (s === 'OK') return '<span class="status-pill ok">OK</span>';
  if (s === 'NOT NEEDED' || s === '') return '<span class="status-pill not-needed">Not needed</span>';
  return `<span class="status-pill not-needed">${status}</span>`;
}

function renderMRP() {
  const data = (typeof MRP_DATA !== 'undefined') ? MRP_DATA : null;
  if (!data) {
    document.getElementById('mrp-orders-body').innerHTML = '<tr><td colspan="7" class="no-data"><div class="big">⚖️</div>No MRP data available.</td></tr>';
    return;
  }
  
  document.getElementById('mrp-title').textContent = data.title || 'Material Requirement Plan';
  
  // Render Orders
  const ordersBody = document.getElementById('mrp-orders-body');
  if (data.orders && data.orders.length > 0) {
    let totalReq = 0, totalProd = 0, totalRem = 0;
    let ohtml = '';
    data.orders.forEach(o => {
      totalReq += o.required;
      totalProd += o.produced;
      totalRem += o.remaining;
      const pct = o.required > 0 ? Math.min(100, o.produced / o.required * 100) : 0;
      ohtml += `<tr>
        <td style="font-family:'DM Mono',monospace">${o.dia}</td>
        <td style="font-weight:500">${o.product}</td>
        <td style="font-size:11px;color:var(--muted)">${o.customer}</td>
        <td>${o.required > 0 ? o.required.toLocaleString() : '—'}</td>
        <td style="font-weight:600;color:var(--navy)">${o.produced > 0 ? o.produced.toLocaleString() : '—'}</td>
        <td style="color:${o.remaining > 0 ? 'var(--orange)' : 'var(--green)'}; font-weight:600">${o.remaining > 0 ? o.remaining.toLocaleString() : '✓ Done'}</td>
        <td style="font-size:11px;color:var(--muted);font-style:italic">${o.remarks || ''}</td>
      </tr>`;
    });
    ohtml += `<tr class="total-row">
      <td colspan="3">TOTAL</td>
      <td>${totalReq.toLocaleString()}</td>
      <td>${totalProd.toLocaleString()}</td>
      <td>${totalRem.toLocaleString()}</td>
      <td></td>
    </tr>`;
    ordersBody.innerHTML = ohtml;
    document.getElementById('mrp-orders-badge').textContent = data.orders.length + ' orders · ' + totalReq.toLocaleString() + ' required';
  }
  
  // Material status badge
  if (data.materials) {
    const shortages = data.materials.filter(m => m.status === 'SHORTAGE').length;
    const lowCount = data.materials.filter(m => m.status === 'LOW').length;
    document.getElementById('mrp-mat-badge').textContent = 
      `${data.materials.length} items` + 
      (shortages > 0 ? ` · 🔴 ${shortages} shortage${shortages > 1 ? 's' : ''}` : '') +
      (lowCount > 0 ? ` · 🟡 ${lowCount} low` : '');
  }
  
  // Render Inks
  const inkBody = document.getElementById('mrp-ink-body');
  if (data.inks && data.inks.length > 0) {
    let ihtml = '';
    data.inks.forEach(ink => {
      ihtml += `<tr>
        <td style="font-family:'DM Mono',monospace;font-size:11px;color:var(--muted)">${ink.id}</td>
        <td style="font-weight:500">${ink.name}</td>
        <td>${ink.avgUse > 0 ? fmtNum(ink.avgUse, ink.uom) + ' ' + ink.uom : '—'}</td>
        <td style="font-weight:600;color:var(--navy)">${fmtNum(ink.stock, ink.uom)} ${ink.uom}</td>
        <td style="color:${ink.daysLeft < 15 ? 'var(--red)' : ink.daysLeft < 30 ? 'var(--orange)' : 'var(--green)'}; font-weight:600">${ink.daysLeft} days</td>
        <td>${getMRPStatusPill(ink.status)}</td>
      </tr>`;
    });
    inkBody.innerHTML = ihtml;
    document.getElementById('mrp-ink-badge').textContent = data.inks.length + ' inks tracked';
  }
  
  filterMRP();
}

function setMRPFilter(f, btn) {
  mrpStatusFilter = f;
  document.querySelectorAll('.mrp-filter-bar .filter-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  filterMRP();
}

function filterMRP() {
  const data = MRP_DATA;
  if (!data || !data.materials) return;
  
  const q = (document.getElementById('mrpSearch').value || '').toLowerCase();
  const tbody = document.getElementById('mrp-mat-body');
  
  let html = '';
  let visible = 0;
  
  data.materials.forEach(m => {
    if (mrpStatusFilter !== 'all' && m.status !== mrpStatusFilter) return;
    if (q && !m.name.toLowerCase().includes(q) && !String(m.id).includes(q) && !m.cat.toLowerCase().includes(q)) return;
    
    // Skip items with zero required and zero stock (truly irrelevant)
    if (m.required === 0 && m.stock === 0) return;
    
    visible++;
    const surplusClass = m.surplus < 0 ? 'deficit-val' : m.surplus > 0 ? 'surplus-val' : '';
    const surplusPrefix = m.surplus < 0 ? '' : '+';
    const surplusDisplay = m.required === 0 ? '—' : `${surplusPrefix}${fmtNum(m.surplus, m.uom)}`;
    
    html += `<tr>
      <td style="font-family:'DM Mono',monospace;font-size:11px;color:var(--muted)">${m.id}</td>
      <td><span style="font-size:10px;font-weight:600;letter-spacing:.5px;text-transform:uppercase;color:var(--muted)">${m.cat}</span></td>
      <td style="font-weight:500">${m.name}${m.products ? `<div style="font-size:10px;color:var(--muted);margin-top:2px">→ ${m.products}</div>` : ''}</td>
      <td>${m.required > 0 ? fmtNum(m.required, m.uom) + ' <span style="font-size:10px;color:var(--muted)">' + m.uom + '</span>' : '—'}</td>
      <td style="font-weight:600;color:var(--navy)">${fmtNum(m.stock, m.uom)} <span style="font-size:10px;color:var(--muted)">${m.uom}</span></td>
      <td class="${surplusClass}">${surplusDisplay} <span style="font-size:10px;color:var(--muted)">${m.required > 0 ? m.uom : ''}</span></td>
      <td>${getMRPStatusPill(m.status)}</td>
    </tr>`;
  });
  
  if (visible === 0) {
    html = '<tr><td colspan="7" class="no-data">No materials match your filters.</td></tr>';
  }
  
  tbody.innerHTML = html;
}
```

---

## UPDATE THE INIT BLOCK

Find the existing init block (around line ~1242):
```javascript
// ─── INIT ───
renderDashboard();
renderGrid();
updateGrandTotal();
```

Replace with:
```javascript
// ─── INIT ───
renderDashboard();
renderGrid();
updateGrandTotal();
renderProductionLog();
renderFGStock();
renderInventory();
renderMRP();
```

---

## IMPORTANT RULES

1. All functions must be **defensively coded** — check `typeof X !== 'undefined'` before accessing data constants, in case update_html.py hasn't run yet.
2. Use the existing `fmtNum()` helper you define in the Inventory section — reuse it in MRP too.
3. All number formatting must use `.toLocaleString()` for display (thousands separators).
4. Status pills use the CSS classes from Prompt 2: `.status-pill.ok`, `.status-pill.low`, `.status-pill.shortage`, `.status-pill.not-needed`, `.status-pill.not-ready`, `.status-pill.in-progress`.
5. Machine badges use: `.machine-badge.press`, `.machine-badge.print`, `.machine-badge.pline`, `.machine-badge.pet`, `.machine-badge.other`.
6. The existing `filterProducts()` function (for Material Calculator) must NOT be affected.
7. Do NOT duplicate the `clearAll()` or `renderGrid()` functions — they already exist for the Material Calculator tab.
8. The MRP Material Requirement section is the **most important feature** — the surplus/deficit column with color coding (green=surplus, red=deficit) lets the user instantly see what needs ordering.
9. The `fmtNum` function should be defined ONCE and reused by both Inventory and MRP renders.
