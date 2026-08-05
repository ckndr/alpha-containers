"""
generate_customer_report.py
===========================
Reads all production data from:
  - Production_Archive.xlsx  (each month tab)
  - Tubex_Aug26.xlsx         (current month Production_Log)

Outputs a self-contained customer_report.html with all data embedded.
Open the HTML in any browser -- no server, no internet needed (Chart.js is inlined).

Usage:
    python Scripts/generate_customer_report.py

Re-run each time you want to refresh the report with new data.
"""

import os, sys, json, datetime, textwrap
import openpyxl

# ============================================================
#  CONFIGURATION
# ============================================================

SOURCES = [
    # (file_path,                                    sheet_name,       month_label,    data_start_row)
    (r"d:\Alpha\Tubex Records\Production_Archive.xlsx", "July 2026",  "July 2026",    3),
    (r"d:\Alpha\Tubex_Aug26.xlsx",                      "Production_Log", "August 2026", 3),
    # Add new months here -- both archive tabs and the current active file
]

OUTPUT_HTML = r"d:\Alpha\customer_report.html"

# ============================================================
#  DATA EXTRACTION
# ============================================================

def extract_records():
    all_records = []
    seen_files = {}

    for path, sheet_name, month_label, start_row in SOURCES:
        if not os.path.exists(path):
            print(f"  [SKIP] File not found: {path}")
            continue

        cache_key = f"{path}::{sheet_name}"
        if cache_key in seen_files:
            continue
        seen_files[cache_key] = True

        print(f"  Reading: {os.path.basename(path)} -> '{sheet_name}' ...")
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            if sheet_name not in wb.sheetnames:
                print(f"    [WARN] Sheet '{sheet_name}' not found")
                wb.close()
                continue

            ws = wb[sheet_name]
            rows_loaded = 0

            for row in ws.iter_rows(min_row=start_row, values_only=True):
                if not row[0]:
                    continue
                cols = list(row) + [None] * 20
                date_val, machine, customer, product, dia, pid, target, good, reject, waste = cols[:10]

                if not customer or good is None:
                    continue

                customer = str(customer).strip()
                machine  = str(machine).strip() if machine else ""
                product  = str(product).strip() if product else ""
                dia_str  = str(dia).strip() if dia else ""

                # Determine type: PET if dia contains 'ml', else TUBE
                prod_type = "PET" if "ml" in dia_str.lower() else "TUBE"

                # Normalise date
                if isinstance(date_val, datetime.datetime):
                    date_str = date_val.strftime("%Y-%m-%d")
                elif isinstance(date_val, str):
                    date_str = date_val[:10]
                else:
                    date_str = ""

                rec = {
                    "date":     date_str,
                    "month":    month_label,
                    "customer": customer,
                    "product":  product,
                    "machine":  machine,
                    "dia":      dia_str,
                    "type":     prod_type,
                    "target":   int(target or 0),
                    "good":     int(good or 0),
                    "reject":   int(reject or 0),
                }
                all_records.append(rec)
                rows_loaded += 1

            wb.close()
            print(f"    Loaded {rows_loaded} records")
        except Exception as e:
            print(f"    [ERROR] {e}")

    return all_records


# ============================================================
#  HTML GENERATION
# ============================================================

HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tubex — Customer Production Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --navy:   #0f1e35;
    --mid:    #1a3355;
    --card:   #1e2d45;
    --accent: #d4af37;
    --green:  #2ecc71;
    --red:    #e74c3c;
    --text:   #e8edf5;
    --muted:  #8899aa;
    --border: rgba(212,175,55,.18);
    --radius: 12px;
  }
  body {
    background: var(--navy);
    color: var(--text);
    font-family: 'Segoe UI', system-ui, sans-serif;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
  }

  /* ── Header ─────────────────────────────────────────────── */
  header {
    background: linear-gradient(135deg, #0a1628 0%, #1a3355 100%);
    border-bottom: 1px solid var(--border);
    padding: 18px 32px;
    display: flex; align-items: center; gap: 16px;
  }
  header .logo { font-size: 26px; font-weight: 800; letter-spacing: 3px; color: var(--accent); }
  header .subtitle { color: var(--muted); font-size: 13px; }
  header .meta { margin-left: auto; font-size: 12px; color: var(--muted); text-align: right; }

  /* ── Main layout ─────────────────────────────────────────── */
  .layout { display: flex; flex: 1; overflow: hidden; }

  /* ── Sidebar ─────────────────────────────────────────────── */
  aside {
    width: 230px; min-width: 230px;
    background: #111c2e;
    border-right: 1px solid var(--border);
    display: flex; flex-direction: column;
    overflow-y: auto;
  }
  .sidebar-title {
    padding: 16px 16px 8px;
    font-size: 10px; font-weight: 700;
    letter-spacing: 2px; text-transform: uppercase;
    color: var(--muted);
    border-bottom: 1px solid var(--border);
  }
  .customer-btn {
    display: block; width: 100%;
    padding: 13px 16px;
    background: transparent; border: none;
    border-bottom: 1px solid rgba(255,255,255,.04);
    color: var(--text); text-align: left;
    font-size: 13px; cursor: pointer;
    transition: all .2s; line-height: 1.4;
  }
  .customer-btn:hover { background: rgba(212,175,55,.08); }
  .customer-btn.active {
    background: linear-gradient(90deg, rgba(212,175,55,.2) 0%, transparent 100%);
    border-left: 3px solid var(--accent);
    color: var(--accent); font-weight: 600;
  }
  .customer-btn .cust-total {
    display: block; font-size: 11px;
    color: var(--muted); margin-top: 2px;
  }
  .customer-btn.active .cust-total { color: rgba(212,175,55,.6); }
  .search-wrap { padding: 12px; border-bottom: 1px solid var(--border); }
  .search-wrap input {
    width: 100%; padding: 8px 12px;
    background: rgba(255,255,255,.06); border: 1px solid var(--border);
    border-radius: 8px; color: var(--text); font-size: 13px; outline: none;
  }
  .search-wrap input::placeholder { color: var(--muted); }
  .search-wrap input:focus { border-color: var(--accent); }

  /* ── Main content ────────────────────────────────────────── */
  main { flex: 1; overflow-y: auto; padding: 24px 28px; }

  /* ── KPI Cards ───────────────────────────────────────────── */
  .kpi-row { display: grid; grid-template-columns: repeat(4,1fr); gap: 14px; margin-bottom: 24px; }
  .kpi-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius); padding: 18px 20px;
    position: relative; overflow: hidden;
  }
  .kpi-card::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px;
  }
  .kpi-card.tube::before  { background: #3a7bd5; }
  .kpi-card.pet::before   { background: #27ae60; }
  .kpi-card.total::before { background: var(--accent); }
  .kpi-card.reject::before{ background: #e74c3c; }
  .kpi-label { font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px; color: var(--muted); }
  .kpi-value { font-size: 28px; font-weight: 800; margin: 6px 0 4px; letter-spacing: -1px; }
  .kpi-value.tube   { color: #5a9bf5; }
  .kpi-value.pet    { color: #2ecc71; }
  .kpi-value.total  { color: var(--accent); }
  .kpi-value.reject { color: #e74c3c; }
  .kpi-sub { font-size: 11px; color: var(--muted); }

  /* ── Filter bar ──────────────────────────────────────────── */
  .filter-bar {
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 22px; flex-wrap: wrap;
  }
  .filter-label { font-size: 12px; color: var(--muted); margin-right: 4px; }
  .filter-btn {
    padding: 6px 14px;
    background: rgba(255,255,255,.06); border: 1px solid var(--border);
    border-radius: 20px; color: var(--text); font-size: 12px;
    cursor: pointer; transition: all .2s;
  }
  .filter-btn:hover { border-color: var(--accent); color: var(--accent); }
  .filter-btn.active {
    background: var(--accent); border-color: var(--accent); color: #000; font-weight: 600;
  }
  .type-toggle {
    margin-left: auto; display: flex; gap: 8px;
  }

  /* ── Chart + Monthly table row ───────────────────────────── */
  .charts-row { display: grid; grid-template-columns: 1fr 400px; gap: 18px; margin-bottom: 24px; }
  .card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: var(--radius); padding: 20px;
  }
  .card-title { font-size: 12px; text-transform: uppercase; letter-spacing: 1.5px; color: var(--muted); margin-bottom: 16px; }
  .chart-wrap { position: relative; height: 240px; }

  /* ── Monthly breakdown table ─────────────────────────────── */
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th {
    padding: 10px 12px; text-align: left;
    font-size: 10px; text-transform: uppercase; letter-spacing: 1px;
    color: var(--muted); border-bottom: 1px solid var(--border);
  }
  td { padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,.04); }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: rgba(255,255,255,.03); }
  td.num { text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }
  td.tube-num { color: #5a9bf5; }
  td.pet-num  { color: #2ecc71; }
  td.total-num{ color: var(--accent); }
  .total-row td { font-weight: 700; border-top: 1px solid var(--border); color: var(--text); }

  /* ── Product breakdown ───────────────────────────────────── */
  .full-width { margin-bottom: 24px; }
  .badge {
    display: inline-block; padding: 2px 8px;
    border-radius: 4px; font-size: 10px; font-weight: 700;
  }
  .badge.TUBE { background: rgba(58,123,213,.2); color: #5a9bf5; }
  .badge.PET  { background: rgba(39,174,96,.2);  color: #2ecc71; }
  .bar-wrap { display: flex; align-items: center; gap: 10px; }
  .bar-bg { flex: 1; height: 6px; background: rgba(255,255,255,.08); border-radius: 3px; overflow: hidden; }
  .bar-fill { height: 100%; border-radius: 3px; }
  .bar-fill.TUBE { background: #3a7bd5; }
  .bar-fill.PET  { background: #27ae60; }
  .bar-pct { font-size: 11px; color: var(--muted); width: 38px; text-align: right; }

  /* ── Empty state ─────────────────────────────────────────── */
  .empty {
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    height: 50vh; color: var(--muted); text-align: center; gap: 12px;
  }
  .empty .icon { font-size: 48px; }
  .empty h3 { font-size: 18px; color: var(--text); }

  /* ── Section label ───────────────────────────────────────── */
  .section-hdr {
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 14px;
  }
  .section-hdr h2 { font-size: 15px; font-weight: 700; }
  .section-hdr .pill {
    padding: 2px 10px; border-radius: 20px; font-size: 11px; font-weight: 600;
    background: rgba(212,175,55,.15); color: var(--accent);
  }

  .updated { font-size: 11px; color: var(--muted); margin-top: 4px; }
</style>
</head>
<body>

<header>
  <div>
    <div class="logo">TUBEX</div>
    <div class="subtitle">Customer Production Report</div>
  </div>
  <div class="meta">
    <div>Generated: <strong>__GENERATED__</strong></div>
    <div class="updated">Covers: __MONTHS_COVERED__</div>
  </div>
</header>

<div class="layout">
  <!-- Sidebar -->
  <aside>
    <div class="sidebar-title">Customers</div>
    <div class="search-wrap">
      <input type="text" id="searchInput" placeholder="Search customer..." oninput="filterSidebar()">
    </div>
    <div id="customerList"></div>
  </aside>

  <!-- Main -->
  <main id="mainContent">
    <div class="empty">
      <div class="icon">👈</div>
      <h3>Select a customer</h3>
      <p>Click any customer on the left to view their production history.</p>
    </div>
  </main>
</div>

<script>
// ============================================================
//  EMBEDDED DATA
// ============================================================
const RAW_DATA = __DATA__;

// ============================================================
//  STATE
// ============================================================
let selectedCustomer = null;
let activeMonthFilter = 'all';
let activeTypeFilter  = 'all';
let chartInstance     = null;

// ============================================================
//  INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  buildSidebar();
});

// ============================================================
//  SIDEBAR
// ============================================================
function getCustomers() {
  const totals = {};
  RAW_DATA.forEach(r => {
    if (!totals[r.customer]) totals[r.customer] = 0;
    totals[r.customer] += r.good;
  });
  return Object.entries(totals).sort((a,b) => b[1]-a[1]);
}

function buildSidebar() {
  const list = document.getElementById('customerList');
  const query = document.getElementById('searchInput').value.toLowerCase();
  list.innerHTML = '';
  getCustomers().forEach(([name, total]) => {
    if (query && !name.toLowerCase().includes(query)) return;
    const btn = document.createElement('button');
    btn.className = 'customer-btn' + (name === selectedCustomer ? ' active' : '');
    btn.dataset.name = name;
    const short = name.length > 28 ? name.substring(0, 26) + '...' : name;
    btn.innerHTML = `${short}<span class="cust-total">${fmtNum(total)} units total</span>`;
    btn.onclick = () => selectCustomer(name);
    list.appendChild(btn);
  });
}

function filterSidebar() { buildSidebar(); }

// ============================================================
//  CUSTOMER SELECT
// ============================================================
function selectCustomer(name) {
  selectedCustomer = name;
  activeMonthFilter = 'all';
  activeTypeFilter  = 'all';
  buildSidebar();
  renderMain();
}

// ============================================================
//  MAIN RENDER
// ============================================================
function getMonths() {
  const s = new Set();
  RAW_DATA.forEach(r => { if (r.customer === selectedCustomer) s.add(r.month); });
  return [...s].sort();
}

function filtered() {
  return RAW_DATA.filter(r => {
    if (r.customer !== selectedCustomer) return false;
    if (activeTypeFilter !== 'all' && r.type !== activeTypeFilter) return false;
    if (activeMonthFilter !== 'all' && r.month !== activeMonthFilter) return false;
    return true;
  });
}

function last3Months() {
  const months = getMonths();
  return months.slice(-3);
}

function renderMain() {
  if (!selectedCustomer) return;

  const months   = getMonths();
  const allRecs  = filtered();

  const tubeTotal   = allRecs.filter(r=>r.type==='TUBE').reduce((s,r)=>s+r.good,0);
  const petTotal    = allRecs.filter(r=>r.type==='PET').reduce((s,r)=>s+r.good,0);
  const rejectTotal = allRecs.reduce((s,r)=>s+r.reject,0);
  const grandTotal  = tubeTotal + petTotal;

  // Month filter labels
  const l3 = last3Months();
  const filterBtns = [
    {key:'all', label:'All Months'},
    {key:'l3',  label:'Last 3 Months'},
    ...months.map(m => ({key:m, label:m}))
  ].map(f => `<button class="filter-btn ${activeMonthFilter===f.key?'active':''}"
    onclick="setMonthFilter('${f.key}')">${f.label}</button>`).join('');

  const typeBtns = ['all','TUBE','PET'].map(t =>
    `<button class="filter-btn ${activeTypeFilter===t?'active':''}"
      onclick="setTypeFilter('${t}')">${t==='all'?'All Types':t}</button>`
  ).join('');

  // Monthly summary data
  const monthSummary = {};
  RAW_DATA.filter(r => r.customer === selectedCustomer).forEach(r => {
    if (!monthSummary[r.month]) monthSummary[r.month] = {TUBE:0, PET:0, reject:0};
    monthSummary[r.month][r.type] += r.good;
    monthSummary[r.month].reject  += r.reject;
  });

  const monthRows = months.map(m => {
    const d = monthSummary[m] || {TUBE:0,PET:0,reject:0};
    const tot = d.TUBE + d.PET;
    const rejPct = tot > 0 ? ((d.reject/tot)*100).toFixed(1) : '0.0';
    return `<tr>
      <td>${m}</td>
      <td class="num tube-num">${fmtNum(d.TUBE)}</td>
      <td class="num pet-num">${fmtNum(d.PET)}</td>
      <td class="num total-num">${fmtNum(tot)}</td>
      <td class="num" style="color:#e74c3c">${rejPct}%</td>
    </tr>`;
  }).join('');

  // Totals row for table
  const allMonthData = Object.values(monthSummary);
  const sumTUBE   = allMonthData.reduce((s,d)=>s+d.TUBE,0);
  const sumPET    = allMonthData.reduce((s,d)=>s+d.PET,0);
  const sumReject = allMonthData.reduce((s,d)=>s+d.reject,0);
  const sumTotal  = sumTUBE + sumPET;
  const sumRejPct = sumTotal > 0 ? ((sumReject/sumTotal)*100).toFixed(1) : '0.0';

  const totalsRow = `<tr class="total-row">
    <td>TOTAL</td>
    <td class="num tube-num">${fmtNum(sumTUBE)}</td>
    <td class="num pet-num">${fmtNum(sumPET)}</td>
    <td class="num total-num">${fmtNum(sumTotal)}</td>
    <td class="num" style="color:#e74c3c">${sumRejPct}%</td>
  </tr>`;

  // Product breakdown
  const products = {};
  allRecs.forEach(r => {
    const k = r.product + '||' + r.type;
    if (!products[k]) products[k] = {product:r.product, type:r.type, good:0};
    products[k].good += r.good;
  });
  const prodMax = Math.max(...Object.values(products).map(p=>p.good), 1);
  const prodRows = Object.values(products)
    .sort((a,b)=>b.good-a.good)
    .map(p => {
      const pct = ((p.good/prodMax)*100).toFixed(0);
      return `<tr>
        <td><span class="badge ${p.type}">${p.type}</span></td>
        <td>${p.product}</td>
        <td class="num" style="color:var(--text)">${fmtNum(p.good)}</td>
        <td style="width:200px">
          <div class="bar-wrap">
            <div class="bar-bg"><div class="bar-fill ${p.type}" style="width:${pct}%"></div></div>
            <span class="bar-pct">${pct}%</span>
          </div>
        </td>
      </tr>`;
    }).join('');

  // L3 quick stat
  const l3Data = RAW_DATA.filter(r => r.customer===selectedCustomer && l3.includes(r.month));
  const l3Total = l3Data.reduce((s,r)=>s+r.good, 0);

  document.getElementById('mainContent').innerHTML = `
    <div class="section-hdr">
      <h2>${selectedCustomer}</h2>
      <span class="pill">${months.length} month${months.length!==1?'s':''} on record</span>
    </div>

    <!-- KPI Cards -->
    <div class="kpi-row">
      <div class="kpi-card tube">
        <div class="kpi-label">TUBE Produced</div>
        <div class="kpi-value tube">${fmtNum(tubeTotal)}</div>
        <div class="kpi-sub">filtered selection</div>
      </div>
      <div class="kpi-card pet">
        <div class="kpi-label">PET Produced</div>
        <div class="kpi-value pet">${fmtNum(petTotal)}</div>
        <div class="kpi-sub">filtered selection</div>
      </div>
      <div class="kpi-card total">
        <div class="kpi-label">Grand Total</div>
        <div class="kpi-value total">${fmtNum(grandTotal)}</div>
        <div class="kpi-sub">Last 3M: ${fmtNum(l3Total)}</div>
      </div>
      <div class="kpi-card reject">
        <div class="kpi-label">Reject / Scrap</div>
        <div class="kpi-value reject">${fmtNum(rejectTotal)}</div>
        <div class="kpi-sub">${grandTotal>0?((rejectTotal/grandTotal)*100).toFixed(2):0}% reject rate</div>
      </div>
    </div>

    <!-- Filters -->
    <div class="filter-bar">
      <span class="filter-label">Period:</span>
      ${filterBtns}
      <div class="type-toggle">
        <span class="filter-label">Type:</span>
        ${typeBtns}
      </div>
    </div>

    <!-- Charts + Monthly table -->
    <div class="charts-row">
      <div class="card">
        <div class="card-title">Monthly Production Trend</div>
        <div class="chart-wrap"><canvas id="trendChart"></canvas></div>
      </div>
      <div class="card">
        <div class="card-title">Monthly Breakdown</div>
        <table>
          <thead><tr><th>Month</th><th>TUBE</th><th>PET</th><th>Total</th><th>Rej%</th></tr></thead>
          <tbody>${monthRows}${totalsRow}</tbody>
        </table>
      </div>
    </div>

    <!-- Product Breakdown -->
    <div class="card full-width">
      <div class="card-title">Product Breakdown (filtered)</div>
      <table>
        <thead><tr><th>Type</th><th>Product</th><th>Qty Produced</th><th>Share</th></tr></thead>
        <tbody>${prodRows || '<tr><td colspan="4" style="color:var(--muted);text-align:center;padding:20px">No records match filter</td></tr>'}</tbody>
      </table>
    </div>
  `;

  renderChart(months, monthSummary);
}

// ============================================================
//  CHART
// ============================================================
function renderChart(months, monthSummary) {
  if (chartInstance) chartInstance.destroy();
  const ctx = document.getElementById('trendChart');
  if (!ctx) return;

  chartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: months,
      datasets: [
        {
          label: 'TUBE',
          data: months.map(m => (monthSummary[m]||{}).TUBE || 0),
          backgroundColor: 'rgba(58,123,213,0.75)',
          borderRadius: 5,
        },
        {
          label: 'PET',
          data: months.map(m => (monthSummary[m]||{}).PET || 0),
          backgroundColor: 'rgba(39,174,96,0.75)',
          borderRadius: 5,
        },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#8899aa', font: {size: 11} } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString()}`
          }
        }
      },
      scales: {
        x: { ticks: { color: '#8899aa' }, grid: { color: 'rgba(255,255,255,.04)' } },
        y: {
          ticks: { color: '#8899aa', callback: v => fmtNumShort(v) },
          grid: { color: 'rgba(255,255,255,.06)' }
        }
      }
    }
  });
}

// ============================================================
//  FILTER HANDLERS
// ============================================================
function setMonthFilter(key) {
  if (key === 'l3') {
    const l3 = last3Months();
    activeMonthFilter = 'l3_special';
    // temp: override filter logic for l3
    window._l3months = l3;
  } else {
    activeMonthFilter = key;
    window._l3months = null;
  }
  activeMonthFilter = key;
  renderMain();
}

function setTypeFilter(t) {
  activeTypeFilter = t;
  renderMain();
}

// Override filtered() to handle l3
const _origFiltered = filtered;
function filtered() {
  const l3months = window._l3months;
  return RAW_DATA.filter(r => {
    if (r.customer !== selectedCustomer) return false;
    if (activeTypeFilter !== 'all' && r.type !== activeTypeFilter) return false;
    if (activeMonthFilter === 'l3' && l3months && !l3months.includes(r.month)) return false;
    if (activeMonthFilter !== 'all' && activeMonthFilter !== 'l3' && r.month !== activeMonthFilter) return false;
    return true;
  });
}

// ============================================================
//  UTILS
// ============================================================
function fmtNum(n) {
  if (n == null || n === '') return '—';
  return Math.round(n).toLocaleString();
}
function fmtNumShort(n) {
  if (n >= 1000000) return (n/1000000).toFixed(1)+'M';
  if (n >= 1000)    return (n/1000).toFixed(0)+'K';
  return n;
}
</script>
</body>
</html>"""


def generate_html(records):
    customers = sorted(set(r["customer"] for r in records))
    months    = sorted(set(r["month"] for r in records))

    data_json = json.dumps(records, ensure_ascii=False, indent=None)
    generated = datetime.datetime.now().strftime("%d %B %Y  %H:%M")
    months_covered = ", ".join(months)

    html = HTML_TEMPLATE \
        .replace("__DATA__", data_json) \
        .replace("__GENERATED__", generated) \
        .replace("__MONTHS_COVERED__", months_covered)

    return html


# ============================================================
#  MAIN
# ============================================================

if __name__ == "__main__":
    print("=" * 58)
    print("  Tubex Customer Report Generator")
    print("=" * 58)

    print("\nExtracting production records...")
    records = extract_records()

    if not records:
        print("\n  No records found. Check SOURCES configuration.")
        sys.exit(1)

    customers = sorted(set(r["customer"] for r in records))
    months    = sorted(set(r["month"] for r in records))
    print(f"\n  Records  : {len(records):,}")
    print(f"  Customers: {len(customers)} -> {', '.join(customers)}")
    print(f"  Months   : {', '.join(months)}")

    print(f"\nGenerating HTML...")
    html = generate_html(records)

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    kb = os.path.getsize(OUTPUT_HTML) // 1024
    print(f"\n  Saved: {OUTPUT_HTML}  ({kb} KB)")
    print("\n  Open customer_report.html in your browser.")
    print("=" * 58)
