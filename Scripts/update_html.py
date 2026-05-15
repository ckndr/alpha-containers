"""
update_html.py
Alpha Containers — Dashboard HTML Auto-Updater
------------------------------------------------
Reads AlphaContainers_vX_XX.xlsx → calculates live KPIs, MTD production,
downtime, and order compliance → injects into AlphaContainers_App.html.

Also rebuilds PRODUCTS array and BOM dict directly from Excel on every run.
New or re-activated products are automatically picked up — no manual HTML edits needed.

Run manually or add to Run All Updates.bat:
    python update_html.py

Author: Sikander / Claude
Version: 2.1
Changes from 2.0:
  - read_orders_dynamic: also captures rows where dispatch > 0 even when ordered = 0
    (fixes PHLOGIN GEL and any future completed-but-dispatched products being missed)
  - DASH_DATA kpi now includes tubeMTDDispatch / petMTDDispatch totals
    requires dispatch.xls date column mapping to compute; add when file structure known)
"""

import os, re, json, glob
from datetime import datetime, date

# ── PATH SETUP ──────────────────────────────────────────────
# Scripts live in AlphaContainers/Scripts/ — Excel and HTML are one level up
DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

excel_pattern = os.path.join(DIR, 'AlphaContainers_v*.xlsx')
excel_files   = sorted(glob.glob(excel_pattern))
if not excel_files:
    raise FileNotFoundError(f"No AlphaContainers_v*.xlsx found in {DIR}")
EXCEL_PATH = excel_files[-1]
HTML_PATH  = os.path.join(DIR, 'AlphaContainers_App.html')

print(f"Reading:  {os.path.basename(EXCEL_PATH)}")
print(f"Updating: {os.path.basename(HTML_PATH)}")

# ── LOAD EXCEL ───────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    raise ImportError("openpyxl not installed. Run: pip install openpyxl --break-system-packages")

wb     = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
ws_pl  = wb['Production_Log']
ws_db  = wb['Tubex_Dashboard']
ws_bom = wb['BOM']
ws_cat = wb['Product_Catalog']

# ── MONTH FILTER ─────────────────────────────────────────────
now        = datetime.now()
cur_month  = now.month
cur_year   = now.year
month_name = now.strftime('%B %Y')

# ── READ PRODUCTION LOG ──────────────────────────────────────
# Cols (1-based): A=Date B=Machine C=Customer D=Product E=Dia F=PID
#                 G=Target H=Good I=Reject J=Waste%
#                 K=Mech L=Elec M=MatShort N=Changeover O=Operations
#                 P=PowerShutdown Q=GasShutdown R=WorkersShort

mtd_by_pid  = {}
yest_tube   = 0
yest_pet    = 0
latest_date = None

downtime_cols = {
    'Mechanical':        11,
    'Electrical':        12,
    'Material Shortage': 13,
    'Changeover':        14,
    'Operations':        15,
    'Power Shutdown':    16,
    'Gas Shutdown':      17,
    'Workers Shortage':  18,
}
dt_totals = {k: 0.0 for k in downtime_cols}

for row in ws_pl.iter_rows(min_row=3, values_only=True):
    row_date  = row[0]
    machine   = row[1]
    prod_name = row[3]
    pid       = row[5]
    good_qty  = row[7]

    if not row_date or not machine: continue
    if not isinstance(row_date, (date, datetime)): continue

    row_dt = row_date if isinstance(row_date, datetime) else datetime(row_date.year, row_date.month, row_date.day)
    if row_dt.month != cur_month or row_dt.year != cur_year: continue

    if latest_date is None or row_dt.date() > latest_date:
        latest_date = row_dt.date()

    for cat, col_idx in downtime_cols.items():
        val = row[col_idx - 1]
        if val and isinstance(val, (int, float)):
            dt_totals[cat] += float(val)

    if not pid or not good_qty: continue
    good_qty = int(good_qty)

    mach_up  = str(machine).upper()
    is_print = mach_up.startswith('PRINT') or mach_up.startswith('PLINE')
    is_pet   = mach_up.startswith('PF') or mach_up.startswith('PET')
    is_varn  = '(VARNISH)' in str(prod_name).upper()

    if (is_print and not is_varn) or is_pet:
        pid_int = int(pid)
        mtd_by_pid[pid_int] = mtd_by_pid.get(pid_int, 0) + good_qty

# Yesterday's production
if latest_date:
    for row in ws_pl.iter_rows(min_row=3, values_only=True):
        row_date  = row[0]
        machine   = row[1]
        prod_name = row[3]
        good_qty  = row[7]
        if not row_date or not machine or not good_qty: continue
        if not isinstance(row_date, (date, datetime)): continue
        row_dt = row_date if isinstance(row_date, datetime) else datetime(row_date.year, row_date.month, row_date.day)
        if row_dt.date() != latest_date: continue
        mach_up  = str(machine).upper()
        is_print = mach_up.startswith('PRINT') or mach_up.startswith('PLINE')
        is_pet   = mach_up.startswith('PF') or mach_up.startswith('PET')
        is_varn  = '(VARNISH)' in str(prod_name).upper()
        if is_print and not is_varn:
            yest_tube += int(good_qty)
        elif is_pet:
            yest_pet += int(good_qty)

tube_mtd = sum(v for k, v in mtd_by_pid.items() if k < 8000)
pet_mtd  = sum(v for k, v in mtd_by_pid.items() if k >= 8000)

print(f"\nKPIs for {month_name}:")
print(f"  Tube MTD: {tube_mtd:,}  |  Yesterday: {yest_tube:,}")
print(f"  PET MTD:  {pet_mtd:,}   |  Yesterday: {yest_pet:,}")

# ── READ ACTIVE ORDERS (DYNAMIC SCAN) ────────────────────────
# Dashboard layout (1-based columns):
#   A=blank  B=Type(TUBE/PET)  C=Customer  D=Product  E=Dia
#   F=ProdID  G=Orders  H=MTDProd  I=Remaining  J=Compliance  K=Dispatch
#
# A row is ACTIVE if col B matches the type AND at least ONE of:
#   • col G (ordered qty)   > 0  — has an open or in-progress order
#   • MTD produced (Log)    > 0  — was produced this month even if order closed
#   • col K (dispatch)      > 0  — dispatched this month (e.g. PHLOGIN ordered=0, dispatch=35,280)
# A row is inactive ONLY when all three are zero simultaneously.
# Scanning rows 11–100 catches any future rows added when more orders come in.

def read_orders_dynamic(ws, order_type):
    orders = []
    is_pet = order_type.upper() == 'PET'
    for r in range(11, 100):
        type_val = ws.cell(row=r, column=2).value   # col B
        pid      = ws.cell(row=r, column=6).value   # col F
        ordered  = ws.cell(row=r, column=7).value   # col G
        dispatch = ws.cell(row=r, column=11).value  # col K

        if not type_val:
            continue
        if str(type_val).strip().upper() != order_type.upper():
            continue
        if not pid:
            continue
        try:
            pid_int = int(pid)
        except (TypeError, ValueError):
            continue
        try:
            ordered_int = int(ordered) if ordered else 0
        except (TypeError, ValueError):
            ordered_int = 0
        try:
            dispatch_int = int(dispatch) if dispatch else 0
        except (TypeError, ValueError):
            dispatch_int = 0

        # Resolve MTD produced here so we can use it in the active check
        produced = mtd_by_pid.get(pid_int, 0)

        # Skip truly inactive rows — must have NO order, NO production, AND NO dispatch
        if ordered_int == 0 and produced == 0 and dispatch_int == 0:
            continue

        product  = ws.cell(row=r, column=4).value
        customer = ws.cell(row=r, column=3).value
        dia      = ws.cell(row=r, column=5).value

        dia_raw = str(dia).replace(' ml', '').replace('ml', '') \
                          .replace(' mm', '').replace('mm', '').strip() if dia else ''
        dia_str = (dia_raw + ('ml' if is_pet else 'mm')) if dia_raw else '—'

        orders.append({
            'pid':      pid_int,
            'product':  str(product) if product else '',
            'customer': str(customer) if customer else '—',
            'dia':      dia_str,
            'ordered':  ordered_int,
            'produced': produced,
            'dispatch': dispatch_int,
        })
    return orders

tube_orders = read_orders_dynamic(ws_db, 'TUBE')
pet_orders  = read_orders_dynamic(ws_db, 'PET')

# ── DISPATCH TOTALS ───────────────────────────────────────────
# Monthly: sum dispatch column across all active order rows
tube_mtd_dispatch = sum(o['dispatch'] for o in tube_orders)
pet_mtd_dispatch  = sum(o['dispatch'] for o in pet_orders)

print(f"\nTube orders found: {len(tube_orders)}")
for o in tube_orders:
    print(f"  PID {o['pid']} — {o['product'][:35]} | ordered={o['ordered']:,} | produced={o['produced']:,} | dispatch={o['dispatch']:,}")
print(f"\nPET orders found: {len(pet_orders)}")
for o in pet_orders:
    print(f"  PID {o['pid']} — {o['product'][:35]} | ordered={o['ordered']:,} | produced={o['produced']:,} | dispatch={o['dispatch']:,}")
print(f"\nDispatch MTD — Tubes: {tube_mtd_dispatch:,}  |  PET: {pet_mtd_dispatch:,}")

# ── BUILD PRODUCTS FROM PRODUCT_CATALOG ──────────────────────
# Shorten long customer names for display in the app
CUST_SHORT = {
    'Brookes Pharma Private Limited':              'Brookes Pharma',
    'Samsol International Private Limited':        'Samsol International',
    'Mablay Beauty PVT LTD.':                      'Mablay Beauty',
    'Golden Pearl Cosmetics (PVT) LTD':            'Golden Pearl Cosmetics',
    'Al-Rehman Group':                             'Al-Rehman Group',
    'Professional Beauty Solution (PVT) LTD.Pakistan': 'Professional Beauty Solution',
    'Professional Beauty Solution (PVT) LTD. Pakistan':'Professional Beauty Solution',
    'Seatle (Private) Limited':                    'Seatle (Private) Limited',
    'Mega Grey':   'Mega Grey',
    'Hola Hair':   'Hola Hair',
    'Adore':       'Adore',
    'Bahadur':     'Bahadur',
    'DTM':         'DTM',
    'Alpha Labs PVT LTD':  'Alpha Labs',
    'Alpha Labs PVT LTD.': 'Alpha Labs',
}

products_list = []
for row in ws_cat.iter_rows(min_row=3, values_only=True):
    pid      = row[0]   # col A
    customer = row[2]   # col C
    name     = row[3]   # col D
    dia      = row[4]   # col E
    if not pid or not name:
        continue
    try:
        pid_int = int(pid)
    except (TypeError, ValueError):
        continue

    dia_str = str(dia).strip() if dia else '?'
    is_pet  = 'ml' in dia_str.lower()
    p_type  = 'pet' if is_pet else 'tube'

    if is_pet:
        dia_out = dia_str.replace(' ', '').lower()   # "120 ml" → "120ml"
    else:
        try:
            d = float(dia_str.replace('mm', '').strip())
            dia_out = str(int(d)) if d == int(d) else str(d)
        except (ValueError, TypeError):
            dia_out = dia_str

    cust_short = CUST_SHORT.get(str(customer), str(customer)) if customer else '—'

    products_list.append({
        'pid':      pid_int,
        'type':     p_type,
        'dia':      dia_out,
        'customer': cust_short,
        'name':     str(name),
    })

print(f"\nProducts loaded from catalog: {len(products_list)}")

# ── BUILD BOM FROM BOM SHEET ─────────────────────────────────
# Categories excluded from the Material Calculator (packing jali / metal frames)
SKIP_CATS = {'JALI', 'IRON AND METAL'}

# UOM normalisation — match what MRP/Inventory sheets use
UOM_NORM = {'KGS': 'kg', 'LTRS': 'L', 'NO': 'pcs', 'PCS': 'pcs'}

bom_dict = {}   # pid_int → [mat, ...]
for row in ws_bom.iter_rows(min_row=3, values_only=True):
    pid        = row[0]    # col A = Product ID
    mat_cat    = row[5]    # col F = Material Category
    item_id    = row[6]    # col G = Item ID
    item_name  = row[7]    # col H = Item Name
    uom        = row[8]    # col I = UOM
    per_1000   = row[9]    # col J = Per 1000 Units
    mat_group  = row[10]   # col K = Material Group
    scrap      = row[11]   # col L = Scrap %
    change_note= row[12]   # col M = Change Note

    if not pid or not mat_cat:
        continue
    try:
        pid_int = int(pid)
    except (TypeError, ValueError):
        continue

    cat_upper = str(mat_cat).strip().upper()
    if cat_upper in SKIP_CATS:
        continue
    if item_id is None:
        continue        # no item assigned yet
    if per_1000 is None:
        continue        # no rate at all — row is completely empty

    # est flag: red-font items are estimated/placeholder
    est = False
    try:
        if int(item_id) >= 9000:
            est = True
    except (TypeError, ValueError):
        pass
    if mat_group and str(mat_group).strip().lower() == 'estimated':
        est = True
    if item_name and '(est)' in str(item_name).lower():
        est = True
    if change_note and ('PLACEHOLDER' in str(change_note).upper() or
                        'EST.' in str(change_note).upper()):
        est = True
    # Zero-rate line = not yet confirmed, treat as estimated
    if per_1000 == 0:
        est = True

    uom_raw = str(uom).strip() if uom else 'kg'
    uom_out = UOM_NORM.get(uom_raw.upper(), uom_raw)

    mat = {
        'cat':  cat_upper,
        'id':   int(item_id),
        'name': str(item_name).strip() if item_name else '',
        'uom':  uom_out,
        'r':    round(float(per_1000), 7),
        's':    round(float(scrap), 4) if scrap else 0.1,
        'est':  est,
    }

    bom_dict.setdefault(pid_int, []).append(mat)

print(f"BOM entries loaded: {sum(len(v) for v in bom_dict.values())} lines across {len(bom_dict)} products")

# ── JS HELPERS ───────────────────────────────────────────────
def js_esc(s):
    return str(s).replace('\\', '\\\\').replace("'", "\\'").replace('\n', '').replace('\r', '')

def build_products_js(pl):
    lines = [
        f"  {{pid:{p['pid']},type:'{p['type']}',dia:'{js_esc(p['dia'])}'"
        f",customer:'{js_esc(p['customer'])}',name:'{js_esc(p['name'])}'}}"
        for p in pl
    ]
    return "const PRODUCTS = [\n" + ",\n".join(lines) + "\n];"

def build_bom_js(bd):
    pid_lines = []
    for pid in sorted(bd.keys()):
        mats = bd[pid]
        mat_parts = []
        for m in mats:
            est_js = 'true' if m['est'] else 'false'
            # Compact number formatting: drop trailing zeros
            r_str = f"{m['r']:.7f}".rstrip('0').rstrip('.')
            s_str = f"{m['s']:.4f}".rstrip('0').rstrip('.')
            mat_parts.append(
                f"{{cat:'{m['cat']}',id:{m['id']},name:'{js_esc(m['name'])}'"
                f",uom:'{m['uom']}',r:{r_str},s:{s_str},est:{est_js}}}"
            )
        pid_lines.append(f"  {pid}:[{','.join(mat_parts)}]")
    return "const BOM = {\n" + ",\n".join(pid_lines) + "\n};"

# ── BUILD DASH_DATA ──────────────────────────────────────────
DOWNTIME_ICONS = {
    'Mechanical':        '🔧',
    'Electrical':        '⚡',
    'Workers Shortage':  '👷',
    'Power Shutdown':    '🔌',
    'Material Shortage': '📦',
    'Gas Shutdown':      '🔥',
    'Changeover':        '🔄',
    'Operations':        '⚙️',
}

dash_data = {
    'month':       month_name,
    'lastUpdated': f"Updated {now.strftime('%d-%b-%Y %H:%M')} from {os.path.basename(EXCEL_PATH)}",
    'kpi': {
        'tubeYest':         yest_tube,
        'tubeMTD':          tube_mtd,
        'petYest':          yest_pet,
        'petMTD':           pet_mtd,
        'tubeMTDDispatch':  tube_mtd_dispatch,
        'petMTDDispatch':   pet_mtd_dispatch,
    },
    'downtime': [
        {'cat': cat, 'icon': DOWNTIME_ICONS.get(cat, '⏱'), 'hrs': round(dt_totals[cat] / 60, 2)}
        for cat in downtime_cols
    ],
    'tubeOrders': tube_orders,
    'petOrders':  pet_orders,
}

# ── INJECT INTO HTML ─────────────────────────────────────────
if not os.path.exists(HTML_PATH):
    raise FileNotFoundError(f"HTML not found: {HTML_PATH}")

with open(HTML_PATH, 'r', encoding='utf-8') as f:
    html = f.read()

# ── SELF-PATCH: add CATALOG markers on first run ──────────────
# This only runs once — subsequent runs find the markers already in place.
if '/* CATALOG_START */' not in html:
    print("  First run: adding CATALOG markers to HTML...")
    # Insert CATALOG_START right after DATA_END (on a new line)
    html = html.replace(
        '/* DATA_END */',
        '/* DATA_END */\n/* CATALOG_START */'
    )
    # Insert CATALOG_END just before const CAT_ICON
    if '\nconst CAT_ICON' in html:
        html = html.replace('\nconst CAT_ICON', '\n/* CATALOG_END */\nconst CAT_ICON', 1)
    else:
        raise RuntimeError("Could not locate 'const CAT_ICON' to place CATALOG_END marker.")

# ── INJECT DASH_DATA ─────────────────────────────────────────
marker_start = '/* DATA_START */'
marker_end   = '/* DATA_END */'
pos_start    = html.find(marker_start)
pos_end      = html.find(marker_end)
if pos_start == -1 or pos_end == -1:
    raise RuntimeError("DATA_START / DATA_END markers not found.")

new_data_block = (
    f"{marker_start}\n"
    f"const DASH_DATA = {json.dumps(dash_data, indent=2, ensure_ascii=False)};\n"
    f"{marker_end}"
)
html = html[:pos_start] + new_data_block + html[pos_end + len(marker_end):]

# Update timestamp comment if present
ts_old = re.search(r'Last generated: .+', html)
if ts_old:
    html = html.replace(ts_old.group(), f"Last generated: {now.strftime('%d-%b-%Y %H:%M')}")

# ── INJECT PRODUCTS + BOM ────────────────────────────────────
cat_start = '/* CATALOG_START */'
cat_end   = '/* CATALOG_END */'
pos_cs    = html.find(cat_start)
pos_ce    = html.find(cat_end)
if pos_cs == -1 or pos_ce == -1:
    raise RuntimeError("CATALOG_START / CATALOG_END markers not found after patching.")

products_js = build_products_js(products_list)
bom_js      = build_bom_js(bom_dict)

new_catalog_block = (
    f"{cat_start}\n"
    f"{products_js}\n\n"
    f"{bom_js}\n"
    f"{cat_end}"
)
html = html[:pos_cs] + new_catalog_block + html[pos_ce + len(cat_end):]

# ── WRITE HTML ───────────────────────────────────────────────
with open(HTML_PATH, 'w', encoding='utf-8') as f:
    f.write(html)

# ── UPDATE SERVICE WORKER CACHE VERSION ──────────────────────
SW_PATH = os.path.join(DIR, 'sw.js')
if os.path.exists(SW_PATH):
    with open(SW_PATH, 'r', encoding='utf-8') as f:
        sw = f.read()
    new_cache = f"alpha-containers-{now.strftime('%Y%m%d%H%M')}"
    sw = re.sub(r"const CACHE_NAME = '[^']+';",
                f"const CACHE_NAME = '{new_cache}';", sw)
    with open(SW_PATH, 'w', encoding='utf-8') as f:
        f.write(sw)
    print(f"  SW cache version → {new_cache}")

print(f"\n✓ HTML updated successfully → {os.path.basename(HTML_PATH)}")
print(f"  Dashboard orders: {len(tube_orders)} tube + {len(pet_orders)} PET")
print(f"  Products in catalog: {len(products_list)}  |  BOM products: {len(bom_dict)}")
print(f"  Open in Chrome on PC or Android to view dashboard.")
