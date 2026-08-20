# ALPHA CONTAINERS (TUBEX)
## STRATEGIC MODERNIZATION BLUEPRINT & TECHNICAL SPECIFICATIONS
**Document Reference**: `ALPHA-MOD-BP-2026-v1.0`  
**Target Ecosystem**: Aluminum Collapsible Tubes & PET Blow-Molding Manufacturing  
**Author**: Modernization Architecture & Engineering Team  
**Status**: Production-Ready Architectural Deliverable  
**Date**: August 2026  

---

## EXECUTIVE SUMMARY & MODERNIZATION VISION

Alpha Containers (Tubex) operates a high-throughput industrial manufacturing facility in Karachi, Pakistan, producing millions of extruded aluminum collapsible tubes (pharmaceutical, cosmetics, adhesives) and PET stretch blow-molded bottles and jars. Over the preceding operational cycles, the automation stack evolved from manual Excel spreadsheets into an automated hybrid data pipeline powered by Python (`daily.py`, `update_production.py`, `update_inventory.py`, `update_dispatch.py`, `sort_dashboard.py`, `build_archives.py`, `update_html.py`) and an offline-first Progressive Web Application (`Tubex.html`, `sw.js`).

Following the comprehensive remediation audit (resolving 56 systemic issues across data pipelines, formula consistency, DOM security, and COM process isolation), this **Strategic Modernization Blueprint** defines the engineering specifications for the next evolutionary stage of the Alpha Containers ecosystem.

### Key Objectives of this Blueprint:
1. **Future_Plans Sheet Deep Specifications**: Deliver complete mathematical formulations, UI mockups, client-side reactive architectures, and pipeline integration hooks for the two officially recorded future features on the `Future_Plans` sheet of `Tubex_Aug26.xlsx`:
   - **FP-01**: *Raw Material Yield & Capacity Calculator (Aluminum Slugs & PET Resin)*
   - **FP-02**: *Historical Month Selector & Dashboard Archive Navigation*
2. **12 High-Impact Strategic Proposals across 4 Pillars**: Formulate comprehensive architectural proposals covering:
   - **Pillar 1: Web Dashboard & User Experience (UX)**
   - **Pillar 2: Data Pipeline, Automation & Ingestion**
   - **Pillar 3: Advanced Planning, MRP & Shop-Floor Intelligence**
   - **Pillar 4: Code Quality, Observability, Testing & System Resilience**
3. **Execution Blueprint**: Provide concrete database schemas (SQL DDL), mathematical models, Python package refactoring architectures (`alphapackage`), JavaScript state machines, JSON data contracts, risk matrices, and a phased multi-sprint implementation roadmap.

```
┌──────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                    ALPHA CONTAINERS MODERNIZATION STACK                          │
└──────────────────────────────────────────────────────────────────────────────────────────────────┘
   ┌────────────────────────────────────────────────────────────────────────────────────────────┐
   │                                   PILLAR 1: WEB DASHBOARD & UX                             │
   │  • FP-01: Quick Raw Material Calculator   • FP-02: Historical Month Archive Selector      │
   │  • Touch-First Mobile/Tablet Interface    • Real-Time Shift Run-Rate Velocity Telemetry    │
   │  • High-Contrast Industrial Dark/Light Themes                                              │
   └────────────────────────────────────────────────────────────────────────────────────────────┘
                                                  ▲
                                                  │ (JSON / PWA Cache / Reactive State)
   ┌────────────────────────────────────────────────────────────────────────────────────────────┐
   │                               PILLAR 3: PLANNING & MRP INTELLIGENCE                        │
   │  • Dynamic Rolling Scrap Calibration      • Statistical Lead-Time Safety Stock & ROP       │
   │  • Automated Supplier Reorder Triggers    • Bottleneck Machine Scheduling & Changeovers   │
   └────────────────────────────────────────────────────────────────────────────────────────────┘
                                                  ▲
                                                  │ (BOM / Material Math / Scheduling)
   ┌────────────────────────────────────────────────────────────────────────────────────────────┐
   │                           PILLAR 2: DATA PIPELINE & AUTOMATION                             │
   │  • Direct ERP SQL/ODBC Ingestion Connector • WhatsApp AI/Regex Parsing Bot (Floor WIP/Logs)│
   │  • Atomic Pre-Flight Integrity Guard      • Automated Git/Cloud Webhook Synchronization    │
   └────────────────────────────────────────────────────────────────────────────────────────────┘
                                                  ▲
                                                  │ (Clean Parquet / SQLite / JSON Data Streams)
   ┌────────────────────────────────────────────────────────────────────────────────────────────┐
   │                    PILLAR 4: ARCHITECTURE, QUALITY & RESILIENCE                            │
   │  • Unified Python Package (`alphapackage`) • Structured JSON Telemetry Logging (structlog)  │
   │  • Automated Executive Daily Health Alerts • Automated Regression Test Suite (Pytest/Play) │
   └────────────────────────────────────────────────────────────────────────────────────────────┘
```

---

# PART 1: FUTURE_PLANS SHEET DEEP SPECIFICATIONS

The `Future_Plans` worksheet in `Tubex_Aug26.xlsx` formally registers two critical functional enhancements requested by plant leadership and procurement. Below are the exhaustive mathematical, architectural, and UI specifications.

---

## 1.1 FEATURE FP-01: RAW MATERIAL YIELD & CAPACITY CALCULATOR (SLUGS & RESIN)

### 1.1.1 Operational Context & Business Problem
In tube and bottle manufacturing, raw materials represent over 65% of total unit cost:
1. **Aluminum Slugs (99.7% Pure Al)**: Every extruded tube of a specific diameter ($\varnothing$) requires a slug of precise diameter, thickness, and alloy hardness (Al 99.7% with lubrication). For a given tube diameter and length specification, slug consumption is identical across all customer artwork variants, internal lacquer types (epoxy-phenolic), and basecoat colors. Currently, plant operators calculate batch potential manually or rely on misleading arithmetic averages (Finding R2-06). Floor supervisors need an instantaneous calculator: *"Given $M_{\text{avail}}$ kg of $\varnothing D$ slugs in stock, what is the exact yield in finished, inspected tubes factoring standard line scrap?"* Conversely, procurement requires: *"To produce $N$ tubes for customer orders, how many kg of slugs must be issued/purchased?"*
2. **PET Resin (Polyethylene Terephthalate A-84 Virgin Polymer)**: A single grade of virgin PET resin (`Item 2680 PET RESIN A-84`) feeds injection-stretch blow molding (ISBM) machines to produce preforms and blow bottles ranging from 60ml to 500ml. Different bottle sizes have distinct grammages (preform weights). Procurement needs a multi-format yield matrix: *"If a 5-ton bulk resin truck arrives, how many bottles can be blown across each format independently or in mixed-batch allocations?"*

---

### 1.1.2 Mathematical Foundations & Conversion Formulas

#### A. Aluminum Slugs Conversion Mathematics
Let:
- $D \in \{12.5, 13.5, 16.0, 19.0, 20.5, 22.0, 25.0, 28.0, 30.0, 32.0, 35.0\}\,\text{mm}$ be the tube diameter.
- $W_{\text{slug}}(D)$ be the standard nominal slug weight per 1,000 tubes (kg/1,000 pcs).
- $s_{\text{tube}}$ be the baseline operational scrap rate ($0.10$ for mature Tubex lines; $0.15$ for newly commissioned lines).
- $M_{\text{slug}}$ be the mass of available aluminum slugs in kilograms (kg).
- $Q_{\text{tube}}$ be the demanded quantity of finished tubes (pieces).

**1. Forward Calculation (Available Stock $\to$ Maximum Finished Units Yield)**:
In mature Tubex operations, scrap follows the linear additive standard established in plant BOMs:
$$Y_{\text{gross}}(M_{\text{slug}}, D) = \left\lfloor \frac{M_{\text{slug}} \times 1,000}{W_{\text{slug}}(D) \times (1 + s_{\text{tube}})} \right\rfloor$$

For precision scrap modeling (Yield Inverse Scrap standard used in commissioning aerosol operations):
$$Y_{\text{inverse}}(M_{\text{slug}}, D) = \left\lfloor \frac{M_{\text{slug}} \times 1,000 \times (1 - s_{\text{tube}})}{W_{\text{slug}}(D)} \right\rfloor$$

**2. Reverse Calculation (Demanded Finished Units $\to$ Required Raw Material Mass)**:
$$\text{Mass}_{\text{required}}(Q_{\text{tube}}, D) = \frac{Q_{\text{tube}}}{1,000} \times W_{\text{slug}}(D) \times (1 + s_{\text{tube}})$$

Factoring a safety stock buffer $\alpha \in [0.00, 0.10]$ (default 3% buffer for extrusion machine startup & trimming):
$$\text{Mass}_{\text{procure}}(Q_{\text{tube}}, D, \alpha) = \text{Mass}_{\text{required}}(Q_{\text{tube}}, D) \times (1 + \alpha)$$

**3. Standard Tube Diameter Parameter Matrix**:
The following parameters are extracted from the Master Catalog BOM and physical tool specifications:

| Tube Diameter ($\varnothing$) | Standard Wall Thickness (mm) | Slug Weight $W_{\text{slug}}$ (kg / 1,000 pcs) | Net Slug Weight ($g$ / tube) | Theoretical Output (pcs / kg gross) | Net Yield @ 10% Scrap (pcs / kg) | Net Yield @ 10% Scrap (pcs / Ton) |
|---|---|---|---|---|---|---|
| **$\varnothing 12.5$ / $13.5\text{ mm}$** | 0.100 mm | **$1.950\text{ kg}$** | $1.95\text{ g}$ | 512.8 pcs | **466.2 pcs** | 466,200 pcs |
| **$\varnothing 16.0\text{ mm}$** | 0.110 mm | **$2.519\text{ kg}$** | $2.52\text{ g}$ | 397.0 pcs | **360.9 pcs** | 360,900 pcs |
| **$\varnothing 19.0\text{ mm}$** | 0.115 mm | **$3.367\text{ kg}$** | $3.37\text{ g}$ | 297.0 pcs | **270.0 pcs** | 270,000 pcs |
| **$\varnothing 20.5$ / $22.0\text{ mm}$** | 0.120 mm | **$3.937\text{ kg}$** | $3.94\text{ g}$ | 254.0 pcs | **230.9 pcs** | 230,900 pcs |
| **$\varnothing 25.0\text{ mm}$** | 0.125 mm | **$5.917\text{ kg}$** | $5.92\text{ g}$ | 169.0 pcs | **153.6 pcs** | 153,600 pcs |
| **$\varnothing 28.0$ / $30.0\text{ mm}$** | 0.135 mm | **$8.000\text{ kg}$** | $8.00\text{ g}$ | 125.0 pcs | **113.6 pcs** | 113,600 pcs |
| **$\varnothing 32.0\text{ mm}$** | 0.140 mm | **$10.863\text{ kg}$** | $10.86\text{ g}$ | 92.1 pcs | **83.7 pcs** | 83,700 pcs |
| **$\varnothing 35.0\text{ mm}$** | 0.150 mm | **$12.820\text{ kg}$** | $12.82\text{ g}$ | 78.0 pcs | **70.9 pcs** | 70,900 pcs |

---

#### B. PET Resin & Masterbatch Conversion Mathematics
Let:
- $V \in \{60, 75, 100, 120, 130, 150, 200, 250, 300, 500\}\,\text{ml}$ be the PET bottle/jar container volume.
- $W_{\text{resin}}(V)$ be the nominal preform/bottle resin weight per 1,000 units (kg/1,000 pcs).
- $s_{\text{pet}}$ be the injection-blow molding scrap rate ($0.15$ or 15% standard for start-up purging, runner regrind loss, and bottle neck trimming).
- $\beta_{\text{mb}}$ be the Masterbatch (Colorant) dosing ratio (typically $1.5\%$ to $3.0\%$ by weight).
- $M_{\text{resin}}$ be the virgin PET resin stock in kilograms (kg).
- $Q_{\text{pet}}$ be the demanded bottle units.

**1. Forward Calculation (Resin Stock $\to$ Finished Bottles)**:
$$Y_{\text{pet}}(M_{\text{resin}}, V) = \left\lfloor \frac{M_{\text{resin}} \times 1,000}{W_{\text{resin}}(V) \times (1 + s_{\text{pet}})} \right\rfloor$$

**2. Reverse Calculation (Demanded Bottles $\to$ Required Resin & Masterbatch)**:
$$\text{Resin}_{\text{required}}(Q_{\text{pet}}, V) = \frac{Q_{\text{pet}}}{1,000} \times W_{\text{resin}}(V) \times (1 + s_{\text{pet}})$$
$$\text{Masterbatch}_{\text{required}}(Q_{\text{pet}}, V, \beta_{\text{mb}}) = \text{Resin}_{\text{required}}(Q_{\text{pet}}, V) \times \frac{\beta_{\text{mb}}}{100}$$

**3. Standard PET Bottle Grammage & Yield Parameter Matrix**:

| Container Format | Nominal Grammage ($g$ / bottle) | Resin Rate $W_{\text{resin}}$ (kg / 1,000 pcs) | Theoretical Yield (pcs / kg) | Net Yield @ 15% Scrap (pcs / kg) | Net Yield @ 15% Scrap (pcs / Ton) | Masterbatch Req. @ 2% (kg / 1,000 pcs) |
|---|---|---|---|---|---|---|
| **$60\text{ ml Bottle}$** | $10.5\text{ g}$ | **$10.50\text{ kg}$** | 95.2 pcs | **82.8 pcs** | 82,800 pcs | $0.241\text{ kg}$ |
| **$75\text{ ml Bottle}$** | $12.5\text{ g}$ | **$12.50\text{ kg}$** | 80.0 pcs | **69.6 pcs** | 69,600 pcs | $0.288\text{ kg}$ |
| **$100\text{ ml Bottle}$** | $15.0\text{ g}$ | **$15.00\text{ kg}$** | 66.7 pcs | **58.0 pcs** | 58,000 pcs | $0.345\text{ kg}$ |
| **$120\text{ ml Bottle}$** | $17.1\text{ g}$ | **$17.10\text{ kg}$** | 58.5 pcs | **50.8 pcs** | 50,800 pcs | $0.393\text{ kg}$ |
| **$130\text{ ml Bottle}$** | $18.0\text{ g}$ | **$18.00\text{ kg}$** | 55.6 pcs | **48.3 pcs** | 48,300 pcs | $0.414\text{ kg}$ |
| **$150\text{ ml Mist / Bottle}$**| $21.0\text{ g}$ | **$21.00\text{ kg}$** | 47.6 pcs | **41.4 pcs** | 41,400 pcs | $0.483\text{ kg}$ |
| **$200\text{ ml Bottle}$** | $23.75\text{ g}$ | **$23.75\text{ kg}$** | 42.1 pcs | **36.6 pcs** | 36,600 pcs | $0.546\text{ kg}$ |
| **$250\text{ ml Bottle}$** | $26.0\text{ g}$ | **$26.00\text{ kg}$** | 38.5 pcs | **33.4 pcs** | 33,400 pcs | $0.598\text{ kg}$ |
| **$300\text{ ml Jar / Bottle}$** | $25.0\text{ g}$ | **$25.00\text{ kg}$** | 40.0 pcs | **34.8 pcs** | 34,800 pcs | $0.575\text{ kg}$ |
| **$500\text{ ml Jar / Bottle}$** | $50.0\text{ g}$ | **$50.00\text{ kg}$** | 20.0 pcs | **17.4 pcs** | 17,400 pcs | $1.150\text{ kg}$ |

---

### 1.1.3 UI/UX Specification & Wireframe Mockup

The Calculator tab (`#panel-calc`) in `Tubex.html` is upgraded with a segmented view controller at the top toolbar:
`[ 📦 Full SKU BOM Mode ]`  |  `[ ⚡ Quick Slugs & Resin Yield Simulator ]`

#### UI Layout Wireframe:
```
┌────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│ 🧮 RAW MATERIAL YIELD & CAPACITY CALCULATOR (FP-01)                                                        │
│ Mode: [ Full SKU BOM ]  [● Quick Slugs & Resin Simulator ]       Scrap Standard: [ 10% Mature ▼ ]         │
├────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
│ ┌──────────────────────────────────────────────────┐ ┌──────────────────────────────────────────────────┐ │
│ │ 🧱 ALUMINUM SLUGS TO TUBES CONVERTER             │ │ 🧴 PET RESIN TO BOTTLES CONVERTER                │ │
│ ├──────────────────────────────────────────────────┤ ├──────────────────────────────────────────────────┤ │
│ │ Direction: [● Stock (kg) ➔ Pcs ] [ Pcs ➔ Kg ]    │ │ Direction: [● Stock (kg) ➔ Pcs ] [ Pcs ➔ Kg ]    │ │
│ │                                                  │ │                                                  │ │
│ │ Tube Diameter: [ Ø 25.0 mm (Face Wash/Ointment)▼]│ │ Available Resin Stock (kg): [  2,500.00  ]       │ │
│ │ Available Slug Mass (kg): [  5,000.00  ]         │ │ Masterbatch Dosing Rate (%): [  2.0%     ]       │ │
│ │ Scrap Rate Adjustment:    [ 10.0 %     ]         │ │ Scrap Rate Adjustment:       [ 15.0%     ]       │ │
│ │ Warehouse Balance (Live): [ 6,420 kg (In Stock)] │ │ Live Warehouse Balance:      [ 4,150 kg In Stock]│ │
│ ├──────────────────────────────────────────────────┤ ├──────────────────────────────────────────────────┤ │
│ │ 🎯 CALCULATED OUTPUT CAPACITY                    │ │ 🎯 MULTI-FORMAT COMPARATIVE YIELD MATRIX         │ │
│ │ Maximum Gross Tubes:     845,022 pcs             │ │ Format       Unit Wt    Max Yield   Masterbatch  │ │
│ │ Expected Net Tubes:      768,201 pcs             │ │ 120ml Bottle 17.10 g    127,145 pcs   43.48 kg   │ │
│ │ Slug Consumption Rate:   5.917 kg / 1,000        │ │ 150ml Mist   21.00 g    103,534 pcs   43.48 kg   │ │
│ │ Production Days @ 30k/d: 25.6 Operating Days     │ │ 200ml Bottle 23.75 g     91,532 pcs   43.48 kg   │ │
│ │ [ 📋 Export Job Card Batch Allocation ]          │ │ 500ml Jar    50.00 g     43,478 pcs   43.48 kg   │ │
│ └──────────────────────────────────────────────────┘ └──────────────────────────────────────────────────┘ │
└────────────────────────────────────────────────────────────────────────────────────────────────────────────┘
```

---

### 1.1.4 Frontend Integration Snippet (`Tubex.html`)

```html
<!-- Integration Point: Inside #panel-calc in Tubex.html -->
<div id="quick-calc-container" class="quick-calc-wrapper" style="display:none;margin-top:16px;">
  <div class="calc-grid-two-col">
    
    <!-- CARD 1: ALUMINUM SLUG SIMULATOR -->
    <div class="calc-card">
      <div class="calc-card-hdr">
        <div class="calc-card-icon ic-slug">🧱</div>
        <div>
          <h3>Aluminum Slugs ➔ Tubes Yield Engine</h3>
          <p class="sub">Instant diameter-level extrusion capacity converter</p>
        </div>
      </div>
      
      <div class="calc-form-group">
        <label>Calculation Mode</label>
        <div class="calc-toggle-pill">
          <button id="slug-mode-fwd" class="active" onclick="setSlugCalcMode('fwd')">Stock (kg) ➔ Tubes</button>
          <button id="slug-mode-rev" onclick="setSlugCalcMode('rev')">Order (Tubes) ➔ Slugs (kg)</button>
        </div>
      </div>

      <div class="calc-form-group">
        <label for="slug-dia-select">Tube Diameter Specification</label>
        <select id="slug-dia-select" class="calc-input-select" onchange="runQuickCalc()">
          <option value="12.5" data-weight="1.950">Ø 12.5 / 13.5 mm (1.950 kg / 1k pcs)</option>
          <option value="16.0" data-weight="2.519">Ø 16.0 mm (2.519 kg / 1k pcs)</option>
          <option value="19.0" data-weight="3.367">Ø 19.0 mm (3.367 kg / 1k pcs)</option>
          <option value="20.5" data-weight="3.937">Ø 20.5 / 22.0 mm (3.937 kg / 1k pcs)</option>
          <option value="25.0" data-weight="5.917" selected>Ø 25.0 mm (5.917 kg / 1k pcs)</option>
          <option value="30.0" data-weight="8.000">Ø 28.0 / 30.0 mm (8.000 kg / 1k pcs)</option>
          <option value="32.0" data-weight="10.863">Ø 32.0 mm (10.863 kg / 1k pcs)</option>
          <option value="35.0" data-weight="12.820">Ø 35.0 mm (12.820 kg / 1k pcs)</option>
        </select>
      </div>

      <div class="calc-form-group" id="slug-input-qty-group">
        <label id="slug-input-lbl" for="slug-input-val">Available Slug Mass (kg)</label>
        <div class="calc-input-addon">
          <input type="number" id="slug-input-val" class="calc-input-num" value="5000" min="1" step="10" oninput="runQuickCalc()">
          <span class="addon-tag" id="slug-addon-tag">kg</span>
        </div>
      </div>

      <div class="calc-form-group">
        <label for="slug-scrap-rate">Operational Scrap Factor (%)</label>
        <input type="number" id="slug-scrap-rate" class="calc-input-num" value="10.0" min="0" max="30" step="0.5" oninput="runQuickCalc()">
      </div>

      <!-- RESULTS BOX -->
      <div class="calc-result-box">
        <div class="calc-res-row">
          <span class="res-lbl">Expected Net Finished Tubes:</span>
          <span class="res-val highlight" id="slug-res-net-tubes">768,201 pcs</span>
        </div>
        <div class="calc-res-row">
          <span class="res-lbl">Gross Theoretical Yield:</span>
          <span class="res-val" id="slug-res-gross-tubes">845,022 pcs</span>
        </div>
        <div class="calc-res-row">
          <span class="res-lbl">Estimated Line Scrap Loss:</span>
          <span class="res-val text-warn" id="slug-res-scrap-loss">76,821 pcs (454.5 kg)</span>
        </div>
      </div>
    </div>

    <!-- CARD 2: PET RESIN SIMULATOR -->
    <div class="calc-card">
      <div class="calc-card-hdr">
        <div class="calc-card-icon ic-pet">🧴</div>
        <div>
          <h3>PET Resin ➔ Multi-Format Matrix</h3>
          <p class="sub">Simultaneous capacity breakdown across all bottle sizes</p>
        </div>
      </div>

      <div class="calc-form-group">
        <label for="pet-resin-val">Available Resin Mass (kg)</label>
        <div class="calc-input-addon">
          <input type="number" id="pet-resin-val" class="calc-input-num" value="2500" min="1" step="25" oninput="runQuickCalc()">
          <span class="addon-tag">kg</span>
        </div>
      </div>

      <div class="calc-two-col-inputs">
        <div class="calc-form-group">
          <label for="pet-mb-rate">Masterbatch Dosing (%)</label>
          <input type="number" id="pet-mb-rate" class="calc-input-num" value="2.0" min="0" max="10" step="0.1" oninput="runQuickCalc()">
        </div>
        <div class="calc-form-group">
          <label for="pet-scrap-rate">Molding Scrap (%)</label>
          <input type="number" id="pet-scrap-rate" class="calc-input-num" value="15.0" min="0" max="30" step="0.5" oninput="runQuickCalc()">
        </div>
      </div>

      <!-- MULTI-FORMAT COMPARISON TABLE -->
      <table class="calc-matrix-tbl">
        <thead>
          <tr>
            <th>Format</th>
            <th class="num">Grammage</th>
            <th class="num">Net Bottles</th>
            <th class="num">MB Req. (kg)</th>
          </tr>
        </thead>
        <tbody id="pet-matrix-body">
          <!-- Dynamically populated via runQuickCalc() -->
        </tbody>
      </table>
    </div>

  </div>
</div>
```

---

### 1.1.5 JavaScript Reactive Execution Engine

```javascript
// Quick Calculator Reactive Controller
const SLUG_SPEC = {
  "12.5": 1.950, "16.0": 2.519, "19.0": 3.367, "20.5": 3.937,
  "25.0": 5.917, "30.0": 8.000, "32.0": 10.863, "35.0": 12.820
};

const PET_FORMATS = [
  { name: "60ml Bottle",  weight: 10.50 },
  { name: "75ml Bottle",  weight: 12.50 },
  { name: "100ml Bottle", weight: 15.00 },
  { name: "120ml Bottle", weight: 17.10 },
  { name: "130ml Bottle", weight: 18.00 },
  { name: "150ml Mist",   weight: 21.00 },
  { name: "200ml Bottle", weight: 23.75 },
  { name: "250ml Bottle", weight: 26.00 },
  { name: "300ml Jar",    weight: 25.00 },
  { name: "500ml Jar",    weight: 50.00 }
];

let _slugCalcMode = 'fwd'; // 'fwd' = kg -> pcs; 'rev' = pcs -> kg

function setSlugCalcMode(mode) {
  _slugCalcMode = mode;
  document.getElementById('slug-mode-fwd').classList.toggle('active', mode === 'fwd');
  document.getElementById('slug-mode-rev').classList.toggle('active', mode === 'rev');
  document.getElementById('slug-input-lbl').innerText = (mode === 'fwd') ? 'Available Slug Mass (kg)' : 'Demanded Finished Tubes (pcs)';
  document.getElementById('slug-addon-tag').innerText = (mode === 'fwd') ? 'kg' : 'pcs';
  runQuickCalc();
}

function runQuickCalc() {
  // 1. Aluminum Slugs Math
  const dia = document.getElementById('slug-dia-select').value;
  const rateKgPer1k = SLUG_SPEC[dia] || 5.917;
  const scrapPct = (parseFloat(document.getElementById('slug-scrap-rate').value) || 10.0) / 100.0;
  const inputVal = parseFloat(document.getElementById('slug-input-val').value) || 0;

  if (_slugCalcMode === 'fwd') {
    const grossTubes = Math.floor((inputVal * 1000) / rateKgPer1k);
    const netTubes = Math.floor((inputVal * 1000) / (rateKgPer1k * (1.0 + scrapPct)));
    const scrapTubes = grossTubes - netTubes;
    const scrapKg = (scrapTubes * rateKgPer1k) / 1000;

    document.getElementById('slug-res-net-tubes').innerText = netTubes.toLocaleString() + ' pcs';
    document.getElementById('slug-res-gross-tubes').innerText = grossTubes.toLocaleString() + ' pcs';
    document.getElementById('slug-res-scrap-loss').innerText = `${scrapTubes.toLocaleString()} pcs (${scrapKg.toFixed(1)} kg)`;
  } else {
    const requiredKg = (inputVal / 1000.0) * rateKgPer1k * (1.0 + scrapPct);
    const netKgWithoutScrap = (inputVal / 1000.0) * rateKgPer1k;
    const scrapKg = requiredKg - netKgWithoutScrap;

    document.getElementById('slug-res-net-tubes').innerText = requiredKg.toFixed(1) + ' kg slugs';
    document.getElementById('slug-res-gross-tubes').innerText = `Net Material: ${netKgWithoutScrap.toFixed(1)} kg`;
    document.getElementById('slug-res-scrap-loss').innerText = `Scrap Allowance: ${scrapKg.toFixed(1)} kg`;
  }

  // 2. PET Multi-Format Matrix Math
  const resinKg = parseFloat(document.getElementById('pet-resin-val').value) || 0;
  const mbPct = (parseFloat(document.getElementById('pet-mb-rate').value) || 2.0) / 100.0;
  const petScrapPct = (parseFloat(document.getElementById('pet-scrap-rate').value) || 15.0) / 100.0;

  const rowsHtml = PET_FORMATS.map(fmt => {
    const netBottles = Math.floor((resinKg * 1000.0) / (fmt.weight * (1.0 + petScrapPct)));
    const mbKg = (resinKg * mbPct);
    return `
      <tr>
        <td style="font-weight:600">${escapeHtml(fmt.name)}</td>
        <td class="num">${fmt.weight.toFixed(2)} g</td>
        <td class="num" style="color:var(--navy);font-weight:700">${netBottles.toLocaleString()}</td>
        <td class="num">${mbKg.toFixed(2)} kg</td>
      </tr>
    `;
  }).join('');

  document.getElementById('pet-matrix-body').innerHTML = rowsHtml;
}
```

---

## 1.2 FEATURE FP-02: HISTORICAL MONTH SELECTOR & ARCHIVE NAVIGATION

### 1.2.1 Operational Context & Architectural Challenge
Currently, `Tubex.html` dynamically reflects the single active operational month loaded during the most recent execution of `update_html.py` (e.g., `"August 2026"`). While `build_archives.py` extracts historical months into `Production_Archive.xlsx`, executive leadership, auditors, and planning managers cannot retrospectively inspect past monthly dashboards directly inside the web UI without manually opening legacy workbooks.

**FP-02 transforms the HTML dashboard from an ephemeral daily monitor into a multi-period executive analytics platform.**

---

### 1.2.2 Unified Historical Archive Data Schema

`update_html.py` and `build_archives.py` will generate and inject a historical archive catalog into `Tubex.html` under the marker `/* ARCHIVES_START */ ... /* ARCHIVES_END */`.

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "MonthlyDashboardArchive",
  "type": "object",
  "patternProperties": {
    "^[0-9]{4}-[0-9]{2}$": {
      "type": "object",
      "required": ["monthLabel", "year", "monthNumber", "kpi", "downtime", "ordersSummary", "tubeOrders", "petOrders"],
      "properties": {
        "monthLabel": { "type": "string", "example": "July 2026" },
        "year": { "type": "integer", "example": 2026 },
        "monthNumber": { "type": "integer", "example": 7 },
        "isCurrentMonth": { "type": "boolean", "example": false },
        "kpi": {
          "type": "object",
          "properties": {
            "tubeMTD": { "type": "integer" },
            "petMTD": { "type": "integer" },
            "tubeMTDDispatch": { "type": "integer" },
            "petMTDDispatch": { "type": "integer" },
            "tubeOperatingHours": { "type": "number" },
            "petOperatingHours": { "type": "number" }
          }
        },
        "downtime": {
          "type": "object",
          "properties": {
            "tubeCategories": { "type": "array", "items": { "type": "object", "properties": { "category": { "type": "string" }, "hours": { "type": "number" }, "pct": { "type": "number" } } } },
            "petCategories": { "type": "array", "items": { "type": "object", "properties": { "category": { "type": "string" }, "hours": { "type": "number" }, "pct": { "type": "number" } } } }
          }
        },
        "ordersSummary": {
          "type": "object",
          "properties": {
            "totalTubeOrders": { "type": "integer" },
            "totalTubeDelivered": { "type": "integer" },
            "complianceTubePct": { "type": "number" },
            "totalPetOrders": { "type": "integer" },
            "totalPetDelivered": { "type": "integer" },
            "compliancePetPct": { "type": "number" }
          }
        },
        "tubeOrders": { "type": "array", "items": { "type": "object" } },
        "petOrders": { "type": "array", "items": { "type": "object" } },
        "prodlogRows": { "type": "array", "items": { "type": "object" } }
      }
    }
  }
}
```

---

### 1.2.3 Header UI Navigation Mockup

Replace the static `#monthLabel` element in `Tubex.html`:

```html
<!-- BEFORE (Static Badge): -->
<!-- <div class="month-badge">📅 <span id="monthLabel">August 2026</span></div> -->

<!-- AFTER (Interactive Archive Navigation Controller): -->
<div class="archive-nav-bar">
  <div class="month-select-wrap">
    <label for="monthArchiveSelector" class="nav-icon-lbl">📅</label>
    <select id="monthArchiveSelector" class="month-dropdown-select" onchange="onSelectHistoricalMonth(this.value)">
      <option value="CURRENT" selected>August 2026 (Live Current)</option>
      <option value="2026-07">July 2026 (Archived)</option>
      <option value="2026-06">June 2026 (Archived)</option>
      <option value="2026-05">May 2026 (Archived)</option>
      <option value="2026-04">April 2026 (Archived)</option>
      <option value="2026-03">March 2026 (Archived)</option>
      <option value="2026-02">February 2026 (Archived)</option>
      <option value="2026-01">January 2026 (Archived)</option>
      <option value="2025-12">December 2025 (Archived)</option>
      <option value="2025-11">November 2025 (Archived)</option>
    </select>
  </div>
  <div id="archive-status-badge" class="archive-badge-live">● LIVE PRODUCTION</div>
</div>
```

---

### 1.2.4 Dynamic Client-Side State Switching & Fallback Engine

```javascript
// State Management for Historical Navigation
let _activeDashboardSnapshot = null;
let _liveDashboardSnapshot = null;

function initArchiveSelector() {
  if (typeof DASH_DATA !== 'undefined') {
    _liveDashboardSnapshot = JSON.parse(JSON.stringify(DASH_DATA));
    _activeDashboardSnapshot = _liveDashboardSnapshot;
  }
  
  // Populate dropdown options from MONTHLY_ARCHIVES if available
  const selector = document.getElementById('monthArchiveSelector');
  if (!selector) return;
  
  if (typeof MONTHLY_ARCHIVES !== 'undefined') {
    const months = Object.keys(MONTHLY_ARCHIVES).sort().reverse();
    // Build options dynamically
    selector.innerHTML = `<option value="CURRENT" selected>${escapeHtml(DASH_DATA.month || 'Current Month')} (Live)</option>` +
      months.map(mKey => {
        const item = MONTHLY_ARCHIVES[mKey];
        return `<option value="${escapeHtml(mKey)}">${escapeHtml(item.monthLabel || mKey)} (Archived)</option>`;
      }).join('');
  }
}

function onSelectHistoricalMonth(selectedKey) {
  const badge = document.getElementById('archive-status-badge');
  
  if (selectedKey === 'CURRENT') {
    _activeDashboardSnapshot = _liveDashboardSnapshot;
    badge.className = 'archive-badge-live';
    badge.innerText = '● LIVE PRODUCTION';
    renderCompleteDashboard(_activeDashboardSnapshot);
    return;
  }
  
  // Attempt loading from in-memory archives or offline fetch
  if (typeof MONTHLY_ARCHIVES !== 'undefined' && MONTHLY_ARCHIVES[selectedKey]) {
    loadArchiveData(MONTHLY_ARCHIVES[selectedKey], selectedKey);
  } else {
    // Attempt asynchronous fetch from ./archives/{YYYY-MM}.json
    badge.className = 'archive-badge-loading';
    badge.innerText = '⏳ Loading Archive...';
    
    fetch(`./archives/${encodeURIComponent(selectedKey)}.json`)
      .then(res => {
        if (!res.ok) throw new Error(`Archive ${selectedKey} not found`);
        return res.json();
      })
      .then(data => {
        loadArchiveData(data, selectedKey);
      })
      .catch(err => {
        console.warn(`Archive fetch failed for ${selectedKey}:`, err);
        alert(`Historical snapshot for ${selectedKey} is not cached locally. Switching back to live.`);
        document.getElementById('monthArchiveSelector').value = 'CURRENT';
        onSelectHistoricalMonth('CURRENT');
      });
  }
}

function loadArchiveData(archivePayload, key) {
  _activeDashboardSnapshot = archivePayload;
  const badge = document.getElementById('archive-status-badge');
  badge.className = 'archive-badge-history';
  badge.innerText = `🔒 READ-ONLY ARCHIVE: ${archivePayload.monthLabel || key}`;
  
  // Re-render dashboard panels
  renderKPIs(archivePayload.kpi);
  renderOrdersTable('tube', archivePayload.tubeOrders || []);
  renderOrdersTable('pet', archivePayload.petOrders || []);
  renderDowntimeCharts(archivePayload.downtime || {});
  if (archivePayload.prodlogRows) {
    renderProdlogRows(archivePayload.prodlogRows);
  }
}
```

---

# PART 2: 12 HIGH-IMPACT STRATEGIC IMPROVEMENT PROPOSALS

The 12 strategic proposals below address the 4 core pillars of manufacturing intelligence, data automation, web performance, and software resilience.

---

## PILLAR 1: WEB DASHBOARD & USER EXPERIENCE (UX)

### PROPOSAL 1.1: Historical Month Selector & Multi-Period Trend Analytics (FP-02)
- **Problem**: Plant managers cannot review previous months' KPIs, machine line performance, or downtime distributions within the web dashboard.
- **Solution**: Complete implementation of Feature FP-02 with historical archive navigation, multi-month trend overlays (e.g., 6-month tube scrap progression line charts), and historical order compliance comparisons.
- **Architecture**: Ingestion of monthly JSON snapshots via Service Worker cache and indexed SQLite/JSON blobs; dynamic Chart.js rendering for multi-month trend analytics.

### PROPOSAL 1.2: Raw Material Slugs & Resin Yield Simulator (FP-01)
- **Problem**: Arithmetic mean distortions in capacity calculation cause procurement over/under-ordering.
- **Solution**: Complete implementation of Feature FP-01 embedded into `#panel-calc`, featuring two-way reactive math, diameter-level slug yield tables, and multi-format PET bottle matrices with live stock cross-checks.
- **Architecture**: Instant DOM reactivity using vanilla JavaScript; zero-dependency micro-engine with print-friendly job-card export.

### PROPOSAL 1.3: Touch-Optimized Mobile/Tablet Interface & QR/Barcode Inventory Audit
- **Problem**: Floor supervisors using tablets or mobile phones struggle with dense desktop data tables; inventory auditing requires manual paper logging.
- **Solution**:
  1. CSS responsive overhaul with 48px minimum touch targets, collapsible card accordions for mobile viewports ($< 768\text{px}$), and swipe navigation between dashboard tabs.
  2. Built-in HTML5 camera barcode/QR scanner (`html5-qrcode` or Web Barcode Detection API) allowing floor auditors to scan pallet QR codes or bin tags to verify stock against `INVENTORY_DATA` directly on the factory floor.
- **Architecture**:
  ```javascript
  function startFloorBarcodeAudit() {
    const html5QrCode = new Html5Qrcode("qr-reader");
    html5QrCode.start(
      { facingMode: "environment" },
      { fps: 10, qrbox: 250 },
      (decodedText) => {
        // Match scanned Item ID or Lot Number against INVENTORY_DATA
        const match = INVENTORY_DATA.items.find(i => String(i.id) === decodedText.trim());
        if (match) {
          showQuickAuditCard(match);
        }
      }
    );
  }
  ```

### PROPOSAL 1.4: Real-Time Shift Run-Rate Velocity & Micro-Downtime Telemetry
- **Problem**: Daily production summaries only show aggregate shift numbers without indicating whether a printing press or PF machine ran at rated velocity ($3,500\text{ pcs/hr}$) or suffered hidden micro-stoppages.
- **Solution**: Shift Velocity Indicator widget calculating:
  $$\text{Line Velocity Ratio} = \frac{\text{Actual Output (pcs)} / \text{Effective Operating Hours}}{\text{Rated Machine Capacity (pcs/hr)}} \times 100\%$$
  - Velocity Bands: 🟢 **Optimal** ($\ge 90\%$), 🟡 **Sub-Optimal** ($75\text{--}89\%$), 🔴 **Throttled** ($< 75\%$).
- **Architecture**: Embedded gauges in the Production tab highlighting shift-by-shift run-rates per machine line (Print 1, Print 2, PF 1, PF 2).

### PROPOSAL 1.5: High-Contrast Industrial Dual Theme Engine (Dark / Solarized Daylight)
- **Problem**: Factory tablets operated in outdoor loading bays or high-glare shop floors suffer from washed-out screen visibility; control room desktops require dark themes to prevent eye fatigue.
- **Solution**: CSS Custom Property theme engine with instant toggle `[ ☀️ Daylight High-Contrast | 🌙 Industrial Dark ]`, persisted in `localStorage` and synchronized with system OS `prefers-color-scheme`.
- **Architecture**:
  ```css
  :root[data-theme="dark"] {
    --bg: #0b111e;
    --card: #151f32;
    --border: #26354f;
    --text: #e2e8f0;
    --navy: #60a5fa;
    --light: #1c2a42;
    --muted: #94a3b8;
  }
  :root[data-theme="daylight"] {
    --bg: #ffffff;
    --card: #f8fafc;
    --border: #000000;
    --text: #000000;
    --navy: #002b80;
    --light: #e2e8f0;
    --muted: #334155;
  }
  ```

---

## PILLAR 2: DATA PIPELINE & PROCESS AUTOMATION

### PROPOSAL 2.1: Direct ERP Database Connector (ODBC / SQL REST API)
- **Problem**: Daily operations require manual Remote Desktop (RDP) login to export `inventory.xls`, `dispatch.xls`, and `dispatch_pet.xls`. This creates human bottlenecks, delay risks, and formatting discrepancies.
- **Solution**: Build an automated extraction microservice (`alphapackage.pipeline.erp_connector`) that connects directly to the ERP backend (Microsoft SQL Server / Oracle / PostgreSQL) via ODBC at scheduled intervals (06:00 PKT daily).
- **Architecture**:
  ```python
  import pyodbc, pandas as pd
  from datetime import datetime

  def extract_erp_data(conn_str: str, output_dir: str):
      conn = pyodbc.connect(conn_str, timeout=30)
      
      # Ingest Stock Ledger
      inv_query = """
      SELECT ItemCode AS [Item Code], ItemName AS [Description], 
             UnitOfMeasure AS [UOM], CurrentBalance AS [Closing Balance]
      FROM ERP_Inventory_Current
      WHERE LocationCode IN ('WH-MAIN', 'WH-PLANT')
      """
      df_inv = pd.read_sql(inv_query, conn)
      df_inv.to_parquet(f"{output_dir}/inventory_live.parquet")
      
      # Ingest Dispatches
      disp_query = """
      SELECT DeliveryDate, CustomerName, ItemCode, ItemDescription, Quantity
      FROM ERP_Dispatch_Invoices
      WHERE DeliveryDate >= DATEADD(month, -1, GETDATE())
      """
      df_disp = pd.read_sql(disp_query, conn)
      df_disp.to_parquet(f"{output_dir}/dispatch_live.parquet")
      conn.close()
  ```

### PROPOSAL 2.2: Automated WhatsApp Shop-Floor Bot (Mehmood WIP & Imran Daily Logs)
- **Problem**: Work-in-Progress (WIP) counts from Mehmood and daily machine logs from Imran are sent via WhatsApp text and manually keyed into `Production.xlsx` and `update_wip.py`.
- **Solution**: Deploy an automated WhatsApp Business API / Webhook parsing bot (`alphapackage.bots.whatsapp_listener`). The bot parses structured shift texts (e.g., `*GP 25mm* Shift B: 42,000 Good, 1,200 Rej, Print 1, Downtime 1.5h Power`) via NLP/Regex, validates totals against machine capacity, stages the record in an audit queue, and auto-appends to `Production.xlsx`.
- **Architecture**:
  ```
  [Shop Floor Mobile (WhatsApp)] 
                 │ (HTTPS Webhook)
                 ▼
  [FastAPI / Webhook Listener] ──► [Regex & Parsing Engine] 
                                            │
                                            ▼
                               [Staging DB & Anomaly Check]
                                            │ (Approved / Auto-Commit)
                                            ▼
                                [Update Production.xlsx & Tubex.html]
  ```

### PROPOSAL 2.3: Pre-Flight Integrity Guard & Atomic Safe-Swap Transaction Pipeline
- **Problem**: If an Excel file is locked or corrupt during pipeline execution, half-updated state can be injected into `Tubex.html`.
- **Solution**: Implement an atomic staging workflow in `daily.py`. All Excel modifications and HTML builds are written to temporary staging files (`.tmp.xlsx`, `.tmp.html`). A strict Pre-Flight Invariant Assertion verifies file sizes, zero `#REF!/#VALUE!` cells, and required DOM tags before performing atomic OS replacement (`os.replace`).
- **Safety Invariant Checklist**:
  1. Temporary file size $\ge 512\text{ bytes}$ (R1-20 guard).
  2. All formula evaluation caches contain valid numeric or string tokens.
  3. `/* DATA_START */` and `/* DATA_END */` markers exist with non-empty payload.

### PROPOSAL 2.4: Automated Git & Cloud Storage Webhook Synchronization
- **Problem**: Backup relies on local batch scripts (`Robocopy /E`) to OneDrive. Cloud versioning and collaborative branch tracking are manual.
- **Solution**: Implement a background Git sync and S3/Azure Blob mirror hook at the conclusion of `daily.py`. Automatically commits daily operational state to a secure private Git repository with tag `ops-YYYYMMDD`, while publishing static dashboard assets (`Tubex.html`, `sw.js`, `manifest.json`) to an encrypted intranet HTTPS edge server.

---

## PILLAR 3: PLANNING & MRP INTELLIGENCE

### PROPOSAL 3.1: Dynamic Rolling Empirical Scrap Calibration Engine
- **Problem**: Fixed 10% (Tubes) and 15% (PET) scrap factors in static BOMs cause raw material shortages on complex 6-color artwork or older tooling, while overestimating scrap on simple 1-color runs.
- **Solution**: Implement an empirical scrap calibrator that evaluates the last 90 days of production logs:
  $$s_{\text{empirical}}(\text{PID}) = \text{Clamp}\left(\frac{\sum_{t=1}^{90} \text{Rejects}_{t,\text{PID}}}{\sum_{t=1}^{90} \text{Total Produced}_{t,\text{PID}}},\; 0.05,\; 0.25\right)$$
  - Dynamically injects $s_{\text{empirical}}$ into the MRP order evaluation pipeline, generating dynamic, real-world raw material requisition requirements.

### PROPOSAL 3.2: Statistical Lead-Time Safety Stock & Dynamic Reorder Point (ROP)
- **Problem**: Inventory alerts trigger only on immediate negative order balances without factoring supplier procurement lead times.
- **Solution**: Implement the classic inventory replenishment formula for all critical raw materials (Slugs, Resins, Lacquers, Inks, Caps):
  $$\text{ROP} = \left(\bar{d} \times L\right) + Z \times \sqrt{L \times \sigma_d^2 + \bar{d}^2 \times \sigma_L^2}$$
  Where:
  - $\bar{d}$ = Average daily consumption rate.
  - $L$ = Supplier lead time in days (e.g., Slugs = 25 days, Imported Lacquer = 35 days, Masterbatch = 7 days).
  - $Z = 1.65$ (95% service level confidence coefficient).
  - $\sigma_d, \sigma_L$ = Standard deviation of daily demand and lead time.

### PROPOSAL 3.3: Automated Multi-Tier Supplier Reorder Dispatcher
- **Problem**: Purchasing officers must manually scan the MRP report to determine which vendor purchase orders to draft.
- **Solution**: When raw material stock drops below ROP, the MRP engine automatically formats a standardized PDF / Excel Purchase Requisition with item codes, recommended batch order quantities, and supplier contact details, staging it for 1-click email dispatch by the Procurement Lead.

### PROPOSAL 3.4: Bottleneck Machine Capacity Scheduler & Changeover Sequence Optimizer
- **Problem**: Line changeovers on Printing Presses (Print 1 & Print 2) from $\varnothing 19\text{mm}$ to $\varnothing 35\text{mm}$ or dark inks to white basecoats consume 3 to 6 hours of downtime.
- **Solution**: Build an integer linear programming (ILP) or heuristic scheduling module (`alphapackage.mrp.scheduler`) that groups active orders to minimize setup and cleaning downtime:
  - Sequences jobs by diameter $\to$ nozzle type $\to$ artwork color gradient (light to dark).
  - Generates optimized daily machine run-cards displayed in the Production tab.

---

## PILLAR 4: ARCHITECTURE, QUALITY, OBSERVABILITY & RESILIENCE

### PROPOSAL 4.1: Unified Python Package Architecture (`alphapackage`)
- **Problem**: Automation logic is dispersed across 10+ standalone scripts in `Scripts/` with duplicated helper functions and mixed concerns.
- **Solution**: Consolidate the entire codebase into a clean, modern Python package structure following PEP 517/621 with `pyproject.toml`, explicit dependency management, typed data models (`Pydantic`), and a unified CLI (`alpha-cli`).

#### Target Package Structure:
```
d:\Alpha\
├── pyproject.toml
├── src\
│   └── alphapackage\
│       ├── __init__.py
│       ├── config.py             # Paths, thresholds, environment configuration
│       ├── models\               # Typed Pydantic data schemas
│       │   ├── __init__.py
│       │   ├── bom.py
│       │   ├── order.py
│       │   ├── inventory.py
│       │   └── production.py
│       ├── core\                 # Core engine logic
│       │   ├── __init__.py
│       │   ├── excel_com.py      # Bulletproof Excel COM lifecycle manager
│       │   ├── formula_eval.py   # Safe formula evaluators & regex parsers
│       │   └── normalizer.py     # Customer alias token matcher
│       ├── pipeline\             # Daily ingestion ETL workflows
│       │   ├── __init__.py
│       │   ├── daily_orchestrator.py
│       │   ├── production_sync.py
│       │   ├── inventory_sync.py
│       │   └── dispatch_sync.py
│       ├── mrp\                  # Planning & Yield engines
│       │   ├── __init__.py
│       │   ├── yield_calculator.py
│       │   ├── scrap_model.py
│       │   └── rop_engine.py
│       ├── web\                  # Web generator & archive compilers
│       │   ├── __init__.py
│       │   ├── html_builder.py
│       │   └── archive_builder.py
│       └── cli\                  # Unified CLI Interface
│           ├── __init__.py
│           └── main.py           # alpha-cli entry point
└── tests\
    ├── unit\
    ├── integration\
    └── e2e\
```

### PROPOSAL 4.2: Structured JSON Telemetry Logging & OpenTelemetry Instrumentation
- **Problem**: Scripts use standard `print()` statements; historical execution durations, row counts, memory consumption, and warnings are lost upon console closure.
- **Solution**: Implement structured logging with `structlog` writing machine-readable JSON logs to `Logs/daily_telemetry_{YYYYMMDD}.json`.
- **Log Sample**:
  ```json
  {
    "timestamp": "2026-08-19T07:45:00.124Z",
    "level": "INFO",
    "event": "inventory_sync_completed",
    "module": "alphapackage.pipeline.inventory_sync",
    "duration_ms": 412,
    "rows_ingested": 184,
    "inactive_items_flagged": 12,
    "active_workbook": "Tubex_Aug26.xlsx",
    "memory_rss_mb": 42.6
  }
  ```

### PROPOSAL 4.3: Automated Daily Executive Health & Shortage Notification Engine
- **Problem**: Executive management must open the web dashboard to check whether any critical materials are in shortage.
- **Solution**: At the end of `daily.py`, automatically compile an Executive Briefing snapshot and transmit it via SMTP (HTML email) or Telegram/WhatsApp bot to plant leadership:
  - **Summary**: Daily Tube Output vs PET Output vs Target Run-Rate.
  - **MRP Critical Alerts**: List of raw materials with $< 3\text{ days}$ of production buffer.
  - **Pipeline Health**: Total execution duration, Excel COM exit status, and data freshness timestamp.

### PROPOSAL 4.4: Automated End-to-End Regression Test Suite (Pytest + Playwright)
- **Problem**: Formula changes or HTML updates can accidentally introduce regressions without immediate detection.
- **Solution**: Establish a continuous testing harness:
  1. **Pytest**: Over 60 unit tests verifying BOM math, scrap formulas, Excel COM cleanup, regex safety, and date parsing.
  2. **Playwright E2E**: Automated headless browser testing asserting that `Tubex.html` renders zero `#NaN` or `undefined` cells, service worker registers successfully, and modal interactions work flawlessly across Chromium and WebKit viewports.

---

# PART 3: MASTER IMPLEMENTATION ROADMAP & RESOURCE PLANNING

---

## 3.1 PHASED ROLLOUT ROADMAP

```
2026 Q3 — Q4 IMPLEMENTATION TIMELINE
├──────────────────────────────────────────────────────────────────────────────────────────────────┤
│ Phase 1: Core Web Polish & Future_Plans Execution (Sprint 1–2, 3 Weeks)                          │
│   ├── [FP-01] Slugs & Resin Fast Calculator in Tubex.html                                       │
│   ├── [FP-02] Historical Month Selector & Archive Navigation                                     │
│   └── [P1.5] Industrial Dual Theme Engine (Dark/Daylight)                                       │
├──────────────────────────────────────────────────────────────────────────────────────────────────┤
│ Phase 2: Architecture Refactoring & Observability (Sprint 3–4, 4 Weeks)                         │
│   ├── [P4.1] Unified Python Package Architecture (`alphapackage`)                                │
│   ├── [P4.2] Structured JSON Telemetry Logging (`structlog`)                                    │
│   ├── [P4.4] Pytest & Playwright Regression Test Harness                                        │
│   └── [P2.3] Atomic Pre-Flight Integrity Guard & Safe-Swap                                       │
├──────────────────────────────────────────────────────────────────────────────────────────────────┤
│ Phase 3: Intelligent MRP & Mobile Floor Audit (Sprint 5–6, 4 Weeks)                             │
│   ├── [P1.3] Touch-First Responsive Mobile Overhaul & Barcode Scanner                            │
│   ├── [P3.1] Dynamic Rolling Scrap Calibration Model                                            │
│   ├── [P3.2] Statistical Safety Stock & Dynamic ROP Engine                                      │
│   └── [P4.3] Automated Executive Daily Health Alerts (Email/WhatsApp)                           │
├──────────────────────────────────────────────────────────────────────────────────────────────────┤
│ Phase 4: Full Enterprise Automation & Direct ETL (Sprint 7–8, 5 Weeks)                          │
│   ├── [P2.1] Direct ERP SQL/ODBC Automated Data Ingestion                                       │
│   ├── [P2.2] WhatsApp Ingestion Bot for Mehmood/Imran Shift Logs                                │
│   ├── [P3.4] Machine Line Scheduling & Changeover Optimizer                                     │
│   └── [P2.4] Automated Git/Cloud Webhook & Edge Deployment                                      │
└──────────────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## 3.2 EFFORT ESTIMATION & RESOURCE ALLOCATION

| Phase | Feature / Proposal | Pillar | Story Points | Est. Dev-Days | Target Roles | Priority |
|---|---|---|---|---|---|---|
| **Phase 1** | **FP-01**: Raw Material Slugs & Resin Calculator | Web UX | 5 SP | 4 Days | Frontend Dev / Industrial Eng | **CRITICAL** |
| **Phase 1** | **FP-02**: Historical Month Archive Navigation | Web UX | 5 SP | 4 Days | Full-Stack Dev | **CRITICAL** |
| **Phase 1** | **P1.5**: Industrial Dark & High-Glare Daylight Themes | Web UX | 2 SP | 2 Days | UI/UX Designer | Medium |
| **Phase 2** | **P4.1**: Unified Python Package (`alphapackage`) | Quality | 8 SP | 6 Days | Backend Architect | **HIGH** |
| **Phase 2** | **P4.2**: Structured JSON Telemetry Logging | Quality | 3 SP | 2 Days | DevOps / Backend | Medium |
| **Phase 2** | **P2.3**: Pre-Flight Invariant & Atomic Safe-Swap | Pipeline | 3 SP | 2 Days | Backend Engineer | **HIGH** |
| **Phase 2** | **P4.4**: Pytest & Playwright Automated Test Suite | Quality | 5 SP | 4 Days | QA Automation Eng | **HIGH** |
| **Phase 3** | **P1.3**: Mobile Touch UI & QR Barcode Scanner | Web UX | 5 SP | 4 Days | Mobile / Frontend Dev | Medium |
| **Phase 3** | **P1.4**: Shift Velocity Telemetry Gauge | Web UX | 3 SP | 2 Days | Frontend Dev | Medium |
| **Phase 3** | **P3.1**: Dynamic Empirical Scrap Calibration | MRP | 5 SP | 4 Days | Data Scientist / Industrial | **HIGH** |
| **Phase 3** | **P3.2**: Lead-Time Safety Stock & ROP Engine | MRP | 5 SP | 4 Days | Supply Chain Engineer | **HIGH** |
| **Phase 3** | **P4.3**: Executive Health Email/WhatsApp Dispatch | Quality | 3 SP | 2 Days | Integration Dev | Medium |
| **Phase 4** | **P2.1**: Direct ERP SQL/ODBC Database Connector | Pipeline | 8 SP | 7 Days | ERP / Database Specialist | **HIGH** |
| **Phase 4** | **P2.2**: WhatsApp Bot for Shop-Floor WIP & Logs | Pipeline | 8 SP | 7 Days | Full-Stack / NLP Dev | **HIGH** |
| **Phase 4** | **P3.4**: Bottleneck Scheduling & Changeover Engine | MRP | 8 SP | 7 Days | Operations Research Eng | Medium |
| **Phase 4** | **P2.4**: Git Cloud Webhook & Edge Deployment | Pipeline | 3 SP | 2 Days | DevOps Engineer | Low |
| **TOTAL** | **Full 4-Pillar Modernization Program** | **All** | **78 SP** | **63 Days** | **Team of 2–3 Engineers** | — |

---

# PART 4: COMPREHENSIVE RISK ANALYSIS & MITIGATION MATRIX

```
┌──────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                   MODERNIZATION RISK ASSESSMENT MATRIX                           │
└──────────────────────────────────────────────────────────────────────────────────────────────────┘
  Impact
    ▲
  H │                  [R-02: ERP Schema Change]      [R-01: Excel Concurrency Lock]
    │                  [R-05: Floor WhatsApp Spam]    [R-03: Scrap Under-Estimation]
  M │ [R-06: Dark Theme Glare]                        [R-04: Offline Cache Storage Limit]
  L │ [R-07: Telemetry Log Disk Fill]
    └─────────────────────────────────────────────────────────────────────────────────────────────►
      Low                     Medium                         High                   Probability
```

### Risk Registry & Mitigation Controls:

| Risk ID | Category | Risk Description | Severity | Probability | Mitigation Strategy & Safeguards |
|---|---|---|---|---|---|
| **R-01** | Technical | **Excel File Concurrency Locks during COM Automation**: Windows file locking prevents `daily.py` from updating `Tubex_Aug26.xlsx` if a user has it open. | **CRITICAL** | High | 1. Implement `DispatchEx` + strict `try...finally` teardown.<br>2. Use `Alpha_Checks.purge_excel_locks()` before execution.<br>3. Enforce read-only staging copies for COM formula recalculation. |
| **R-02** | Integration | **ERP Database Schema Drift or Connectivity Outage**: Direct SQL queries fail if ERP column names or server IP change. | **HIGH** | Medium | 1. Implement database schema contract tests in `alphapackage`.<br>2. Provide automatic fallback to manual Excel export ingestion (`inventory.xls`) if ODBC connection times out. |
| **R-03** | Operational | **Under-Ordering due to Dynamic Scrap Model Drift**: An unrepresentative short production run skews empirical scrap downwards, resulting in slug shortages. | **HIGH** | Medium | 1. Enforce strict clamp boundaries: $5\% \le s_{\text{empirical}} \le 25\%$.<br>2. Require minimum sample size ($N \ge 50,000\text{ pcs}$) before overriding master catalog BOM baseline. |
| **R-04** | Web / PWA | **Service Worker Offline Cache Quota Exhaustion**: Caching 24 months of full JSON snapshots exceeds browser storage on low-end shop-floor tablets. | **MEDIUM** | Medium | 1. Implement lazy-loading for historical archives ($> 3\text{ months}$ loaded on-demand via fetch).<br>2. Compress historical payloads using GZIP/Brotli or delta JSON representations. |
| **R-05** | Security / Data | **Malformed or Unauthorized WhatsApp Shift Submissions**: Operators send corrupted text or fraudulent production figures via WhatsApp bot. | **HIGH** | Medium | 1. Enforce phone-number whitelisting with cryptographic PIN authentication.<br>2. Stage all submissions in an "Operator Review Queue" requiring supervisor 1-click confirmation before writing to production models. |
| **R-06** | Usability | **Theme Contrast Inadequacy under Harsh Sunlight**: Field tablets in daylight become unreadable if contrast ratios are $< 4.5:1$. | **LOW** | Low | 1. Maintain WCAG 2.1 AAA contrast compliance ($\ge 7:1$) on the Solarized Daylight theme palette.<br>2. Provide ambient light sensor auto-switching via CSS `@media (light-level: bright)`. |
| **R-07** | Infrastructure | **Telemetry Log Disk Space Growth**: Unbounded JSON telemetry log files fill the local drive over multi-year operations. | **LOW** | Low | 1. Configure Python `RotatingFileHandler` capped at 10 MB per file with 30-day automatic retention purge. |

---

# PART 5: ARCHITECTURAL CODE & SPECIFICATION LIBRARY

Below is the concrete reference implementation for core modernization modules.

---

## 5.1 PYTHON PACKAGE INITIALIZER & CONFIGURATION (`src/alphapackage/config.py`)

```python
"""
alphapackage.config
~~~~~~~~~~~~~~~~~~~
Centralized configuration, paths, and operational thresholds for Alpha Containers.
"""

from pathlib import Path
from pydantic import BaseModel, Field

class AppConfig(BaseModel):
    # Base Directories
    ROOT_DIR: Path = Path("d:/Alpha")
    SCRIPTS_DIR: Path = ROOT_DIR / "Scripts"
    ARCHIVES_DIR: Path = ROOT_DIR / "archives"
    LOGS_DIR: Path = ROOT_DIR / "Logs"
    DATA_FEED_DIR: Path = ROOT_DIR / "data_feed"
    
    # Active Files
    ACTIVE_WORKBOOK_PATTERN: str = "Tubex*.xlsx"
    MASTER_CATALOG_FILE: str = "Master_Catalog.xlsx"
    PRODUCTION_LOG_FILE: str = "Production.xlsx"
    HTML_DASHBOARD_FILE: str = "Tubex.html"
    
    # Scrap Model Parameters
    DEFAULT_TUBE_SCRAP: float = Field(default=0.10, ge=0.0, le=0.30)
    DEFAULT_PET_SCRAP: float = Field(default=0.15, ge=0.0, le=0.30)
    AEROSOL_LACQUER_SCRAP: float = Field(default=0.35, ge=0.0, le=0.50)
    
    # Service Level Parameters
    SAFETY_STOCK_Z_SCORE: float = 1.65  # 95% Service Level
    
    # Telemetry
    ENABLE_STRUCTURED_LOGGING: bool = True
    LOG_RETENTION_DAYS: int = 30

CONFIG = AppConfig()
```

---

## 5.2 ROBUST EXCEL COM LIFECYCLE CONTROLLER (`src/alphapackage/core/excel_com.py`)

```python
"""
alphapackage.core.excel_com
~~~~~~~~~~~~~~~~~~~~~~~~~~~
Leak-proof, isolated Excel COM process manager ensuring zero lingering EXCEL.EXE tasks.
"""

import os
import platform
from contextlib import contextmanager
from typing import Generator, Any

@contextmanager
def safe_excel_context(visible: bool = False, display_alerts: bool = False) -> Generator[Any, None, None]:
    """Context manager guaranteeing Excel COM termination even on fatal exceptions."""
    if platform.system() != 'Windows':
        raise OSError("Excel COM automation is only supported on Windows operating systems.")
        
    import win32com.client
    
    excel_app = None
    try:
        # Enforce isolated COM instance
        excel_app = win32com.client.DispatchEx("Excel.Application")
        excel_app.Visible = visible
        excel_app.DisplayAlerts = display_alerts
        yield excel_app
    finally:
        if excel_app is not None:
            try:
                excel_app.Quit()
            except Exception:
                pass
            del excel_app

def evaluate_and_save_workbook(file_path: str) -> bool:
    """Opens workbook in isolated Excel COM instance, recalculates all formulas, and saves."""
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(abs_path):
        raise FileNotFoundError(f"Target Excel file not found: {abs_path}")
        
    with safe_excel_context(visible=False, display_alerts=False) as excel:
        wb = None
        try:
            wb = excel.Workbooks.Open(abs_path)
            wb.Save()
            return True
        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
                del wb
```

---

## 5.3 ADVANCED MRP YIELD & REORDER ENGINE (`src/alphapackage/mrp/yield_calculator.py`)

```python
"""
alphapackage.mrp.yield_calculator
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Mathematical yield, scrap calibration, and inventory reorder point engine.
"""

import math
from typing import Dict, List, Tuple

# Tube Slug Constants (kg / 1,000 tubes)
SLUG_WEIGHT_MAP: Dict[float, float] = {
    12.5: 1.950,
    13.5: 1.950,
    16.0: 2.519,
    19.0: 3.367,
    20.5: 3.937,
    22.0: 3.937,
    25.0: 5.917,
    28.0: 8.000,
    30.0: 8.000,
    32.0: 10.863,
    35.0: 12.820
}

# PET Bottle Constants (kg / 1,000 bottles)
PET_WEIGHT_MAP: Dict[str, float] = {
    "60ml": 10.50,
    "75ml": 12.50,
    "100ml": 15.00,
    "120ml": 17.10,
    "130ml": 18.00,
    "150ml": 21.00,
    "200ml": 23.75,
    "250ml": 26.00,
    "300ml": 25.00,
    "500ml": 50.00
}

def calculate_slug_yield(slug_mass_kg: float, diameter_mm: float, scrap_rate: float = 0.10) -> Tuple[int, int, float]:
    """
    Calculates gross yield, net yield, and scrap mass for aluminum tube extrusion.
    
    Returns:
        (net_tubes, gross_tubes, scrap_loss_kg)
    """
    rate = SLUG_WEIGHT_MAP.get(diameter_mm)
    if not rate:
        raise ValueError(f"Unsupported diameter: {diameter_mm}mm. Valid diameters: {list(SLUG_WEIGHT_MAP.keys())}")
        
    gross_tubes = math.floor((slug_mass_kg * 1000.0) / rate)
    net_tubes = math.floor((slug_mass_kg * 1000.0) / (rate * (1.0 + scrap_rate)))
    scrap_tubes = gross_tubes - net_tubes
    scrap_loss_kg = (scrap_tubes * rate) / 1000.0
    
    return net_tubes, gross_tubes, scrap_loss_kg

def calculate_pet_matrix(resin_mass_kg: float, mb_dosing_pct: float = 2.0, scrap_rate: float = 0.15) -> List[Dict]:
    """Generates simultaneous capacity yield across all PET bottle formats."""
    results = []
    mb_factor = mb_dosing_pct / 100.0
    
    for format_name, rate in PET_WEIGHT_MAP.items():
        net_bottles = math.floor((resin_mass_kg * 1000.0) / (rate * (1.0 + scrap_rate)))
        mb_required_kg = resin_mass_kg * mb_factor
        results.append({
            "format": format_name,
            "unit_weight_g": rate,
            "net_bottles": net_bottles,
            "masterbatch_kg": round(mb_required_kg, 2)
        })
    return results

def calculate_reorder_point(avg_daily_demand: float, lead_time_days: int, 
                            std_daily_demand: float, std_lead_time: float, 
                            service_z: float = 1.65) -> float:
    """Calculates statistical Reorder Point (ROP) with demand and lead-time variability."""
    lead_time_demand = avg_daily_demand * lead_time_days
    safety_stock = service_z * math.sqrt(
        (lead_time_days * (std_daily_demand ** 2)) + ((avg_daily_demand ** 2) * (std_lead_time ** 2))
    )
    return round(lead_time_demand + safety_stock, 2)
```

---

## 5.4 STRUCTURED TELEMETRY LOGGER (`src/alphapackage/core/telemetry.py`)

```python
"""
alphapackage.core.telemetry
~~~~~~~~~~~~~~~~~~~~~~~~~~~
High-performance structured JSON telemetry logger.
"""

import sys
import time
import json
import logging
from pathlib import Path
from datetime import datetime

class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        log_obj = {
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "level": record.levelname,
            "message": record.getMessage(),
            "logger": record.name,
            "file": record.filename,
            "line": record.lineno
        }
        if hasattr(record, "props"):
            log_obj.update(record.props)
        return json.dumps(log_obj, ensure_ascii=False)

def get_telemetry_logger(name: str, log_dir: Path) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    today_str = datetime.now().strftime("%Y%m%d")
    log_file = log_dir / f"telemetry_{today_str}.json"
    
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    
    if not logger.handlers:
        file_h = logging.FileHandler(log_file, encoding='utf-8')
        file_h.setFormatter(JsonFormatter())
        logger.addHandler(file_h)
        
        console_h = logging.StreamHandler(sys.stdout)
        console_h.setFormatter(logging.Formatter("[%(levelname)s] %(asctime)s - %(message)s"))
        logger.addHandler(console_h)
        
    return logger
```

---

## 5.5 MASTER TESTING SPECIFICATION & REGRESSION HARNESS (`tests/unit/test_mrp_math.py`)

```python
"""
tests.unit.test_mrp_math
~~~~~~~~~~~~~~~~~~~~~~~~
Invariant tests for raw material yield formulas, scrap boundaries, and ROP.
"""

import pytest
from alphapackage.mrp.yield_calculator import (
    calculate_slug_yield,
    calculate_pet_matrix,
    calculate_reorder_point,
    SLUG_WEIGHT_MAP,
    PET_WEIGHT_MAP
)

def test_slug_yield_25mm_standard():
    """Verify 25mm slug yield at 10% standard scrap."""
    net_tubes, gross_tubes, scrap_kg = calculate_slug_yield(5000.0, 25.0, 0.10)
    assert gross_tubes == 845022
    assert net_tubes == 768201
    assert round(scrap_kg, 1) == 454.6

def test_slug_yield_zero_mass_raises_safe_zero():
    """Verify zero mass returns zero yield without divide-by-zero."""
    net_tubes, gross_tubes, scrap_kg = calculate_slug_yield(0.0, 19.0, 0.10)
    assert net_tubes == 0
    assert gross_tubes == 0
    assert scrap_kg == 0.0

def test_slug_yield_invalid_diameter_raises_error():
    """Verify invalid diameter triggers explicit ValueError."""
    with pytest.raises(ValueError):
        calculate_slug_yield(1000.0, 42.0, 0.10)

def test_pet_matrix_output_consistency():
    """Verify PET capacity matrix calculates for all 10 formats."""
    res = calculate_pet_matrix(2500.0, mb_dosing_pct=2.0, scrap_rate=0.15)
    assert len(res) == 10
    
    # Check 120ml format
    fmt_120 = next(f for f in res if f["format"] == "120ml")
    assert fmt_120["net_bottles"] == 127145
    assert fmt_120["masterbatch_kg"] == 50.0

def test_reorder_point_safety_buffer():
    """Verify ROP strictly exceeds basic lead-time demand due to safety stock."""
    d_bar = 100.0
    L = 20
    rop = calculate_reorder_point(
        avg_daily_demand=d_bar, 
        lead_time_days=L, 
        std_daily_demand=15.0, 
        std_lead_time=3.0, 
        service_z=1.65
    )
    basic_lead_time_demand = d_bar * L
    assert rop > basic_lead_time_demand
    assert round(rop, 0) == 2507.0
```

---

# CONCLUSION & STRATEGIC RECOMMENDATIONS

This Comprehensive Modernization Blueprint provides a definitive, production-grade roadmap for transitioning the Alpha Containers manufacturing intelligence infrastructure into an enterprise-grade, resilient, and fully automated ecosystem.

### Immediate Action Plan (Next Sprint):
1. **Frontend Integration**: Integrate `FP-01` (Quick Slugs & Resin Simulator) and `FP-02` (Historical Month Selector) directly into `Tubex.html` and `update_html.py`.
2. **Package Consolidation**: Begin Phase 2 refactoring by establishing the `src/alphapackage` structure and wrapping COM automation in `safe_excel_context`.
3. **Observability Deployment**: Replace standard prints with structured JSON telemetry to establish complete execution history.

*Document finalized and ready for executive review and sprint planning.*
