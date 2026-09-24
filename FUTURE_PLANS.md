# Alpha Containers — Future Plans & Feature Specifications (FP)

> **Quick Invocation**: To implement any feature below, simply tell Antigravity / Gemini Flash:  
> **"Implement FP-01"**, **"Do FP-04"**, etc. The agent will read this specification and execute it directly.

---

## Active Feature Registry

| ID | Feature Name | Pillar | Target Files | Status | Priority |
|:---|:---|:---|:---|:---:|:---:|
| [**FP-01**](#fp-01-raw-material-yield--capacity-simulator-slugs--resin) | Raw Material Yield & Capacity Simulator | Web Dashboard & UX | `Tubex.html` (`#panel-calc`) | **Ready** | High / Quick Win |
| [**FP-04**](#fp-04-dark--light-theme-toggle--mobile-touch-navigation) | Dark / Light Theme & Touch Navigation | Web Dashboard & UX | `Tubex.html` (CSS & Touch JS) | **Ready** | Medium / Polish |
| [**FP-05**](#fp-05-one-click-executive-pdf--csv-report-exporter) | 1-Click Executive PDF & CSV Exporter | Web Dashboard & UX | `Tubex.html` (`@media print` & JS) | **Ready** | Medium / Quick Win |
| [**FP-08**](#fp-08-automated-daily-morning-executive-briefing) | WhatsApp Executive Briefing Generator | Pipeline Automation | `Scripts/daily.py` (Step 7) | **Ready** | High / Time Saver |
| [**FP-11**](#fp-11-smart-printing-line-changeover-scheduler) | Smart Printing Changeover Scheduler | Floor Scheduling | `Tubex.html` (Production Tab) | **Ready** | Strategic / Efficiency |

---

### FP-01: Raw Material Yield & Capacity Simulator (Slugs & Resin)

- **Target Files**: [Tubex.html](file:///d:/Alpha/Tubex.html) (inside the `🧮 Material Calc` tab / `#panel-calc`), [sw.js](file:///d:/Alpha/sw.js)
- **Objective**: Instant dual-mode interactive converter on the web dashboard to calculate how many tubes or bottles can be produced from available raw material stock (kg), or reverse-calculate required kg from piece targets.
- **Specification**:
  1. **Dual Mode Switcher**:
     - Render `#yieldSimulator` card at the top of `#panel-calc` with tab/toggle buttons: `[ 🧱 Aluminum Slugs (Tubes) | 🧴 PET Resin (Bottles) ]`.
  2. **Aluminum Slugs Converter**:
      - Plant Slug Weight Standards:
        - 12.5 / 13.5 mm: 1.950 g
        - 16.0 mm: 2.519 g
        - 19.0 mm: 3.367 g
        - 20.5 / 22.0 mm: 3.937 g
        - 25.0 mm: 5.917 g
        - 28.0 / 30.0 mm: 8.000 g
        - 32.0 mm: 10.863 g
        - 35.0 mm: 12.820 g
      - Forward Calculation:  
        `Net Tubes = Floor((Available kg * 1,000) / (Slug Weight (g) * (1 + Scrap Factor)))`  
        *(Default Scrap = 10%, adjustable with slider 5%–20%)*
      - Reverse Calculation:  
        `Required kg = Round((Target Pieces / 1,000) * Slug Weight (g) * (1 + Scrap Factor), 2)`
  3. **PET Resin Converter**:
     - Preform Grammages: 60ml (10.5g), 75ml (12.5g), 100ml (15.0g), 120ml (17.1g), 130ml (18.0g), 150ml (21.0g), 200ml (23.75g), 250ml (26.0g), 300ml (25.0g), 500ml (50.0g).
     - Standard PET Scrap = 15%, Masterbatch Dosing = 2.0%.
     - Input available PET resin (kg) -> renders comparison grid showing producible bottles and required Masterbatch (kg) across all standard formats.
  4. **Live Stock Auto-Fill**:
     - Provide a `"⚡ Load Current Store Stock"` button that pulls the actual physical balance directly from `DASH_DATA.inventory` for the selected diameter / resin.
  5. **Safety Constraints**:
     - 100% reactive client-side JavaScript. Zero Excel modifications or backend dependencies.

---

### FP-04: Dark / Light Theme Toggle & Mobile Touch Navigation

- **Target Files**: [Tubex.html](file:///d:/Alpha/Tubex.html) (CSS `:root`, header toolbar, touch JS event listeners), [sw.js](file:///d:/Alpha/sw.js)
- **Objective**: Provide a clean, high-contrast dark theme for night shifts and low-light factory floor inspections, and smooth mobile touch navigation.
- **Specification**:
  1. **Theme Engine**:
     - Implement `[data-theme="dark"]` CSS variables:
       - `--bg: #0d1526;`
       - `--card: #15203b;`
       - `--border: #233554;`
       - `--text: #e2e8f0;`
       - `--muted: #8c9cb8;`
       - `--navy: #1e293b;`
       - `--light: #1b2845;`
       - `--accent2: #f5c842;`
     - Header icon button `🌓` in `.header-top` to toggle theme.
     - Persist choice in `localStorage.setItem('tubex_theme', mode)` and sync on startup with system preference (`prefers-color-scheme`).
  2. **Mobile Touch Navigation**:
     - Attach `touchstart` and `touchend` listeners to the main container.
     - Gesture rule: horizontal swipe distance `>= 60px` and horizontal delta `>= 2 * vertical delta`.
     - Exclude touch events originating inside horizontal scroll wrappers (`.tbl-wrap`, `input`, `select`).
     - Swipe Left -> Switch to Next Tab; Swipe Right -> Switch to Previous Tab across:
       `[📊 Dashboard ⇄ 👥 Customer Report ⇄ 🧮 Material Calc ⇄ 🏭 Production ⇄ 📋 FG Stock ⇄ 📦 Inventory]`.
  3. **Safety Constraints**:
     - Completely responsive; does not alter click navigation on desktop.

---

### FP-05: One-Click Executive PDF & CSV Report Exporter

- **Target Files**: [Tubex.html](file:///d:/Alpha/Tubex.html) (`@media print` stylesheet, action buttons, client-side CSV downloaders), [sw.js](file:///d:/Alpha/sw.js)
- **Objective**: 1-click printable single-page executive summary PDF for management sharing, and instant CSV table exports for Inventory, Orders, and FG Stock.
- **Specification**:
  1. **Executive 1-Page Printable Briefing**:
     - Add `@media print` rules optimizing the Dashboard view into an executive single-page printable briefing.
     - Content included: Header banner with Plant Date & Shift Timestamps, MTD KPIs (Tube & PET output, Dispatches), Printing Line OTIF status (22K benchmark), Top Active Orders with progress bars, and Yesterday Downtime causes.
     - Automatically hide navigation bars, tab selectors, search filters, and action buttons during print.
     - Add `"📄 Print Executive Summary"` button in header action bar calling `window.print()`.
  2. **UTF-8 CSV Table Downloaders**:
     - Add `"📥 Export CSV"` buttons to Orders table, Inventory table, and FG Stock table.
     - Build pure client-side CSV blob generator with UTF-8 Byte Order Mark (`\uFEFF`) so customer names and product descriptions open cleanly in Microsoft Excel.
  3. **Safety Constraints**:
     - 100% client-side `Blob` and `URL.createObjectURL`. Zero external library bloat.

---

### FP-08: Automated Daily Morning Executive Briefing (WhatsApp Text Generator)

- **Target Files**: [Scripts/daily.py](file:///d:/Alpha/Scripts/daily.py) (Step 7 `step_screenshot`), [Logs/daily_briefing.txt](file:///d:/Alpha/Logs/)
- **Objective**: Automatically compose a formatted, emoji-enhanced morning executive text summary during `daily.py` and copy it directly to the Windows clipboard for instant WhatsApp group pasting.
- **Specification**:
  1. **Briefing Payload Composition** (hooked right after screenshot in Step 7):
     - Compute from active workbook (`Tubex_Sep26.xlsx`) and Imran's `Production.xlsx`:
       - 🏭 **Plant Date & Operating Shift Timestamps**
       - 📦 **Yesterday Production**: Tube pieces (Printing-03 + Printing-04) & PET bottles (PF Machine).
       - 📈 **MTD vs Monthly Plan**: MTD output vs Plan (% compliance & required run-rate).
       - 🎯 **Printing 22K Target**: Status of Printing-03 & Printing-04 against the 22,000 pcs/day OTIF benchmark.
       - 🚨 **Critical Shortages**: Any active MRP items with deficit balance (`Balance <= 0`).
       - ⏱️ **Top Downtimes**: Top 3 machine downtime causes and hours.
       - 🔗 **Live Link**: GitHub Pages URL for phone access.
  2. **Delivery Mechanism**:
     - Save text copy to `Logs/daily_briefing.txt`.
     - Automatically copy text to Windows clipboard via `win32clipboard` / `powershell`.
     - Print console notice: `"[OK] Morning briefing copied to clipboard — press Ctrl+V in WhatsApp"`.
  3. **Safety Constraints**:
     - 100% reliable clipboard copy. Zero brittle WhatsApp Web scrapers or risk of account bans.

---

### FP-11: Smart Printing Line Changeover & Sequence Scheduler

- **Target Files**: [Tubex.html](file:///d:/Alpha/Tubex.html) (under Production tab or dedicated schedule card), optional [Scripts/schedule_advisory.py](file:///d:/Alpha/Scripts/)
- **Objective**: Algorithmic advisory tool sequencing open tube orders by diameter and base coat color gradient to minimize mechanical tooling swaps (saving 3–4 hours) and ink washdowns (saving 1–2 hours) on Print 1 and Print 2.
- **Specification**:
  1. **Machine Line Constraints**:
     - Printing-03: Small/medium diameters (12.5mm, 13.5mm, 16mm, 19mm, 20.5mm, 25mm).
     - Printing-04: Medium/large diameters (25mm, 30mm, 32mm, 35mm).
  2. **Sequencing Optimization Algorithm**:
     - Ingest open unfulfilled orders from `Tubex_Dashboard` and `MRP`.
     - Group jobs by **Diameter** (eliminates mechanical tooling changeovers).
     - Sub-sequence within diameter by **Nozzle / Shoulder Type**.
     - Sub-sequence by **Base Coat Color Gradient** (`White` -> `Clear/Varnish` -> `Dark/Red/Black`) to eliminate intermediate wash-down cycles.
  3. **Advisory Deliverable in Dashboard**:
     - Render interactive recommended schedule card under Production tab in `Tubex.html`.
     - Display estimated downtime hours saved and extra piece capacity gained.
     - Allow supervisor to flag priority rush jobs with visual indicator of changeover penalty.
  4. **Safety Constraints**:
     - Non-destructive advisory view. Does NOT overwrite master Excel job orders.
