# Tubex Lahore — Production Planning Rules (Tubes & PET)

Authoritative reference for production scheduling, capacity planning, changeover sequencing, material constraints, and automated plan tracking at Alpha Containers (Tubex) Lahore.

---

## 1. Aluminum Tubes Planning Rules

### 1.1 Machine Lines & Diameter Allocations
- **Printing / Decoration Lines (`Pline-01`..`Pline-05` / `Printing-01`..`05`)**:
  - `Printing-03` (`Pline-03`): Small & medium diameters — `12.5 mm`, `13.5 mm`, `16 mm`, `19 mm`, `20.5 mm`, `25 mm` (plus `30 mm` / `35 mm` overflow).
  - `Printing-04` (`Pline-04`): Medium & large diameters — `19 mm`, `20.5 mm`, `25 mm`, `30 mm`, `32 mm`, `35 mm`.
  - `Printing-01` / `Printing-02` / `Printing-05`: Auxiliary / historical offset lines (`20.5 mm`, `25 mm`, `30 mm`, `32 mm`, `35 mm`).
- **Upstream / Downstream Tube Lines**:
  - `Press-01`, `Press-04`, `Press-06`: Extrusion (`25 mm`, `30 mm`, `32 mm` currently active).
  - `Lacquer-03` (`30 mm`), `Lacquer-04` (`25 mm`, `32 mm`): Internal lacquering.
  - `Latex-01` (`30 mm`, `32 mm`): End-sealing latex application.

### 1.2 Throughput & Daily Benchmarks
- **OTIF Target Benchmark**: **22,000 pcs/day** per active printing line.
- **Historical YTD Actual Average**: **~13,800 – 15,900 pcs/day** (single 8-hour shift `480 min`, reflecting tooling swaps, colour washdowns, and line downtime).

### 1.3 Changeover & Job Sequencing Hierarchy
Always sequence open tube orders on a printing line in the following priority order to avoid **3–4 hour mechanical tooling swaps** and **1–2 hour ink/roller washdowns**:
1. **Diameter (`Dia`)**: Group all jobs of the same tube diameter together first.
2. **Nozzle / Shoulder Type**: Sub-group within the same diameter by nozzle/shoulder specification.
3. **Base Coat Color Gradient (Light-to-Dark)**: Sequence from **White → Clear/Varnish → Dark/Red/Black** to minimize full roller washdowns between jobs.

### 1.4 Material Constraints & Scrap Standards
- **Hard Scheduling Gates**: Verify `MRP` availability before releasing a job to the floor — **Aluminum Slug**, **Cap**, **Base Coat**, **Lacquer**, and **Inks** must cover the order's remaining balance.
- **Plant Slug Weight Standards (g / tube)**:
  - `12.5 / 13.5 mm`: `1.950 g`
  - `16.0 mm`: `2.519 g`
  - `19.0 mm`: `3.367 g`
  - `20.5 / 22.0 mm`: `3.937 g`
  - `25.0 mm`: `5.917 g`
  - `28.0 / 30.0 mm`: `8.000 g`
  - `32.0 mm`: `10.863 g`
  - `35.0 mm`: `12.820 g`
- **Scrap Model (Linear Additive)**: $\text{Gross} = \text{Net} \times (1 + \text{Scrap})$.
  - **10% Scrap (`0.10`)**: Aluminum Slugs, Base Coat, Lacquer, Ink, Latex, Zinc Stearate, Cartons, Thinner, Packing Tape.
  - **2% Scrap (`0.02`)**: Caps.

### 1.5 Varnish Pass Exclusion Policy
- Overprint `(Varnish)` passes have no separate BOM or PID and are logged with `(Varnish)` in the product name.
- All dashboard KPIs, order completion formulas, and plan actuals **must exclude `(Varnish)` rows** (`ISERROR(SEARCH("(Varnish)", Product))`) so good tube output is never double-counted.

---

## 2. PET Bottles Planning Rules (`PF Machine`)

### 2.1 Capacity & Shift Assumptions
- **Planning Rate**: **10,000 good bottles/day** assuming a **12-hour effective run** / double shift (`A+B (8+8) = 960 min` logged in `Production.xlsx`).
- Single 8-hour shift (`480 min`) typically yields ~5,000–7,000 good bottles depending on startup/purging waste.

### 2.2 Mould Swaps & Colour Sequencing
- **Mould Swaps (`CHANGEOVER`)**: Changing bottle volume or mould shape requires **6–8 hours** of mechanical changeover; schedule **1 full `CHANGEOVER` day (`0` planned qty)** between different mould campaigns (e.g., `150 ml TRP → 150 ml Capsule → 130 ml White → 120 ml → 200 ml → 150 ml TRP`).
- **Colour Sequencing & Purge**: Group same-mould colours together and avoid jumping directly from **Black** to **White/Transparent** without allocating extra purge time/material (e.g., `Mould Swap & Purge: 120ml Black -> 200ml White`, `200ml Black -> 150ml Clear`).

### 2.3 Non-Working Days & Holidays
- **Sundays**: `WEEKLY OFF` (`0` planned production).
- **Pakistan National Holidays**: `PUBLIC HOLIDAY` (`0` planned production), including `2026-12-25` (Quaid-e-Azam Day / Christmas), `2027-02-05` (Kashmir Day), and gazetted Eid/national holidays.

### 2.4 PET Material & Scrap Standards
- **Standard PET BOM Scrap**: **15% (`0.15`)** on PET Resin (`2680`), Masterbatch (`4002` Yellow, `5816` White, `7412` Black), Master Cartons, Poly Bags, Packing Tape, and Plastic Doori.
- **Masterbatch Dosing**: Typically **2%–5%** by weight depending on opacity/colour (`0.54–1.25 kg` per `1,000` bottles in `TableBOM`).

### 2.5 Baseline vs. Forecast Tracking (`PET_Plan` & `Scripts/update_plans.py`)
- **Static Baseline**: The `Planned Qty` column in `PET_Plan` (`Tubex_Nov26.xlsx`) represents the frozen baseline commitment and is **not overwritten** during daily updates (unless explicitly invoked with `--replan`).
- **Daily Actuals & Status**: `Scripts/update_plans.py` automatically populates `Actual Good`, `Wastage`, `Variance`, and `Status` (`Scheduled`, `Done`, `Ahead (+X)`, `Behind (-X)`, `Off/Changeover`, `Unplanned: <PID>`) from `Production_Log` (`PF Machine`) for all dates `<= max_prod_date`.
- **Dynamic Forecast Finish**: `Forecast Finish` in the `PET_Plan` summary table is re-simulated forward every run from `max_prod_date + 1` at `10,000/day` using live `MRP` remaining balances (skipping Sundays, national holidays, and remaining changeovers).
- **Shareable Export**: `update_plans.py` also keeps `<MonthName>_Plan.xlsx` (e.g., `November_Plan.xlsx`) synced for management sharing, falling back to `<MonthName>_Plan_LATEST.xlsx` if the primary file is open in Excel.
