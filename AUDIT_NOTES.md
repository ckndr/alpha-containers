# Alpha Containers — Operational Rules & Audit Notes

This document serves as a permanent reference explaining intentional design decisions, business rules, and domain-specific rationale across the Alpha Containers codebase and data pipeline. Future audits and tools should consult these notes before flagging intentional behaviors as anomalies.

---

## 1. Inventory Management & Zeroing Policy (Rule R1-02)

- **Behavior**: When updating `Inventory` from `inventory.xls` via `update_inventory.py`, any item present in the master workbook but absent from the daily ERP export has its Opening, Received, and Issued quantities set to `0.0` and its status in Column K marked as `"Not active in ERP"`.
- **Operational Rationale**:
  - The ERP "Item Wise Consolidated Report" (`inventory.xls`) dynamically filters out inactive items and items with zero balance/movement during the period.
  - In a prior operational incident, an item (30kg of white masterbatch) became inactive and was dropped by the ERP report. Because older scripts preserved existing workbook values when items were missing from the export, the dashboard continued showing 30kg of available stock for several weeks when physical inventory was actually 0 kg, causing a production halt when the shortage was discovered on the shop floor.
  - Setting missing items to `0.0` reflects actual ERP state and prevents "phantom stock".
- **Guardrails**:
  - To prevent accidental zeroing if an operator exports a partially filtered category (e.g. only Slugs), `update_inventory.py` includes a sanity check requiring at least 5 items / 20% catalog coverage before proceeding.

---

## 2. Dispatch Reporting & Previous-Day Cutoff Policy (Rule R1-06)

- **Behavior**: `update_dispatch.py` excludes dispatches with today's date when computing Month-to-Date (MTD) customer dispatch quantities for the daily dashboard.
- **Operational Rationale**:
  - The daily management dashboard and plant KPIs are designed to report on **closed, verified dispatches through the previous operating day**.
  - Current-day dispatches are frequently in progress (goods being staged, gate passes being signed, trucks in transit), and ERP entries during the day are partial.
  - Reporting strictly up to the previous day ensures daily dispatch figures match signed reconciliation records and prevents partial intraday numbers from misrepresenting daily closing metrics.

---

## 3. Product Cataloging & Unassigned PID Tracking (Rule R1-01)

- **Behavior**: When `update_production.py` encounters a product alias in `Production.xlsx` that is not mapped in `ALIASES` (and not a varnish pass), it prompts the operator to input a PID # or assign `PID = 0` (unassigned).
- **Operational Rationale**:
  - Assigning `PID = 0` logs the produced quantity and machine hours in `Production_Log` immediately so production is not lost or excluded from total shift metrics, while clearly flagging the SKU for formal PID catalog assignment.
  - Varnish passes do not require a PID and are tracked by name with `(Varnish)` suffix.

---

## 4. Machine String Matching & Line Identifiers (Rule R1-04)

- **Behavior**: Production line identifiers in `Production.xlsx` may appear as `"Print 1"`, `"Print 2"`, or `"PLINE 1"`, `"PLINE 2"`.
- **Operational Rationale**:
  - Both naming conventions refer to the tube offset printing lines.
  - All aggregation formulas and Python sorting logic treat `"Print*"` and `"PLINE*"` as printing lines, and `"PF*"` / `"PET*"` as PET bottle injection-blow molding lines.

---

## 5. Critical Material-Gated Missing Item Alerts (Rule R1-16)

- **Behavior**: When reviewing inventory discrepancies in `daily.py` and `update_inventory.py`, missing raw material items are only flagged as alerts if the item is a critical core raw material (**Slugs** or **Resin**). All other inactive items are quietly zeroed out (`0.0`), styled in RED, and marked with `"Not active in ERP"` in Column K of the Excel sheet without cluttering daily console output or error summaries.
- **Operational Rationale**:
  - Many auxiliary or non-urgent raw materials (cartons, caps, lacquers, inks) exist in catalog master records that are not active in daily ERP exports. Alerting on all unneeded missing items causes alert fatigue and obscures genuine critical material shortages.
  - Slugs and Resin represent the primary core constraints and are alerted immediately if missing. All other items can be reviewed directly in the Inventory sheet.

---

## 6. Non-Blocking Freshness Warnings (Rule R1-19)

- **Behavior**: `check_freshness()` in `alpha_checks.py` prints a warning when an ERP export is older than 26 hours, but does not halt the pipeline.
- **Operational Rationale**:
  - On weekends, holidays, or days without new dispatches, operators legitimately proceed using the previous day's verified data exports without interruption.

---

## 7. Plant Separation & Scrap Models (Rule R2-07)

- **Behavior**: Tubex Plant and Aerosol Plant operate as two completely separate manufacturing facilities with distinct products, equipment, and BOM scrap modeling:
  - **Tubex (Mature Plant)**: Uses the linear additive model $\text{Gross} = \text{Net} \times (1 + \text{Scrap})$ tailored for mature aluminum tube extrusion and printing lines.
  - **Aerosol (Commissioning Plant)**: Uses the yield inverse model $\text{Gross} = \frac{\text{Net}}{1 - \text{Scrap}}$ tailored for aerosol can drawing, washing, and coating lines.

---

## 8. Shop-Floor Production Data Ownership (Rules R2-08, R2-09, R2-10)

- **Behavior**: `Production.xlsx` is owned and edited solely by the shop-floor data entry operator (Imran).
- **Operational Rationale**:
  - Automation scripts and analysts treat `Production.xlsx` strictly as a read-only input source to prevent file collisions or overwriting floor entries.
  - Internal formula quirks, subtotal methods, or legacy scrap tables within `Production.xlsx` are managed on the shop floor and communicated verbally if adjustments are needed.

---

## 9. Executive Dashboard Downtime Filtering (Rule R2-14)

- **Behavior**: `sort_dashboard.py` and the Dashboard layout omit machine downtime categories with 0.0 MTD hours.
- **Operational Rationale**:
  - Filtering out 0-hour categories saves vertical screen space on the executive overview, highlighting only active machine stoppages requiring management attention.

---

## 10. Pipeline Failure Recovery & Deployment Gating (Rules R4-01 & R4-02)

- **Behavior**:
  - If any pipeline sub-script fails with a non-zero exit code during interactive execution, `daily.py` pauses and prompts the operator whether to continue or stop. In non-interactive runs, it halts immediately.
  - Deployment tasks (OneDrive cloud backup and GitHub push) only execute if the core pipeline succeeds (`success == True`). Minor warnings or cross-check discrepancies do not block deployment.

---

## 11. Unified Non-Destructive Cloud Backups (Rules R4-05 & R4-06)

- **Behavior**: All automated backup routines target `C:\Users\HP\OneDrive\Alpha` using additive copy flags (`/E /COPY:DAT /DCOPY:DAT`) with lockfile exclusions (`/XF "~$*"`).
- **Operational Rationale**:
  - Eliminates destructive mirroring (`/MIR`) so local file cleanups never purge historical cloud backup copies.



