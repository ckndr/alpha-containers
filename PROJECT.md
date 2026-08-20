# Project Architecture & Master Plan: Alpha Containers (Tubex) Post-Remediation Audit & Modernization

## Architecture
Alpha Containers (Tubex) operates an industrial manufacturing ERP synchronization and reporting pipeline.
The system integrates shop-floor production logs, ERP inventory/dispatch exports, master catalog BOM definitions, automated Excel data models, and a modern offline-capable Progressive Web Application (PWA).

```
[Shop Floor / ERP Exports]
 (Production.xlsx, inventory.xls, dispatch.xls, dispatch_pet.xls)
             │
             ▼
    [Scripts/daily.py] ── Execution Pipeline ──┐
             │                                │
             ├── 1. update_production.py      │
             ├── 2. update_inventory.py       ├─── [Alpha_Checks / Normalization]
             ├── 3. update_dispatch.py        │    (safe copy, lockfile purge,
             ├── 4. sort_dashboard.py         │     customer aliases, UTF-8)
             ├── 5. build_archives.py         │
             └── 6. update_html.py ───────────┘
                         │
                         ▼
        [Excel Models: Tubex_Aug26.xlsx, Master_Catalog.xlsx, Stock.xlsx]
                         │
                         ▼
        [Web Dashboard / PWA: Tubex.html, index.html, sw.js]
```

---

## Feature Inventory

| # | Feature / Item | Description | Milestone | Source |
|---|---|---|---|---|
| 1 | R1-01 | Interactive & Default PID Assignment in update_production.py | M1 | AUDIT_REPORT.md / Scripts |
| 2 | R1-02 | Safe Inventory 8-Col Ingestion & Non-Active Marking | M1 | AUDIT_REPORT.md / Scripts |
| 3 | R1-03 | Regex Formula Rewriting Protection in sort_dashboard.py | M1 | AUDIT_REPORT.md / Scripts |
| 4 | R1-04 | Machine String Matching Parity (Print & PLINE) | M1 | AUDIT_REPORT.md / Scripts |
| 5 | R1-05 | Dynamic Production Log Row Bounds | M1 | AUDIT_REPORT.md / Scripts |
| 6 | R1-06 | Dispatch Date Parsing & Previous-Day Cutoff | M1 | AUDIT_REPORT.md / Scripts |
| 7 | R1-07 | Dynamic Dispatch Header Discovery | M1 | AUDIT_REPORT.md / Scripts |
| 8 | R1-08 | Dynamic FG Stock Header Discovery | M1 | AUDIT_REPORT.md / Scripts |
| 9 | R1-09 | Deterministic Date Parsing with dayfirst=True | M1 | AUDIT_REPORT.md / Scripts |
| 10 | R1-10 | 8-Column Inventory Layout Default | M1 | AUDIT_REPORT.md / Scripts |
| 11 | R1-11 | Inventory Title Date Range Formatting | M1 | AUDIT_REPORT.md / Scripts |
| 12 | R1-12 | Full Column FG Stock Wiping | M1 | AUDIT_REPORT.md / Scripts |
| 13 | R1-13 | Catalog-Driven Product Type Resolution | M1 | AUDIT_REPORT.md / Scripts |
| 14 | R1-14 | Script Execution Order Harmonization | M1 | AUDIT_REPORT.md / PIPELINE.md |
| 15 | R1-15 | Explicit UTF-8 Encoding across File Operations | M1 | AUDIT_REPORT.md / Scripts |
| 16 | R1-16 | MRP-Gated Shortage Visibility & Persistence | M1 | AUDIT_REPORT.md / Scripts |
| 17 | R1-17 | Dynamic Summary Label Cross-Checks | M1 | AUDIT_REPORT.md / Scripts |
| 18 | R1-18 | Missing Input File Safety Assertion Policy | M1 | AUDIT_REPORT.md / Scripts |
| 19 | R1-19 | Non-Blocking Freshness Warning Policy | M1 | AUDIT_REPORT.md / Scripts |
| 20 | R1-20 | Safe Copy Replacement Guard (>=512 bytes) | M1 | AUDIT_REPORT.md / Scripts |
| 21 | R1-21 | Bounded Customer Normalization & Token Matching | M1 | AUDIT_REPORT.md / Scripts |
| 22 | R1-22 | Standardized Active Workbook Version Sorting | M1 | AUDIT_REPORT.md / Scripts |
| 23 | R2-01 | Dashboard Order Formula Range Expansion (G12:G56) | M1 | AUDIT_REPORT.md / Tubex_Aug26.xlsx |
| 24 | R2-02 | Catalog Formula Offsets Remediation (J50:P55) | M1 | AUDIT_REPORT.md / Master_Catalog.xlsx |
| 25 | R2-03 | Aerosol Lacquer Waste Factor (35% TDS Standard) | M1 | AUDIT_REPORT.md / Master_Catalog.xlsx |
| 26 | R2-04 | Job Card Double Scrap Tolerance Elimination | M1 | AUDIT_REPORT.md / Daily_Job_Card.xlsx |
| 27 | R2-05 | Job Card 12-Color Ink Formula Limitation | M1 | AUDIT_REPORT.md / Commissioning |
| 28 | R2-06 | Inventory AVERAGEIF Slugs/Resin Distortion | M1 | AUDIT_REPORT.md / Stock.xlsx |
| 29 | R2-07 | Linear vs Yield Inverse Scrap Consistency | M1 | AUDIT_REPORT.md / AUDIT_NOTES.md |
| 30 | R2-08 | Production.xlsx Non-Numeric Shift Data Handling | M1 | AUDIT_REPORT.md / Production.xlsx |
| 31 | R2-09 | Production.xlsx Line Rate Scrap Assumptions | M1 | AUDIT_REPORT.md / Production.xlsx |
| 32 | R2-10 | Production.xlsx Manual Downtime Discrepancies | M1 | AUDIT_REPORT.md / Production.xlsx |
| 33 | R2-11 | Historical Baseline #VALUE! Error Elimination | M1 | AUDIT_REPORT.md / Tubex_v10_30.xlsx |
| 34 | R2-12 | August Plan PET Total Row 9 Sum Correction | M1 | AUDIT_REPORT.md / Plan_2026.xlsx |
| 35 | R2-13 | FG Stock Cap ID Lookup Exact Match Correction | M1 | AUDIT_REPORT.md / Tubex_Aug26.xlsx |
| 36 | R2-14 | Dashboard Downtime 0-Hour Category Filtering | M1 | AUDIT_REPORT.md / Tubex_Aug26.xlsx |
| 37 | R2-15 | Inventory Formula Row Offset J63 Alignment | M1 | AUDIT_REPORT.md / Stock.xlsx |
| 38 | R2-16 | Historical Fragile Formula Additions in Pending Orders | M1 | AUDIT_REPORT.md / Pending.xlsx |
| 39 | R3-01 | EscapeHtml DOM XSS Protection on Orders & FG Stock | M1 | AUDIT_REPORT.md / Tubex.html |
| 40 | R3-02 | Data-Attribute Event Binding for Customer/Period Chips | M1 | AUDIT_REPORT.md / Tubex.html |
| 41 | R3-03 | EscapeHtml DOM Protection across all Summary Tables | M1 | AUDIT_REPORT.md / Tubex.html |
| 42 | R3-04 | Service Worker HTTP 200 Caching Guard | M1 | AUDIT_REPORT.md / sw.js |
| 43 | R3-05 | Service Worker HTTP/HTTPS Scheme Validation | M1 | AUDIT_REPORT.md / sw.js |
| 44 | R3-06 | Controllerchange Immediate Live Refresh | M1 | AUDIT_REPORT.md / sw.js & Tubex.html |
| 45 | R3-07 | Standard ISO-8601 Timestamp Generation & Parsing | M1 | AUDIT_REPORT.md / Scripts & HTML |
| 46 | R3-08 | Template Injection Marker Integrity (Zero Duplicate Tags) | M1 | AUDIT_REPORT.md / Tubex.html |
| 47 | R3-09 | PWA Offline Root Navigation & Asset Fallbacks | M1 | AUDIT_REPORT.md / sw.js & index.html |
| 48 | R4-01 | Interactive Pipeline Failure Prompt & Halt | M1 | AUDIT_REPORT.md / Scripts/daily.py |
| 49 | R4-02 | Deployment Gating on Pipeline Success | M1 | AUDIT_REPORT.md / Scripts/daily.py |
| 50 | R4-03 | Excel COM Leak Elimination via DispatchEx & Finally | M1 | AUDIT_REPORT.md / Scripts |
| 51 | R4-04 | Persistent MRP Shortage Alert Tracking | M1 | AUDIT_REPORT.md / Scripts/daily.py |
| 52 | R4-05 | Unified OneDrive Destination Path | M1 | AUDIT_REPORT.md / Scripts & Batches |
| 53 | R4-06 | Non-Destructive Robocopy /E Additive Backup | M1 | AUDIT_REPORT.md / Scripts & Batches |
| 54 | R4-07 | Startup Lockfile Purge & Robocopy Exclusion | M1 | AUDIT_REPORT.md / alpha_checks.py |
| 55 | R4-08 | Synchronized Execution Sequence in Documentation | M1 | AUDIT_REPORT.md / PIPELINE.md |
| 56 | D1 | End-to-End Daily Pipeline Script Dry Run | M2 | Daily Workflow |
| 57 | D2 | Process Cleanliness & Zero EXCEL.EXE COM Leaks | M2 | Daily Workflow |
| 58 | D3 | Workbook Formula Integrity & Zero #REF!/#VALUE! | M2 | Daily Workflow |
| 59 | D4 | Operational Guarantee for Next Daily Update Cycle | M2 | Daily Workflow |
| 60 | M1 | FP-01: Raw Material Slugs & Resin Yield Calculator Spec | M3 | Future_Plans / Tubex_Aug26.xlsx |
| 61 | M2 | FP-02: Historical Month Selector & Archive Navigation Spec | M3 | Future_Plans / Tubex_Aug26.xlsx |
| 62 | M3 | Direct ERP Database Connector (Proposal 1) | M3 | Modernization Blueprint |
| 63 | M4 | Automated WhatsApp Production Parser Bot (Proposal 2) | M3 | Modernization Blueprint |
| 64 | M5 | Dynamic Machine-Specific Scrap Model (Proposal 3) | M3 | Modernization Blueprint |
| 65 | M6 | MRP Lead-Time Reorder Point Triggering (Proposal 4) | M3 | Modernization Blueprint |
| 66 | M7 | Package Architecture & CLI Tooling Refactor (Proposal 5)| M3 | Modernization Blueprint |
| 67 | M8 | Structured JSON Telemetry & Health Monitoring (Proposal 6)| M3 | Modernization Blueprint |
| 68 | M9 | Touch-Optimized Mobile Navigation & Shift Velocity (Proposal 7)| M3 | Modernization Blueprint |
| 69 | M10| High-Contrast Dark & Light Industrial Themes (Proposal 8)| M3 | Modernization Blueprint |
| 70 | R-DOC | Comprehensive Master POST_REMEDIATION_AUDIT_REPORT.md | M4 | Synthesis & Handoff |

---

## Milestones

| # | Name | Scope | Dependencies | Status |
|---|---|---|---|---|
| M1 | 56-Finding Post-Remediation Evidence Matrix | Compile verified code evidence, formulas, line numbers, and proofs for all 56 findings (R1-01 to R1-22, R2-01 to R2-16, R3-01 to R3-09, R4-01 to R4-08) | Survey Complete | **DONE** |
| M2 | End-to-End Daily Workflow Dry Run & Process Health | Execute and verify daily update pipeline, confirm zero Excel COM leaks, zero #REF!/#VALUE! formula errors, zero Windows encoding crashes, and assert operational guarantee | M1 | **DONE** |
| M3 | Modernization Blueprint & Technical Specifications | Formulate comprehensive technical specifications for FP-01 (Slugs/Resin Calc), FP-02 (Historical Month Selector), and 8 high-impact proposals across the 4 pillars | M1 | **DONE** |
| M4 | Master Report Assembly & Multi-Agent Gate Verification | Synthesize master deliverable `POST_REMEDIATION_AUDIT_REPORT.md`, conduct Reviewer, Challenger, and Forensic Auditor verification gates, and deliver final handoff | M1, M2, M3 | **DONE** |

---

## Interface Contracts

### Script Pipeline ↔ Excel Workbooks
- `update_production.py`: Reads `Production.xlsx`, writes to active `Tubex_*.xlsx` sheets `Production_Log` and `FG_Stock`.
- `update_inventory.py`: Ingests 8-column `inventory.xls`, writes to active `Tubex_*.xlsx` sheet `Inventory` (clearing unused rows, highlighting inactive items).
- `update_dispatch.py`: Ingests `dispatch.xls` & `dispatch_pet.xls`, writes to active `Tubex_*.xlsx` sheet `Dispatch_Log`.
- `sort_dashboard.py`: Cleans and sorts `Dashboard` sheet in active `Tubex_*.xlsx`, rewriting formulas safely via regex without corrupting table refs.
- `build_archives.py`: Evaluates and snapshots historical monthly production logs using clean COM isolation (`DispatchEx` + `try...finally`).
- `update_html.py`: Evaluates active workbook via COM or openpyxl, extracts metrics into `DASH_DATA` JSON, and injects into `Tubex.html` between exact `/* DATA_START */` and `/* DATA_END */` markers.

### Web Dashboard ↔ Service Worker & Browser
- `sw.js`: Intercepts `GET` requests, caches HTTP 200 responses, provides navigation fallback to `./Tubex.html`.
- `Tubex.html`: Listens to `controllerchange` to reload on update, sanitizes all strings with `escapeHtml()`, binds event handlers via `data-*` attributes.

---

## Code Layout

- `Scripts/`: Core automation pipeline (`daily.py`, `update_production.py`, `update_inventory.py`, `update_dispatch.py`, `sort_dashboard.py`, `build_archives.py`, `update_html.py`, `alpha_checks.py`, `customer_normalization.py`, `update_wip.py`).
- Root Excel Workbooks: `Tubex_Aug26.xlsx` (Active Model), `Master_Catalog.xlsx` (BOM Catalog), `Daily_Job_Card.xlsx`, `Plan_2026.xlsx`, `Stock.xlsx`, `Pending.xlsx`.
- Web / PWA Assets: `Tubex.html`, `index.html`, `sw.js`, `manifest.json`.
- Documentation & Reports: `AUDIT_REPORT.md`, `AUDIT_NOTES.md`, `PIPELINE.md`, `DAILY_WORKFLOW.md`, `POST_REMEDIATION_AUDIT_REPORT.md`.
- Metadata & Coordination: `.agents/` (agent briefings, progress, evidence handoffs).
