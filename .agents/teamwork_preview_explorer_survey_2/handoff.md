# Handoff Report: Requirement R2 (Excel Models, Formulas & BOM Consistency Audit)

**Agent**: Survey Explorer 2 (`teamwork_preview_explorer_survey_2`)  
**Recipient**: Orchestrator (`orchestrator_1`)  
**Date**: 2026-08-19  
**Full Investigation Report**: `d:\Alpha\.agents\teamwork_preview_explorer_survey_2\r2_excel_bom_audit.md`

---

## 1. Observation

Direct forensic observations from openpyxl inspection of all Excel workbooks across `d:\Alpha`:

1. **`Tubex_Aug26.xlsx` — `Tubex_Dashboard!G12:G56`**:
   - Formula verbatim:
     `=IFERROR(INDEX(MRP!$F$3:$F$3,MATCH(Tubex_Dashboard!F12,MRP!$D$3:$D$3,0)),0)`
   - Array locked to single cell `$F$3:$F$3` instead of `$F$3:$F$100`, causing 37 out of 38 tube SKUs to report `0` required orders.
2. **`Tubex_Aug26.xlsx` — `Product_Catalog!J50:P55`**:
   - Formulas verbatim on Row 51 (PID 8013):
     `=IF(I50="","",IFERROR(SUMPRODUCT((TableBOM[Product ID]=A50)*(TableBOM[Material Category]="SLUG")*TableBOM[Per 1000 Units]*(1+TableBOM[Scrap %]))*I50/1000,0))`
   - References `A50/I50` instead of `A51/I51`. Displaced across rows 50–55 (PIDs 8013, 2909, 4227, 5389, 6151) by -1 to -2 row offsets.
3. **`Tubex_Aug26.xlsx` — `Inventory!J3:J111` & `Inventory!J63`**:
   - Formula verbatim on Row 3:
     `=IFERROR(IF(AVERAGEIF(TableBOM[Item ID],A3,TableBOM[Per 1000 Units])=0,"-",ROUND((H3+I3)/(AVERAGEIF(TableBOM[Item ID],A3,TableBOM[Per 1000 Units])/1000),0)),"-")`
   - Item 2680 (`PET RESIN A-84`) has rates from 17.1 to 50.0 kg/1000; unweighted arithmetic mean (23.54 kg/1000) causes capacity errors between -27% and +112%.
   - Cell J63 references `A62` instead of `A63`.
4. **`Tubex_Aug26.xlsx` — `FG Stock!I4:I99`**:
   - Formula verbatim on Row 4:
     `=IFERROR(SUMPRODUCT((TableBOM[Product ID]=B4)*(TableBOM[Material Category]="CAP")*TableBOM[Item ID]),0)`
   - Multiplies Item ID numerically, summing multi-component IDs into nonexistent IDs.
5. **`Production.xlsx` — `Summary 14-08-2026!B13, B24` & `Production Day wise!N3:N73, N1`**:
   - Cell B13: `=B11/B12` evaluates to cached `#DIV/0!` because target dispatch B12 is 0.
   - Cell N3: `=IFERROR(L3/M3,"0%")` divides Wastage by Good Production instead of Total Production (`L3/K3`) and returns text `"0%"`.
   - Cell N1: `=SUBTOTAL(101,N3:N28442)` averages percentage scrap across unequal batches.
6. **`Production.xlsx` — `Sheet3!J3:P29`**:
   - Contains broken external link `[1]!TableBOM` and typo `"LECQUER"` in Column L.
7. **`Aerosol BOM.xlsx` — `Theoretical BOM!K6:K7, L6:L7`**:
   - Internal Lacquer (Items 504, 505) assigned flat 10% scrap (`0.1`), whereas technical documentation and TDS confirm internal airless spray loss is 30%–40% (transfer efficiency 60%–70%), causing a 27.8% raw material deficit.
8. **`Aerosol_Job_Card.xlsx` — `Job Card!E12:E36` & `B12:F32`**:
   - Cell E12: `=IFERROR(VLOOKUP(...) * ($B$8*(1+$D$8)) / 1000, "")` multiplies Gross Qty by `(1 + Tolerance)`, double-counting waste.
   - Pulls all 12 inks (3.36 kg/1000) for every job regardless of artwork color count (4 colors = 1.12 kg/1000).
9. **`August_Plan.xlsx` — `August Plan PET!K10:M10`**:
   - Formulas `=SUM(K6:K8)`, `=SUM(L6:L8)`, `=SUM(M6:M8)` omit Row 9 (`Samsol Yellow 120ml`, 37,160 units).
10. **`Tubex_v10_30.xlsx` — `MRP!F118:G121`**:
    - Formulas `=IF(E111=0,0,ROUND(H111/E111*30,1))` divide text strings of product names by numbers, evaluating to cached `#VALUE!`.

---

## 2. Logic Chain

1. **Observation 1 & 2 $\rightarrow$ Requirement Blindness & Wrong BOM Execution**:
   - Hardcoding `$F$3:$F$3` masks 37 open orders in `Tubex_Dashboard`. Planners cannot see demand.
   - Row-index offsets in `Product_Catalog` (Obs 2) link PIDs 8013, 2909, 4227, 5389, and 6151 to incorrect product IDs, directly generating wrong slug and cap requisitions.
2. **Observation 3 & 4 $\rightarrow$ Flawed Planning & Inventory Allocation**:
   - Using `AVERAGEIF` on multi-BOM raw materials (Obs 3) introduces massive variance (-27% to +112%) between projected pieces and physical capacity.
   - `SUMPRODUCT` on Item IDs (Obs 4) corrupts inventory tracking for multi-cap SKUs.
3. **Observation 7, 8 $\rightarrow$ Severe Material Shortages & Aerosol Cost Inflation**:
   - Under-estimating internal spray loss at 10% instead of 35% (Obs 7) causes an immediate 335 kg lacquer shortage on a 750k can run.
   - Job Card calculations (Obs 8) double-count scrap on tolerance and triple-count ink usage by pulling all 12 colors.
4. **Observation 5, 6, 9, 10 $\rightarrow$ Reporting Corruption & Silent Failures**:
   - `#DIV/0!` and `#VALUE!` errors propagate to dashboard summary cards.
   - PET monthly plan (Obs 9) under-reports demand by 37,160 units due to a truncated SUM range.
   - Production wastage calculation `L3/M3` violates scrap math fundamentals.

---

## 3. Caveats

- Historical workbooks in `Tubex Records/` prior to November 2025 were spot-checked but not individually re-tabulated.
- VBA macros: No `.xlsm` VBA macro files exist in the operational root; models rely strictly on Excel worksheet formulas and external Python ETL.
- All observations reflect direct static parsing of the `.xlsx` XML structures.

---

## 4. Conclusion

The Excel model architecture contains **3 Critical, 8 High, and 5 Medium** severity failure modes. Primary risks include:
1. **Critical order requirement blindness** in `Tubex_Dashboard` (`G12:G56`).
2. **Catalog BOM misattribution** in `Product_Catalog` (`J50:P55`).
3. **Severe lacquer under-budgeting (27.8%)** in `Aerosol BOM.xlsx`.
4. **Multi-ink inflation (200%-300%)** and **double-counted tolerance** in `Aerosol_Job_Card.xlsx`.
5. **Planning total truncation (37,160 units)** in `August_Plan.xlsx`.

All issues are mathematically documented with exact cell coordinates and drop-in formula remediations in `r2_excel_bom_audit.md`.

---

## 5. Verification Method

To independently verify all findings:
1. Run Python inspection scripts against the target workbooks:
   ```powershell
   python d:\Alpha\.agents\teamwork_preview_explorer_survey_2\deep_inspect.py
   python d:\Alpha\.agents\teamwork_preview_explorer_survey_2\audit_tubex_aug26.py
   python d:\Alpha\.agents\teamwork_preview_explorer_survey_2\audit_full_forensic_data.py
   ```
2. Inspect target cells directly in Excel:
   - `Tubex_Aug26.xlsx` $\rightarrow$ `Tubex_Dashboard!G12`, `Product_Catalog!J51:P55`, `Inventory!J63`.
   - `Production.xlsx` $\rightarrow$ `Summary 14-08-2026!B13`, `Production Day wise!N3`.
   - `Aerosol/Aerosol BOM.xlsx` $\rightarrow$ `Theoretical BOM!K6:L7`.
   - `Aerosol/Aerosol_Job_Card.xlsx` $\rightarrow$ `Job Card!E12`.
   - `August_Plan.xlsx` $\rightarrow$ `August Plan PET!K10:M10`.
