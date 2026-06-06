# Tubex Daily Workflow — Sikander's Process

> **Purpose**: Documents the exact daily process for updating the Tubex dashboard.
> Anyone reading this (AI or human) should be able to follow these steps.

---

## Prerequisites
- Windows PC with Remote Desktop access to ERP server
- WhatsApp (for receiving files from Imran and Mehmood)
- `d:\Alpha\` folder with all scripts set up
- Git configured for GitHub push

---

## Step-by-Step Daily Process

### Phase 1: ERP Data Extraction (~5 min)

**Step 1.1 — Connect to ERP**
- Open the Remote Desktop shortcut on desktop
- Enable: ✅ Clipboard sharing, ✅ Drive redirection (so local D: drive is accessible inside RDP)
- Connect to the ERP server

**Step 1.2 — Extract Reports from ERP**
Inside the Remote Desktop session:
- Login to ERP application
- Navigate to Reports section
- Export 3 reports:
  1. **Inventory report** → "Item Wise Consolidated" → saves as `inventory.xls`
  2. **Dispatch report (TUBEX-ALUM)** → "Dispatch Report (Date Wise)" → saves as `dispatch.xls`
  3. **Dispatch report (TUBEX-PET)** → "Dispatch Report (Date Wise)" → saves as `dispatch_pet.xls`

**Step 1.3 — Copy to Shared Drive**
- Copy all 3 files to the shared drive folder where `d:\Alpha\` is accessible
- Rename them as copies: `inventory - copy.xls`, `dispatch - copy.xls`, `dispatch_pet - copy.xls`
- The Python scripts auto-detect `- copy` files and replace the old ones

> **Why copies?** The scripts look for `filename - copy.ext` pattern. When found, they rename
> the copy to replace the original, ensuring you always have a fresh export.

---

### Phase 2: Production Data (~3 min)

**Step 2.1 — Get Production.xlsx from Imran**
- Imran Manzoor sends the daily production file via WhatsApp
- Download the file from WhatsApp (it may have a random name like `document_2026-06-05.xlsx`)
- Rename it to `Production.xlsx`
- Move/copy it to `d:\Alpha\` folder (replacing the old one)

**Step 2.2 — Get WIP Message from Mehmood**
- Mehmood sends a WhatsApp message with current WIP (Work In Process) slug weights
- Format example: `#19mm 10kg #30mm 125kg #32mm 25kg`
- Copy this message text to clipboard

**Step 2.3 — Run WIP Updater**
- Open command prompt in `d:\Alpha\Scripts\`
- Run: `python update_wip.py` (or double-click `Update_WIP.bat`)
- Paste the WIP message when prompted
- Script updates Inventory sheet col I with slug WIP weights

---

### Phase 3: Run Pipeline (~2 min)

**Step 3.1 — Close Excel**
- Make sure NO Tubex Excel files are open (the scripts need exclusive file access)

**Step 3.2 — Run Full Update**
- Double-click `d:\Alpha\Scripts\Run_All_Updates.bat`
- Wait for all 5 steps to complete
- Check that all steps say "OK" (no red `!!` errors)
- The batch file:
  1. Backs up Excel to `Logs/`
  2. Updates Production_Log + FG Stock from Production.xlsx
  3. Updates Inventory from inventory.xls
  4. Updates Dispatch from dispatch.xls + dispatch_pet.xls
  5. Sorts Dashboard (active products to top)
  6. Generates Tubex.html + pushes to GitHub

---

### Phase 4: Verification & Sharing (~3 min)

**Step 4.1 — Cross-Check Data**
- Open the updated `Tubex_v10_XX.xlsx` in Excel
- Press `Ctrl+Shift+F9` to recalculate all formulas
- Compare Dashboard numbers against Imran's Production.xlsx:
  - Total good production per machine should match
  - MTD tube total should match
  - MTD PET total should match
  - Any mismatches = data entry error (Imran's or script's)

**Step 4.2 — Share Dashboard**
- Open `Tubex.html` in Chrome (or check on phone via GitHub Pages URL)
- Take a screenshot of the Dashboard tab showing:
  - MTD Production numbers
  - Dispatch numbers
  - Order compliance
- Share the screenshot to the WhatsApp group for bosses

---

## Timing Summary

| Phase | Steps | Time | Can Automate? |
|---|---|---|---|
| ERP Extraction | 1.1–1.3 | ~5 min | Partially (RDP is manual) |
| Production Data | 2.1–2.3 | ~3 min | Partially (WhatsApp is manual) |
| Run Pipeline | 3.1–3.2 | ~2 min | Fully |
| Verify & Share | 4.1–4.2 | ~3 min | Partially |
| **Total** | | **~13 min** | |

---

## File Flow Diagram

```
ERP Server (via RDP)
  ├── inventory.xls ─────────┐
  ├── dispatch.xls ──────────┤  copied as "- copy" variants
  └── dispatch_pet.xls ──────┤  to d:\Alpha\
                              │
WhatsApp (Imran)              │
  └── Production.xlsx ────────┤  renamed & moved to d:\Alpha\
                              │
WhatsApp (Mehmood)            │
  └── WIP text message ───────┤  pasted into update_wip.py
                              │
                              ▼
                    Run_All_Updates.bat
                              │
                    ┌─────────┴─────────┐
                    │                   │
              Tubex_v10_XX.xlsx    Tubex.html
              (updated Excel)     (web dashboard)
                    │                   │
                    ▼                   ▼
              Cross-check         GitHub Pages
              with Imran          (phone access)
                    │
                    ▼
              Screenshot → WhatsApp group
```

---

## Key People

| Person | Role | What They Send | Format |
|---|---|---|---|
| **Imran Manzoor** | Production Supervisor | Daily production file | `.xlsx` via WhatsApp |
| **Mehmood (Aurangzeb)** | Floor Worker | WIP slug weights | Text message via WhatsApp |
| **Sikander** | You — runs the pipeline | Dashboard screenshot | Screenshot via WhatsApp |
| **Bosses** | Recipients | View dashboard | WhatsApp group |
