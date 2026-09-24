# Alpha Containers Project

This repository contains project files, documentation, and automation tools for two plants:

1. **Tubex Plant** (Production): A fully mature plant operating for 40+ years. Its data processing pipeline, scripts, and logs are located in the root directory.
2. **Aerosol Plant** (Commissioning): A new plant currently in the commissioning phase. It is managed in an independent project directory at `D:\Aerosol` (ignored by this repository).

---

## Tubex - Excel Workbook Versioning

When editing an Excel workbook:
- Do not overwrite the existing file. Save the edited workbook as a new version instead.
- For normal incremental workbook edits, increase the final minor number. Example: `Tubex_v10_28.xlsx` becomes `Tubex_v10_29.xlsx`.
- For major structural or logic changes, increase the main version. Example: `Tubex_v10_28.xlsx` becomes `Tubex_v11_1.xlsx`.
- Keep the previous workbook file in place so it remains available as a backup/reference.

## Temporary Files Cleanup

- Always delete any temporary scripts, log dumps, or automation files (e.g. `fix_mrp_formulas.py` or `fix_mrp_v28_formulas.py`) created in the repository root directory before finishing a task.
- Keep the workspace clean and tidy to prevent cluttering version control.

## GitHub Sync Workflow

Use `Scripts\Push.bat` when you want to upload the current project files to GitHub without running the daily update workflow.

`Scripts\Push.bat` stages changed repository files, commits them with a timestamp, and pushes directly to GitHub `main`.

Use `Scripts\Pull.bat` before starting work on another computer. It downloads the latest `main` branch from GitHub into the local `D:\Alpha` folder. If local uncommitted changes exist, it stops and asks you to push or clean up first.

Important: the main versioned workbook files, such as `Tubex_v10_27.xlsx`, are included in GitHub sync so they can move between work and home. Raw Excel/data exports such as `dispatch.xls`, `inventory.xls`, `Production.xlsx`, and other `.xls`/`.xlsx` files remain ignored unless intentionally allowed.

## Monthly Rollover

To roll over to a new operating month, run `new_month.py`:
```cmd
python new_month.py --month <Mon> --year <YYYY> [--dry-run]
```
