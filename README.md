# Tubex

## Excel Workbook Versioning

When editing an Excel workbook, do not overwrite the existing file. Save the edited workbook as a new version instead.

- For normal incremental workbook edits, increase the final minor number. Example: `Tubex_v10_26.xlsx` becomes `Tubex_v10_27.xlsx`.
- For major structural or logic changes, increase the main version. Example: `Tubex_v10_26.xlsx` becomes `Tubex_v11_1.xlsx`.
- Keep the previous workbook file in place so it remains available as a backup/reference.

## GitHub Sync Workflow

Use `Scripts\Push.bat` when you want to upload the current project files to GitHub without running the daily update workflow.

`Scripts\Push.bat` also copies the project folder to `C:\Users\HP\OneDrive\Tubex` before pushing to GitHub. It copies new and changed files, but does not delete extra files already in OneDrive. It excludes the `.git` folder, local logs, and temporary Excel lock files.

Use `Scripts\Pull.bat` before starting work on another computer. It downloads the latest `main` branch from GitHub into the local `D:\Alpha` folder. If local uncommitted changes exist, it stops and asks you to push or clean up first.

Important: the main versioned workbook files, such as `Tubex_v10_27.xlsx`, are included in GitHub sync so they can move between work and home. Raw Excel/data exports such as `dispatch.xls`, `inventory.xls`, `Production.xlsx`, and other `.xls`/`.xlsx` files remain ignored unless intentionally allowed.

## Hourly Auto Push

Run `Scripts\Install_Hourly_Push.bat` once on any PC where you want automatic backups. It creates a Windows scheduled task named `Tubex Hourly Push` that runs `Scripts\Push.bat /auto` every hour.

Before relying on the scheduled task, run `Scripts\Push.bat` manually once on that PC so GitHub credentials are saved.

Hourly auto push only works while the PC is powered on, Windows is running, and internet/GitHub access is available. If there is a blackout or the PC is off, the task cannot push during that time; it will try again on the next scheduled run after the PC is back on.

Logs are saved at `Logs\hourly_push.log`. To remove the scheduled task, run `Scripts\Remove_Hourly_Push.bat`.
