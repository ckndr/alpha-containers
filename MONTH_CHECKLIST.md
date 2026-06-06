# Month-End Transition Checklist

> **When to use**: On the **1st working day of every new month** (or last day of previous month).
> Print this or keep it open on your phone. Check off each step as you go.

---

## 🔴 BEFORE the New Month (Last working day of old month)

### Step 1: Final Daily Run
- [ ] Download fresh ERP exports (dispatch.xls, dispatch_pet.xls, inventory.xls, Production.xlsx)
- [ ] Run `Run_All_Updates.bat` one final time
- [ ] Open Tubex.html, verify all numbers look correct for end-of-month
- [ ] Screenshot the Dashboard totals (Tube MTD, PET MTD, Dispatch) — save in WhatsApp/Notes as backup

### Step 2: Archive the Old Month (Optional but Recommended)
- [ ] Note down the current month's TOTAL production and dispatch numbers
- [ ] The Production.xlsx file from Imran will be overwritten — no action needed, old data stays in git history

---

## 🟢 START of New Month (1st working day)

### Step 3: Imran's Production File
Imran clears `Production Day wise` sheet and starts fresh for the new month.

- [ ] Confirm with Imran: "Did you start a new month in Production.xlsx?"
- [ ] Verify: Open Production.xlsx → `Production Day wise` tab → dates should be the new month only
- [ ] Verify: `FG Stock In hand` tab → Imran should have a new date row for today

> **What if Imran forgot?** The scripts will still work, but Production_Log will contain old month data mixed with new month. Ask Imran to clear old rows.

### Step 4: Update MRP Orders in Excel
This is the **most important manual step**. The MRP sheet drives what the factory needs to produce this month.

- [ ] Open `Tubex_v10_XX.xlsx` → `MRP` sheet
- [ ] **Tube Orders section** (rows 3–14):
  - [ ] Delete completed orders from last month (orders where Remaining = 0)
  - [ ] Update quantities for continuing orders (check with Sales/Dispatch)
  - [ ] Add new orders: Fill in Dia, Customer, Product Name, Product ID, Job Order #, Required Qty
  - [ ] The `Produced` column (G) is a formula — don't touch it, it auto-calculates
  - [ ] The `Remaining` column (H) is `=F-G` — don't touch it either
- [ ] **PET Orders section** (rows ~106–111):
  - [ ] Same process: remove done, update continuing, add new
- [ ] **DO NOT touch** the Material Requirement rows (19–101, 116–123) — they calculate automatically from orders

> **Common mistake**: Forgetting to update the `Date` cell in MRP (H1). This cell controls which month the production formulas filter by. Set it to any date in the new month (e.g., `01-Jul-2026`).

- [ ] Update cell `H1` in MRP sheet to a date in the new month

### Step 5: Update Dashboard Orders
The Dashboard has a separate "Orders" column that some products use.

- [ ] Check `Tubex_Dashboard` column G — if any products have hardcoded order quantities (not formulas), update them for the new month
- [ ] Products with `=IFERROR(INDEX(...)...)` formulas in column G pull from MRP automatically — no action needed

### Step 6: Download Fresh ERP Exports
- [ ] Download `dispatch.xls` (TUBEX-ALUM) from ERP — this will be empty or near-empty for day 1
- [ ] Download `dispatch_pet.xls` (TUBEX-PET) from ERP
- [ ] Download `inventory.xls` from ERP (Item Wise Consolidated, new month dates)
- [ ] Get fresh `Production.xlsx` from Imran

### Step 7: First Run of the New Month
- [ ] Close Excel (all Tubex files)
- [ ] Run `Run_All_Updates.bat`
- [ ] Check the console output for:
  - [ ] No `!! PRODUCTS NOT MATCHED` warnings
  - [ ] No `!! WARNING` about stale files
  - [ ] All 5 steps show `OK`
- [ ] Open `Tubex.html` in browser
- [ ] Verify: Month name in header shows the new month
- [ ] Verify: Production MTD shows 0 or very small numbers (new month)
- [ ] Verify: MRP tab shows the new orders with correct quantities

### Step 8: Push to GitHub
- [ ] Run `Scripts\Push.bat` to deploy the new month's dashboard

---

## 🟡 Troubleshooting

| Problem | Fix |
|---|---|
| Dashboard still shows last month's production | MRP cell H1 not updated to new month. Fix it and re-run. |
| "No Tubex*.xlsx found" error | Excel file is open. Close it and re-run. |
| New product not appearing on Dashboard | Add it to `Product_Catalog` sheet first, then to MRP orders. |
| Dispatch numbers seem wrong | Check dispatch.xls is the new month's export, not last month's. |
| FG Stock showing old date | Imran hasn't entered new date rows in `FG Stock In hand` tab. |
| WIP not updating | Make sure Aurangzeb's message format is correct: `#19mm 10kg #30mm 125kg` |

---

## 📅 Quick Reference: What Changes Monthly

| Item | Who Changes It | Where |
|---|---|---|
| MRP Orders (Tube) | **You** | MRP sheet rows 3–14 |
| MRP Orders (PET) | **You** | MRP sheet rows 106–111 |
| MRP Date Filter | **You** | MRP sheet cell H1 |
| Production Day wise | **Imran** | Production.xlsx |
| FG Stock In hand | **Imran** | Production.xlsx |
| Dispatch exports | **ERP** | Download fresh from ERP |
| Inventory export | **ERP** | Download fresh from ERP |
| Scripts / HTML | **Nobody** | Auto-generated, don't touch |
