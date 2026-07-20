import sys
sys.path.append("D:/Alpha/scratch")
import pivot_breakdowns

df_monthly = pivot_breakdowns.df_monthly
df_overall = pivot_breakdowns.df_overall

output_file = "D:/Alpha/scratch/samsol_breakdown.txt"

with open(output_file, "w") as f:
    f.write("=== MONTHLY BREAKDOWN ===\n")
    months_sorted = sorted(df_monthly['Month'].unique())
    for m in months_sorted:
        df_m = df_monthly[df_monthly['Month'] == m]
        f.write(f"\n--- Month: {m} ---\n")
        tot_p = 0
        tot_d = 0
        for idx, row in df_m.iterrows():
            p_name = row['Product Name']
            cust = row['Customer']
            prod = row['Produced Qty']
            disp = row['Dispatched Qty']
            tot_p += prod
            tot_d += disp
            f.write(f"  {p_name:<40} ({cust}): Produced={prod:7,d} | Dispatched={disp:7,d}\n")
        f.write(f"  {'Total':<40}: Produced={tot_p:7,d} | Dispatched={tot_d:7,d}\n")
        
    f.write("\n=== TOTAL BREAKDOWN (NOV TO JUL) ===\n")
    tot_p = 0
    tot_d = 0
    for idx, row in df_overall.iterrows():
        p_name = row['Product Name']
        cust = row['Customer']
        prod = row['Produced Qty']
        disp = row['Dispatched Qty']
        tot_p += prod
        tot_d += disp
        f.write(f"  {p_name:<40} ({cust}): Produced={prod:9,d} | Dispatched={disp:9,d}\n")
    f.write(f"  {'Total':<40}: Produced={tot_p:9,d} | Dispatched={tot_d:9,d}\n")

print("Breakdown written to:", output_file)
