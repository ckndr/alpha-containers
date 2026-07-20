import sys
sys.path.append("D:/Alpha/scratch")
import pivot_breakdowns

df = pivot_breakdowns.df_monthly

for m in ['2025-11', '2025-12']:
    df_m = df[df['Month'] == m]
    print(f"\n--- Month: {m} ---")
    tot_p = 0
    tot_d = 0
    for idx, row in df_m.iterrows():
        p_name = row['Product Name']
        cust = row['Customer']
        prod = row['Produced Qty']
        disp = row['Dispatched Qty']
        tot_p += prod
        tot_d += disp
        print(f"  {p_name:<40} ({cust}): Produced={prod:7,d} | Dispatched={disp:7,d}")
    print(f"  {'Total':<40}: Produced={tot_p:7,d} | Dispatched={tot_d:7,d}")
