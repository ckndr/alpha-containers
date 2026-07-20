import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
disp_path = "D:/Alpha/Tubex Records/dispatch nov to jul.xls"

def inspect_file(file_path, sheet_name):
    print(f"\n=================== File: {file_path} (Sheet: {sheet_name}) ===================")
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    print("Shape:", df.shape)
    for idx, row in df.head(40).iterrows():
        # drop nan to see what's actually there
        non_nan = row.dropna().tolist()
        print(f"Row {idx:2d}: {non_nan}")

inspect_file(prod_path, 'Prod_Rp_DPR_Detail_Tubx.rpt')
inspect_file(disp_path, 'Dly_Rpt_PWR0.rpt')
