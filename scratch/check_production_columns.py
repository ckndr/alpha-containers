import pandas as pd

prod_path = "D:/Alpha/Tubex Records/production nov to jul.xls"
df = pd.read_excel(prod_path, header=None)

# Let's inspect rows starting from row 10
data_rows = df.iloc[10:]

# Filter out rows that don't have a valid date in Col 1
valid_rows = data_rows[data_rows[1].apply(lambda x: isinstance(x, pd.Timestamp) or hasattr(x, 'date'))]

print(f"Total valid data rows: {len(valid_rows)}")

# Let's count how many have non-zero or non-NaN in Col 12 (OT Production) and Col 13 (OT Wastage)
ot_prod_non_zero = valid_rows[valid_rows[12] > 0]
print(f"Rows with non-zero Overtime Production (Col 12): {len(ot_prod_non_zero)}")
if len(ot_prod_non_zero) > 0:
    print("Sample OT production rows:")
    print(ot_prod_non_zero[[1, 6, 7, 9, 10, 12, 13]].head(5))

ot_wast_non_zero = valid_rows[valid_rows[13] > 0]
print(f"Rows with non-zero Overtime Wastage (Col 13): {len(ot_wast_non_zero)}")

# Let's see some details on what columns we should extract
# Col 1: Date
# Col 6: Product Name
# Col 8: Dia
# Col 5: POF #
# Col 7: Order Qty (PO Qty)
# Col 9: Normal Good Qty
# Col 10: Normal Wastage
# Col 12: OT Good Qty
# Col 13: OT Wastage
