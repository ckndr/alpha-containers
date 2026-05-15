import json
from pathlib import Path
import pandas as pd
import csv

# ==================================================
# AUTOMATIC FILE SELECTION & DIRECTORY PATH
# ==================================================
TARGET_DIR = Path(r"D:\Alpha\BOMs")

print("=" * 60)
print(f"SCANNING FOLDER: {TARGET_DIR}")
print("=" * 60)

if not TARGET_DIR.exists():
    print(f"❌ ERROR: The folder {TARGET_DIR} does not exist.")
    input("Press Enter to exit...")
    exit()

# Find all XLS files
xls_files = list(TARGET_DIR.glob("*.xls"))

if not xls_files:
    print(f"❌ ERROR: No .xls files found in {TARGET_DIR}")
    input("Press Enter to exit...")
    exit()

all_boms = []

for file_path in xls_files:
    print(f"Processing: {file_path.name}")
    bom_info = {
        "file_name": file_path.name,
        "metadata": {},
        "items": []
    }
    
    rows = []
    
    # Try reading as a true binary .xls file first
    try:
        # header=None ensures we get every row, exactly like a CSV reader
        df = pd.read_excel(file_path, header=None, engine="xlrd")
        df = df.fillna("") # Replace empty NaNs with blank strings
        rows = df.values.tolist()
        
    except Exception as e:
        # FALLBACK: If the ERP system exported a CSV but named it .xls
        try:
            with open(file_path, mode='r', encoding='utf-8-sig', errors='ignore') as f:
                reader = csv.reader(f)
                for r in reader:
                    rows.append(r)
        except Exception as fallback_error:
            print(f"  -> ❌ Failed to read {file_path.name}. Skipping.")
            continue

    in_table = False
    headers = []
    
    for row in rows:
        # Clean row (strip whitespace)
        clean_row = [str(cell).strip() for cell in row]
        
        # Skip completely empty rows
        if not any(clean_row):
            continue
            
        if not in_table:
            # --- METADATA EXTRACTION ---
            for i, cell in enumerate(clean_row):
                if cell.endswith(":") or cell.replace(" ", "") in ["Dia", "Lenght", "Colorscount", "Basecoat", "Cap", "Lecquer"]:
                    value = None
                    for j in range(i + 1, len(clean_row)):
                        if clean_row[j] and not str(clean_row[j]).endswith(":"):
                            value = clean_row[j]
                            break
                    
                    if value:
                        key = str(cell).strip(" :")
                        bom_info["metadata"][key] = value
                        
            # Detect the start of the formulation table
            if "Sr #" in clean_row:
                in_table = True
                headers = [col for col in clean_row if col] 
        else:
            # --- TABLE EXTRACTION ---
            if clean_row[0]:
                item_data = {}
                actual_values = [val for val in clean_row if val]
                
                for idx, val in enumerate(actual_values):
                    if idx < len(headers):
                        item_data[headers[idx]] = val
                        
                if item_data:
                    bom_info["items"].append(item_data)
                    
    all_boms.append(bom_info)

output_file = TARGET_DIR / "Bom_scrap.json"

with open(output_file, 'w', encoding='utf-8') as f:
    json.dump(all_boms, f, indent=4)

print("=" * 60)
print(f"SUCCESS: Processed {len(xls_files)} BOM files.")
print(f"Saved compiled data to: {output_file}")
print("=" * 60)
input("Press Enter to close...")
