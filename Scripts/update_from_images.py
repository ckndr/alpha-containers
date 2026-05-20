import os
import json
import re
import ast
import glob
import pandas as pd
from PIL import Image
from google import genai

# Scripts live in Tubex/Scripts/
# All data files (Excel, images) live in Tubex/ (one level up)
DATA_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# --- CONFIGURATION ---
API_KEY = "AIzaSyCiFKtV61J3Of22GzUR0pOAKAG2BMXXAlA"

# Auto-detect latest Tubex Excel in parent folder
_excel_files = sorted(glob.glob(os.path.join(DATA_DIR, "Tubex*.xlsx")))
EXCEL_FILE = _excel_files[-1] if _excel_files else ""

# Image files expected in parent folder (copy them there before running)
PRODUCTION_IMG = os.path.join(DATA_DIR, "Production.jpeg")
DOWNTIME_IMG   = os.path.join(DATA_DIR, "Downtime.jpeg")


def load_aliases_from_main_script():
    aliases = {}
    try:
        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'update_production.py')
        if os.path.exists(script_path):
            with open(script_path, 'r', encoding='utf-8') as f:
                content = f.read()

            match = re.search(r'ALIASES\s*=\s*(\{.*?\n\})', content, re.DOTALL | re.IGNORECASE)
            if match:
                dict_str = match.group(1)
                dict_str = re.sub(r'#.*', '', dict_str)
                try:
                    aliases = ast.literal_eval(dict_str)
                    print(f"Successfully loaded {len(aliases)} aliases from update_production.py")
                except Exception as eval_err:
                    print(f"Could not parse ALIASES dict: {eval_err}")
    except Exception as e:
        print(f"Error loading aliases from script: {e}")

    if not aliases:
        print("Using fallback internal aliases...")
        aliases = {
            ("vivid h.c", "35"): ("V- HC BROWN", "6530"),
            ("vivid hc", "35"): ("V- HC BROWN", "6530"),
            ("hola hair", "30"): ("H.H 100GM", "6515"),
            ("hola hair (varnish)", "30"): ("H.H 100GM", "6515"),
            ("samsol men blue", "25"): ("TUBES MEN BLUE", "6506")
        }
    return aliases


def main():
    if not EXCEL_FILE:
        print(f"Error: No Tubex*.xlsx found in {DATA_DIR}")
        return

    print(f"Excel file: {os.path.basename(EXCEL_FILE)}")

    if not os.path.exists(PRODUCTION_IMG) or not os.path.exists(DOWNTIME_IMG):
        print(f"Error: Make sure both Production.jpeg and Downtime.jpeg are in:")
        print(f"  {DATA_DIR}")
        return

    ALIASES = load_aliases_from_main_script()

    print("Initializing new GenAI Vision Client...")
    client = genai.Client(api_key=API_KEY)

    print("Reading screenshots...")
    img_prod = Image.open(PRODUCTION_IMG)
    img_down = Image.open(DOWNTIME_IMG)

    prompt = """
    You are a data extraction assistant. Analyze the two provided images: a daily production log and a downtime summary pivot table.
    Combine the data based on the 'Machines' name.
    
    Return the data strictly as a raw JSON list of objects. Do not use markdown blocks, just raw JSON.
    Format each object exactly with these keys:
    "Date" (YYYY-MM-DD format based on the images), 
    "Machine", 
    "Customer", 
    "Product Name", 
    "Dia / Volume",
    "Target Quantity" (integer, map this from 'Total Production' or 'Available time' * capacity if standard), 
    "Good Quantity Produced" (integer), 
    "Reject/Scrap Quantity" (integer, map this from 'Wastage'),
    "Mechanical DT" (integer or 0), 
    "Electrical DT" (integer or 0), 
    "Material Shortage DT" (integer or 0),
    "Changeover DT" (integer or 0), 
    "Operations DT" (integer or 0), 
    "Power Shutdown DT" (integer or 0),
    "Gas Shutdown DT" (integer or 0), 
    "Workers Shortage DT" (integer or 0).
    """

    print("Sending images to AI for extraction (this takes 10-15 seconds)...")
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[prompt, img_prod, img_down]
        )
    except Exception as e:
        print(f"API Error: {e}")
        return

    json_text = response.text.strip()
    if json_text.startswith("```json"):
        json_text = json_text[7:-3]
    elif json_text.startswith("```"):
        json_text = json_text[3:-3]

    try:
        extracted_data = json.loads(json_text)
        print(f"Successfully extracted {len(extracted_data)} production rows.")
    except json.JSONDecodeError:
        print("Error parsing AI output. Check the raw response:")
        print(response.text)
        return

    df_new = pd.DataFrame(extracted_data)

    if 'Machines' in df_new.columns:
        df_new.rename(columns={'Machines': 'Machine'}, inplace=True)

    df_new['Date'] = pd.to_datetime(df_new['Date'], errors='coerce').dt.date

    good = pd.to_numeric(df_new.get('Good Quantity Produced', 0), errors='coerce').fillna(0)
    reject = pd.to_numeric(df_new.get('Reject/Scrap Quantity', 0), errors='coerce').fillna(0)
    total = good + reject
    df_new['Waste%'] = (reject / total.replace(0, pd.NA)).fillna(0)

    print(f"Loading {os.path.basename(EXCEL_FILE)} for PID and Customer mapping...")
    try:
        df_catalog = pd.read_excel(EXCEL_FILE, sheet_name="Product_Catalog", header=1)
        pid_to_customer = pd.Series(
            df_catalog['Customer'].values,
            index=df_catalog['Product ID'].dropna().astype(str).str.strip()
        ).to_dict()
    except Exception as e:
        print(f"Error reading {EXCEL_FILE}: {e}")
        return

    def get_pid_and_customer(row):
        prod_name = str(row.get('Product Name', '')).strip().lower()
        dia = str(row.get('Dia / Volume', '')).strip().lower()
        ai_customer = str(row.get('Customer', '')).strip()

        pid = None

        alias_key = (prod_name, dia)
        if alias_key in ALIASES:
            val = ALIASES[alias_key]
            if isinstance(val, (tuple, list)):
                pid = str(val[1])
            else:
                pid = str(val)
        else:
            match = df_catalog[(df_catalog['Product Name'].astype(str).str.strip().str.lower() == prod_name) &
                               (df_catalog['Dia (mm)'].astype(str).str.strip().str.lower() == dia)]
            if not match.empty:
                pid = str(match.iloc[0]['Product ID'])

        if pid and pid != 'nan':
            customer = pid_to_customer.get(pid, ai_customer)
            return pid, customer

        return "", ai_customer

    mapping_results = df_new.apply(get_pid_and_customer, axis=1)
    df_new['Product ID'] = [res[0] for res in mapping_results]
    df_new['Customer'] = [res[1] for res in mapping_results]

    final_cols = [
        'Date', 'Machine', 'Customer', 'Product Name', 'Dia / Volume', 'Product ID',
        'Target Quantity', 'Good Quantity Produced', 'Reject/Scrap Quantity', 'Waste%',
        'Mechanical DT', 'Electrical DT', 'Material Shortage DT', 'Changeover DT',
        'Operations DT', 'Power Shutdown DT', 'Gas Shutdown DT', 'Workers Shortage DT'
    ]

    for col in final_cols:
        if col not in df_new.columns:
            df_new[col] = ""
    df_new = df_new[final_cols]

    int_cols = [
        'Product ID', 'Target Quantity', 'Good Quantity Produced', 'Reject/Scrap Quantity',
        'Mechanical DT', 'Electrical DT', 'Material Shortage DT', 'Changeover DT',
        'Operations DT', 'Power Shutdown DT', 'Gas Shutdown DT', 'Workers Shortage DT'
    ]
    for col in int_cols:
        df_new[col] = pd.to_numeric(df_new[col], errors='coerce').astype('Int64')

    df_new = df_new[
        (df_new['Machine'].astype(str).str.strip() != "") &
        (df_new['Machine'].astype(str).str.lower() != "nan")
    ]

    print("Calculating true row count...")
    try:
        df_existing = pd.read_excel(EXCEL_FILE, sheet_name='Production_Log')
        df_valid = df_existing.dropna(subset=['Date', 'Machine'], how='all')
        true_last_row = len(df_valid) + 1

        print(f"\nExtracted Data to Append ({len(df_new)} rows):")
        print(df_new[['Date', 'Machine', 'Customer', 'Product ID']].to_string(index=False))
        print(f"\nAppending to Excel at row {true_last_row + 1}...")

        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df_new.to_excel(writer, sheet_name='Production_Log', startrow=true_last_row, index=False, header=False)
        print("Done! Excel sheet updated successfully.")

    except PermissionError:
        print(f"Error: {os.path.basename(EXCEL_FILE)} is open. Please close it and try again.")
    except Exception as e:
        print(f"Error saving to Excel: {e}")


if __name__ == "__main__":
    main()
