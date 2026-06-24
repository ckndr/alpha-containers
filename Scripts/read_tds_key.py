import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')

# Print the product data sheets (not the SDS safety sheets)
pages_to_print = [8, 22]  # LUBRIMET GR8 product info, ALULIQUID 13 product info

for pg in pages_to_print:
    i = pg - 1  # 0-indexed
    text = doc[i].get_text()
    print(f'\n{"="*60}')
    print(f'Page {pg}')
    print(f'{"="*60}')
    print(text)
