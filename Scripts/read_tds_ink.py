import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')

for pg in [6, 7]:
    i = pg - 1
    text = doc[i].get_text()
    print(f'\n{"="*60}')
    print(f'Page {pg}')
    print(f'{"="*60}')
    print(text)
