import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')

# Only show pages for LUBRIMET, ALULIQUID, 400 9 901, 422 0 903, 422 9 900
target_pages = {
    'LUBRIMET': [],
    'ALULIQUID': [],
    '400 9 901': [],
    '422 0 903': [],
    '422 9 900': [],
}

for i in range(len(doc)):
    text = doc[i].get_text()
    text_upper = text.upper()
    for kw in target_pages:
        if kw.upper() in text_upper:
            target_pages[kw].append(i+1)

print("Pages found for each product:")
for kw, pages in target_pages.items():
    print(f"  {kw}: Pages {pages}")

# Print only LUBRIMET and ALULIQUID pages (the ones we haven't seen yet)
for i in range(len(doc)):
    text = doc[i].get_text()
    text_upper = text.upper()
    if 'LUBRIMET' in text_upper or 'ALULIQUID' in text_upper:
        print(f'\n{"="*60}')
        print(f'Page {i+1}')
        print(f'{"="*60}')
        print(text)
