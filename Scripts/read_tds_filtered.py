import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')
print('Total Pages:', len(doc))

# Search for specific product pages
keywords = [
    'LUBRIMET', 'ALULIQUID', '400 9 901', '422 0 903', '422 9 900',
    'OPV GLOSSY', 'INT PROT', 'CLEANER', 'SunAltec', 'ALULIQUID 13',
    'wet film weight', 'application rate', 'consumption', 'dosage'
]

for i in range(len(doc)):
    text = doc[i].get_text()
    text_upper = text.upper()
    
    # Check if any of our BOM products are on this page
    found = []
    for kw in keywords:
        if kw.upper() in text_upper:
            found.append(kw)
    
    if found:
        print(f'\n{"="*60}')
        print(f'Page {i+1} - Found keywords: {found}')
        print(f'{"="*60}')
        print(text)
