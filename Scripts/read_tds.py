import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import subprocess
try:
    import fitz
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pymupdf'])
    import fitz

doc = fitz.open(r'd:\Alpha\TDS.pdf')
print('Total Pages:', len(doc))
for i in range(len(doc)):
    print(f'\n--- Page {i+1} ---')
    text = doc[i].get_text()
    print(text)
