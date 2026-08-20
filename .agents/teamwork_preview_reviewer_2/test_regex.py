import re

orders_formula_1 = '=IFERROR(INDEX(MRP!$F$3:$F$100,MATCH(Tubex_Dashboard!F12,MRP!$D$3:$D$100,0)),0)'
orders_formula_2 = '=IFERROR(INDEX(MRP!$F$3:$F$100,MATCH(F12,MRP!$D$3:$D$100,0)),0)'
orders_formula_3 = '=VLOOKUP(F12,MRP!$D$3:$F$50,3,FALSE)'

r = 25

def update_orders_formula(formula, r):
    # If the formula is a standard template, rewrite cleanly:
    # Or use regex that replaces only relative F{row} or D{row} within the active sheet
    # Pattern: Match optional (Tubex_Dashboard!) followed by F or D and digits, ensuring NOT preceded by '$' or another sheet like 'MRP!'
    pattern = r'((?:Tubex_Dashboard!)?)(?<!\$)(?<!\w)(?<![A-Za-z0-9_]!)([FD])(\d+)\b'
    # Wait, let's test a simple parser or targeted substitution:
    # Notice: MRP!$D$3:$F$50 has $D and $F (preceded by $).
    # If not preceded by $, and either preceded by start/operator or Tubex_Dashboard!
    return re.sub(r'(Tubex_Dashboard!|\b)([FD])(\d+)\b', 
                  lambda m: f"{m.group(1)}{m.group(2)}{r}" if not (m.start() > 0 and formula[m.start()-1] == '$') and not (m.start() > 0 and formula[m.start()-1] == '!' and not m.group(1)) else m.group(0), 
                  formula)

print('Test 1:', update_orders_formula(orders_formula_1, 25))
print('Test 2:', update_orders_formula(orders_formula_2, 25))
print('Test 3:', update_orders_formula(orders_formula_3, 25))
