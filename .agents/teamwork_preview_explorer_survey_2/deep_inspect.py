import json
import os

with open(r'd:\Alpha\.agents\teamwork_preview_explorer_survey_2\raw_audit_findings.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

print("=== AUDIT SUMMARY ===")
for fpath, res in data.items():
    print('==============================')
    print('FILE:', fpath)
    if 'error' in res:
        print('  ERROR:', res['error'])
        continue
    for sname, sinfo in res['sheets'].items():
        print(f"  Sheet: {sname} (max_row={sinfo['max_row']}, max_col={sinfo['max_col']}, formulas={sinfo['formula_count']})")
    print(f"  Broken formulas (in formula text): {len(res['broken_formulas'])}")
    for bf in res['broken_formulas']:
        print(f"    [{bf['sheet']}!{bf['cell']}] {bf['formula']}")
    print(f"  Cached error values: {len(res['cached_errors'])}")
    for ce in res['cached_errors'][:20]:
        print(f"    [{ce['sheet']}!{ce['cell']}] Val: {ce['cached_value']} | Formula: {ce['formula']}")
    if len(res['cached_errors']) > 20:
        print(f"    ... and {len(res['cached_errors']) - 20} more cached errors")
    print(f"  External links: {len(res['external_links'])}")
    for el in res['external_links'][:15]:
        print(f"    [{el['sheet']}!{el['cell']}] Target: {el['external_target']} | Formula: {el['formula']}")
    if len(res['external_links']) > 15:
        print(f"    ... and {len(res['external_links']) - 15} more external links")
    print(f"  Volatile formulas: {len(res['volatile_formulas'])}")
    for vf in res['volatile_formulas'][:10]:
        print(f"    [{vf['sheet']}!{vf['cell']}] {vf['formula']}")
    if len(res['volatile_formulas']) > 10:
        print(f"    ... and {len(res['volatile_formulas']) - 10} more volatile formulas")
