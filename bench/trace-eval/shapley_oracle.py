"""LibreOffice as the oracle for Shapley contributions.

For each fixture: every subset of its edits is undone in the new workbook,
LibreOffice recalculates, and each traced cell is read. Shapley values from
those runs are compared with spreadiff's.
Usage: python3 shapley_oracle.py SPEC.json WORK_DIR (exits 1 on any mismatch)
"""
import json, math, os, sys
from itertools import combinations
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
spec_path = os.path.abspath(sys.argv[1])
os.makedirs(sys.argv[2], exist_ok=True)
os.chdir(sys.argv[2])
import lo

spec = json.load(open(spec_path))
tot_ok = tot = 0
for case in spec:
    edits, cells = case['edits'], case['cells']
    k = len(edits)
    f = {}
    for mask in range(1 << k):
        doc = lo.open_doc(case['file'])
        try:
            for i, e in enumerate(edits):
                if mask & (1 << i):
                    continue  # applied: as in the new file
                if e['kind'] == 'formula' and e['oldFormula']:
                    c = lo.cell(doc, e['sheet'], e['cell'])
                    c.setFormula(lo.api_formula(e['oldFormula']))
                elif e['oldValue'] is None:
                    lo.cell(doc, e['sheet'], e['cell']).setString('')
                else:
                    lo.set_value(doc, e['sheet'], e['cell'], e['oldValue'])
            doc.calculateAll()
            f[mask] = [lo.cell(doc, c['sheet'], c['cell']).getValue() for c in cells]
        finally:
            lo.close(doc)
    full = (1 << k) - 1
    ok = 0
    bad = []
    for ci, c in enumerate(cells):
        change = f[full][ci] - f[0][ci]
        good = True
        for i, e in enumerate(edits):
            phi = 0.0
            for mask in range(1 << k):
                if mask & (1 << i):
                    continue
                s = bin(mask).count('1')
                w = math.factorial(s) * math.factorial(k - s - 1) / math.factorial(k)
                phi += w * (f[mask | (1 << i)][ci] - f[mask][ci])
            ours = c['shares'].get(e.get('key'), 0.0)
            tol = 1e-7 * max(1.0, abs(change), abs(f[full][ci]))
            if abs(phi - ours) > tol:
                good = False
                bad.append(f"{c['sheet']}!{c['cell']} ← {e.get('key')}: ours {ours}, LibreOffice {phi}")
        ok += good
    tot_ok += ok
    tot += len(cells)
    print(f"{case['dir']}: {ok}/{len(cells)} cells match LibreOffice ({k} edits, {1 << k} recalculations)")
    for b in bad[:3]:
        print('   ', b)
lo.stop()
print(f"TOTAL {tot_ok}/{tot} cells: every edit's Shapley value equal to LibreOffice's")
sys.exit(0 if tot and tot_ok == tot else 1)
