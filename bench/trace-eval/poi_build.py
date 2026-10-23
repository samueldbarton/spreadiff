"""Build cases from real Excel files: pick inputs that have dependents, edit them, recalc."""
import json, os, random, sys, shutil
import lo, build

survey = json.load(open('survey.json'))
picked = [r for r in survey if r['formulas'] >= 8 and r['app'].startswith('Microsoft') and r['cached'] == r['formulas']]
print(len(picked), 'candidate files')
os.makedirs('poi-cases', exist_ok=True)
build.OUT = 'poi-cases'
made = []
for r in picked:
    path = 'poi/test-data/spreadsheet/' + r['file']
    name = r['file'][:-5].replace(' ', '_')
    try:
        doc = lo.open_doc(path)
    except Exception as e:
        print('open failed', name, e); continue
    rnd = random.Random(name)
    inputs, formulas = [], []
    try:
        for (sh, rr, cc, f) in lo.used_cells(doc):
            a = build.addr(rr, cc)
            if f.startswith('='):
                if len(formulas) < 400: formulas.append((sh, a, f))
            else:
                try: v = float(f)
                except ValueError: continue
                inputs.append((sh, a, v))
        rnd.shuffle(inputs); rnd.shuffle(formulas)
        edits = []
        for (sh, a, v) in inputs:
            if len(edits) >= 3: break
            if len(lo.dependents(doc, sh, a)) > 1:
                edits.append(('value', sh, a, round(v * 1.37 + 1, 4)))
        for (sh, a, f) in formulas:
            if len(lo.dependents(doc, sh, a)) > 1 and 'INDIRECT' not in f.upper():
                edits.append(('formula', sh, a, '=(' + f[1:] + ')*1.5'))
                break
    finally:
        lo.close(doc)
    if not edits:
        continue
    try:
        build.build(name, path, edits)
        # Engine round trip: original Excel file vs the same workbook recalculated by LibreOffice.
        shutil.copy(path, os.path.join('poi-cases', name, 'excel-original.xlsx'))
        made.append(name)
    except Exception as e:
        print('build failed', name, repr(e)[:200])
print('made', len(made))
lo.stop()
