import zipfile, re, glob, json, sys, collections
rows = []
for p in sorted(glob.glob('poi/test-data/spreadsheet/*.xlsx')):
    try:
        z = zipfile.ZipFile(p)
    except Exception as e:
        continue
    names = z.namelist()
    try:
        [z.read(n) for n in names if n.endswith('.xml')]
    except Exception:
        continue  # deliberately corrupt test files
    app = ''
    if 'docProps/app.xml' in names:
        m = re.search(rb'<Application>([^<]*)</Application>', z.read('docProps/app.xml'))
        app = m.group(1).decode(errors='replace') if m else ''
    wbx = z.read('xl/workbook.xml') if 'xl/workbook.xml' in names else b''
    calc = re.search(rb'<calcPr[^>]*>', wbx)
    calc = calc.group(0).decode() if calc else ''
    f = fv = arr = shared = dt = ext = dyn = 0
    for n in names:
        if not re.match(r'xl/worksheets/sheet\d+\.xml', n): continue
        x = z.read(n)
        for m in re.finditer(rb'<c\b[^>]*>(<f\b[^>]*/>|<f\b[^>]*>[^<]*</f>)(<v>)?', x):
            body = m.group(0)
            f += 1
            if m.group(2): fv += 1
            if b't="array"' in body: arr += 1
            if b't="shared"' in body: shared += 1
            if b't="dataTable"' in body: dt += 1
            if re.search(rb'\[\d+\]', body): ext += 1
            if b'_xlfn' in body: dyn += 1
    rows.append(dict(file=p.split('/')[-1], app=app, formulas=f, cached=fv, array=arr, shared=shared, dataTable=dt, external=ext, xlfn=dyn, manual='calcMode="manual"' in calc, fullCalcOnLoad='fullCalcOnLoad="1"' in calc, size=len(open(p,'rb').read())))
json.dump(rows, open('survey.json','w'), indent=1)
withf = [r for r in rows if r['formulas']]
print('files', len(rows), 'with formulas', len(withf))
print('apps (files with formulas):', collections.Counter(r['app'].split(' ')[0] or '?' for r in withf).most_common(8))
print('formula cells', sum(r['formulas'] for r in withf), 'with cached value', sum(r['cached'] for r in withf))
print('files with uncached formulas', sum(1 for r in withf if r['cached'] < r['formulas']))
for k in ['array','shared','dataTable','external','xlfn']: print(k, 'files', sum(1 for r in withf if r[k]), 'cells', sum(r[k] for r in withf))
print('manual calc files', sum(r['manual'] for r in withf), 'fullCalcOnLoad', sum(r['fullCalcOnLoad'] for r in withf))
big = sorted(withf, key=lambda r: -r['formulas'])[:12]
for r in big: print(r['formulas'], r['cached'], r['app'][:22], r['file'])
