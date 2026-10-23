"""LibreOffice (UNO) helpers: open, edit, recalculate and save .xlsx workbooks.

LibreOffice is the ground-truth engine for the trace evaluation: it really
recalculates every formula, so a file it saves has correct cached values.
"""
import os
import subprocess
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellRangeAddress, CellAddress

PORT = 2002
_proc = None
_desktop = None


def _prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def start():
    """Start a headless soffice (once) and return the desktop."""
    global _proc, _desktop
    if _desktop is not None:
        return _desktop
    profile = os.path.abspath('lo-profile')  # a private LibreOffice profile in the work folder
    _proc = subprocess.Popen(
        ['soffice', '--headless', '--invisible', '--norestore', '--nologo', '--nodefault',
         f'-env:UserInstallation=file://{profile}',
         f'--accept=socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext'],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
    for _ in range(60):
        try:
            ctx = resolver.resolve(f'uno:socket,host=127.0.0.1,port={PORT};urp;StarOffice.ComponentContext')
            break
        except Exception:
            time.sleep(0.5)
    else:
        raise RuntimeError('soffice did not start')
    _desktop = ctx.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
    return _desktop


def stop():
    global _proc, _desktop
    try:
        if _desktop is not None:
            _desktop.terminate()
    except Exception:
        pass
    if _proc is not None:
        _proc.wait(timeout=20)
    _proc = _desktop = None


def open_doc(path):
    url = uno.systemPathToFileUrl(os.path.abspath(path))
    return start().loadComponentFromURL(url, '_blank', 0, (_prop('Hidden', True),))


def recalc_save(doc, out):
    """Hard-recalculate every formula, then save as .xlsx."""
    doc.calculateAll()
    url = uno.systemPathToFileUrl(os.path.abspath(out))
    doc.storeToURL(url, (_prop('FilterName', 'Calc MS Excel 2007 XML'),))


def close(doc):
    doc.close(True)


def sheet(doc, name):
    return doc.Sheets.getByName(name)


def cell(doc, name, addr):
    return sheet(doc, name).getCellRangeByName(addr)


def set_value(doc, name, addr, value):
    c = cell(doc, name, addr)
    if isinstance(value, str):
        c.setString(value)
    else:
        c.setValue(value)


def api_formula(formula):
    """Excel syntax → LibreOffice API syntax: argument separators are ';'."""
    out, quoted = [], False
    for ch in formula:
        if ch == '"':
            quoted = not quoted
        out.append(';' if ch == ',' and not quoted else ch)
    return ''.join(out)


def set_formula(doc, name, addr, formula):
    c = cell(doc, name, addr)
    c.setFormula(api_formula(formula))
    doc.calculateAll()
    if c.getError():
        raise ValueError(f'formula edit {name}!{addr} = {formula} gives LibreOffice error {c.getError()}')


def insert_rows(doc, name, row0, count=1):
    """Insert `count` rows before 0-based row `row0`; formulas adjust, as in Excel."""
    sheet(doc, name).Rows.insertByIndex(row0, count)


def delete_rows(doc, name, row0, count=1):
    sheet(doc, name).Rows.removeByIndex(row0, count)


def used_cells(doc):
    """Yield (sheet, addr, kind, formula, value) for every non-empty cell."""
    for s in doc.Sheets:
        cur = s.createCursor()
        cur.gotoEndOfUsedArea(False)
        last_r, last_c = cur.RangeAddress.EndRow, cur.RangeAddress.EndColumn
        if last_r > 5000:
            last_r = 5000
        data = s.getCellRangeByPosition(0, 0, last_c, last_r)
        formulas = data.getFormulaArray()
        for r, row in enumerate(formulas):
            for c, f in enumerate(row):
                if f == '':
                    continue
                yield s.Name, r, c, f


def dependents(doc, name, addr):
    """Cells that depend on this cell, transitively (LibreOffice's own analysis)."""
    c = cell(doc, name, addr)
    out = []
    for rng in c.queryDependents(True):
        a = rng.RangeAddress
        sh = doc.Sheets.getByIndex(a.Sheet).Name
        for r in range(a.StartRow, a.EndRow + 1):
            for col in range(a.StartColumn, a.EndColumn + 1):
                out.append((sh, r, col))
    return out
