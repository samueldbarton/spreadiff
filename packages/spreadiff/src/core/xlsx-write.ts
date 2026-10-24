/**
 * Write review decisions back into an .xlsx: the new file, with each
 * rejected cell change set back to its old value or formula. Everything
 * else in the file (styles, charts, other sheets) is left as it was.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { columnLabel } from './csv.js';
import { isoToSerial } from './serial.js';
import { attr, decode, decodeRef, eachElement, resolvePath } from './xml.js';
import { tokenizeFormula } from './formula.js';
import { externalSource } from './external.js';
import { parseXlsx } from './xlsx.js';
import { diffWorkbooks } from './diff.js';
import { traceWorkbook } from './trace.js';
import { ROW_STATUS, type CellValue, type TableDiff, type WorkbookDiff } from './types.js';
import type { SheetDecisions } from './review.js';

/** A decision {@link applyDecisionsToXlsx} couldn't write, and why. */
export interface SkippedDecision {
  /** Sheet name. */
  sheet: string;
  /** The cell (`B5`), row (`row 7`) or column (`column Price`). */
  target: string;
  /** Why it wasn't applied. */
  reason: string;
}

const xmlEscape = (s: string) => s.replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'));
const ERROR = /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/;

/** Cell XML for a value or formula, keeping the existing cell's style. */
function cellXml(p: string, ref: string, style: string | undefined, value: CellValue, formula: string | null, date1904: boolean): string {
  const s = style !== undefined ? ` s="${style}"` : '';
  if (formula) return `<${p}c r="${ref}"${s}><${p}f>${xmlEscape(formula.replace(/^=/, ''))}</${p}f></${p}c>`;
  if (value === null || value === '') return `<${p}c r="${ref}"${s}/>`;
  if (typeof value === 'number') return `<${p}c r="${ref}"${s}><${p}v>${value}</${p}v></${p}c>`;
  if (typeof value === 'boolean') return `<${p}c r="${ref}"${s} t="b"><${p}v>${value ? 1 : 0}</${p}v></${p}c>`;
  if (ERROR.test(value)) return `<${p}c r="${ref}"${s} t="e"><${p}v>${value}</${p}v></${p}c>`;
  // Dates are read as ISO text: write the serial back, so the cell's date format applies.
  const serial = /^\d/.test(value) ? isoToSerial(value, date1904) : null;
  if (serial !== null) return `<${p}c r="${ref}"${s}><${p}v>${serial}</${p}v></${p}c>`;
  return `<${p}c r="${ref}"${s} t="inlineStr"><${p}is><${p}t xml:space="preserve">${xmlEscape(value)}</${p}t></${p}is></${p}c>`;
}

/** Sheet name → its part, from workbook.xml and its relationships. */
function sheetParts(files: Record<string, Uint8Array>): Map<string, string> {
  const text = (p: string) => (files[p] ? strFromU8(files[p]) : '');
  const strip = (x: string) => x.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');
  const rels = new Map<string, string>();
  eachElement(strip(text('xl/_rels/workbook.xml.rels')), 'Relationship', (a) => {
    const id = attr(a, 'Id');
    const target = attr(a, 'Target');
    if (id && target) rels.set(id, target.startsWith('/') ? target.slice(1) : resolvePath('xl/workbook.xml', target));
  });
  const out = new Map<string, string>();
  const lower = new Map(Object.keys(files).map((k) => [k.toLowerCase(), k]));
  eachElement(strip(text('xl/workbook.xml')), 'sheet', (a) => {
    const name = attr(a, 'name');
    const rid = attr(a, 'r:id') ?? /\s[\w.-]+:id\s*=\s*["']([^"']*)/.exec(a)?.[1];
    const path = rid ? rels.get(rid) : undefined;
    if (name && path) out.set(decode(name), files[path] ? path : (lower.get(path.toLowerCase()) ?? path));
  });
  return out;
}

/**
 * Apply review decisions to the new .xlsx: every rejected cell change is
 * written back to its old value or formula, in place, keeping the cell's
 * style. Formulas are recalculated when the file is next opened (the file
 * is marked for a full recalculation and its calculation chain removed).
 *
 * Row and column decisions can't be written this way (inserting or
 * deleting rows would mean rewriting every reference in the workbook);
 * they're listed in `skipped`, as are cells inside array formulas or table
 * headers. Apply them to tables with `applyDecisions` instead.
 *
 * `decisions[i]` is for `diff.sheets[i]`; leave entries out for sheets
 * without decisions.
 */
export function applyDecisionsToXlsx(
  newFile: ArrayBuffer | Uint8Array,
  diff: WorkbookDiff,
  decisions: (SheetDecisions | undefined)[],
  options: { recalculate?: boolean } = {},
): { file: Uint8Array; applied: number; recalculated: number; skipped: SkippedDecision[] } {
  const original = newFile instanceof Uint8Array ? newFile : new Uint8Array(newFile);
  const files = unzipSync(original);
  const parts = sheetParts(files);
  const skipped: SkippedDecision[] = [];
  const date1904 = diff.date1904?.new ?? false;
  let applied = 0;
  diff.sheets.forEach((sheet, si) => {
    const dec = decisions[si];
    if (!dec) return;
    const d = sheet.diff;
    const w = d.columns.length;
    const sheetRow = (r: number) => (d.newRows[r] >= 0 ? d.newRows[r] + (d.newTable.origin ?? { row: 1 }).row : d.oldRows[r] + (d.oldTable.origin ?? { row: 1 }).row) + 1;
    for (const [r, v] of dec.rows) if (v === 'reject') skipped.push({ sheet: sheet.name, target: `row ${sheetRow(r)}`, reason: d.status[r] === ROW_STATUS.added ? 'deleting a row would shift the references to the rows below it' : 'restoring a row would shift the references to the rows below it' });
    for (const [c, v] of dec.columns) if (v === 'reject') skipped.push({ sheet: sheet.name, target: `column ${d.columns[c].name}`, reason: 'columns can only be undone in a table (applyDecisions)' });
    const rejected = [...dec.cells].filter(([, v]) => v === 'reject').map(([k]) => k);
    if (!rejected.length) return;
    const path = parts.get(sheet.name);
    if (!path || !files[path]) {
      for (const k of rejected) skipped.push({ sheet: sheet.name, target: `cell ${k}`, reason: 'sheet not found in the new file' });
      return;
    }
    let xml = strFromU8(files[path]);
    const p = /<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(xml)?.[1] ?? '';
    const nt = d.newTable;
    const ot = d.oldTable;
    const no = nt.origin ?? { row: 1, col: 0 };
    const headers = (diff.tables?.new ?? []).filter((t) => t.sheet === sheet.name);
    const unshared = new Set<string>();
    for (const k of rejected) {
      const r = Math.floor(k / w);
      const ci = k % w;
      const col = d.columns[ci];
      const ni = d.newRows[r];
      const oi = d.oldRows[r];
      if (ni < 0 || oi < 0 || col.newIndex < 0 || col.oldIndex < 0) continue;
      const row0 = ni + no.row;
      const col0 = col.newIndex + no.col;
      const ref = `${columnLabel(col0)}${row0 + 1}`;
      if (nt.arrays?.some((b) => row0 >= b.r0 && row0 <= b.r1 && col0 >= b.c0 && col0 <= b.c1)) {
        skipped.push({ sheet: sheet.name, target: ref, reason: 'part of an array formula' });
        continue;
      }
      if (headers.some((t) => t.headerRows > 0 && row0 >= t.range.r0 && row0 < t.range.r0 + t.headerRows && col0 >= t.range.c0 && col0 <= t.range.c1)) {
        skipped.push({ sheet: sheet.name, target: ref, reason: 'a table header (the table would no longer match it)' });
        continue;
      }
      const value = ot.rows[oi][col.oldIndex] ?? null;
      const oldFormula = ot.formulas?.[oi]?.[col.oldIndex] ?? null;
      const newFormula = nt.formulas?.[ni]?.[col.newIndex] ?? null;
      if (oldFormula && oldFormula === newFormula) {
        // Same formula, new value: outside data keeps the old saved value; anything else is a downstream result.
        if (externalSource(oldFormula)) {
          xml = setResult(xml, p, ref, value);
          applied++;
        } else skipped.push({ sheet: sheet.name, target: ref, reason: 'a downstream result: it follows its inputs; reject the edits that moved it' });
        continue;
      }
      // The old formula's references, moved to where those cells are now.
      const formula = oldFormula ? translateFormula(oldFormula, sheet.name, diff) : null;
      if (oldFormula && formula === null) {
        skipped.push({ sheet: sheet.name, target: ref, reason: `its old formula ${oldFormula} reads a cell the new file no longer has` });
        continue;
      }
      // The existing cell, if any.
      const cellRe = new RegExp(`<${p}c\\b[^>]*?\\sr\\s*=\\s*["']${ref}["'][^>]*?(?:/>|>[\\s\\S]*?</${p}c>)`);
      const m = cellRe.exec(xml);
      if (m) {
        const open = /^<[^>]*/.exec(m[0])![0];
        const style = attr(open, 's');
        // A shared formula's master: give every cell of its group its own formula first.
        const si = /<(?:[\w.-]+:)?f\b[^>]*\bt\s*=\s*["']shared["'][^>]*\bsi\s*=\s*["'](\d+)["']/.exec(m[0])?.[1] ?? /<(?:[\w.-]+:)?f\b[^>]*\bsi\s*=\s*["'](\d+)["'][^>]*\bt\s*=\s*["']shared["']/.exec(m[0])?.[1];
        if (si !== undefined && !unshared.has(si)) {
          xml = unshare(xml, p, si, (rr, cc) => nt.formulas?.[rr - no.row]?.[cc - no.col] ?? null);
          unshared.add(si);
        }
        const again = cellRe.exec(xml)!;
        xml = xml.slice(0, again.index) + cellXml(p, ref, style, value, formula, date1904) + xml.slice(again.index + again[0].length);
      } else {
        xml = insertCell(xml, p, row0, col0, cellXml(p, ref, undefined, value, formula, date1904));
      }
      applied++;
    }
    files[path] = strToU8(xml);
  });
  if (!applied) return { file: original, applied, recalculated: 0, skipped };
  markForRecalc(files);
  let file = zipSync(files, { level: 6 });
  // Apps that read the saved values (LibreOffice by default, Google Sheets,
  // pandas) don't recalculate on open: write the results that moved too.
  const recalculated = options.recalculate === false ? 0 : refreshResults(files, parts, original, file);
  if (recalculated) file = zipSync(files, { level: 6 });
  return { file, applied, recalculated, skipped };
}

/**
 * Recalculate what the written cells feed, and store the new results as the
 * formulas' saved values. Uses tracing between the new file and the patched
 * one: only formulas downstream of the written cells, and only ones the
 * engine reproduces (calibration), are updated. Returns how many.
 */
function refreshResults(files: Record<string, Uint8Array>, parts: Map<string, string>, before: Uint8Array, after: Uint8Array): number {
  const a = parseXlsx(before, { computeMissing: false });
  const b = parseXlsx(after, { computeMissing: false });
  const wd = diffWorkbooks(a, b, { keys: 'none' });
  const trace = traceWorkbook(wd);
  let n = 0;
  wd.sheets.forEach((sheet, si) => {
    const stale = [...trace.sheet(si)].filter(([, t]) => t.stale);
    const path = parts.get(sheet.name);
    if (!stale.length || !path || !files[path]) return;
    const d = sheet.diff;
    const w = d.columns.length;
    const o = d.newTable.origin ?? { row: 1, col: 0 };
    let xml = strFromU8(files[path]);
    const p = /<([A-Za-z_][\w.-]*:)?worksheet\b/.exec(xml)?.[1] ?? '';
    for (const [k, t] of stale) {
      const ni = d.newRows[Math.floor(k / w)];
      const ci = d.columns[k % w].newIndex;
      if (ni < 0 || ci < 0) continue;
      const next = setResult(xml, p, `${columnLabel(ci + o.col)}${ni + o.row + 1}`, t.stale!.recalculated);
      if (next !== xml) n++;
      xml = next;
    }
    files[path] = strToU8(xml);
  });
  return n;
}

/** Set a formula cell's saved result (`<v>` and its type), leaving its formula and style. */
function setResult(xml: string, p: string, ref: string, value: CellValue): string {
  const re = new RegExp(`<${p}c\\b[^>]*?\\sr\\s*=\\s*["']${ref}["'][^>]*?>[\\s\\S]*?</${p}c>`);
  const m = re.exec(xml);
  if (!m || !new RegExp(`<${p}f\\b`).test(m[0])) return xml;
  const open = /^<[^>]*>/.exec(m[0])![0];
  const inner = m[0].slice(open.length, -`</${p}c>`.length).replace(new RegExp(`<${p}v\\b[^>]*?(?:/>|>[\\s\\S]*?</${p}v>)`), '');
  let type = '';
  let v: string | null;
  if (value === null) v = null;
  else if (typeof value === 'number') v = String(value);
  else if (typeof value === 'boolean') ((type = 'b'), (v = value ? '1' : '0'));
  else if (ERROR.test(value)) ((type = 'e'), (v = value));
  else ((type = 'str'), (v = xmlEscape(value)));
  const attrs = open.slice(0, -1).replace(/\st\s*=\s*["'][^"']*["']/, '') + (type ? ` t="${type}"` : '');
  return xml.slice(0, m.index) + `${attrs}>${inner}${v === null ? '' : `<${p}v>${v}</${p}v>`}</${p}c>` + xml.slice(m.index + m[0].length);
}

/** Give each cell of shared-formula group `si` its own formula, from the parsed table. */
function unshare(xml: string, p: string, si: string, formulaAt: (row: number, col: number) => string | null): string {
  // A cell that isn't self-closing, up to its own </c>, holding an <f> of the group.
  const re = new RegExp(`<${p}c\\b(?:[^>]*[^/>])?>(?:(?!</${p}c>|<${p}c\\b)[\\s\\S])*?<${p}f\\b[^>]*\\bsi\\s*=\\s*["']${si}["'][^>]*?(?:/>|>[^<]*</${p}f>)(?:(?!<${p}c\\b)[\\s\\S])*?</${p}c>`, 'g');
  return xml.replace(re, (cell) => {
    if (!/\bt\s*=\s*["']shared["']/.test(cell)) return cell;
    const ref = attr(/^<[^>]*/.exec(cell)![0], 'r');
    if (!ref) return cell;
    const [r, c] = decodeRef(ref);
    const f = formulaAt(r, c);
    if (!f) return cell;
    return cell.replace(new RegExp(`<${p}f\\b[^>]*?(?:/>|>[^<]*</${p}f>)`), `<${p}f>${xmlEscape(f.replace(/^=/, ''))}</${p}f>`);
  });
}

/** Insert a cell into its row (creating the row if needed), in column order. */
function insertCell(xml: string, p: string, row0: number, col0: number, cell: string): string {
  const rowRe = new RegExp(`<${p}row\\b[^>]*?\\sr\\s*=\\s*["']${row0 + 1}["'][^>]*?(/>|>[\\s\\S]*?</${p}row>)`);
  const m = rowRe.exec(xml);
  if (m) {
    if (m[1] === '/>') return xml.slice(0, m.index) + m[0].slice(0, -2) + `>${cell}</${p}row>` + xml.slice(m.index + m[0].length);
    // Before the first cell to its right.
    const body = m[0];
    const cells = [...body.matchAll(new RegExp(`<${p}c\\b[^>]*?\\sr\\s*=\\s*["']([A-Z]+)\\d+["']`, 'g'))];
    const after = cells.find((c) => decodeRef(`${c[1]}1`)[1] > col0);
    const at = after ? after.index! : body.lastIndexOf(`</${p}row>`);
    return xml.slice(0, m.index) + body.slice(0, at) + cell + body.slice(at) + xml.slice(m.index + body.length);
  }
  const rows = [...xml.matchAll(new RegExp(`<${p}row\\b[^>]*?\\sr\\s*=\\s*["'](\\d+)["']`, 'g'))];
  const next = rows.find((r) => Number(r[1]) > row0 + 1);
  const rowXml = `<${p}row r="${row0 + 1}">${cell}</${p}row>`;
  if (next) return xml.slice(0, next.index!) + rowXml + xml.slice(next.index!);
  const end = xml.indexOf(`</${p}sheetData>`);
  if (end >= 0) return xml.slice(0, end) + rowXml + xml.slice(end);
  // An empty sheet: <sheetData/>.
  return xml.replace(new RegExp(`<${p}sheetData\\s*/>`), `<${p}sheetData>${rowXml}</${p}sheetData>`);
}

/** Ask Excel to recalculate everything on open, and drop the calculation chain (it lists formulas that may be gone). */
function markForRecalc(files: Record<string, Uint8Array>): void {
  const wb = strFromU8(files['xl/workbook.xml']);
  const p = /<([A-Za-z_][\w.-]*:)?workbook\b/.exec(wb)?.[1] ?? '';
  let next: string;
  const calc = new RegExp(`<${p}calcPr\\b([^>]*?)(/?)>`);
  if (calc.test(wb)) next = wb.replace(calc, (_m, attrs: string, close: string) => `<${p}calcPr${attrs.replace(/\sfullCalcOnLoad\s*=\s*["'][^"']*["']/, '')} fullCalcOnLoad="1"${close}>`);
  else {
    // The schema fixes the order: calcPr comes right after definedNames (or what precedes it).
    const after = ['definedNames', 'externalReferences', 'functionGroups', 'sheets']
      .map((t) => {
        const close = wb.indexOf(`</${p}${t}>`);
        return close >= 0 ? close + `</${p}${t}>`.length : -1;
      })
      .find((i) => i >= 0);
    next = after === undefined ? wb : wb.slice(0, after) + `<${p}calcPr fullCalcOnLoad="1"/>` + wb.slice(after);
  }
  files['xl/workbook.xml'] = strToU8(next);
  if (files['xl/calcChain.xml']) {
    delete files['xl/calcChain.xml'];
    const rels = 'xl/_rels/workbook.xml.rels';
    if (files[rels]) files[rels] = strToU8(strFromU8(files[rels]).replace(/<(?:[\w.-]+:)?Relationship\b[^>]*calcChain[^>]*\/>/g, ''));
    const ct = '[Content_Types].xml';
    if (files[ct]) files[ct] = strToU8(strFromU8(files[ct]).replace(/<(?:[\w.-]+:)?Override\b[^>]*calcChain[^>]*\/>/g, ''));
  }
}

/** Old sheet row/column → new, through one sheet's alignment; null for one that was removed. */
function coordinateMap(d: TableDiff): { row(r: number): number | null; col(c: number): number | null } {
  const oo = d.oldTable.origin ?? { row: 1, col: 0 };
  const no = d.newTable.origin ?? { row: 1, col: 0 };
  const rows = new Int32Array(d.oldTable.rows.length).fill(-1);
  for (let a = 0; a < d.rowCount; a++) if (d.oldRows[a] >= 0) rows[d.oldRows[a]] = d.newRows[a];
  const cols = new Int32Array(d.oldTable.columns.length).fill(-1);
  for (const c of d.columns) if (c.oldIndex >= 0) cols[c.oldIndex] = c.newIndex;
  const along = (x: number, origin0: number, origin1: number, map: Int32Array, len1: number) => {
    const t = x - origin0;
    if (t < 0) return x - origin0 + origin1; // above or left of the data: moves with its start
    if (t >= map.length) return x - (origin0 + map.length) + origin1 + len1; // past the end: moves with its end
    return map[t] < 0 ? null : map[t] + origin1;
  };
  return {
    row: (r) => along(r, oo.row, no.row, rows, d.newTable.rows.length),
    col: (c) => along(c, oo.col, no.col, cols, d.newTable.columns.length),
  };
}

/**
 * An old formula with its A1 references moved to the new file's rows and
 * columns (`=B4*2` → `=B5*2` after a row went in above). Null when one
 * points at a cell the new file doesn't have.
 */
export function translateFormula(formula: string, sheet: string, diff: WorkbookDiff): string | null {
  const maps = new Map<string, ReturnType<typeof coordinateMap> | null>();
  const mapFor = (name: string) => {
    const k = name.toUpperCase();
    if (!maps.has(k)) {
      const s = diff.sheets.find((x) => x.name.toUpperCase() === k);
      maps.set(k, s ? coordinateMap(s.diff) : null);
    }
    return maps.get(k)!;
  };
  const t = tokenizeFormula(formula);
  let out = '=';
  for (const tok of t.tokens) {
    if (tok.type !== 'ref' || tok.range.sheetTo !== undefined) {
      out += tok.text;
      continue;
    }
    const m = mapFor(tok.range.sheet ?? sheet);
    if (!m) {
      out += tok.text; // a sheet outside the diff: left as written
      continue;
    }
    const bang = tok.text.lastIndexOf('!');
    const prefix = bang >= 0 ? tok.text.slice(0, bang + 1) : '';
    let failed = false;
    const body = tok.text.slice(prefix.length).replace(/(\$?)([A-Z]{1,3})?(\$?)(\d+)?/gi, (all, d1: string, letters: string | undefined, d2: string, digits: string | undefined) => {
      if (!letters && !digits) return all;
      let text = '';
      if (letters) {
        const c = m.col(decodeRef(`${letters.toUpperCase()}1`)[1]);
        if (c === null) failed = true;
        else text += d1 + columnLabel(c);
      } else text += d1;
      if (digits) {
        const r = m.row(Number(digits) - 1);
        if (r === null) failed = true;
        else text += d2 + String(r + 1);
      } else text += d2;
      return text;
    });
    if (failed) return null;
    out += prefix + body;
  }
  return out;
}
