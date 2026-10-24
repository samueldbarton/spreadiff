/**
 * Workbook objects beyond cells: charts and pivot tables (read from
 * .xlsx), and changes to them and to defined names between two workbooks.
 * A chart pointed at a different range, a pivot reading a new source, a
 * name redefined: none of these change a cell, and each changes what the
 * workbook says.
 */
import { attr, decode, eachElement, resolvePath } from './xml.js';
import type { Workbook } from './types.js';

/** A chart, as {@link Workbook.charts} lists it. */
export interface ChartInfo {
  /** The sheet (or chart sheet) it sits on. */
  sheet: string;
  /** Its title, or `Chart N` (N counts the sheet's charts from 1). */
  title: string;
  /** Chart types, in plot order (`bar`, `line`, … ; several for a combination chart). */
  types: string[];
  /** Data series. */
  series: { name?: string; categories?: string; values?: string }[];
}

/** A pivot table, as {@link Workbook.pivots} lists it. */
export interface PivotInfo {
  /** The sheet it sits on. */
  sheet: string;
  /** Its name (`PivotTable1`). */
  name: string;
  /** Where it sits (`A3:D7`). */
  location: string;
  /** What it summarizes: `Data!A1:C100`, a table or a name, or `external` for an outside connection. */
  source: string;
  /** Its value fields (`Sum of Revenue`). */
  values: string[];
}

/** A change to a defined name, chart or pivot table, from {@link diffWorkbookObjects}. */
export interface ObjectChange {
  /** What it is. */
  kind: 'name' | 'chart' | 'pivot';
  /** `added`, `removed` or `modified`. */
  change: 'added' | 'removed' | 'modified';
  /** The name, `Sheet › chart title`, or `Sheet › PivotTable1`. */
  name: string;
  /** What changed, in words (`series 2 values: Data!$B$2:$B$13 → Data!$B$2:$B$14`). One entry for added or removed objects. */
  details: string[];
}

const strip = (xml: string) => xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, '<$1');

/** The text of an element's `<f>` reference, or its literal `<v>`. */
function refOrValue(body: string | null): string | undefined {
  if (!body) return undefined;
  let out: string | undefined;
  eachElement(body, 'f', (_, b) => (out ??= b ? decode(b) : undefined));
  if (out === undefined) eachElement(body, 'v', (_, b) => (out ??= b ? decode(b) : undefined));
  return out;
}

function parseChart(xml: string, sheet: string, n: number): ChartInfo {
  const x = strip(xml);
  const types: string[] = [];
  for (const m of x.matchAll(/<(\w+?(?:3D)?)Chart[\s>]/g)) if (!types.includes(m[1])) types.push(m[1]);
  // The chart's own title comes before the plot area; titles after it are the axes'.
  const head = x.slice(0, Math.max(0, x.indexOf('<plotArea')) || x.length);
  let title = '';
  eachElement(head, 'title', (_, body) => {
    if (title || !body) return;
    const f = refOrValue(/<tx>([\s\S]*?)<\/tx>/.exec(body)?.[1] ?? null);
    let runs = '';
    eachElement(body, 't', (_a, b) => (runs += b ? decode(b) : ''));
    title = runs || f || '';
  });
  const series: ChartInfo['series'] = [];
  eachElement(x, 'ser', (_, body) => {
    if (!body) return;
    const part = (tag: string) => {
      let v: string | undefined;
      eachElement(body, tag, (_a, b) => (v ??= refOrValue(b)));
      return v;
    };
    const s: ChartInfo['series'][number] = {};
    const name = part('tx');
    const categories = part('cat') ?? part('xVal');
    const values = part('val') ?? part('yVal');
    if (name !== undefined) s.name = name;
    if (categories !== undefined) s.categories = categories;
    if (values !== undefined) s.values = values;
    series.push(s);
  });
  return { sheet, title: title || `Chart ${n}`, types, series };
}

function parsePivot(xml: string, cacheXml: string, sheet: string): PivotInfo {
  const x = strip(xml);
  const c = strip(cacheXml);
  let name = '';
  eachElement(x, 'pivotTableDefinition', (a) => (name ||= decode(attr(a, 'name') ?? '')));
  let location = '';
  eachElement(x, 'location', (a) => (location ||= attr(a, 'ref') ?? ''));
  const values: string[] = [];
  eachElement(x, 'dataField', (a) => values.push(decode(attr(a, 'name') ?? '')));
  let source = 'external';
  eachElement(c, 'worksheetSource', (a) => {
    const ref = attr(a, 'ref');
    const sh = attr(a, 'sheet');
    const nm = attr(a, 'name');
    source = nm ? decode(nm) : `${sh ? decode(sh) + '!' : ''}${ref ?? ''}`;
  });
  return { sheet, name, location, source, values };
}

/**
 * Charts and pivot tables of an .xlsx, by sheet. `text(path)` reads a part
 * of the package; `sheets` maps each sheet's name to its part.
 */
export function parseObjects(text: (path: string) => string, sheets: { name: string; path: string }[]): { charts: ChartInfo[]; pivots: PivotInfo[] } {
  const charts: ChartInfo[] = [];
  const pivots: PivotInfo[] = [];
  const relsOf = (path: string) => {
    const out: { type: string; target: string }[] = [];
    eachElement(strip(text(path.replace(/([^/]+)$/, '_rels/$1.rels'))), 'Relationship', (a) => {
      const target = attr(a, 'Target');
      const type = attr(a, 'Type') ?? '';
      if (target && attr(a, 'TargetMode') !== 'External') out.push({ type: type.slice(type.lastIndexOf('/') + 1), target: resolvePath(path, target) });
    });
    return out;
  };
  for (const { name, path } of sheets) {
    let n = 0;
    for (const r of relsOf(path)) {
      if (r.type === 'drawing') {
        for (const d of relsOf(r.target)) if (d.type === 'chart') charts.push(parseChart(text(d.target), name, ++n));
      } else if (r.type === 'pivotTable') {
        const cache = relsOf(r.target).find((p) => p.type === 'pivotCacheDefinition');
        pivots.push(parsePivot(text(r.target), cache ? text(cache.target) : '', name));
      }
    }
  }
  return { charts, pivots };
}

/** What changed between two charts, in words. */
function chartDetails(a: ChartInfo, b: ChartInfo): string[] {
  const out: string[] = [];
  if (a.types.join() !== b.types.join()) out.push(`type: ${a.types.join(' + ') || '?'} → ${b.types.join(' + ') || '?'}`);
  const n = Math.max(a.series.length, b.series.length);
  for (let i = 0; i < n; i++) {
    const x = a.series[i];
    const y = b.series[i];
    const label = `series ${i + 1}${(y ?? x)?.name ? ` (${(y ?? x)!.name})` : ''}`;
    if (!x) out.push(`${label} added: ${y!.values ?? ''}`);
    else if (!y) out.push(`${label} removed`);
    else
      for (const k of ['name', 'categories', 'values'] as const) {
        if (x[k] !== y[k]) out.push(`${label} ${k}: ${x[k] ?? '—'} → ${y[k] ?? '—'}`);
      }
  }
  return out;
}

/**
 * Changes to defined names, charts and pivot tables between two workbooks.
 * Charts pair by sheet and title (then by position on the sheet), pivots by
 * sheet and name. References are compared as written: a chart whose range
 * moved because rows were inserted shows as changed.
 */
export function diffWorkbookObjects(a: Workbook, b: Workbook): ObjectChange[] {
  const out: ObjectChange[] = [];
  // Defined names: workbook-scoped and sheet-scoped (shown as Sheet!Name).
  const namesOf = (wb: Workbook) => {
    const m = new Map<string, string>(Object.entries(wb.names ?? {}));
    for (const t of wb.sheets) for (const [k, v] of Object.entries(t.names ?? {})) m.set(`${t.name}!${k}`, v);
    return m;
  };
  const na = namesOf(a);
  const nb = namesOf(b);
  for (const [k, v] of nb) {
    const old = na.get(k);
    if (old === undefined) out.push({ kind: 'name', change: 'added', name: k, details: [`= ${v}`] });
    else if (old !== v) out.push({ kind: 'name', change: 'modified', name: k, details: [`${old} → ${v}`] });
  }
  for (const [k, v] of na) if (!nb.has(k)) out.push({ kind: 'name', change: 'removed', name: k, details: [`was ${v}`] });

  const pair = <T,>(xs: T[], ys: T[], key: (t: T) => string, fallback: (t: T, i: number) => string) => {
    const pairs: [T | undefined, T | undefined][] = [];
    const used = new Set<number>();
    const keyed = new Map<string, number[]>();
    ys.forEach((y, i) => keyed.set(key(y), [...(keyed.get(key(y)) ?? []), i]));
    const unmatched: T[] = [];
    for (const x of xs) {
      const i = keyed.get(key(x))?.find((j) => !used.has(j));
      if (i === undefined) unmatched.push(x);
      else {
        used.add(i);
        pairs.push([x, ys[i]]);
      }
    }
    // Untitled or renamed: pair what's left by position on the same sheet.
    const rest = ys.map((y, i) => ({ y, i })).filter(({ i }) => !used.has(i));
    for (const x of unmatched) {
      const k = rest.findIndex(({ y }) => fallback(y, 0) === fallback(x, 0));
      if (k >= 0) {
        pairs.push([x, rest[k].y]);
        rest.splice(k, 1);
      } else pairs.push([x, undefined]);
    }
    for (const { y } of rest) pairs.push([undefined, y]);
    return pairs;
  };
  const ca = a.charts ?? [];
  const cb = b.charts ?? [];
  for (const [x, y] of pair(ca, cb, (c) => `${c.sheet}\u0000${c.title}`, (c) => c.sheet)) {
    const t = (y ?? x)!;
    const name = `${t.sheet} › ${t.title}`;
    if (!x) out.push({ kind: 'chart', change: 'added', name, details: [t.series.map((s) => s.values).filter(Boolean).join(', ')] });
    else if (!y) out.push({ kind: 'chart', change: 'removed', name, details: [x.series.map((s) => s.values).filter(Boolean).join(', ')] });
    else {
      const details = chartDetails(x, y);
      if (x.title !== y.title) details.unshift(`title: ${x.title} → ${y.title}`);
      if (details.length) out.push({ kind: 'chart', change: 'modified', name, details });
    }
  }
  const pa = a.pivots ?? [];
  const pb = b.pivots ?? [];
  for (const [x, y] of pair(pa, pb, (p) => `${p.sheet}\u0000${p.name}`, (p) => p.sheet)) {
    const t = (y ?? x)!;
    const name = `${t.sheet} › ${t.name}`;
    if (!x) out.push({ kind: 'pivot', change: 'added', name, details: [`of ${t.source}`] });
    else if (!y) out.push({ kind: 'pivot', change: 'removed', name, details: [`was of ${x.source}`] });
    else {
      const details: string[] = [];
      if (x.name !== y.name) details.push(`name: ${x.name} → ${y.name}`);
      if (x.source !== y.source) details.push(`source: ${x.source} → ${y.source}`);
      if (x.location !== y.location) details.push(`location: ${x.location} → ${y.location}`);
      if (x.values.join('\u0000') !== y.values.join('\u0000')) details.push(`values: ${x.values.join(', ') || '—'} → ${y.values.join(', ') || '—'}`);
      if (details.length) out.push({ kind: 'pivot', change: 'modified', name, details });
    }
  }
  return out;
}
