/**
 * spreadiff/html — a dependency-free renderer that turns a diff into an HTML
 * string (or plain text). No canvas, no DOM APIs, no workers: render on the
 * server, in an email, a PR comment, or a static page. Shows only changes
 * plus a little context, so it's meant for reviewing edits, not for
 * scrolling a million-row sheet (use `DiffGrid` for that).
 *
 * ```ts
 * import { diffTables, parseCsv, renderHTML, HTML_CSS } from 'spreadiff/html';
 * declare const before: string, after: string, container: HTMLElement;
 *
 * const diff = diffTables(parseCsv(before), parseCsv(after));
 * container.innerHTML = `<style>${HTML_CSS}</style>` + renderHTML(diff, { layout: 'split' });
 * ```
 */
import { columnLabel } from '../core/csv.js';
import { changeKind } from '../core/kind.js';
import { diffText, formatCell, numericDelta } from '../core/text.js';
import type { ChangeKind } from '../core/trace.js';
import { ROW_STATUS, type CellValue, type TableDiff } from '../core/types.js';
import { lightDarkFallback } from '../core/lightdark.js';
import { rowLabelColumn } from '../core/label.js';

export { diffTables, cellsEqual, hasChanges } from '../core/diff.js';
export { parseCsv } from '../core/csv.js';
export { formatCell, numericDelta, diffText } from '../core/text.js';
export { changeKind } from '../core/kind.js';
export type { TableDiff, Table, CellValue, DiffOptions } from '../core/types.js';

/**
 * - `split`: old | new side by side, one row per line, like a code review.
 * - `unified`: one table; modified rows as a − line and a + line, like `git diff`.
 */
export type HtmlLayout = 'split' | 'unified';
/** How a changed cell reads; same meanings as the canvas grid's {@link CellDisplay}. */
export type HtmlDisplay = 'stacked' | 'new' | 'delta' | 'inline';

/** Options for {@link renderHTML}. All are optional. */
export interface HtmlOptions {
  /** `split`: old | new side by side, one row per line. `unified`: −/+ line pairs, like `git diff`. Default `split`. */
  layout?: HtmlLayout;
  /** How changed cells read. Default `stacked` (unified: a − line then a + line). */
  display?: HtmlDisplay;
  /** Unchanged rows shown around each change. Default 2. */
  context?: number;
  /** Only show columns that changed, plus key columns (or, when rows were matched by order, the first mostly-text column, which names the rows). Default false. */
  changedColumnsOnly?: boolean;
  /** Stop after this many rows of output (a note says how many were cut). Default 2000. */
  maxRows?: number;
  /** Override the change kind of a cell (e.g. from a full trace). */
  kindOf?: (row: number, col: number) => ChangeKind | undefined;
  /** Row numbers start at this for row index 0. Default: from the table origin (sheet rows). */
  rowNumberOffset?: number;
}

const esc = (s: string) =>
  s.replace(/[&<>"]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;'));

type Item = { type: 'row'; row: number } | { type: 'gap'; from: number; to: number };

/** Changed rows plus context, with collapsed gaps between them. */
function hunks(diff: TableDiff, context: number): Item[] {
  const n = diff.rowCount;
  const keep = new Uint8Array(n);
  for (let r = 0; r < n; r++) {
    if (diff.status[r] === ROW_STATUS.equal && !diff.moved?.[r]) continue;
    for (let k = Math.max(0, r - context); k <= Math.min(n - 1, r + context); k++) keep[k] = 1;
  }
  const out: Item[] = [];
  let r = 0;
  while (r < n) {
    if (keep[r]) {
      out.push({ type: 'row', row: r++ });
      continue;
    }
    const from = r;
    while (r < n && !keep[r]) r++;
    out.push({ type: 'gap', from, to: r });
  }
  return out;
}

function visibleColumns(diff: TableDiff, changedOnly: boolean): number[] {
  const keys = new Set(diff.keyColumns);
  const label = changedOnly ? rowLabelColumn(diff) : -1; // without keys, keep what names the rows
  const cols: number[] = [];
  diff.columns.forEach((c, i) => {
    if (!changedOnly || c.status !== 'equal' || c.changes > 0 || keys.has(c.name) || i === label) cols.push(i);
  });
  return cols;
}

function deltaLabel(a: CellValue | undefined, b: CellValue | undefined): { text: string; up: boolean } | null {
  const d = numericDelta(a ?? null, b ?? null);
  if (!d || d.delta === 0) return null;
  const up = d.delta > 0;
  const s = up ? '+' : '−';
  const abs = formatCell(Math.abs(d.delta));
  if (d.ratio === null) return { text: `${s}${abs}`, up };
  const pct = Math.abs(d.ratio * 100);
  return { text: `${s}${abs} · ${s}${pct >= 100 ? pct.toFixed(0) : pct >= 10 ? pct.toFixed(1) : pct.toFixed(2)}%`, up };
}

/** Old/new text with only the differing characters wrapped, like intra-line code highlights. */
const NUMBERISH = /^[-+]?[\d.,]+%?$/;

function intra(before: string, after: string, side: 'old' | 'new'): string {
  if (before.length + after.length > 400) return esc(side === 'old' ? before : after);
  // Digit-level highlights in numbers (marking "248" inside 124800) are noise.
  if (NUMBERISH.test(before) && NUMBERISH.test(after)) return `<mark>${esc(side === 'old' ? before : after)}</mark>`;
  const spans = diffText(before, after);
  // Mostly different: highlighting fragments is noise; mark the whole value.
  const same = spans.filter((s) => s.kind === 'equal').reduce((n, s) => n + s.text.length, 0);
  if (same < Math.min(before.length, after.length) * 0.4) return `<mark>${esc(side === 'old' ? before : after)}</mark>`;
  return spans
    .filter((s) => s.kind === 'equal' || s.kind === (side === 'old' ? 'removed' : 'added'))
    .map((s) => (s.kind === 'equal' ? esc(s.text) : `<mark>${esc(s.text)}</mark>`))
    .join('');
}

/** Render a table diff as an HTML string. Pair it with `HTML_CSS`. */
export function renderHTML(diff: TableDiff, options: HtmlOptions = {}): string {
  const layout = options.layout ?? 'split';
  const display = options.display ?? 'stacked';
  const cols = visibleColumns(diff, options.changedColumnsOnly ?? false);
  const items = hunks(diff, options.context ?? 2);
  const maxRows = options.maxRows ?? 2000;
  const kindOf = (r: number, c: number) => options.kindOf?.(r, c) ?? changeKind(diff, r, c);
  const offOld = options.rowNumberOffset ?? (diff.oldTable.origin?.row ?? 1) + 1;
  const offNew = options.rowNumberOffset ?? (diff.newTable.origin?.row ?? 1) + 1;
  const numeric = cols.map((ci) => isNumericColumn(diff, ci));
  const moveMark = (r: number) => (diff.moved?.[r] ? `<span class="sdl-mv" title="Moved from row ${diff.oldRows[r] + offOld}">↕</span>` : '');
  const val = (r: number, ci: number, side: 'old' | 'new'): CellValue | undefined => {
    const c = diff.columns[ci];
    const ri = side === 'old' ? diff.oldRows[r] : diff.newRows[r];
    const idx = side === 'old' ? c.oldIndex : c.newIndex;
    if (ri < 0 || idx < 0) return undefined;
    return (side === 'new' ? diff.newTable : diff.oldTable).rows[ri][idx] ?? null;
  };

  const header = (side: 'old' | 'new' | 'unified') =>
    cols
      .map((ci) => {
        const c = diff.columns[ci];
        const missing = (side === 'old' && c.status === 'added') || (side === 'new' && c.status === 'removed');
        const idx = side === 'old' ? c.oldIndex : c.newIndex >= 0 ? c.newIndex : c.oldIndex;
        const name = c.status === 'renamed' && side !== 'new' ? (side === 'old' ? c.oldName! : `${c.oldName} → ${c.name}`) : c.name;
        const cls = missing ? 'sdl-gap' : c.status !== 'equal' ? `sdl-col-${c.status}` : '';
        return `<th class="${cls}"${numeric[cols.indexOf(ci)] ? ' data-num' : ''}><span class="sdl-letter">${missing ? '' : columnLabel(idx)}</span>${missing ? '' : esc(name)}</th>`;
      })
      .join('');

  /** One side of one row: the cells. */
  const cells = (r: number, side: 'old' | 'new' | 'unified', mode: 'old' | 'new' | 'merged') => {
    const st = diff.status[r];
    const mask = diff.changedCells.get(r);
    return cols
      .map((ci, k) => {
        const c = diff.columns[ci];
        const num = numeric[k] ? ' data-num' : '';
        const colMissing = (mode === 'old' && c.status === 'added') || (mode === 'new' && c.status === 'removed');
        if (colMissing) return `<td class="sdl-gap"></td>`;
        const ov = val(r, ci, 'old');
        const nv = val(r, ci, 'new');
        const changed = st === ROW_STATUS.modified && mask?.[ci] === 1;
        if (!changed) {
          const v = mode === 'old' ? ov : mode === 'new' ? nv : nv !== undefined ? nv : ov;
          return `<td${num}>${esc(formatCell(v ?? null))}</td>`;
        }
        const kind = kindOf(r, ci);
        const ot = formatCell(ov ?? null);
        const nt = formatCell(nv ?? null);
        const k2 = ` data-kind="${kind}"`;
        if (mode === 'old') return `<td class="sdl-c-old"${k2}${num}>${intra(ot, nt, 'old')}</td>`;
        const d = deltaLabel(ov, nv);
        const deltaHtml = d ? `<small class="sdl-delta" data-up="${d.up}">${d.text}</small>` : '';
        if (mode === 'new' && side !== 'unified') {
          return `<td class="sdl-c-new"${k2}${num}>${intra(ot, nt, 'new')}${display === 'delta' ? deltaHtml : ''}</td>`;
        }
        // Unified, single line per row.
        if (display === 'inline') return `<td class="sdl-c-new"${k2}${num}><s>${esc(ot)}</s><span class="sdl-arrow">→</span>${esc(nt)}</td>`;
        if (display === 'delta') return `<td class="sdl-c-new"${k2}${num}>${esc(nt)}${deltaHtml || `<small class="sdl-was">was ${esc(ot) || '∅'}</small>`}</td>`;
        return `<td class="sdl-c-new"${k2}${num}>${intra(ot, nt, 'new')}</td>`;
      })
      .join('');
  };

  const width = cols.length;
  let out = '';
  let emitted = 0;
  let truncated = 0;
  const kindSummary = (r: number) => {
    const mask = diff.changedCells.get(r);
    if (!mask) return '';
    let computed = false;
    let direct = false;
    cols.forEach((ci) => {
      if (!mask[ci]) return;
      if (kindOf(r, ci) === 'computed') computed = true;
      else direct = true;
    });
    return computed && !direct ? ' data-computed-only' : '';
  };

  if (layout === 'split') {
    out += `<table class="sdl sdl-split"><thead><tr><th class="sdl-ln">old</th>${header('old')}<td class="sdl-mid" aria-hidden="true"></td><th class="sdl-ln">new</th>${header('new')}</tr></thead><tbody>`;
    for (const it of items) {
      if (emitted >= maxRows) {
        if (it.type === 'row') truncated++;
        continue;
      }
      if (it.type === 'gap') {
        out += `<tr class="sdl-hunk"><td colspan="${2 * width + 3}">⋯ ${(it.to - it.from).toLocaleString()} unchanged row${it.to - it.from === 1 ? '' : 's'}</td></tr>`;
        continue;
      }
      emitted++;
      const r = it.row;
      const st = diff.status[r];
      const oi = diff.oldRows[r];
      const ni = diff.newRows[r];
      const left = st === ROW_STATUS.added ? `<td class="sdl-ln"></td><td class="sdl-gap" colspan="${width}"></td>` : `<td class="sdl-ln">${oi + offOld}</td>${cells(r, 'old', 'old')}`;
      const right = st === ROW_STATUS.removed ? `<td class="sdl-ln"></td><td class="sdl-gap" colspan="${width}"></td>` : `<td class="sdl-ln">${moveMark(r)}${ni + offNew}</td>${cells(r, 'new', 'new')}`;
      out += `<tr class="sdl-r sdl-${statusName(st)}${diff.moved?.[r] ? ' sdl-moved' : ''}"${kindSummary(r)}>${left}<td class="sdl-mid"></td>${right}</tr>`;
    }
  } else {
    out += `<table class="sdl sdl-unified"><thead><tr><th class="sdl-ln">old</th><th class="sdl-ln">new</th><th class="sdl-sign"><span class="sdl-sr">change</span></th>${header('unified')}</tr></thead><tbody>`;
    for (const it of items) {
      if (emitted >= maxRows) {
        if (it.type === 'row') truncated++;
        continue;
      }
      if (it.type === 'gap') {
        out += `<tr class="sdl-hunk"><td colspan="${width + 3}">⋯ ${(it.to - it.from).toLocaleString()} unchanged row${it.to - it.from === 1 ? '' : 's'}</td></tr>`;
        continue;
      }
      emitted++;
      const r = it.row;
      const st = diff.status[r];
      const oi = diff.oldRows[r];
      const ni = diff.newRows[r];
      if (st === ROW_STATUS.modified && display === 'stacked') {
        // Code-diff convention: the old line, then the new line.
        out += `<tr class="sdl-r sdl-modified sdl-minus"${kindSummary(r)}><td class="sdl-ln">${oi + offOld}</td><td class="sdl-ln"></td><td class="sdl-sign">−</td>${cells(r, 'unified', 'old')}</tr>`;
        out += `<tr class="sdl-r sdl-modified sdl-plus${diff.moved?.[r] ? ' sdl-moved' : ''}"${kindSummary(r)}><td class="sdl-ln"></td><td class="sdl-ln">${moveMark(r)}${ni + offNew}</td><td class="sdl-sign">+</td>${cells(r, 'unified', 'new')}</tr>`;
        continue;
      }
      const sign = diff.moved?.[r] ? moveMark(r) : st === ROW_STATUS.added ? '+' : st === ROW_STATUS.removed ? '−' : st === ROW_STATUS.modified ? '•' : '';
      out += `<tr class="sdl-r sdl-${statusName(st)}${diff.moved?.[r] ? ' sdl-moved' : ''}"${kindSummary(r)}><td class="sdl-ln">${oi >= 0 ? oi + offOld : ''}</td><td class="sdl-ln">${ni >= 0 ? ni + offNew : ''}</td><td class="sdl-sign">${sign}</td>${cells(r, 'unified', st === ROW_STATUS.removed ? 'old' : 'merged')}</tr>`;
    }
  }
  out += '</tbody></table>';
  if (truncated) out += `<p class="sdl-note">${truncated.toLocaleString()} more changed rows not shown.</p>`;
  if (emitted === 0 && truncated === 0) out = `<p class="sdl-note">No changes.</p>`;
  return `<div class="sdl-wrap">${out}</div>`;
}

function statusName(st: number): string {
  return st === ROW_STATUS.added ? 'added' : st === ROW_STATUS.removed ? 'removed' : st === ROW_STATUS.modified ? 'modified' : 'equal';
}

function isNumericColumn(diff: TableDiff, ci: number): boolean {
  const c = diff.columns[ci];
  let num = 0;
  let total = 0;
  for (const [table, idx] of [[diff.newTable, c.newIndex], [diff.oldTable, c.oldIndex]] as const) {
    if (idx < 0) continue;
    for (let r = 0; r < table.rows.length && total < 200; r++) {
      const v = table.rows[r][idx];
      if (v === null || v === '' || v === undefined) continue;
      total++;
      if (typeof v === 'number') num++;
    }
  }
  return total > 0 && num / total >= 0.8;
}

/**
 * Plain-text unified diff with aligned columns, for terminals, logs and
 * markdown code blocks. Lines start with `+` / `-` (ASCII), so a ```diff
 * fence or a terminal diff highlighter colors them:
 *
 *       id  price   qty
 *     ⋯ 3 unchanged rows
 *   -  P-1003  65.4    28
 *   +  P-1003  70.63   28
 *
 * A row that moved (rows matched by key) is marked `~`, with the row it
 * moved from: `~  P-1007  12.5  3  (moved from row 2)`.
 */
export function renderText(diff: TableDiff, options: Pick<HtmlOptions, 'context' | 'changedColumnsOnly' | 'maxRows'> & { maxWidth?: number } = {}): string {
  const cols = visibleColumns(diff, options.changedColumnsOnly ?? false);
  const items = hunks(diff, options.context ?? 2);
  const maxW = options.maxWidth ?? 24;
  const lines: { sign: string; cells: string[]; note?: string }[] = [];
  const offOld = (diff.oldTable.origin?.row ?? 1) + 1;
  const movedNote = (r: number) => (diff.moved?.[r] ? `(moved from row ${diff.oldRows[r] + offOld})` : undefined);
  const value = (r: number, ci: number, side: 'old' | 'new') => {
    const c = diff.columns[ci];
    const ri = side === 'old' ? diff.oldRows[r] : diff.newRows[r];
    const idx = side === 'old' ? c.oldIndex : c.newIndex;
    const t = side === 'old' ? diff.oldTable : diff.newTable;
    return ri >= 0 && idx >= 0 ? formatCell(t.rows[ri][idx] ?? null).replace(/\s+/g, ' ') : '';
  };
  let emitted = 0;
  for (const it of items) {
    if (emitted >= (options.maxRows ?? 2000)) break;
    if (it.type === 'gap') {
      lines.push({ sign: '⋯', cells: [`${it.to - it.from} unchanged row${it.to - it.from === 1 ? '' : 's'}`] });
      continue;
    }
    emitted++;
    const r = it.row;
    const st = diff.status[r];
    if (st === ROW_STATUS.modified) {
      lines.push({ sign: '-', cells: cols.map((ci) => value(r, ci, 'old')) });
      lines.push({ sign: '+', cells: cols.map((ci) => value(r, ci, 'new')), note: movedNote(r) });
    } else if (st === ROW_STATUS.added) lines.push({ sign: '+', cells: cols.map((ci) => value(r, ci, 'new')) });
    else if (st === ROW_STATUS.removed) lines.push({ sign: '-', cells: cols.map((ci) => value(r, ci, 'old')) });
    else lines.push({ sign: diff.moved?.[r] ? '~' : ' ', cells: cols.map((ci) => value(r, ci, 'new')), note: movedNote(r) });
  }
  const head = cols.map((ci) => diff.columns[ci].name);
  const widths = head.map((h, k) =>
    Math.min(maxW, Math.max(h.length, ...lines.filter((l) => l.sign !== '⋯').map((l) => [...(l.cells[k] ?? '')].length))),
  );
  const fit = (s: string, w: number) => {
    const chars = [...s];
    return chars.length > w ? chars.slice(0, w - 1).join('') + '…' : s + ' '.repeat(w - chars.length);
  };
  const row = (sign: string, cells: string[]) => `${sign}  ${cells.map((c, k) => fit(c, widths[k])).join('  ')}`.trimEnd();
  const out = [row(' ', head)];
  for (const l of lines) out.push(l.sign === '⋯' ? `⋯  ${l.cells[0]}` : row(l.sign, l.cells) + (l.note ? `  ${l.note}` : ''));
  return out.join('\n') + '\n';
}

const HTML_BASE = /* css */ `
.sdl-wrap {
  --_bg: var(--sd-bg, light-dark(#fff, #0b0b0c));
  --_fg: var(--sd-fg, light-dark(#0b0b0c, #ededef));
  --_muted: color-mix(in srgb, var(--_fg) 64%, var(--_bg)); /* AA on the background and the panel */
  --_line: color-mix(in srgb, var(--_fg) 9%, var(--_bg));
  --_panel: color-mix(in srgb, var(--_fg) 3%, var(--_bg));
  --_add: var(--sd-added, light-dark(#0dbe4e, #5ecc71));
  --_del: var(--sd-removed, light-dark(#ff2e3f, #ff6762));
  --_add-fg: var(--sd-added-fg, light-dark(color-mix(in srgb, var(--_add) 62%, #000), color-mix(in srgb, var(--_add) 70%, var(--_fg))));
  --_del-fg: var(--sd-removed-fg, light-dark(color-mix(in srgb, var(--_del) 62%, #000), color-mix(in srgb, var(--_del) 70%, var(--_fg))));
  --_mod: var(--sd-modified, light-dark(#009fff, #69b1ff));
  --_mod-fg: var(--sd-modified-fg, light-dark(color-mix(in srgb, var(--_mod) 62%, #000), var(--_mod)));
  --_cmp: var(--sd-computed, light-dark(#e08a00, #f5b83d));
  --_fx: var(--sd-formula, light-dark(#7c3aed, #a78bfa));
  color-scheme: light dark;
  overflow: auto;
  background: var(--_bg);
  color: var(--_fg);
  font: 12px/1.5 var(--sd-mono-font-family, ui-monospace, "SF Mono", Menlo, Consolas, monospace);
}
.sdl { border-collapse: separate; border-spacing: 0; min-width: 100%; }
.sdl th, .sdl td { padding: 3px 10px; white-space: pre; text-align: left; vertical-align: top; border-bottom: 1px solid var(--_line); }
.sdl th { position: sticky; top: 0; z-index: 1; background: var(--_panel); font-weight: 600; font-family: var(--sd-font-family, ui-sans-serif, system-ui, sans-serif); }
.sdl [data-num] { text-align: right; font-variant-numeric: tabular-nums; }
.sdl .sdl-letter { color: var(--_muted); font-weight: 400; font-size: 10px; margin-right: 6px; }
.sdl .sdl-ln { color: var(--_muted); text-align: right; background: var(--_panel); user-select: none; min-width: 3ch; font-size: 11px; }
.sdl th.sdl-ln { font-size: 10px; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; }
.sdl .sdl-mid { padding: 0; width: 1px; background: var(--_line); }
.sdl .sdl-sign { color: var(--_muted); width: 1ch; padding: 3px 4px; }
.sdl .sdl-hunk td { background: color-mix(in srgb, var(--_mod) 7%, var(--_bg)); color: var(--_muted); font-family: var(--sd-font-family, ui-sans-serif, system-ui, sans-serif); padding: 4px 12px; }
.sdl .sdl-gap { background: repeating-linear-gradient(135deg, transparent 0 5px, var(--_line) 5px 6px); }
.sdl .sdl-added td:not(.sdl-ln):not(.sdl-gap):not(.sdl-mid) { background: color-mix(in srgb, var(--_add) 9%, var(--_bg)); }
.sdl-unified .sdl-removed td:not(.sdl-ln), .sdl-split .sdl-removed td:not(.sdl-ln):not(.sdl-gap):not(.sdl-mid) { background: color-mix(in srgb, var(--_del) 9%, var(--_bg)); }
.sdl-unified .sdl-removed td:not(.sdl-ln):not(.sdl-sign) { text-decoration: line-through; text-decoration-color: color-mix(in srgb, var(--_del) 60%, transparent); }
.sdl .sdl-added .sdl-sign { color: var(--_add-fg); }
.sdl .sdl-removed .sdl-sign, .sdl .sdl-minus .sdl-sign { color: var(--_del-fg); }
.sdl .sdl-plus .sdl-sign { color: var(--_add-fg); }
.sdl .sdl-modified:not(.sdl-minus):not(.sdl-plus) .sdl-sign { color: var(--_mod-fg); }
.sdl .sdl-mv { color: var(--_mod-fg); font-weight: 600; margin-right: 3px; cursor: help; }
.sdl .sdl-modified > td:first-child { box-shadow: inset 3px 0 var(--_mod); }
.sdl .sdl-modified[data-computed-only] > td:first-child { box-shadow: inset 3px 0 var(--_cmp); }
.sdl .sdl-c-old { background: color-mix(in srgb, var(--_del) 14%, var(--_bg)); }
.sdl .sdl-c-new { background: color-mix(in srgb, var(--_add) 16%, var(--_bg)); }
.sdl .sdl-c-old[data-kind="computed"] { background: var(--_panel); color: var(--_muted); }
.sdl .sdl-c-new[data-kind="computed"] { background: color-mix(in srgb, var(--_cmp) 13%, var(--_bg)); }
.sdl .sdl-c-new[data-kind="formula"]::before { content: "ƒ"; color: var(--_fx); font-style: italic; font-weight: 700; margin-right: 6px; }
.sdl mark { background: color-mix(in srgb, var(--_add) 38%, transparent); color: inherit; border-radius: 2px; }
.sdl .sdl-c-old mark { background: color-mix(in srgb, var(--_del) 34%, transparent); }
.sdl [data-kind="computed"] mark { background: color-mix(in srgb, var(--_cmp) 30%, transparent); }
.sdl .sdl-c-old[data-kind="computed"] mark { background: transparent; }
.sdl s { color: var(--_muted); text-decoration-color: color-mix(in srgb, var(--_del) 70%, transparent); }
.sdl .sdl-arrow { color: var(--_muted); margin: 0 6px; }
.sdl small { display: block; font-size: 10px; line-height: 1.3; }
.sdl .sdl-delta[data-up="true"] { color: var(--_add-fg); }
.sdl .sdl-delta[data-up="false"] { color: var(--_del-fg); }
.sdl .sdl-was { color: var(--_muted); }
.sdl th.sdl-col-added { color: var(--_add-fg); }
.sdl th.sdl-col-removed { color: var(--_del-fg); text-decoration: line-through; }
.sdl th.sdl-col-renamed { color: var(--_mod-fg); }
.sdl .sdl-sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; }
.sdl-note { color: var(--_muted); font-family: var(--sd-font-family, ui-sans-serif, system-ui, sans-serif); padding: 8px 12px; margin: 0; }
`;

/** Styles for `renderHTML`. Themed by the same `--sd-*` variables as the canvas grid. Includes a fallback for browsers and mail clients without `light-dark()`. */
export const HTML_CSS =
  HTML_BASE +
  lightDarkFallback(/\.sdl-wrap \{([\s\S]*?)\n\}/.exec(HTML_BASE)![1].replace(/\/\*[\s\S]*?\*\//g, ''), {
    light: '.sdl-wrap',
    darkWhenSystem: '.sdl-wrap',
  });
