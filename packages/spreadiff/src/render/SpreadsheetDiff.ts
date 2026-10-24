import { diffFiles, readSpreadsheetInput, type Comparison, type CompareOptions, type DiffFilesOptions, type SheetView, type SpreadsheetInput } from '../core/compare.js';
import { formatCell } from '../core/text.js';
import { displayMagnitude, displayValue, formatCodeAt } from '../core/display.js';
import { applyDecisions, decideRange, decisionStats, noDecisions, tableToCsv, type Decision, type SheetDecisions } from '../core/review.js';
import { applyDecisionsToXlsx, type SkippedDecision } from '../core/xlsx-write.js';
import type { CellChange } from '../core/report.js';
import type { CellRef } from '../core/trace.js';
import type { CellValue, Table } from '../core/types.js';
import { DiffGrid, type DiffGridOptions, type GridRange } from './DiffGrid.js';
import { STYLE } from './theme.js';

/** Options for {@link SpreadsheetDiff}: every `DiffGrid` option, plus these. */
export interface SpreadsheetDiffOptions<A = unknown> extends Omit<DiffGridOptions<A>, 'trace' | 'decisions'> {
  /** How to read, diff and trace the files (`diff`, `parse`, `trace`). */
  compare?: CompareOptions;
  /** Do the work in a Web Worker (see `DiffFilesOptions.worker`). */
  worker?: DiffFilesOptions['worker'];
  /**
   * The outputs panel: each changed final result, what drove it and by how
   * much (a waterfall). `'auto'` (default) shows it when tracing found
   * results downstream of the edits.
   */
  outputs?: boolean | 'auto';
  /**
   * Review mode: accept or reject the selected changes (buttons, or `a` /
   * `r` / `u` in the grid), and download the merged file. Default false.
   */
  review?: boolean;
  /** Content for the header of each sheet (a badge, a link): an element or text. */
  renderSheetHeader?: (sheet: SheetView, comparison: Comparison) => HTMLElement | string | null | undefined;
  /** Called when the files have been compared (and after each new `render` with files). */
  onLoad?: (comparison: Comparison) => void;
  /** Called when reading or comparing failed. The component shows the message too. */
  onError?: (error: Error) => void;
  /** Called when the shown sheet changes. */
  onSheetChange?: (sheet: SheetView) => void;
  /** Called when review decisions change, with one entry per sheet. */
  onDecisionsChange?: (decisions: SheetDecisions[]) => void;
}

/** Arguments to {@link SpreadsheetDiff.render}: two files, or a comparison you already have. */
export type SpreadsheetDiffRenderProps =
  | {
      /** The old version: a File, Blob, URL, bytes or `{ name, data }` (see `SpreadsheetInput`). */
      oldFile: SpreadsheetInput;
      /** The new version. */
      newFile: SpreadsheetInput;
      /** Element to render into. The component fills it: give it a height. */
      container: HTMLElement;
    }
  | {
      /** From `diffFiles` or `compareWorkbooks`. */
      comparison: Comparison;
      /** The new file's bytes, for downloading the merged .xlsx in review mode. Optional. */
      newFile?: { name: string; data: ArrayBuffer | Uint8Array };
      /** Element to render into. */
      container: HTMLElement;
    };

const CSS = /* css */ `
.sx { position: absolute; inset: 0; display: flex; flex-direction: column; }
.sx-bar { display: flex; align-items: center; gap: 6px; padding: 6px 8px; border-bottom: 1px solid var(--sdi-border); background: var(--sdi-header-bg); min-height: 36px; box-sizing: border-box; flex-wrap: wrap; }
.sx-tabs { display: flex; gap: 2px; overflow-x: auto; scrollbar-width: thin; min-width: 0; flex: 1 1 auto; }
.sx-tab { all: unset; cursor: pointer; padding: 4px 10px; border-radius: 6px; white-space: nowrap; color: var(--sdi-fg-muted); display: inline-flex; gap: 6px; align-items: center; font-size: 12px; }
.sx-tab:hover { background: var(--sdi-hover); color: var(--sdi-fg); }
.sx-tab[aria-selected="true"] { background: var(--sdi-bg); color: var(--sdi-fg); box-shadow: 0 0 0 1px var(--sdi-border-strong); }
.sx-tab:focus-visible, .sx-btn:focus-visible, .sx-out:focus-visible { outline: 2px solid var(--sdi-selection); outline-offset: 1px; }
.sx-count { font-family: var(--sdi-mono-font-family); font-size: 10.5px; color: var(--sdi-modified-fg); }
.sx-tab[data-status="added"] .sx-count { color: var(--sdi-added-fg); }
.sx-tab[data-status="removed"] .sx-count { color: var(--sdi-removed-fg); }
.sx-tab[data-status="equal"] .sx-count { color: var(--sdi-fg-faint); }
.sx-head { display: flex; align-items: center; gap: 6px; }
.sx-review { display: flex; align-items: center; gap: 4px; margin-left: auto; }
.sx-btn { all: unset; cursor: pointer; padding: 3px 9px; border-radius: 6px; border: 1px solid var(--sdi-border-strong); font-size: 12px; color: var(--sdi-fg); }
.sx-btn:hover { background: var(--sdi-hover); }
.sx-btn[data-kind="accept"] { color: var(--sdi-added-fg); }
.sx-btn[data-kind="reject"] { color: var(--sdi-removed-fg); }
.sx-btn:disabled { opacity: .45; cursor: default; }
.sx-tally { font-size: 11px; color: var(--sdi-fg-muted); white-space: nowrap; }
.sx-body { flex: 1 1 auto; display: flex; min-height: 0; }
.sx-grid { flex: 1 1 auto; position: relative; min-width: 0; }
.sx-side { flex: 0 0 320px; border-left: 1px solid var(--sdi-border); overflow: auto; background: var(--sdi-bg); font-size: 12px; }
.sx-side h2 { font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: .04em; color: var(--sdi-fg-muted); margin: 12px 12px 6px; }
.sx-out { all: unset; display: block; box-sizing: border-box; width: 100%; padding: 8px 12px; cursor: pointer; border-top: 1px solid var(--sdi-border); }
.sx-out:hover { background: var(--sdi-hover); }
.sx-out[aria-expanded="true"] { background: var(--sdi-header-bg); }
.sx-out-name { color: var(--sdi-fg); font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sx-out-id { font-family: var(--sdi-mono-font-family); font-size: 10.5px; color: var(--sdi-fg-faint); }
.sx-out-val { font-family: var(--sdi-mono-font-family); font-size: 11.5px; margin-top: 2px; }
.sx-out-val s { color: var(--sdi-fg-muted); }
.sx-up { color: var(--sdi-added-fg); }
.sx-down { color: var(--sdi-removed-fg); }
.sx-wf { margin: 8px 0 2px; display: grid; grid-template-columns: minmax(0, 1fr) 96px auto; gap: 3px 8px; align-items: center; font-size: 11px; }
.sx-wf-label { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--sdi-fg-muted); }
.sx-wf-track { position: relative; height: 10px; }
.sx-wf-bar { position: absolute; top: 0; height: 10px; border-radius: 2px; min-width: 1px; }
.sx-wf-bar[data-kind="total"] { background: var(--sdi-border-strong); }
.sx-wf-bar[data-kind="up"] { background: var(--sdi-added); }
.sx-wf-bar[data-kind="down"] { background: var(--sdi-removed); }
.sx-wf-bar[data-kind="other"] { background: var(--sdi-fg-faint); }
.sx-wf-num { font-family: var(--sdi-mono-font-family); text-align: right; white-space: nowrap; }
.sx-objects { border-top: 1px solid var(--sdi-border); padding: 6px 10px; font-size: 12px; background: var(--sdi-header-bg); max-height: 30%; overflow: auto; }
.sx-objects summary { cursor: pointer; color: var(--sdi-fg-muted); }
.sx-objects li { margin: 2px 0; }
.sx-objects code { font-family: var(--sdi-mono-font-family); font-size: 11px; }
.sx-msg { padding: 24px; color: var(--sdi-fg-muted); }
.sx-msg[data-error] { color: var(--sdi-removed-fg); }
.sx-more { all: unset; cursor: pointer; display: block; padding: 8px 12px; color: var(--sdi-modified-fg); }
`;

const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] => {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined) e.textContent = text;
  return e;
};

/**
 * The whole review in one component: give it two spreadsheets (Files,
 * URLs, bytes), and it parses, diffs and traces them (in a worker if you
 * pass one) and shows every sheet with tabs, the grid with downstream
 * changes traced, an outputs panel with a waterfall of what drove each
 * changed result, and changes to names, charts and pivot tables. In review
 * mode it accepts and rejects changes and downloads the merged file.
 *
 * ```ts
 * import { SpreadsheetDiff } from 'spreadiff';
 * declare const container: HTMLElement;
 *
 * const view = new SpreadsheetDiff({ layout: 'split' });
 * await view.render({ oldFile: new URL('/v1.xlsx', location.href), newFile: new URL('/v2.xlsx', location.href), container });
 * ```
 *
 * `DiffGrid` is the piece underneath, for when you want to lay out the
 * rest yourself.
 */
export class SpreadsheetDiff<A = unknown> {
  /** The element the component renders into its container. */
  readonly host: HTMLElement;
  /** The grid inside, for its methods (`nextChange`, `getRange`, …). */
  readonly grid: DiffGrid<A>;
  private options: SpreadsheetDiffOptions<A>;
  private shadow: ShadowRoot;
  private root: HTMLDivElement;
  private tabs: HTMLDivElement;
  private head: HTMLDivElement;
  private review: HTMLDivElement;
  private gridBox: HTMLDivElement;
  private side: HTMLDivElement;
  private objects: HTMLDivElement;
  private msg: HTMLDivElement;
  private cmp: Comparison | null = null;
  private index = 0;
  private decisions: SheetDecisions[] = [];
  private newBytes: { name: string; data: Uint8Array } | null = null;
  private expanded: CellChange | null = null;
  private outputsShown = 100;
  private headerSlot: HTMLElement | null = null;
  private loadId = 0;
  private range: GridRange | null = null;

  constructor(options: SpreadsheetDiffOptions<A> = {}) {
    this.options = options;
    this.host = el('div', 'spreadiff-workbook');
    this.host.style.cssText = 'display:block;position:relative;width:100%;height:100%;';
    this.shadow = this.host.attachShadow({ mode: 'open' });
    const adopt = 'adoptedStyleSheets' in ShadowRoot.prototype && 'replaceSync' in CSSStyleSheet.prototype;
    if (adopt) {
      const sheet = new CSSStyleSheet();
      sheet.replaceSync(STYLE + CSS);
      this.shadow.adoptedStyleSheets = [sheet];
    } else this.shadow.append(el('style', undefined, STYLE + CSS));
    this.root = el('div', 'sx sd-root');
    const bar = el('div', 'sx-bar');
    this.tabs = el('div', 'sx-tabs');
    this.tabs.setAttribute('role', 'tablist');
    this.tabs.setAttribute('aria-label', 'Sheets');
    this.head = el('div', 'sx-head');
    this.head.append(Object.assign(document.createElement('slot'), { name: 'header' }));
    this.review = el('div', 'sx-review');
    bar.append(this.tabs, this.head, this.review);
    const body = el('div', 'sx-body');
    this.gridBox = el('div', 'sx-grid');
    this.gridBox.setAttribute('role', 'tabpanel');
    this.side = el('div', 'sx-side');
    this.side.setAttribute('aria-label', 'Changed outputs');
    body.append(this.gridBox, this.side);
    this.objects = el('div', 'sx-objects');
    this.msg = el('div', 'sx-msg');
    this.root.append(bar, body, this.objects, this.msg);
    this.shadow.append(this.root);
    // The grid renders into a slot-less box inside the shadow root; it has its own shadow root.
    this.grid = new DiffGrid<A>(this.gridOptions());
    this.host.addEventListener('keydown', (e) => this.onKey(e));
    this.applyTheme();
  }

  /** The comparison shown, once loaded. */
  get comparison(): Comparison | null {
    return this.cmp;
  }

  /** Index of the sheet shown. */
  get sheet(): number {
    return this.index;
  }

  /**
   * Mount into `container` and show two files (compared here) or a
   * comparison. Resolves once it's shown; rejects if reading or comparing
   * fails (the component shows the error too).
   */
  async render(props: SpreadsheetDiffRenderProps): Promise<void> {
    if (this.host.parentElement !== props.container) props.container.appendChild(this.host);
    const id = ++this.loadId;
    if ('comparison' in props) {
      const nb = props.newFile;
      this.newBytes = nb ? { name: nb.name, data: nb.data instanceof Uint8Array ? nb.data : new Uint8Array(nb.data) } : null;
      this.show(props.comparison);
      return;
    }
    this.message('Comparing…');
    try {
      const parsed = (x: SpreadsheetInput) => !!x && typeof x === 'object' && ('sheets' in x || 'columns' in x) && !(x instanceof Blob) && !(x instanceof URL);
      const read = async (x: SpreadsheetInput, fallback: string) => (parsed(x) ? x : await readSpreadsheetInput(x as Exclude<SpreadsheetInput, { sheets: unknown } | { columns: unknown }>, fallback));
      const [a, b] = await Promise.all([read(props.oldFile, 'old'), read(props.newFile, 'new')]);
      if (id !== this.loadId) return;
      const nb = b as { name?: string; data?: ArrayBuffer };
      // Keep the new file's bytes for the merged download (the worker takes the buffers it's given).
      this.newBytes = nb.data instanceof ArrayBuffer && nb.name ? { name: nb.name, data: new Uint8Array(nb.data.slice(0)) } : null;
      const cmp = await diffFiles(a as SpreadsheetInput, b as SpreadsheetInput, { ...this.options.compare, worker: this.options.worker });
      if (id !== this.loadId) return;
      this.show(cmp);
    } catch (e) {
      if (id !== this.loadId) return;
      const err = e instanceof Error ? e : new Error(String(e));
      this.message(err.message, true);
      this.options.onError?.(err);
      throw err;
    }
  }

  /** Show sheet `index` of the comparison. */
  showSheet(index: number): void {
    const cmp = this.cmp;
    if (!cmp || !(index >= 0 && index < cmp.sheets.length)) return;
    const changed = index !== this.index;
    this.index = index;
    this.range = null;
    this.grid.setOptions(this.gridOptions());
    this.grid.render({ sheet: cmp.sheets[index], container: this.gridBox });
    this.renderTabs();
    this.renderHeader();
    this.renderReview();
    if (changed) this.options.onSheetChange?.(cmp.sheets[index]);
  }

  /** Show a cell: switch to its sheet, select it and scroll it into view. */
  reveal(ref: CellRef): void {
    if (ref.sheet !== this.index) this.showSheet(ref.sheet);
    this.grid.select(ref.row, ref.col);
  }

  /** Review decisions, one entry per sheet. */
  getDecisions(): SheetDecisions[] {
    return this.decisions;
  }

  /** Replace the review decisions (one entry per sheet; missing entries mean none). */
  setDecisions(decisions: SheetDecisions[]): void {
    const n = this.cmp?.sheets.length ?? decisions.length;
    this.decisions = Array.from({ length: n }, (_, i) => decisions[i] ?? noDecisions());
    this.grid.setOptions({ decisions: this.decisions[this.index] });
    this.renderReview();
    this.options.onDecisionsChange?.(this.decisions);
  }

  /** Accept or reject every change in the selected range (`undefined` clears their decisions). */
  decideSelection(decision: Decision | undefined): void {
    const cmp = this.cmp;
    const range = this.grid.getRange();
    if (!cmp || !range) return;
    const next = this.decisions.slice();
    next[this.index] = decideRange(cmp.sheets[this.index].diff, this.decisions[this.index] ?? noDecisions(), range.rows, range.columns, decision);
    this.setDecisions(next);
  }

  /**
   * The new .xlsx with the decisions applied (see `applyDecisionsToXlsx`).
   * Needs the new file's bytes: render with files, or pass `newFile`.
   */
  exportXlsx(): { file: Uint8Array; applied: number; recalculated: number; skipped: SkippedDecision[] } {
    if (!this.cmp || !this.newBytes) throw new Error('SpreadsheetDiff: no new .xlsx to apply decisions to');
    return applyDecisionsToXlsx(this.newBytes.data, this.cmp.workbook, this.decisions);
  }

  /** A sheet (default: the one shown) with the decisions applied, as CSV. */
  exportCsv(sheet = this.index): string {
    if (!this.cmp) throw new Error('SpreadsheetDiff: nothing loaded');
    return tableToCsv(this.mergedTable(sheet));
  }

  /** A sheet with the decisions applied. */
  mergedTable(sheet = this.index): Table {
    if (!this.cmp) throw new Error('SpreadsheetDiff: nothing loaded');
    return applyDecisions(this.cmp.sheets[sheet].diff, this.decisions[sheet] ?? noDecisions());
  }

  /** Merge options; the grid's go to the grid. */
  setOptions(next: Partial<SpreadsheetDiffOptions<A>>): void {
    const prev = this.options;
    this.options = { ...prev, ...next };
    if (next.theme !== undefined && next.theme !== prev.theme) this.applyTheme();
    this.grid.setOptions(this.gridOptions());
    if (!this.cmp) return;
    if ('outputs' in next || 'review' in next || 'renderSheetHeader' in next) {
      this.renderSide();
      this.renderReview();
      this.renderHeader();
    }
  }

  /** Unmount and release everything. */
  cleanUp(): void {
    this.loadId++;
    this.grid.cleanUp();
    this.headerSlot?.remove();
    this.host.remove();
  }

  // ---- Internals ---------------------------------------------------------

  private gridOptions(): DiffGridOptions<A> {
    const { compare: _c, worker: _w, outputs: _o, review: _r, renderSheetHeader: _h, onLoad: _l, onError: _e, onSheetChange: _s, onDecisionsChange: _d, ...grid } = this.options;
    return { ...grid, trace: this.cmp?.sheets[this.index]?.trace, decisions: this.decisions[this.index], onRangeChange: this.onRange };
  }

  /** One function for the grid's lifetime, so setOptions doesn't see a new one each time. */
  private onRange = (r: GridRange | null): void => {
    this.range = r;
    this.renderReview();
    this.options.onRangeChange?.(r);
  };

  private applyTheme(): void {
    const t = this.options.theme;
    if (t && t !== 'system') this.root.setAttribute('data-theme', t);
    else this.root.removeAttribute('data-theme');
  }

  private message(text: string, error = false): void {
    this.msg.textContent = text;
    this.msg.toggleAttribute('data-error', error);
    this.msg.hidden = false;
  }

  private show(cmp: Comparison): void {
    this.cmp = cmp;
    this.msg.hidden = true;
    this.decisions = cmp.sheets.map(() => noDecisions());
    this.expanded = null;
    this.outputsShown = 100;
    this.index = -1;
    this.showSheet(cmp.firstChanged);
    this.renderSide();
    this.renderObjects();
    this.options.onLoad?.(cmp);
  }

  private renderTabs(): void {
    const cmp = this.cmp!;
    const buttons = cmp.sheets.map((s, i) => {
      const b = el('button', 'sx-tab');
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(i === this.index));
      b.tabIndex = i === this.index ? 0 : -1;
      b.dataset.status = s.status;
      const st = s.diff.stats;
      const n = st.cellsChanged + st.rowsAdded + st.rowsRemoved;
      const count = s.status === 'added' ? 'new' : s.status === 'removed' ? 'removed' : n ? n.toLocaleString('en-US') : '';
      b.append(el('span', undefined, s.name));
      if (count) b.append(el('span', 'sx-count', count));
      b.title = s.status === 'equal' ? `${s.name}: no changes` : `${s.name}: ${st.cellsChanged} cells changed, ${st.rowsAdded} rows added, ${st.rowsRemoved} removed`;
      b.addEventListener('click', () => this.showSheet(i));
      b.addEventListener('keydown', (e) => {
        if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
        const j = (i + (e.key === 'ArrowRight' ? 1 : -1) + cmp.sheets.length) % cmp.sheets.length;
        this.showSheet(j);
        (this.tabs.children[j] as HTMLElement | undefined)?.focus();
        e.preventDefault();
      });
      return b;
    });
    this.tabs.replaceChildren(...buttons);
  }

  private renderHeader(): void {
    this.headerSlot?.remove();
    this.headerSlot = null;
    const cmp = this.cmp;
    const render = this.options.renderSheetHeader;
    if (!cmp || !render) return;
    const content = render(cmp.sheets[this.index], cmp);
    if (content === null || content === undefined || content === '') return;
    // Light DOM, projected into the header: page CSS styles it.
    const slot = typeof content === 'string' ? el('span', undefined, content) : content;
    slot.slot = 'header';
    this.host.append(slot);
    this.headerSlot = slot;
  }

  private renderReview(): void {
    const cmp = this.cmp;
    if (!cmp || !this.options.review) {
      this.review.replaceChildren();
      return;
    }
    const d = cmp.sheets[this.index].diff;
    const stats = decisionStats(d, this.decisions[this.index] ?? noDecisions());
    const range = this.range;
    const btn = (label: string, kind: string, title: string, fn: () => void, enabled = true) => {
      const b = el('button', 'sx-btn', label);
      b.dataset.kind = kind;
      b.title = title;
      b.disabled = !enabled;
      b.addEventListener('click', fn);
      return b;
    };
    const hasSel = !!range;
    const tally = el('span', 'sx-tally', `${stats.rejected} rejected · ${stats.accepted} accepted · ${stats.changes} change${stats.changes === 1 ? '' : 's'}`);
    const download = btn(this.newBytes && /\.xls[xm]$/i.test(this.newBytes.name) ? 'Download .xlsx' : 'Download .csv', 'download', 'The new file with rejected changes undone', () => this.download());
    this.review.replaceChildren(
      tally,
      btn('Accept', 'accept', 'Accept the selected changes (a)', () => this.decideSelection('accept'), hasSel),
      btn('Reject', 'reject', 'Reject the selected changes (r)', () => this.decideSelection('reject'), hasSel),
      btn('Undo', 'clear', 'Clear decisions on the selection (u)', () => this.decideSelection(undefined), hasSel),
      download,
    );
  }

  private download(): void {
    const xlsx = this.newBytes && /\.xls[xm]$/i.test(this.newBytes.name);
    let blob: Blob;
    let name: string;
    if (xlsx) {
      const { file } = this.exportXlsx();
      blob = new Blob([file.slice().buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      name = this.newBytes!.name.replace(/(\.\w+)$/, ' (reviewed)$1');
    } else {
      blob = new Blob([this.exportCsv()], { type: 'text/csv' });
      name = `${this.cmp!.sheets[this.index].name} (reviewed).csv`;
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  private onKey(e: KeyboardEvent): void {
    if (!this.options.review || e.metaKey || e.ctrlKey || e.altKey) return;
    // Only from the grid itself, not from buttons or the side panel.
    if (!e.composedPath().includes(this.gridBox)) return;
    const d: Record<string, Decision | null> = { a: 'accept', r: 'reject', u: null };
    if (!(e.key in d)) return;
    this.decideSelection(d[e.key] ?? undefined);
    e.preventDefault();
  }

  private renderSide(): void {
    const cmp = this.cmp!;
    const mode = this.options.outputs ?? 'auto';
    const report = cmp.trace ? cmp.report() : null;
    const outputs = report ? report.downstream.filter((c) => c.output) : [];
    const show = mode === true || (mode === 'auto' && outputs.length > 0);
    this.side.hidden = !show;
    if (!show) return;
    const title = el('h2', undefined, outputs.length ? `Changed outputs · ${outputs.length.toLocaleString('en-US')}` : 'No outputs changed');
    const items = outputs.slice(0, this.outputsShown).map((c) => this.outputItem(c));
    const more = outputs.length > this.outputsShown ? el('button', 'sx-more', `Show ${Math.min(100, outputs.length - this.outputsShown)} more`) : null;
    more?.addEventListener('click', () => {
      this.outputsShown += 100;
      this.renderSide();
    });
    this.side.replaceChildren(title, ...items, ...(more ? [more] : []));
  }

  /** The format code of a report entry's cell, for showing its values as the sheet does. */
  private codeOf(c: CellChange): { code?: string; date1904: boolean } {
    const s = this.cmp!.sheets[c.ref.sheet].diff;
    const ni = s.newRows[c.ref.row];
    const ci = s.columns[c.ref.col].newIndex;
    if (this.options.numberFormats === false || ni < 0 || ci < 0) return { date1904: false };
    return { code: formatCodeAt(s.newTable, ni, ci), date1904: !!s.newTable.numberFormats?.date1904 };
  }

  private outputItem(c: CellChange): HTMLElement {
    const b = el('button', 'sx-out');
    const open = this.expanded === c;
    b.setAttribute('aria-expanded', String(open));
    const { code, date1904 } = this.codeOf(c);
    const show = (v: CellValue) => (v === null ? '∅' : displayValue(v, code, date1904));
    const name = el('div', 'sx-out-name', c.label || c.id);
    const id = el('div', 'sx-out-id', c.id);
    const val = el('div', 'sx-out-val');
    val.append(el('s', undefined, show(c.before)), ' → ', el('b', undefined, show(c.after)));
    if (c.delta !== undefined && c.delta !== 0) {
      const up = c.delta > 0;
      const pct = c.ratio !== undefined ? ` (${up ? '+' : '−'}${Math.abs(c.ratio * 100).toFixed(Math.abs(c.ratio) < 0.1 ? 2 : 1)}%)` : '';
      val.append(' ', el('span', up ? 'sx-up' : 'sx-down', `${up ? '+' : '−'}${displayMagnitude(Math.abs(c.delta), code)}${pct}`));
    }
    b.append(name, id, val);
    if (open) b.append(this.waterfall(c, code));
    b.addEventListener('click', () => {
      this.expanded = open ? null : c;
      this.reveal(c.ref);
      this.renderSide();
    });
    return b;
  }

  /**
   * Start value, each cause's contribution, anything the contributions don't
   * account for (edits interacting: a price and a volume both moved), and
   * the end value, as floating bars.
   */
  private waterfall(c: CellChange, code: string | undefined): HTMLElement {
    const box = el('div', 'sx-wf');
    if (typeof c.before !== 'number' || typeof c.after !== 'number') {
      box.append(el('span', 'sx-wf-label', 'Not a number: no waterfall.'));
      return box;
    }
    const causes = (c.causes ?? []).filter((x) => typeof x.contribution === 'number' && x.contribution !== 0).slice(0, 8);
    type Step = { label: string; from: number; to: number; kind: 'total' | 'up' | 'down' | 'other' };
    const steps: Step[] = [{ label: 'Before', from: 0, to: c.before, kind: 'total' }];
    let at = c.before;
    for (const x of causes) {
      const v = x.contribution as number;
      steps.push({ label: x.label || x.id, from: at, to: at + v, kind: v > 0 ? 'up' : 'down' });
      at += v;
    }
    const rest = c.after - at;
    if (Math.abs(rest) > 1e-9 * Math.max(1, Math.abs(c.after))) {
      const listed = (c.causes ?? []).length;
      const label = (c.causeCount ?? 0) > causes.length || listed > causes.length ? 'Other causes, together' : 'Interaction of the edits';
      steps.push({ label, from: at, to: c.after, kind: 'other' });
    }
    steps.push({ label: 'After', from: 0, to: c.after, kind: 'total' });
    // The axis covers what the steps move through, not down to zero: a 3% change would be a sliver.
    const ends = steps.flatMap((s) => (s.kind === 'total' ? [s.to] : [s.from, s.to]));
    let lo = Math.min(...ends);
    let hi = Math.max(...ends);
    const pad = (hi - lo || Math.abs(hi) || 1) * 0.15;
    lo = lo >= 0 && lo - pad < 0 ? 0 : lo - pad;
    hi += pad;
    const span = hi - lo || 1;
    const fmt = (v: number) => (code ? displayValue(v, code) : formatCell(Number(v.toPrecision(12))));
    for (const s of steps) {
      const track = el('div', 'sx-wf-track');
      const bar = el('div', 'sx-wf-bar');
      bar.dataset.kind = s.kind;
      // Totals run from the axis start to their value.
      const a = s.kind === 'total' ? lo : Math.min(s.from, s.to);
      const b = s.kind === 'total' ? s.to : Math.max(s.from, s.to);
      bar.style.left = `${((a - lo) / span) * 100}%`;
      bar.style.width = `${((b - a) / span) * 100}%`;
      track.append(bar);
      const d = s.to - s.from;
      const num = s.kind === 'total' ? fmt(s.to) : `${d > 0 ? '+' : '−'}${displayMagnitude(Math.abs(d), code)}`;
      box.append(el('span', 'sx-wf-label', s.label), track, el('span', 'sx-wf-num', num));
    }
    box.setAttribute('role', 'img');
    box.setAttribute('aria-label', steps.map((s) => `${s.label} ${s.kind === 'total' ? fmt(s.to) : formatCell(Number((s.to - s.from).toPrecision(12)))}`).join(', '));
    return box;
  }

  private renderObjects(): void {
    const list = this.cmp?.workbook.objects ?? [];
    this.objects.hidden = !list.length;
    if (!list.length) return;
    const det = el('details');
    det.append(el('summary', undefined, `${list.length} change${list.length === 1 ? '' : 's'} to names, charts or pivot tables`));
    const ul = el('ul');
    const noun = { name: 'Name', chart: 'Chart', pivot: 'Pivot table' };
    for (const o of list) {
      const li = el('li');
      li.append(`${noun[o.kind]} `, el('code', undefined, o.name), ` ${o.change}${o.details.filter(Boolean).length ? ': ' + o.details.filter(Boolean).join('; ') : ''}`);
      ul.append(li);
    }
    det.append(ul);
    this.objects.replaceChildren(det);
  }
}
