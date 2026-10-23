import { columnLabel } from '../core/csv.js';
import { diffText, formatCell, numericDelta } from '../core/text.js';
import { changeKind } from '../core/kind.js';
import type { CellRef, CellTrace, ChangeKind, WorkbookTrace } from '../core/trace.js';
import {
  ROW_STATUS,
  ROW_STATUS_NAMES,
  type AlignedColumn,
  type CellValue,
  type RowStatus,
  type TableDiff,
  type WorkbookDiff,
} from '../core/types.js';
import {
  computeColumnLayout,
  computeRowLayout,
  findChange,
  rowAt,
  ROW_KIND,
  type ColumnLayout,
  type RowLayout,
} from './layout.js';
import { resolveTheme, STYLE, type Theme } from './theme.js';

export type DiffLayout = 'unified' | 'split';

export interface CellAnnotation<A = unknown> {
  /** Aligned row index (`0 … diff.rowCount - 1`). */
  row: number;
  /** Column name to mark with a corner indicator. Optional. */
  column?: string;
  data: A;
}

export interface CellInfo {
  /** Aligned row index. */
  row: number;
  column: AlignedColumn;
  columnIndex: number;
  status: RowStatus;
  changed: boolean;
  oldRow: number;
  newRow: number;
  oldValue: CellValue | undefined;
  newValue: CellValue | undefined;
  oldFormula?: string | null;
  newFormula?: string | null;
  /** input / formula / computed, when the cell changed. */
  kind?: ChangeKind;
  /** Root causes and impacts, when a trace is attached. */
  trace?: CellTrace;
}

/**
 * How a changed cell shows its change:
 *  - `stacked`  old value struck through above the new one (taller rows)
 *  - `new`      just the new value on a tinted background
 *  - `delta`    the new value, with the absolute and % change beneath it
 *  - `inline`   `old → new` on one line
 */
export type CellDisplay = 'stacked' | 'new' | 'delta' | 'inline';

/** A trace of the whole workbook, plus which sheet of it this grid shows. */
export interface TraceContext {
  trace: WorkbookTrace;
  workbook: WorkbookDiff;
  sheet: number;
}

export interface DiffGridOptions<A = unknown> {
  /** `unified` (one grid, old/new stacked in changed cells) or `split` (side by side). */
  layout?: DiffLayout;
  /** Collapse runs of unchanged rows into expandable separators. Default true. */
  collapseUnchanged?: boolean;
  /** Unchanged rows kept around each change when collapsing. Default 2. */
  context?: number;
  /** Hide columns with no changes (key columns stay). Default false. */
  hideUnchangedColumns?: boolean;
  /** Columns pinned to the left. `'keys'` pins the diff's key columns. Default `'keys'`. */
  frozenColumns?: number | 'keys';
  /**
   * Added to the 0-based row index for display. Defaults to the table's
   * `origin` (real sheet row numbers), or 2 (1-based, after a header row).
   */
  rowNumberOffset?: number;
  /** Force a color scheme. Default follows the page (`system`). */
  theme?: 'system' | 'light' | 'dark';
  /** Change overview strip on the right edge. Default true. */
  minimap?: boolean;
  /** Old → new popover when hovering a changed cell. Default true. */
  tooltips?: boolean;
  /** Fixed widths for columns, by name. */
  columnWidths?: Record<string, number>;
  annotations?: CellAnnotation<A>[];
  /** Render a full-width element below an annotated row. */
  renderAnnotation?: (annotation: CellAnnotation<A>) => HTMLElement | null | undefined;
  onCellClick?: (cell: CellInfo, event: MouseEvent) => void;
  onCellHover?: (cell: CellInfo | null) => void;
  onSelectionChange?: (cell: CellInfo | null) => void;
  /** How changed cells show old vs. new. Default `stacked`. */
  cellDisplay?: CellDisplay;
  /**
   * Color changes that are formula results of other changes (downstream)
   * amber, apart from edited inputs. Default true.
   */
  highlightComputed?: boolean;
  /** Attach a trace to show root causes in tooltips and on selection. */
  trace?: TraceContext;
  /** When a traced cell is selected, outline its causes/impacts and dim the rest. Default true. */
  traceFocus?: boolean;
}

export interface DiffGridRenderProps {
  diff: TableDiff;
  /** Element to render into. The grid fills it; give it a height. */
  container: HTMLElement;
}

const HEADER_HEIGHT = 36;
const SEPARATOR_HEIGHT = 30;
const MINIMAP_WIDTH = 12;
const SPLIT_DIVIDER = 1;
const CELL_PAD = 10;
const MIN_COL = 56;
const MAX_COL = 360;
const BAR_WIDTH = 3;

interface Hit {
  area: 'header' | 'gutter' | 'cell' | 'separator' | 'annotation' | 'minimap' | 'none';
  pane: 0 | 1;
  display: number;
  /** Index into the visible column layout. */
  vcol: number;
  /** Column resize handle under the pointer (visible column index). */
  resize?: number;
}

interface PaneGeometry {
  x: number;
  width: number;
  gutter: number;
  side: 'old' | 'new' | 'unified';
}

/**
 * A canvas-rendered, virtualized spreadsheet diff.
 *
 * ```ts
 * const grid = new DiffGrid({ layout: 'split' });
 * grid.render({ diff: diffTables(before, after), container });
 * ```
 */
export class DiffGrid<A = unknown> {
  readonly host: HTMLElement;
  private options: DiffGridOptions<A>;
  private diff: TableDiff | null = null;
  private container: HTMLElement | null = null;
  private shadow: ShadowRoot;
  private scroller: HTMLDivElement;
  private spacer: HTMLDivElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private minimap: HTMLCanvasElement;
  private mctx: CanvasRenderingContext2D;
  private overlay: HTMLDivElement;
  private tooltip: HTMLDivElement;
  private theme!: Theme;
  private hatch: CanvasPattern | null = null;
  private rows!: RowLayout;
  private cols!: ColumnLayout;
  private widths: number[] = [];
  private widthOverrides = new Map<number, number>();
  private numericCols: Uint8Array = new Uint8Array(0);
  private expanded = new Set<string>();
  private annotationEls = new Map<CellAnnotation<A>, { el: HTMLElement; height: number; ro: ResizeObserver }>();
  private cellMarks = new Set<string>();
  private width = 0;
  private height = 0;
  private dpr = 1;
  private frame = 0;
  private hover: Hit | null = null;
  private selected: { row: number; col: number } | null = null;
  private tooltipTimer = 0;
  private resizing: { vcol: number; startX: number; startW: number } | null = null;
  private lastPointer: { x: number; y: number } | null = null;
  /** Cells related to the selection through the trace: `row:col` → role. */
  private focus: Map<string, 'cause' | 'impact' | 'parent'> | null = null;
  private focusKey = '';
  private themeFrame = 0;
  private textCache = new Map<string, string>();
  private minimapCache: { key: string; colors: (string | null)[] } | null = null;
  private cleanups: (() => void)[] = [];

  constructor(options: DiffGridOptions<A> = {}) {
    this.options = options;
    this.host = document.createElement('div');
    this.host.className = 'spreadiff';
    this.host.style.cssText = 'display:block;position:relative;width:100%;height:100%;';
    this.shadow = this.host.attachShadow({ mode: 'open' });
    const style = document.createElement('style');
    style.textContent = STYLE;
    this.scroller = el('div', 'sd-scroller');
    this.scroller.tabIndex = 0;
    this.scroller.setAttribute('role', 'grid');
    this.scroller.setAttribute('aria-label', 'Spreadsheet diff');
    this.spacer = el('div', 'sd-spacer');
    this.scroller.appendChild(this.spacer);
    this.canvas = el('canvas', 'sd-canvas');
    this.minimap = el('canvas', 'sd-minimap');
    this.overlay = el('div', 'sd-overlay');
    // Annotation content lives in the light DOM (so page CSS styles it) and
    // is projected into the overlay through this slot.
    this.overlay.appendChild(document.createElement('slot'));
    this.tooltip = el('div', 'sd-tooltip');
    this.tooltip.setAttribute('role', 'tooltip');
    this.shadow.append(style, this.scroller, this.canvas, this.minimap, this.overlay, this.tooltip);
    this.ctx = this.canvas.getContext('2d', { alpha: false })!;
    this.mctx = this.minimap.getContext('2d')!;
    this.applyThemeAttr();
  }

  // ---- Public API ------------------------------------------------------

  render({ diff, container }: DiffGridRenderProps): void {
    if (this.container !== container) {
      this.detach();
      this.container = container;
      container.appendChild(this.host);
      this.attach();
    }
    this.setDiff(diff);
  }

  setDiff(diff: TableDiff): void {
    if (this.diff === diff) return;
    this.diff = diff;
    this.expanded.clear();
    this.selected = null;
    this.hover = null;
    this.widthOverrides.clear();
    this.scroller.scrollTop = 0;
    this.scroller.scrollLeft = 0;
    if (!this.theme) this.refreshTheme();
    this.measureColumns();
    this.syncAnnotations();
    this.relayout();
  }

  /**
   * Merge options. Cheap to call on every render (as the React wrapper does):
   * only options that actually changed trigger re-measuring or re-layout.
   */
  setOptions(next: Partial<DiffGridOptions<A>>): void {
    const prev = this.options;
    const changed = (Object.keys(next) as (keyof DiffGridOptions<A>)[]).filter((k) =>
      // Inline `columnWidths={{…}}` objects are new every render; compare contents.
      k === 'columnWidths' ? !shallowEqual(next[k], prev[k]) : next[k] !== prev[k],
    );
    if (!changed.length) return;
    this.options = { ...prev, ...next };
    const has = (...keys: (keyof DiffGridOptions<A>)[]) => keys.some((k) => changed.includes(k));
    if (has('theme')) {
      this.applyThemeAttr();
      this.refreshTheme();
    }
    if (!this.diff) return;
    if (has('columnWidths', 'cellDisplay')) this.measureColumns();
    if (has('annotations', 'renderAnnotation')) this.syncAnnotations();
    if (has('trace', 'traceFocus')) this.updateFocus();
    if (has('layout', 'collapseUnchanged', 'context', 'hideUnchangedColumns', 'frozenColumns', 'columnWidths', 'annotations', 'renderAnnotation', 'rowNumberOffset', 'minimap', 'cellDisplay')) {
      this.relayout();
    } else {
      this.schedule();
    }
  }

  getOptions(): Readonly<DiffGridOptions<A>> {
    return this.options;
  }

  /** Scroll to the next block of changes. Returns false at the end. */
  nextChange(): boolean {
    return this.jumpChange(1);
  }

  prevChange(): boolean {
    return this.jumpChange(-1);
  }

  /** Scroll an aligned row into view, expanding its collapsed region if needed. */
  scrollToRow(row: number, align: 'start' | 'center' | 'nearest' = 'center'): void {
    if (!this.diff || !(row >= 0 && row < this.diff.rowCount)) return;
    if (this.rows.displayOf[row] < 0) {
      const region = this.rows.regions.find((g) => row >= g.start && row < g.end);
      if (region) {
        this.expanded.add(`${region.start}:${region.end}`);
        this.relayout();
      }
    }
    const d = this.rows.displayOf[row];
    if (d >= 0) this.scrollToDisplay(d, align);
  }

  /** Expand every collapsed region. */
  expandAll(): void {
    if (!this.diff) return;
    for (const g of this.rows.regions) this.expanded.add(`${g.start}:${g.end}`);
    this.relayout();
  }

  collapseAll(): void {
    this.expanded.clear();
    if (this.diff) this.relayout();
  }

  /** Select a cell by aligned row and column index. Returns false (and changes nothing) when out of range. */
  select(row: number | null, columnIndex = 0): boolean {
    if (row !== null && !this.isCell(row, columnIndex)) return false;
    this.selected = row === null ? null : { row, col: columnIndex };
    this.options.onSelectionChange?.(row === null ? null : this.cellInfo(row, columnIndex));
    if (row !== null) this.scrollToRow(row, 'nearest');
    this.schedule();
    return true;
  }

  private isCell(row: number, columnIndex: number): boolean {
    const d = this.diff;
    return !!d && Number.isInteger(row) && Number.isInteger(columnIndex) && row >= 0 && row < d.rowCount && columnIndex >= 0 && columnIndex < d.columns.length;
  }

  /** Re-read CSS variables (call after changing them programmatically). */
  refreshTheme(): void {
    // Computed styles of a detached element are empty; resolve on attach.
    if (!this.host.isConnected && this.theme) return;
    this.theme = resolveTheme(this.shadow, this.host);
    this.hatch = this.makeHatch();
    this.textCache.clear();
    this.minimapCache = null;
    this.schedule();
  }

  cleanUp(): void {
    this.detach();
    for (const info of this.annotationEls.values()) info.ro.disconnect();
    this.annotationEls.clear();
    this.host.remove();
    this.container = null;
    this.diff = null;
  }

  // ---- Lifecycle ---------------------------------------------------------

  private attach(): void {
    const on = <K extends keyof HTMLElementEventMap>(
      target: HTMLElement,
      type: K,
      fn: (e: HTMLElementEventMap[K]) => void,
      opts?: AddEventListenerOptions,
    ) => {
      target.addEventListener(type, fn as EventListener, opts);
      this.cleanups.push(() => target.removeEventListener(type, fn as EventListener, opts));
    };
    on(this.scroller, 'scroll', () => {
      this.hideTooltip();
      // Content moved under a still pointer: re-hit-test so the hover
      // highlight follows the cursor, not the row it was over.
      if (this.lastPointer && !this.resizing) this.hover = this.hitTest(this.lastPointer.x, this.lastPointer.y);
      this.schedule();
    }, { passive: true });
    on(this.scroller, 'pointermove', (e) => this.onPointerMove(e));
    on(this.scroller, 'pointerleave', () => {
      if (this.resizing) return;
      this.lastPointer = null;
      this.hover = null;
      this.hideTooltip();
      this.options.onCellHover?.(null);
      this.schedule();
    });
    on(this.scroller, 'pointerdown', (e) => this.onPointerDown(e));
    on(this.scroller, 'keydown', (e) => this.onKeyDown(e));

    // Observe the scroller, not the host: its client box shrinks when a
    // scrollbar appears, which the host's size never reflects.
    const ro = new ResizeObserver(() => this.resize());
    ro.observe(this.scroller);
    this.cleanups.push(() => ro.disconnect());

    // Web fonts finishing late would leave canvas text in the fallback face.
    const fonts = document.fonts;
    if (fonts) {
      const onFonts = () => {
        this.textCache.clear();
        if (this.diff) {
          this.measureColumns();
          this.relayoutColumns();
        }
      };
      fonts.addEventListener('loadingdone', onFonts);
      this.cleanups.push(() => fonts.removeEventListener('loadingdone', onFonts));
      void fonts.ready.then(() => this.container && onFonts());
    }

    // Re-resolve colors when the page's scheme or theme attributes change
    // (coalesced: apps may touch <html> attributes many times per frame).
    const mq = matchMedia('(prefers-color-scheme: dark)');
    const onScheme = () => {
      if (this.themeFrame) return;
      this.themeFrame = requestAnimationFrame(() => {
        this.themeFrame = 0;
        this.refreshTheme();
      });
    };
    mq.addEventListener('change', onScheme);
    this.cleanups.push(() => mq.removeEventListener('change', onScheme));
    const mo = new MutationObserver(onScheme);
    const watch = { attributes: true, attributeFilter: ['class', 'style', 'data-theme'] };
    mo.observe(document.documentElement, watch);
    if (this.container) mo.observe(this.container, watch);
    this.cleanups.push(() => mo.disconnect());
    // Colors resolved while detached (or under another container) are stale.
    if (this.theme) this.refreshTheme();
    this.resize();
  }

  private detach(): void {
    for (const fn of this.cleanups.splice(0)) fn();
    cancelAnimationFrame(this.frame);
    cancelAnimationFrame(this.themeFrame);
    this.frame = 0;
    this.themeFrame = 0;
    this.hideTooltip();
  }

  private applyThemeAttr(): void {
    const t = this.options.theme ?? 'system';
    if (t === 'system') this.host.removeAttribute('data-theme');
    else this.host.setAttribute('data-theme', t);
  }

  private resize(): void {
    const w = this.scroller.clientWidth;
    const h = this.scroller.clientHeight;
    const dpr = window.devicePixelRatio || 1;
    if (w === this.width && h === this.height && dpr === this.dpr) return;
    const dprChanged = dpr !== this.dpr;
    this.width = w;
    this.height = h;
    this.dpr = dpr;
    if (dprChanged && this.theme) this.hatch = this.makeHatch();
    for (const [c, cw] of [
      [this.canvas, w],
      [this.minimap, MINIMAP_WIDTH],
    ] as const) {
      c.width = Math.max(1, Math.round(cw * dpr));
      c.height = Math.max(1, Math.round(h * dpr));
      c.style.width = `${cw}px`;
      c.style.height = `${h}px`;
    }
    this.minimap.style.left = `${w - MINIMAP_WIDTH}px`;
    this.minimapCache = null;
    if (this.diff) this.updateSpacer();
    this.drawNow();
  }

  // ---- Layout ------------------------------------------------------------

  private get layout(): DiffLayout {
    return this.options.layout ?? 'unified';
  }

  private get minimapWidth(): number {
    return this.options.minimap === false ? 0 : MINIMAP_WIDTH;
  }

  private get display(): CellDisplay {
    return this.options.cellDisplay ?? 'stacked';
  }

  /** Stacked and delta displays need two lines per changed row. */
  private get tallRows(): boolean {
    return this.display === 'stacked' || this.display === 'delta';
  }

  /** input / formula / computed for a changed cell. */
  private kindOf(r: number, ci: number): ChangeKind {
    const d = this.diff!;
    const tc = this.options.trace;
    if (tc) {
      const k = tc.trace.get({ sheet: tc.sheet, row: r, col: ci })?.kind;
      if (k) return k;
    }
    return changeKind(d, r, ci);
  }

  private get rowHeight(): number {
    return this.theme?.rowHeight ?? 30;
  }

  private relayout(): void {
    const diff = this.diff!;
    const unified = this.layout === 'unified';
    const lh = this.rowHeight;
    const annotationsAfter = new Map<number, number[]>();
    for (const [a, info] of this.annotationEls) {
      const list = annotationsAfter.get(a.row) ?? [];
      list.push(info.height);
      annotationsAfter.set(a.row, list);
    }
    this.rows = computeRowLayout({
      diff,
      collapse: this.options.collapseUnchanged ?? true,
      context: this.options.context ?? 2,
      expanded: this.expanded,
      rowHeight: lh,
      modifiedHeight: unified && this.tallRows ? Math.round(lh * 1.6) : lh,
      separatorHeight: SEPARATOR_HEIGHT,
      annotationsAfter,
    });
    this.minimapCache = null;
    this.relayoutColumns();
  }

  /** Recompute column geometry only (resizing a column must not walk every row). */
  private relayoutColumns(): void {
    const diff = this.diff!;
    const keys = new Set(diff.keyColumns);
    const visible: number[] = [];
    diff.columns.forEach((c, i) => {
      if (!this.options.hideUnchangedColumns || c.status !== 'equal' || c.changes > 0 || keys.has(c.name)) {
        visible.push(i);
      }
    });
    // Pinned columns first, in their original order.
    const frozenOpt = this.options.frozenColumns ?? 'keys';
    let frozenCount: number;
    if (frozenOpt === 'keys') {
      const keyCols = visible.filter((i) => keys.has(diff.columns[i].name));
      const rest = visible.filter((i) => !keys.has(diff.columns[i].name));
      visible.splice(0, visible.length, ...keyCols, ...rest);
      frozenCount = keyCols.length;
    } else {
      frozenCount = Math.max(0, frozenOpt);
    }
    this.cols = computeColumnLayout(visible, (i) => this.widthOverrides.get(i) ?? this.widths[i], frozenCount);
    this.updateSpacer();
    this.schedule();
  }

  private updateSpacer(): void {
    const contentW = this.cols.offsets[this.cols.offsets.length - 1] - this.cols.frozenWidth;
    const pane = this.panes()[0];
    const viewW = pane.width - pane.gutter - this.cols.frozenWidth;
    const maxX = Math.max(0, contentW - viewW);
    this.spacer.style.width = `${this.width + maxX}px`;
    this.spacer.style.height = `${HEADER_HEIGHT + this.rows.offsets[this.rows.count] + 8}px`;
    this.positionAnnotations();
  }

  private panes(): PaneGeometry[] {
    const w = this.width - this.minimapWidth;
    const digits = String(Math.max(
      (this.diff?.oldTable.rows.length ?? 0) + this.rowOffset('old'),
      (this.diff?.newTable.rows.length ?? 0) + this.rowOffset('new'),
    )).length;
    const numW = Math.max(28, digits * 7.5 + 14);
    if (this.layout === 'split') {
      const pw = Math.floor((w - SPLIT_DIVIDER) / 2);
      const gutter = numW + BAR_WIDTH + 4;
      return [
        { x: 0, width: pw, gutter, side: 'old' },
        { x: pw + SPLIT_DIVIDER, width: w - pw - SPLIT_DIVIDER, gutter, side: 'new' },
      ];
    }
    return [{ x: 0, width: w, gutter: numW * 2 + BAR_WIDTH + 16, side: 'unified' }];
  }

  /** Display row number = index + this. Real sheet rows when the table knows its origin. */
  private rowOffset(side: 'old' | 'new'): number {
    if (this.options.rowNumberOffset !== undefined) return this.options.rowNumberOffset;
    const t = side === 'old' ? this.diff?.oldTable : this.diff?.newTable;
    return t?.origin ? t.origin.row + 1 : 2;
  }

  private columnLetter(col: AlignedColumn, side: 'old' | 'new'): string {
    const diff = this.diff!;
    const useOld = side === 'old' || col.newIndex < 0;
    const idx = useOld ? col.oldIndex : col.newIndex;
    const origin = (useOld ? diff.oldTable : diff.newTable).origin?.col ?? 0;
    return columnLabel(idx + origin);
  }

  private measureColumns(): void {
    const diff = this.diff!;
    const ctx = this.ctx;
    const t = this.theme;
    const rows = this.sampleRows(400);
    this.numericCols = new Uint8Array(diff.columns.length);
    this.widths = diff.columns.map((c, i) => {
      const fixed = this.options.columnWidths?.[c.name];
      ctx.font = `600 ${t.fontSize - 1}px ${t.fontFamily}`;
      let header = ctx.measureText(c.status === 'renamed' ? `${c.oldName} → ${c.name}` : c.status === 'equal' ? c.name : `+ ${c.name}`).width;
      if (c.status === 'renamed') header += 20;
      ctx.font = `500 ${t.fontSize - 3}px ${t.monoFontFamily}`;
      header += ctx.measureText(this.columnLetter(c, 'new')).width + 6;
      header += CELL_PAD * 2 + 4 + (c.changes ? 30 : 0);
      ctx.font = `${t.fontSize}px ${t.fontFamily}`;
      const ws: number[] = [];
      let numeric = 0;
      let nonEmpty = 0;
      for (const r of rows) {
        for (const [table, idx, ri] of [
          [diff.oldTable, c.oldIndex, diff.oldRows[r]],
          [diff.newTable, c.newIndex, diff.newRows[r]],
        ] as const) {
          if (idx < 0 || ri < 0) continue;
          const v = table.rows[ri][idx] ?? null;
          if (v === null || v === '') continue;
          nonEmpty++;
          if (typeof v === 'number' || (typeof v === 'string' && NUMERIC.test(v))) numeric++;
          const s = formatCell(v);
          ws.push(s.length > 80 ? 80 * 7 : ctx.measureText(s).width);
        }
        // Changed cells in inline / delta displays need room for more than the value.
        if (diff.changedCells.get(r)?.[i] && c.oldIndex >= 0 && c.newIndex >= 0 && (this.display === 'inline' || this.display === 'delta')) {
          const ov = diff.oldTable.rows[diff.oldRows[r]][c.oldIndex] ?? null;
          const nv = diff.newTable.rows[diff.newRows[r]][c.newIndex] ?? null;
          const ot = formatCell(ov);
          const nt = formatCell(nv);
          if (ot.length + nt.length < 80) {
            if (this.display === 'inline') {
              ctx.font = `${t.fontSize - 1}px ${t.fontFamily}`;
              const ow = ctx.measureText(ot).width;
              ctx.font = `${t.fontSize}px ${t.fontFamily}`;
              ws.push(ow + 14 + ctx.measureText(nt).width);
            } else {
              const dl = numericDelta(ov, nv);
              if (dl && dl.delta) {
                ctx.font = `500 ${t.fontSize - 2}px ${t.monoFontFamily}`;
                ws.push(ctx.measureText(formatDelta(dl).label).width);
                ctx.font = `${t.fontSize}px ${t.fontFamily}`;
              }
            }
          }
        }
      }
      if (nonEmpty && numeric / nonEmpty >= 0.8) this.numericCols[i] = 1;
      if (fixed) return fixed;
      ws.sort((a, b) => a - b);
      // Small samples: fit everything. Large: ignore the longest 5% outliers.
      const p90 = ws.length === 0 ? 0 : ws.length < 40 ? ws[ws.length - 1] : ws[Math.floor(ws.length * 0.95)];
      return Math.round(Math.min(MAX_COL, Math.max(MIN_COL, header, p90 + CELL_PAD * 2 + 2)));
    });
  }

  private sampleRows(limit: number): number[] {
    const diff = this.diff!;
    if (diff.rowCount <= limit) return Array.from({ length: diff.rowCount }, (_, i) => i);
    const out: number[] = [];
    for (const r of diff.changedCells.keys()) {
      if (out.length >= limit / 2) break;
      out.push(r);
    }
    const step = Math.max(1, Math.floor(diff.rowCount / (limit - out.length || 1)));
    for (let r = 0; r < diff.rowCount && out.length < limit; r += step) out.push(r);
    return out;
  }

  // ---- Annotations -----------------------------------------------------

  private syncAnnotations(): void {
    const list = this.options.annotations ?? [];
    const render = this.options.renderAnnotation;
    const keep = new Set(list);
    for (const [a, info] of this.annotationEls) {
      if (!keep.has(a) || !render) {
        info.ro.disconnect();
        info.el.remove();
        this.annotationEls.delete(a);
      }
    }
    this.cellMarks.clear();
    for (const a of list) {
      if (a.column) this.cellMarks.add(`${a.row}:${a.column}`);
      if (!render || this.annotationEls.has(a)) continue;
      const content = render(a);
      if (!content) continue;
      const wrap = el('div', 'sd-annotation');
      wrap.style.cssText = 'position:absolute;left:0;top:0;box-sizing:border-box;pointer-events:auto;';
      wrap.appendChild(content);
      this.host.appendChild(wrap);
      const ro = new ResizeObserver(() => {
        const h = Math.max(1, wrap.offsetHeight);
        if (h !== info.height) {
          info.height = h;
          if (this.diff) this.relayout();
        }
      });
      const info = { el: wrap, height: Math.max(1, wrap.offsetHeight || 40), ro };
      this.annotationEls.set(a, info);
      ro.observe(wrap);
    }
  }

  private positionAnnotations(): void {
    if (!this.annotationEls.size || !this.rows) return;
    const top = this.scroller.scrollTop;
    const byRow = new Map<number, HTMLElement[]>();
    for (const [a, info] of this.annotationEls) {
      const list = byRow.get(a.row) ?? [];
      list.push(info.el);
      byRow.set(a.row, list);
    }
    const width = this.width - this.minimapWidth;
    for (const [row, els] of byRow) {
      const d = this.rows.displayOf[row] ?? -1; // out-of-range rows: hidden
      els.forEach((e, k) => {
        if (d < 0) {
          e.style.display = 'none';
          return;
        }
        const y = HEADER_HEIGHT + this.rows.offsets[d + 1 + k] - top;
        e.style.display = y > this.height || y < -2000 ? 'none' : '';
        e.style.width = `${width}px`;
        e.style.transform = `translateY(${y}px)`;
      });
    }
    // Annotations sit under the header.
    this.overlay.style.clipPath = `inset(${HEADER_HEIGHT}px 0 0 0)`;
  }

  // ---- Drawing -----------------------------------------------------------

  private schedule(): void {
    if (this.frame || !this.container) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.drawNow();
    });
  }

  /** Draw synchronously. Exposed for benchmarks. */
  drawNow(): void {
    if (!this.diff || !this.theme || this.width === 0) return;
    // Window moved to a display with a different pixel ratio.
    if ((window.devicePixelRatio || 1) !== this.dpr) {
      this.width = 0;
      this.resize();
      return;
    }
    const key = this.selected ? `${this.selected.row}:${this.selected.col}` : '';
    if (key !== this.focusKey) this.updateFocus();
    const ctx = this.ctx;
    const t = this.theme;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = t.bg;
    ctx.fillRect(0, 0, this.width, this.height);
    ctx.textBaseline = 'middle';
    const panes = this.panes();
    for (const p of panes) this.drawPane(p);
    if (panes.length === 2) {
      ctx.fillStyle = t.borderStrong;
      ctx.fillRect(panes[1].x - SPLIT_DIVIDER, 0, SPLIT_DIVIDER, this.height);
    }
    this.drawMinimap();
    this.positionAnnotations();
  }

  private drawPane(p: PaneGeometry): void {
    const ctx = this.ctx;
    const t = this.theme;
    const diff = this.diff!;
    const rows = this.rows;
    const cols = this.cols;
    const scrollY = this.scroller.scrollTop;
    const scrollX = this.scroller.scrollLeft;
    const bodyTop = HEADER_HEIGHT;
    const dataX = p.x + p.gutter;
    const scrollStart = dataX + cols.frozenWidth;
    const right = p.x + p.width;

    ctx.save();
    ctx.beginPath();
    ctx.rect(p.x, 0, p.width, this.height);
    ctx.clip();

    // Visible column range in the scrolled region.
    const firstScrolled = Math.max(cols.frozen, colAt(cols, cols.frozenWidth + scrollX));
    const lastScrolled = colAt(cols, cols.frozenWidth + scrollX + (right - scrollStart));
    const colX = (vi: number) =>
      vi < cols.frozen ? dataX + cols.offsets[vi] : scrollStart + cols.offsets[vi] - cols.frozenWidth - scrollX;

    // Visible rows.
    const first = Math.max(0, rowAt(rows, scrollY));
    const yEnd = scrollY + this.height - bodyTop;

    // ---- Body
    for (let d = first; d < rows.count && rows.offsets[d] < yEnd; d++) {
      const y = bodyTop + rows.offsets[d] - scrollY;
      const h = rows.offsets[d + 1] - rows.offsets[d];
      const kind = rows.kind[d];
      if (kind === ROW_KIND.separator) {
        this.drawSeparator(p, d, y, h);
        continue;
      }
      if (kind === ROW_KIND.annotation) {
        ctx.fillStyle = t.gutterBg;
        ctx.fillRect(p.x, y, p.width, h);
        hline(ctx, p.x, right, y + h, t.border);
        continue;
      }
      const r = rows.ref[d];
      const status = diff.status[r];
      const missing = (p.side === 'old' && status === ROW_STATUS.added) || (p.side === 'new' && status === ROW_STATUS.removed);
      // Row background.
      let rowBg = t.bg;
      if (missing) rowBg = t.bg;
      else if (status === ROW_STATUS.added) rowBg = t.addedBg;
      else if (status === ROW_STATUS.removed) rowBg = t.removedBg;
      ctx.fillStyle = rowBg;
      ctx.fillRect(p.x, y, p.width, h);
      if (missing && this.hatch) {
        ctx.fillStyle = this.hatch;
        ctx.fillRect(dataX, y, right - dataX, h);
      }

      // Scrolled cells first, then frozen ones on top.
      ctx.save();
      ctx.beginPath();
      ctx.rect(scrollStart, y, right - scrollStart, h);
      ctx.clip();
      for (let vi = firstScrolled; vi <= lastScrolled && vi < cols.cols.length; vi++) {
        this.drawCell(p, r, status, missing, vi, colX(vi), y, cols.widths[vi], h);
      }
      ctx.restore();
      for (let vi = 0; vi < cols.frozen; vi++) {
        ctx.fillStyle = rowBg;
        ctx.fillRect(colX(vi), y, cols.widths[vi], h);
        this.drawCell(p, r, status, missing, vi, colX(vi), y, cols.widths[vi], h);
      }
      if (cols.frozen) vline(ctx, scrollStart, y, y + h, t.borderStrong);

      // Hover / selection.
      if (this.hover && this.hover.display === d && (this.hover.area === 'cell' || this.hover.area === 'gutter')) {
        ctx.fillStyle = t.hover;
        ctx.fillRect(p.x, y, p.width, h);
      }
      this.drawGutter(p, r, status, missing, y, h, d);
      hline(ctx, p.x, right, y + h, t.border);
    }

    // Selection ring (drawn after rows so neighbors don't paint over it).
    if (this.selected) {
      const d = rows.displayOf[this.selected.row];
      const vi = cols.cols.indexOf(this.selected.col);
      if (d >= 0 && vi >= 0) {
        const y = bodyTop + rows.offsets[d] - scrollY;
        const h = rows.offsets[d + 1] - rows.offsets[d];
        const x = colX(vi);
        ctx.save();
        ctx.beginPath();
        const clipX = vi < cols.frozen ? dataX : scrollStart;
        ctx.rect(clipX, bodyTop, right - clipX, this.height - bodyTop);
        ctx.clip();
        ctx.strokeStyle = t.selection;
        ctx.lineWidth = 2;
        roundRect(ctx, x + 1, y + 1, cols.widths[vi] - 2, h - 2, 3);
        ctx.stroke();
        ctx.restore();
      }
    }

    if (this.focus) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(dataX, bodyTop, right - dataX, this.height - bodyTop);
      ctx.clip();
      ctx.lineWidth = 2;
      for (const [k, role] of this.focus) {
        const [row, col] = k.split(':').map(Number);
        const d = rows.displayOf[row];
        const vi = cols.cols.indexOf(col);
        if (d < 0 || vi < 0) continue;
        const y = bodyTop + rows.offsets[d] - scrollY;
        if (y > this.height || y + 100 < bodyTop) continue;
        const h = rows.offsets[d + 1] - rows.offsets[d];
        ctx.strokeStyle = role === 'impact' ? t.computed : t.modified;
        ctx.setLineDash(role === 'impact' ? [4, 3] : role === 'parent' ? [1.5, 2.5] : []);
        roundRect(ctx, colX(vi) + 1.5, y + 1.5, cols.widths[vi] - 3, h - 3, 3);
        ctx.stroke();
      }
      ctx.setLineDash([]);
      ctx.restore();
    }

    this.drawHeader(p, firstScrolled, lastScrolled, colX);
    ctx.restore();
  }

  private drawCell(
    p: PaneGeometry,
    r: number,
    status: number,
    missing: boolean,
    vi: number,
    x: number,
    y: number,
    w: number,
    h: number,
  ): void {
    const ctx = this.ctx;
    const t = this.theme;
    const diff = this.diff!;
    const ci = this.cols.cols[vi];
    const col = diff.columns[ci];
    const right = x + w;
    vline(ctx, right, y, y + h, t.border);
    if (missing) return;

    const colMissing = (p.side === 'old' && col.status === 'added') || (p.side === 'new' && col.status === 'removed');
    if (colMissing) {
      if (this.hatch) {
        ctx.fillStyle = this.hatch;
        ctx.fillRect(x, y, w - 1, h);
      }
      return;
    }
    const oi = diff.oldRows[r];
    const ni = diff.newRows[r];
    const oldV = oi >= 0 && col.oldIndex >= 0 ? diff.oldTable.rows[oi][col.oldIndex] ?? null : undefined;
    const newV = ni >= 0 && col.newIndex >= 0 ? diff.newTable.rows[ni][col.newIndex] ?? null : undefined;
    const changed = status === ROW_STATUS.modified && diff.changedCells.get(r)![ci] === 1;
    const align = this.numericCols[ci] ? 'right' : 'left';
    const font = `${t.fontSize}px ${t.fontFamily}`;

    // Column-level add/remove tint (only on rows that exist on both sides).
    if (status === ROW_STATUS.equal || status === ROW_STATUS.modified) {
      if (col.status === 'added' && p.side !== 'old') {
        ctx.fillStyle = t.addedBg;
        ctx.fillRect(x, y, w - 1, h);
      } else if (col.status === 'removed' && p.side !== 'new') {
        ctx.fillStyle = t.removedBg;
        ctx.fillRect(x, y, w - 1, h);
      }
    }

    let oldText = formatCell(oldV ?? null);
    let newText = formatCell(newV ?? null);
    if (changed && oldText === newText) {
      // Same display text, different cell: a formula-only change, or digits
      // beyond display precision. Show what actually differs.
      const of = oi >= 0 ? diff.oldTable.formulas?.[oi]?.[col.oldIndex] ?? null : null;
      const nf = ni >= 0 ? diff.newTable.formulas?.[ni]?.[col.newIndex] ?? null : null;
      if (of !== nf) {
        oldText = of ?? oldText;
        newText = nf ?? newText;
      } else {
        oldText = String(oldV);
        newText = String(newV);
      }
    }
    if (changed) {
      this.drawChange(p, r, ci, x, y, w, h, oldText, newText, oldV, newV, align);
    } else {
      const v = p.side === 'old' ? oldV : p.side === 'new' ? newV : newV !== undefined ? newV : oldV;
      let color = t.fg;
      if (status === ROW_STATUS.removed || (p.side === 'unified' && col.status === 'removed')) color = t.fgMuted;
      this.text(formatCell(v ?? null), x, y + h / 2 + 0.5, w, align, font, color, p.side === 'unified' && status === ROW_STATUS.removed);
    }
    if (this.focus && !(this.selected?.row === r && this.selected.col === ci) && !this.focus.has(`${r}:${ci}`)) {
      ctx.fillStyle = t.dim;
      ctx.fillRect(x, y, w, h);
    }

    if (this.cellMarks.size && this.cellMarks.has(`${r}:${col.name}`)) {
      ctx.fillStyle = t.modified;
      ctx.beginPath();
      ctx.moveTo(right - 8, y);
      ctx.lineTo(right - 1, y);
      ctx.lineTo(right - 1, y + 7);
      ctx.closePath();
      ctx.fill();
    }
  }

  /** Recompute which cells the selection's trace points at. */
  private updateFocus(): void {
    const sel = this.selected;
    this.focusKey = sel ? `${sel.row}:${sel.col}` : '';
    this.focus = null;
    const tc = this.options.trace;
    if (!sel || !tc || this.options.traceFocus === false) return;
    const ref = { sheet: tc.sheet, row: sel.row, col: sel.col };
    const t = tc.trace.get(ref);
    if (!t) return;
    const m = new Map<string, 'cause' | 'impact' | 'parent'>();
    for (const c of t.causes) if (c.sheet === tc.sheet) m.set(`${c.row}:${c.col}`, 'cause');
    for (const c of t.impacts) if (c.sheet === tc.sheet) m.set(`${c.row}:${c.col}`, 'impact');
    if (t.kind === 'computed') {
      for (const c of tc.trace.parents(ref)) if (c.sheet === tc.sheet && !m.has(`${c.row}:${c.col}`)) m.set(`${c.row}:${c.col}`, 'parent');
    }
    // Relations only on other sheets: don't dim everything here.
    this.focus = m.size ? m : null;
    this.schedule();
  }

  /** `Sheet!B7 price` for a traced cell, with its change. */
  private describeRef(ref: CellRef): { label: string; delta: string; sign: '+' | '-' | '' } {
    const tc = this.options.trace!;
    const sheet = tc.workbook.sheets[ref.sheet];
    const d = sheet.diff;
    const c = d.columns[ref.col];
    const ni = d.newRows[ref.row];
    const oi = d.oldRows[ref.row];
    const side = ni >= 0 ? 'new' : 'old';
    const table = side === 'new' ? d.newTable : d.oldTable;
    const ri = side === 'new' ? ni : oi;
    const ci = side === 'new' ? c.newIndex : c.oldIndex;
    const origin = table.origin ?? { row: 1, col: 0 };
    const prefix = ref.sheet === tc.sheet ? '' : `${sheet.name}!`;
    // Name the row by its first cell when that's a label ("Price"), else the column.
    const first = table.rows[ri]?.[0];
    const rowName = ref.col !== 0 && typeof first === 'string' && first ? first : c.name;
    const label = `${prefix}${columnLabel(ci + origin.col)}${ri + origin.row + 1} ${rowName}`;
    const ov = oi >= 0 && c.oldIndex >= 0 ? d.oldTable.rows[oi][c.oldIndex] ?? null : null;
    const nv = ni >= 0 && c.newIndex >= 0 ? d.newTable.rows[ni][c.newIndex] ?? null : null;
    if (ni < 0) return { label, delta: 'row removed', sign: '-' };
    if (oi < 0) return { label, delta: 'row added', sign: '+' };
    const dl = numericDelta(ov, nv);
    if (dl && dl.delta !== 0) {
      const f = formatDelta(dl);
      return { label, delta: f.label, sign: f.up ? '+' : '-' };
    }
    return { label, delta: `${formatCell(ov) || '∅'} → ${formatCell(nv) || '∅'}`, sign: '' };
  }

  /** The delta label, or a shorter form (% only, then absolute only) that fits. */
  private fitDelta(d: { label: string; pct?: string; abs?: string }, w: number, font: string): string {
    this.ctx.font = font;
    const room = w - CELL_PAD * 2;
    for (const s of [d.label, d.pct, d.abs]) if (s && this.ctx.measureText(s).width <= room) return s;
    return d.pct ?? d.label;
  }

  /** A changed cell, in the configured display. */
  private drawChange(
    p: PaneGeometry,
    r: number,
    ci: number,
    x: number,
    y: number,
    w: number,
    h: number,
    oldText: string,
    newText: string,
    oldV: CellValue | undefined,
    newV: CellValue | undefined,
    align: 'left' | 'right',
  ): void {
    const ctx = this.ctx;
    const t = this.theme;
    const kind = this.options.highlightComputed === false ? 'input' : this.kindOf(r, ci);
    const computed = kind === 'computed';
    const font = `${t.fontSize}px ${t.fontFamily}`;
    const small = `${t.fontSize - 1}px ${t.fontFamily}`;
    const tiny = `500 ${t.fontSize - 2}px ${t.monoFontFamily}`;
    const newBg = computed ? t.computedEmphasis : t.addedEmphasis;
    // Computed: the old value isn't "wrong", it was just recalculated; keep it quiet.
    const oldBg = computed ? t.separatorBg : t.removedEmphasis;
    const oldFg = computed ? t.fgMuted : t.removedFg;
    const mode = this.display;
    const delta = numericDelta(oldV ?? null, newV ?? null);
    const deltaText = delta && delta.delta !== 0 ? formatDelta(delta) : null;

    if (p.side === 'old') {
      ctx.fillStyle = oldBg;
      ctx.fillRect(x, y, w - 1, h);
      this.text(oldText, x, y + h / 2 + 0.5, w, align, font, t.fg);
    } else if (p.side === 'new') {
      ctx.fillStyle = newBg;
      ctx.fillRect(x, y, w - 1, h);
      if (mode === 'delta' && deltaText) {
        // Split rows are one line: the change sits on the opposite side of the value.
        this.text(newText, x, y + h / 2 + 0.5, w, align, font, t.fg);
        this.text(this.fitDelta(deltaText, w / 2, tiny), x, y + h / 2 + 0.5, w, align === 'right' ? 'left' : 'right', tiny, deltaText.up ? t.added : t.removed);
      } else {
        this.text(newText, x, y + h / 2 + 0.5, w, align, font, t.fg);
      }
    } else if (mode === 'stacked') {
      const half = Math.floor(h / 2);
      ctx.fillStyle = oldBg;
      ctx.fillRect(x, y, w - 1, half);
      ctx.fillStyle = newBg;
      ctx.fillRect(x, y + half, w - 1, h - half);
      this.text(oldText, x, y + half / 2 + 0.5, w, align, small, oldFg, true);
      this.text(newText, x, y + half + (h - half) / 2 + 0.5, w, align, font, t.fg);
    } else if (mode === 'new') {
      ctx.fillStyle = newBg;
      ctx.fillRect(x, y, w - 1, h);
      this.text(newText, x, y + h / 2 + 0.5, w, align, font, t.fg);
    } else if (mode === 'delta') {
      ctx.fillStyle = newBg;
      ctx.fillRect(x, y, w - 1, h);
      const top = y + h * 0.36;
      const bottom = y + h * 0.72;
      this.text(newText, x, top, w, align, font, t.fg);
      if (deltaText) this.text(this.fitDelta(deltaText, w, tiny), x, bottom, w, align, tiny, deltaText.up ? t.added : t.removed);
      else this.text(`was ${oldText || '∅'}`, x, bottom, w, align, small, oldFg);
    } else {
      // inline: old → new on one line; falls back to the new value when narrow.
      ctx.fillStyle = newBg;
      ctx.fillRect(x, y, w - 1, h);
      const cy = y + h / 2 + 0.5;
      ctx.font = small;
      const ow = ctx.measureText(oldText).width;
      ctx.font = font;
      const nw = ctx.measureText(newText).width;
      const arrow = 14;
      if (ow + nw + arrow + 2 * CELL_PAD <= w) {
        const start = align === 'right' ? x + w - CELL_PAD - nw - arrow - ow : x + CELL_PAD;
        this.text(oldText, start - CELL_PAD, cy, ow + 2 * CELL_PAD + 1, 'left', small, oldFg, true);
        this.text('→', start + ow - CELL_PAD + 2, cy, arrow + 2 * CELL_PAD, 'left', small, t.fgFaint);
        this.text(newText, start + ow + arrow - CELL_PAD, cy, nw + 2 * CELL_PAD + 1, 'left', font, t.fg);
      } else {
        this.text(newText, x, cy, w, align, font, t.fg);
      }
    }
    // Formula edits get an ƒ in the corner; computed cells a thin amber edge.
    if (kind === 'formula' && p.side !== 'old') {
      ctx.font = `italic 600 ${t.fontSize - 3}px ${t.fontFamily}`;
      ctx.fillStyle = t.formula;
      ctx.textAlign = 'left';
      ctx.fillText('ƒ', x + 3, y + 8);
    } else if (computed && p.side !== 'old') {
      ctx.fillStyle = t.computed;
      ctx.fillRect(x, y, 2, h);
    }
  }

  private drawGutter(p: PaneGeometry, r: number, status: number, missing: boolean, y: number, h: number, d: number): void {
    const ctx = this.ctx;
    const t = this.theme;
    const diff = this.diff!;
    const offOld = this.rowOffset('old');
    const offNew = this.rowOffset('new');
    const bg =
      missing ? t.gutterBg
      : status === ROW_STATUS.added ? t.addedBg
      : status === ROW_STATUS.removed ? t.removedBg
      : t.gutterBg;
    ctx.fillStyle = bg;
    ctx.fillRect(p.x, y, p.gutter, h);
    if (missing && this.hatch) {
      ctx.fillStyle = this.hatch;
      ctx.fillRect(p.x, y, p.gutter, h);
    }
    const accent =
      status === ROW_STATUS.added ? t.added
      : status === ROW_STATUS.removed ? t.removed
      : status === ROW_STATUS.modified ? t.modified
      : null;
    if (accent && !missing) {
      ctx.fillStyle = accent;
      ctx.fillRect(p.x, y, BAR_WIDTH, h);
    }
    const selected = this.selected?.row === r;
    const hovered = this.hover?.display === d;
    const numColor = selected || hovered ? t.fg : t.fgFaint;
    const mono = `${t.fontSize - 2}px ${t.monoFontFamily}`;
    const cy = y + h / 2 + 0.5;
    const oi = diff.oldRows[r];
    const ni = diff.newRows[r];
    if (p.side === 'unified') {
      const numW = (p.gutter - BAR_WIDTH - 16) / 2;
      if (oi >= 0) this.text(String(oi + offOld), p.x + BAR_WIDTH, cy, numW, 'right', mono, numColor);
      if (ni >= 0) this.text(String(ni + offNew), p.x + BAR_WIDTH + numW, cy, numW, 'right', mono, numColor);
      const sign = status === ROW_STATUS.added ? '+' : status === ROW_STATUS.removed ? '−' : status === ROW_STATUS.modified ? '~' : '';
      if (sign) {
        ctx.font = `600 ${t.fontSize}px ${t.monoFontFamily}`;
        ctx.fillStyle = accent!;
        ctx.textAlign = 'center';
        ctx.fillText(sign, p.x + p.gutter - 9, cy);
      }
    } else if (!missing) {
      const idx = p.side === 'old' ? oi : ni;
      const off = p.side === 'old' ? offOld : offNew;
      if (idx >= 0) this.text(String(idx + off), p.x + BAR_WIDTH, cy, p.gutter - BAR_WIDTH, 'right', mono, numColor);
    }
    vline(this.ctx, p.x + p.gutter, y, y + h, t.border);
  }

  private drawSeparator(p: PaneGeometry, d: number, y: number, h: number): void {
    const ctx = this.ctx;
    const t = this.theme;
    const region = this.rows.regions[this.rows.ref[d]];
    const hovered = this.hover?.area === 'separator' && this.hover.display === d;
    ctx.fillStyle = t.separatorBg;
    ctx.fillRect(p.x, y, p.width, h);
    if (hovered) {
      ctx.fillStyle = t.hover;
      ctx.fillRect(p.x, y, p.width, h);
    }
    hline(ctx, p.x, p.x + p.width, y, t.border);
    hline(ctx, p.x, p.x + p.width, y + h, t.border);
    const n = region.end - region.start;
    const label = `${n.toLocaleString()} unchanged ${n === 1 ? 'row' : 'rows'}`;
    const cy = y + h / 2 + 0.5;
    // Expand glyph: two small chevrons.
    const gx = p.x + p.gutter / 2;
    ctx.strokeStyle = hovered ? t.fg : t.fgMuted;
    ctx.lineWidth = 1.25;
    ctx.beginPath();
    ctx.moveTo(gx - 4, cy - 2);
    ctx.lineTo(gx, cy - 6);
    ctx.lineTo(gx + 4, cy - 2);
    ctx.moveTo(gx - 4, cy + 2);
    ctx.lineTo(gx, cy + 6);
    ctx.lineTo(gx + 4, cy + 2);
    ctx.stroke();
    this.text(label, p.x + p.gutter + 2, cy, 400, 'left', `${t.fontSize - 1}px ${t.fontFamily}`, hovered ? t.fg : t.fgMuted);
  }

  private drawHeader(p: PaneGeometry, firstScrolled: number, lastScrolled: number, colX: (vi: number) => number): void {
    const ctx = this.ctx;
    const t = this.theme;
    const cols = this.cols;
    const dataX = p.x + p.gutter;
    const scrollStart = dataX + cols.frozenWidth;
    const right = p.x + p.width;
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(p.x, 0, p.width, HEADER_HEIGHT);
    ctx.save();
    ctx.beginPath();
    ctx.rect(scrollStart, 0, right - scrollStart, HEADER_HEIGHT);
    ctx.clip();
    for (let vi = firstScrolled; vi <= lastScrolled && vi < cols.cols.length; vi++) {
      this.drawHeaderCell(p, vi, colX(vi), cols.widths[vi]);
    }
    ctx.restore();
    for (let vi = 0; vi < cols.frozen; vi++) {
      ctx.fillStyle = t.headerBg;
      ctx.fillRect(colX(vi), 0, cols.widths[vi], HEADER_HEIGHT);
      this.drawHeaderCell(p, vi, colX(vi), cols.widths[vi]);
    }
    if (cols.frozen) vline(ctx, scrollStart, 0, HEADER_HEIGHT, t.borderStrong);
    // Gutter header: pane label.
    ctx.fillStyle = t.headerBg;
    ctx.fillRect(p.x, 0, p.gutter, HEADER_HEIGHT);
    const labelFont = `600 ${t.fontSize - 4}px ${t.fontFamily}`;
    const ly = HEADER_HEIGHT / 2 + 0.5;
    // Labels sit right-aligned over their number columns (no cell padding).
    const label = (s: string, rightX: number, color: string) => {
      ctx.font = labelFont;
      ctx.fillStyle = color;
      ctx.textAlign = 'right';
      ctx.fillText(s, rightX, ly);
    };
    if (p.side !== 'unified') {
      label(p.side === 'old' ? 'OLD' : 'NEW', p.x + p.gutter - CELL_PAD, p.side === 'old' ? t.removed : t.added);
    } else {
      const numW = (p.gutter - BAR_WIDTH - 16) / 2;
      label('OLD', p.x + BAR_WIDTH + numW - CELL_PAD, t.removed);
      label('NEW', p.x + BAR_WIDTH + numW * 2 - CELL_PAD, t.added);
    }
    vline(ctx, p.x + p.gutter, 0, HEADER_HEIGHT, t.border);
    hline(ctx, p.x, right, HEADER_HEIGHT, t.borderStrong);
  }

  private drawHeaderCell(p: PaneGeometry, vi: number, x: number, w: number): void {
    const ctx = this.ctx;
    const t = this.theme;
    const ci = this.cols.cols[vi];
    const col = this.diff!.columns[ci];
    const cy = HEADER_HEIGHT / 2 + 0.5;
    const missing = (p.side === 'old' && col.status === 'added') || (p.side === 'new' && col.status === 'removed');
    if (missing) {
      if (this.hatch) {
        ctx.fillStyle = this.hatch;
        ctx.fillRect(x, 0, w, HEADER_HEIGHT);
      }
    } else if (col.status === 'added') {
      ctx.fillStyle = t.addedBg;
      ctx.fillRect(x, 0, w, HEADER_HEIGHT);
    } else if (col.status === 'removed') {
      ctx.fillStyle = t.removedBg;
      ctx.fillRect(x, 0, w, HEADER_HEIGHT);
    }
    if (this.hover?.area === 'header' && this.hover.vcol === vi) {
      ctx.fillStyle = t.hover;
      ctx.fillRect(x, 0, w, HEADER_HEIGHT);
    }
    if (this.selected?.col === ci) {
      ctx.fillStyle = t.selection;
      ctx.fillRect(x, HEADER_HEIGHT - 2, w, 2);
    }
    vline(ctx, x + w, 0, HEADER_HEIGHT, t.border);
    if (missing) return;
    const letter = this.columnLetter(col, p.side === 'old' ? 'old' : 'new');
    const small = `500 ${t.fontSize - 3}px ${t.monoFontFamily}`;
    ctx.font = small;
    const lw = ctx.measureText(letter).width;
    ctx.fillStyle = t.fgFaint;
    ctx.textAlign = 'left';
    ctx.fillText(letter, x + CELL_PAD, cy);
    let tx = x + CELL_PAD + lw + 6;
    const bold = `600 ${t.fontSize - 1}px ${t.fontFamily}`;
    let avail = x + w - tx - CELL_PAD;
    if (col.changes > 0) avail -= 28;
    const accent = col.status === 'added' ? t.addedFg : col.status === 'removed' ? t.removedFg : t.fg;
    if (col.status === 'renamed' && p.side !== 'new') {
      const oldName = col.oldName ?? '';
      if (p.side === 'old') {
        this.text(oldName, tx - CELL_PAD, cy, avail + CELL_PAD * 2, 'left', bold, t.fg);
      } else {
        ctx.font = bold;
        const ow = Math.min(ctx.measureText(oldName).width, avail * 0.45);
        this.text(oldName, tx - CELL_PAD, cy, ow + CELL_PAD * 2, 'left', bold, t.fgMuted, true);
        tx += ow + 4;
        this.text('→', tx - CELL_PAD, cy, 14 + CELL_PAD * 2, 'left', bold, t.modified);
        tx += 16;
        this.text(col.name, tx - CELL_PAD, cy, x + w - tx - (col.changes ? 28 : 0), 'left', bold, t.fg);
      }
    } else {
      const prefix = col.status === 'added' && p.side === 'unified' ? '+ ' : col.status === 'removed' && p.side === 'unified' ? '− ' : '';
      this.text(prefix + col.name, tx - CELL_PAD, cy, avail + CELL_PAD * 2, 'left', bold, accent, col.status === 'removed' && p.side === 'unified');
    }
    if (col.changes > 0) {
      const label = col.changes > 999 ? `${Math.round(col.changes / 100) / 10}k` : String(col.changes);
      ctx.font = `600 ${t.fontSize - 3}px ${t.fontFamily}`;
      const bw = Math.max(18, ctx.measureText(label).width + 10);
      const bx = x + w - CELL_PAD - bw + 4;
      ctx.fillStyle = mixAlpha(t.modified, 0.16);
      roundRect(ctx, bx, cy - 8, bw, 16, 8);
      ctx.fill();
      ctx.fillStyle = t.modified;
      ctx.textAlign = 'center';
      ctx.fillText(label, bx + bw / 2, cy + 0.5);
    }
  }

  private drawMinimap(): void {
    const mw = this.minimapWidth;
    this.minimap.style.display = mw ? '' : 'none';
    if (!mw) return;
    const ctx = this.mctx;
    const t = this.theme;
    const h = this.height;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = t.gutterBg;
    ctx.fillRect(0, 0, mw, h);
    vline(ctx, 0, 0, h, t.border);
    const trackTop = HEADER_HEIGHT;
    const trackH = h - trackTop;
    // Scale to at least the viewport, so a short diff isn't stretched to
    // fill the strip (4 rows would otherwise look like the whole file).
    const total = Math.max(this.rows.offsets[this.rows.count], trackH);
    if (this.rows.count === 0 || trackH <= 0) return;
    const key = `${this.rows.count}:${total}:${trackH}`;
    if (this.minimapCache?.key !== key) {
      const colors: (string | null)[] = new Array(Math.ceil(trackH)).fill(null);
      const rank = (s: number) => (s === ROW_STATUS.removed ? 3 : s === ROW_STATUS.added ? 2 : s === ROW_STATUS.modified ? 1 : 0);
      const ranks = new Uint8Array(colors.length);
      for (let d = 0; d < this.rows.count; d++) {
        if (this.rows.kind[d] !== ROW_KIND.data) continue;
        const s = this.diff!.status[this.rows.ref[d]];
        const k = rank(s);
        if (!k) continue;
        const y0 = Math.floor((this.rows.offsets[d] / total) * trackH);
        const y1 = Math.max(y0 + 1, Math.ceil((this.rows.offsets[d + 1] / total) * trackH));
        for (let y = y0; y < y1 && y < colors.length; y++) {
          if (k > ranks[y]) {
            ranks[y] = k;
            colors[y] = k === 3 ? t.removed : k === 2 ? t.added : t.modified;
          }
        }
      }
      this.minimapCache = { key, colors };
    }
    const colors = this.minimapCache.colors;
    let y = 0;
    while (y < colors.length) {
      const c = colors[y];
      let e = y + 1;
      while (e < colors.length && colors[e] === c) e++;
      if (c) {
        ctx.fillStyle = c;
        ctx.fillRect(3, trackTop + y, mw - 5, e - y);
      }
      y = e;
    }
    // Viewport indicator.
    const vt = (this.scroller.scrollTop / total) * trackH;
    const vh = Math.max(6, ((h - HEADER_HEIGHT) / total) * trackH);
    ctx.fillStyle = t.hover;
    ctx.fillRect(1, trackTop + vt, mw - 1, Math.min(vh, trackH));
    ctx.strokeStyle = t.borderStrong;
    ctx.lineWidth = 1;
    ctx.strokeRect(1.5, trackTop + vt + 0.5, mw - 2.5, Math.min(vh, trackH) - 1);
  }

  /** Draw single-line text clipped to a cell with an ellipsis. */
  private text(
    s: string,
    x: number,
    cy: number,
    w: number,
    align: 'left' | 'right',
    font: string,
    color: string,
    strike = false,
  ): void {
    if (!s) return;
    const ctx = this.ctx;
    const max = w - CELL_PAD * 2;
    if (max <= 4) return;
    ctx.font = font;
    // Never more than fits in a cell; also keeps huge cells out of the cache key.
    if (s.length > 400) s = s.slice(0, 400);
    const nl = s.indexOf('\n');
    if (nl >= 0) s = s.slice(0, nl) + ' ↵';
    const key = `${font}|${max | 0}|${s}`;
    let fitted = this.textCache.get(key);
    if (fitted === undefined) {
      fitted = fitText(ctx, s, max);
      if (this.textCache.size > 50_000) this.textCache.clear();
      this.textCache.set(key, fitted);
    }
    ctx.fillStyle = color;
    ctx.textAlign = align;
    const tx = align === 'left' ? x + CELL_PAD : x + w - CELL_PAD;
    ctx.fillText(fitted, tx, cy);
    if (strike) {
      const tw = ctx.measureText(fitted).width;
      const sx = align === 'left' ? tx : tx - tw;
      ctx.fillRect(sx, Math.round(cy) - 0.5, tw, 1);
    }
  }

  private makeHatch(): CanvasPattern | null {
    const size = 8 * this.dpr;
    const c = document.createElement('canvas');
    c.width = c.height = Math.max(1, Math.round(size));
    const g = c.getContext('2d')!;
    g.strokeStyle = this.theme.hatch;
    g.lineWidth = 1.25 * this.dpr;
    g.beginPath();
    for (let i = -1; i <= 1; i++) {
      g.moveTo(i * size, size);
      g.lineTo(i * size + size, 0);
    }
    g.stroke();
    const pattern = this.ctx.createPattern(c, 'repeat');
    pattern?.setTransform(new DOMMatrix().scale(1 / this.dpr));
    return pattern;
  }

  // ---- Interaction -------------------------------------------------------

  private hitTest(clientX: number, clientY: number): Hit {
    const rect = this.scroller.getBoundingClientRect();
    const x = clientX - rect.left;
    const y = clientY - rect.top;
    const none: Hit = { area: 'none', pane: 0, display: -1, vcol: -1 };
    if (x < 0 || y < 0 || x > this.width || y > this.height) return none;
    if (this.minimapWidth && x >= this.width - this.minimapWidth) return { ...none, area: 'minimap' };
    const panes = this.panes();
    const pi = panes.length === 2 && x >= panes[1].x ? 1 : 0;
    const p = panes[pi];
    const cols = this.cols;
    const dataX = p.x + p.gutter;
    let vcol = -1;
    let resize: number | undefined;
    if (x >= dataX) {
      const cx = x - dataX;
      const contentX = cx < cols.frozenWidth ? cx : cx + this.scroller.scrollLeft;
      vcol = colAt(cols, contentX);
      if (vcol >= cols.cols.length) vcol = -1;
      if (vcol >= 0) {
        const edge = cols.offsets[vcol + 1];
        const leftEdge = cols.offsets[vcol];
        if (Math.abs(contentX - edge) <= 4) resize = vcol;
        else if (vcol > 0 && Math.abs(contentX - leftEdge) <= 4) resize = vcol - 1;
      }
    }
    if (y < HEADER_HEIGHT) return { area: 'header', pane: pi as 0 | 1, display: -1, vcol, resize };
    const cy = y - HEADER_HEIGHT + this.scroller.scrollTop;
    if (cy >= this.rows.offsets[this.rows.count]) return none;
    const d = rowAt(this.rows, cy);
    const kind = this.rows.kind[d];
    if (kind === ROW_KIND.separator) return { area: 'separator', pane: pi as 0 | 1, display: d, vcol };
    if (kind === ROW_KIND.annotation) return { area: 'annotation', pane: pi as 0 | 1, display: d, vcol };
    if (x < dataX) return { area: 'gutter', pane: pi as 0 | 1, display: d, vcol: -1 };
    return { area: vcol >= 0 ? 'cell' : 'none', pane: pi as 0 | 1, display: d, vcol };
  }

  private onPointerMove(e: PointerEvent): void {
    this.lastPointer = { x: e.clientX, y: e.clientY };
    if (this.resizing) {
      const w = Math.max(MIN_COL / 2, this.resizing.startW + e.clientX - this.resizing.startX);
      this.widthOverrides.set(this.cols.cols[this.resizing.vcol], Math.round(w));
      this.relayoutColumns();
      return;
    }
    const hit = this.hitTest(e.clientX, e.clientY);
    const prev = this.hover;
    this.hover = hit;
    this.scroller.style.cursor =
      hit.area === 'header' && hit.resize !== undefined ? 'col-resize'
      : hit.area === 'separator' || hit.area === 'minimap' ? 'pointer'
      : 'default';
    if (!prev || prev.display !== hit.display || prev.vcol !== hit.vcol || prev.area !== hit.area || prev.pane !== hit.pane) {
      this.schedule();
      const info = this.cellAt(hit);
      this.options.onCellHover?.(info);
      this.hideTooltip();
      if (info?.changed && this.options.tooltips !== false) {
        this.tooltipTimer = window.setTimeout(() => this.showTooltip(info, hit), 260);
      }
    }
  }

  private onPointerDown(e: PointerEvent): void {
    const hit = this.hitTest(e.clientX, e.clientY);
    if (hit.area === 'header' && hit.resize !== undefined) {
      e.preventDefault();
      this.scroller.setPointerCapture(e.pointerId);
      this.resizing = { vcol: hit.resize, startX: e.clientX, startW: this.cols.widths[hit.resize] };
      const up = () => {
        this.resizing = null;
        for (const t of ['pointerup', 'pointercancel', 'lostpointercapture'] as const) this.scroller.removeEventListener(t, up);
      };
      for (const t of ['pointerup', 'pointercancel', 'lostpointercapture'] as const) this.scroller.addEventListener(t, up);
      return;
    }
    if (hit.area === 'minimap') {
      const rect = this.scroller.getBoundingClientRect();
      const frac = (e.clientY - rect.top - HEADER_HEIGHT) / (this.height - HEADER_HEIGHT);
      const total = Math.max(this.rows.offsets[this.rows.count], this.height - HEADER_HEIGHT);
      this.scroller.scrollTop = Math.max(0, frac * total - (this.height - HEADER_HEIGHT) / 2);
      return;
    }
    if (hit.area === 'separator') {
      const g = this.rows.regions[this.rows.ref[hit.display]];
      this.expanded.add(`${g.start}:${g.end}`);
      this.relayout();
      return;
    }
    const info = this.cellAt(hit);
    if (info) {
      this.selected = { row: info.row, col: info.columnIndex };
      this.options.onSelectionChange?.(info);
      this.options.onCellClick?.(info, e);
      this.schedule();
    }
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (e.metaKey || e.ctrlKey) {
      if (e.key === 'c' && this.selected) {
        const info = this.cellInfo(this.selected.row, this.selected.col);
        const v = info.newValue !== undefined ? info.newValue : info.oldValue;
        // Copy the exact value, not the display-rounded one.
        void navigator.clipboard?.writeText(typeof v === 'number' ? String(v) : formatCell(v ?? null));
        e.preventDefault();
      }
      return;
    }
    const firstDataRowFrom = (d: number) => {
      while (d < this.rows.count && this.rows.kind[d] !== ROW_KIND.data) d++;
      return d < this.rows.count ? d : -1;
    };
    const move = (dr: number, dc: number) => {
      if (!this.cols.cols.length) return;
      let d = this.selected ? this.rows.displayOf[this.selected.row] : -1;
      const col = this.selected?.col ?? this.cols.cols[0];
      if (d < 0) {
        // No (visible) selection yet: start at the first row in view.
        d = firstDataRowFrom(Math.max(0, rowAt(this.rows, this.scroller.scrollTop)));
        if (d < 0) return;
        dr = 0;
      }
      const sel = { row: this.rows.ref[d], col };
      if (dr) {
        do d += dr;
        while (d >= 0 && d < this.rows.count && this.rows.kind[d] !== ROW_KIND.data);
        if (d < 0 || d >= this.rows.count) return;
      }
      let vi = this.cols.cols.indexOf(sel.col) + dc;
      vi = Math.max(0, Math.min(this.cols.cols.length - 1, vi));
      this.selected = { row: this.rows.ref[d], col: this.cols.cols[vi] };
      this.options.onSelectionChange?.(this.cellInfo(this.selected.row, this.selected.col));
      this.scrollToDisplay(d, 'nearest');
      this.scrollToColumn(vi);
      this.schedule();
    };
    switch (e.key) {
      case 'n':
      case 'j':
        this.nextChange();
        break;
      case 'p':
      case 'k':
        this.prevChange();
        break;
      case 'ArrowDown':
        move(1, 0);
        break;
      case 'ArrowUp':
        move(-1, 0);
        break;
      case 'ArrowLeft':
        move(0, -1);
        break;
      case 'ArrowRight':
        move(0, 1);
        break;
      // Tab is deliberately not handled: it must move focus out of the grid.
      case 'Escape':
        this.select(null);
        break;
      default:
        return;
    }
    e.preventDefault();
  }

  private jumpChange(dir: 1 | -1): boolean {
    if (!this.diff) return false;
    let from: number;
    if (this.selected && this.rows.displayOf[this.selected.row] >= 0) {
      from = this.rows.displayOf[this.selected.row];
    } else {
      const top = this.scroller.scrollTop;
      from = rowAt(this.rows, top + (dir === 1 ? 1 : 0)) - (dir === 1 && top === 0 ? 1 : 0);
    }
    const d = findChange(this.rows, this.diff, from, dir);
    if (d < 0) return false;
    const r = this.rows.ref[d];
    const mask = this.diff.changedCells.get(r);
    let col = this.selected?.col ?? this.cols.cols[0];
    if (mask) {
      const vi = Array.from(this.cols.cols).findIndex((c) => mask[c]);
      if (vi >= 0) {
        col = this.cols.cols[vi];
        this.scrollToColumn(vi);
      }
    }
    this.selected = { row: r, col };
    this.options.onSelectionChange?.(this.cellInfo(r, col));
    this.scrollToDisplay(d, 'center');
    this.schedule();
    return true;
  }

  private scrollToDisplay(d: number, align: 'start' | 'center' | 'nearest'): void {
    const top = this.rows.offsets[d];
    const h = this.rows.offsets[d + 1] - top;
    const view = this.height - HEADER_HEIGHT;
    const cur = this.scroller.scrollTop;
    let next = cur;
    if (align === 'start') next = top;
    else if (align === 'center') next = top - (view - h) / 2;
    else if (top < cur) next = top;
    else if (top + h > cur + view) next = top + h - view;
    this.scroller.scrollTop = Math.max(0, next);
  }

  private scrollToColumn(vi: number): void {
    if (vi < this.cols.frozen) return;
    const p = this.panes()[0];
    const view = p.width - p.gutter - this.cols.frozenWidth;
    const left = this.cols.offsets[vi] - this.cols.frozenWidth;
    const right = left + this.cols.widths[vi];
    const cur = this.scroller.scrollLeft;
    if (left < cur) this.scroller.scrollLeft = left;
    else if (right > cur + view) this.scroller.scrollLeft = right - view;
  }

  private cellAt(hit: Hit): CellInfo | null {
    if (hit.area !== 'cell' || hit.display < 0 || hit.vcol < 0) return null;
    return this.cellInfo(this.rows.ref[hit.display], this.cols.cols[hit.vcol]);
  }

  /** Everything about one cell of the diff. */
  cellInfo(row: number, columnIndex: number): CellInfo {
    if (!this.isCell(row, columnIndex)) throw new RangeError(`No cell at row ${row}, column ${columnIndex}`);
    const diff = this.diff!;
    const column = diff.columns[columnIndex];
    const oi = diff.oldRows[row];
    const ni = diff.newRows[row];
    const status = ROW_STATUS_NAMES[diff.status[row]];
    return {
      row,
      column,
      columnIndex,
      status,
      changed: diff.changedCells.get(row)?.[columnIndex] === 1,
      oldRow: oi,
      newRow: ni,
      oldValue: oi >= 0 && column.oldIndex >= 0 ? diff.oldTable.rows[oi][column.oldIndex] ?? null : undefined,
      newValue: ni >= 0 && column.newIndex >= 0 ? diff.newTable.rows[ni][column.newIndex] ?? null : undefined,
      oldFormula: oi >= 0 && column.oldIndex >= 0 ? diff.oldTable.formulas?.[oi]?.[column.oldIndex] : undefined,
      newFormula: ni >= 0 && column.newIndex >= 0 ? diff.newTable.formulas?.[ni]?.[column.newIndex] : undefined,
      kind: diff.changedCells.get(row)?.[columnIndex] === 1 ? this.kindOf(row, columnIndex) : undefined,
      trace: this.options.trace?.trace.get({ sheet: this.options.trace.sheet, row, col: columnIndex }),
    };
  }

  private showTooltip(info: CellInfo, hit: Hit): void {
    const before = formatCell(info.oldValue ?? null);
    const after = formatCell(info.newValue ?? null);
    // Character highlights help in text; inside numbers they're noise.
    const numeric = typeof info.oldValue === 'number' && typeof info.newValue === 'number';
    const spans = !numeric && before.length + after.length < 2000 ? diffText(before, after) : null;
    const esc = (s: string) => s.replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'));
    const side = (kind: 'removed' | 'added') =>
      spans
        ? spans
            .filter((s) => s.kind === 'equal' || s.kind === kind)
            .map((s) => (s.kind === 'equal' ? esc(s.text) : kind === 'removed' ? `<del>${esc(s.text)}</del>` : `<ins>${esc(s.text)}</ins>`))
            .join('')
        : esc(kind === 'removed' ? before : after);
    const rowLabel = info.newRow >= 0 ? `row ${info.newRow + this.rowOffset('new')}` : `row ${info.oldRow + this.rowOffset('old')}`;
    let html = `<div class="sd-tt-head"><span class="sd-tt-col">${esc(info.column.name)}</span><span>${rowLabel}</span></div>`;
    html += `<div class="sd-tt-row sd-tt-old"><span class="sd-tt-sign">−</span><span class="sd-tt-val">${side('removed') || '<i>empty</i>'}</span></div>`;
    html += `<div class="sd-tt-row sd-tt-new"><span class="sd-tt-sign">+</span><span class="sd-tt-val">${side('added') || '<i>empty</i>'}</span></div>`;
    if (info.oldFormula || info.newFormula) {
      html +=
        info.kind === 'computed' || info.oldFormula === info.newFormula
          ? `<div class="sd-tt-formula">ƒ ${esc(info.newFormula ?? info.oldFormula ?? '')}</div>`
          : `<div class="sd-tt-formula">${esc(info.oldFormula ?? '—')} → ${esc(info.newFormula ?? '—')}</div>`;
    }
    const delta = numericDelta(info.oldValue ?? null, info.newValue ?? null);
    if (delta && delta.delta !== 0) {
      const sign = delta.delta > 0 ? '+' : '-';
      const abs = formatCell(Math.abs(delta.delta));
      const pct = delta.ratio === null ? '' : ` (${sign}${Math.abs(delta.ratio * 100).toFixed(Math.abs(delta.ratio) < 0.1 ? 2 : 1)}%)`;
      html += `<div class="sd-tt-delta">Δ <b data-sign="${sign}">${sign}${abs}${pct}</b></div>`;
    }
    const tc = this.options.trace;
    const kind = info.kind;
    if (kind && this.options.highlightComputed !== false) {
      const kindLabel = kind === 'computed' ? 'downstream' : kind === 'formula' ? 'formula edit' : 'input edit';
      let trace = `<div class="sd-tt-trace"><div class="sd-tt-trace-title"><span class="sd-tt-kind" data-kind="${kind}">${kindLabel}</span>`;
      const t = info.trace;
      if (tc && t) {
        const list = kind === 'computed' ? t.causes : t.impacts;
        const count = kind === 'computed' ? t.causeCount : t.impactCount;
        if (kind === 'computed') {
          trace += count
            ? ` changed because of ${count} edit${count === 1 ? '' : 's'}:`
            : t.unexplained === 'volatile' ? ' volatile function (NOW, RAND, …)' : t.unexplained === 'dynamic' ? ' dynamic reference (INDIRECT, OFFSET, …)' : ' cause not found';
        } else {
          trace += count ? ` flows into ${count} cell${count === 1 ? '' : 's'}:` : ' nothing downstream changed';
        }
        trace += '</div>';
        for (const ref of list.slice(0, 6)) {
          const d = this.describeRef(ref);
          trace += `<div class="sd-tt-cause"><span>${kind === 'computed' ? '←' : '→'}</span><span class="sd-tt-ref">${esc(d.label)}</span><span class="sd-tt-d" data-sign="${d.sign}">${esc(d.delta)}</span></div>`;
        }
        if (count > 6) trace += `<div class="sd-tt-trace-title">+ ${count - 6} more</div>`;
      } else {
        trace += '</div>';
      }
      html += trace + '</div>';
    }
    this.tooltip.innerHTML = html;
    // Position under the cell, flipping above near the bottom edge.
    const panes = this.panes();
    const p = panes[hit.pane];
    const cols = this.cols;
    const vx = hit.vcol < cols.frozen ? p.x + p.gutter + cols.offsets[hit.vcol] : p.x + p.gutter + cols.offsets[hit.vcol] - this.scroller.scrollLeft;
    const y = HEADER_HEIGHT + this.rows.offsets[hit.display + 1] - this.scroller.scrollTop;
    const tw = this.tooltip.offsetWidth;
    const th = this.tooltip.offsetHeight;
    const left = Math.max(8, Math.min(vx, this.width - this.minimapWidth - tw - 8));
    const rowH = this.rows.offsets[hit.display + 1] - this.rows.offsets[hit.display];
    const top = y + th + 8 > this.height ? y - rowH - th - 6 : y + 6;
    this.tooltip.style.left = `${left}px`;
    this.tooltip.style.top = `${top}px`;
    this.tooltip.setAttribute('data-open', '');
  }

  private hideTooltip(): void {
    clearTimeout(this.tooltipTimer);
    this.tooltip.removeAttribute('data-open');
  }
}

// ---- helpers ---------------------------------------------------------------

const NUMERIC = /^\s*[-+]?(\d[\d,]*\.?\d*|\.\d+)([eE][-+]?\d+)?%?\s*$/;

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  e.className = cls;
  return e;
}

/** `+5.23 · +8.0%` (or `+5.23` when the old value was 0). */
function formatDelta(d: { delta: number; ratio: number | null }): { label: string; up: boolean; abs: string; pct?: string } {
  const up = d.delta > 0;
  const sign = up ? '+' : '−';
  const abs = `${sign}${formatCell(Math.abs(d.delta))}`;
  if (d.ratio === null) return { label: abs, up, abs };
  const pct = Math.abs(d.ratio * 100);
  const p = `${sign}${pct >= 100 ? pct.toFixed(0) : pct >= 10 ? pct.toFixed(1) : pct.toFixed(2)}%`;
  return { label: `${abs} · ${p}`, up, abs, pct: p };
}

function shallowEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object') return false;
  const ka = Object.keys(a);
  const kb = Object.keys(b);
  return ka.length === kb.length && ka.every((k) => (a as Record<string, unknown>)[k] === (b as Record<string, unknown>)[k]);
}

function colAt(cols: ColumnLayout, x: number): number {
  const o = cols.offsets;
  let lo = 0;
  let hi = cols.cols.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (o[mid + 1] <= x) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function hline(ctx: CanvasRenderingContext2D, x0: number, x1: number, y: number, color: string): void {
  ctx.fillStyle = color;
  ctx.fillRect(x0, Math.round(y) - 1, x1 - x0, 1);
}

function vline(ctx: CanvasRenderingContext2D, x: number, y0: number, y1: number, color: string): void {
  ctx.fillStyle = color;
  ctx.fillRect(Math.round(x) - 1, y0, 1, y1 - y0);
}

function roundRect(ctx: CanvasRenderingContext2D, x: number, y: number, w: number, h: number, r: number): void {
  ctx.beginPath();
  if (ctx.roundRect) ctx.roundRect(x, y, w, h, r);
  else ctx.rect(x, y, w, h);
}

function mixAlpha(rgb: string, alpha: number): string {
  const m = /rgba?\((\d+),(\d+),(\d+)/.exec(rgb);
  return m ? `rgba(${m[1]},${m[2]},${m[3]},${alpha})` : rgb;
}

/** Longest prefix of `s` (plus an ellipsis) that fits in `max` px. */
function fitText(ctx: CanvasRenderingContext2D, s: string, max: number): string {
  if (s.length > 400) s = s.slice(0, 400);
  if (ctx.measureText(s).width <= max) return s;
  const ell = '…';
  let lo = 0;
  let hi = s.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (ctx.measureText(s.slice(0, mid) + ell).width <= max) lo = mid;
    else hi = mid - 1;
  }
  return lo === 0 ? '' : s.slice(0, lo).trimEnd() + ell;
}
