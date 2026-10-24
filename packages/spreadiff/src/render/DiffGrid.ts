import { columnLabel } from '../core/csv.js';
import { diffText, formatCell, numericDelta } from '../core/text.js';
import { displayCell, displayMagnitude, formatCodeAt } from '../core/display.js';
import { changeKind } from '../core/kind.js';
import { rowLabelColumn } from '../core/label.js';
import { sheetLabels, type RowLabel, type SheetLabels } from '../core/labels.js';
import type { CellRef, CellTrace, ChangeKind, TraceContext, WorkbookTrace } from '../core/trace.js';
export type { TraceContext } from '../core/trace.js';
import type { SheetView } from '../core/compare.js';
import type { Decision, SheetDecisions } from '../core/review.js';
import {
  ROW_STATUS,
  ROW_STATUS_NAMES,
  type AlignedColumn,
  type CellValue,
  type RowStatus,
  type Table,
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

/**
 * - `unified`: one grid; changed cells show old and new together (see {@link CellDisplay}).
 * - `split`: old on the left, new on the right, rows aligned; missing rows and columns are hatched.
 */
export type DiffLayout = 'unified' | 'split';

/** A note attached to a row (and optionally a cell) of the diff. */
export interface CellAnnotation<A = unknown> {
  /** Aligned row index (`0 … diff.rowCount - 1`). The annotation renders below this row. */
  row: number;
  /** Column name to mark with a corner triangle. Optional. */
  column?: string;
  /** Your payload, passed back to `renderAnnotation`. */
  data: A;
  /**
   * Stable identity across renders. Without it, annotations are matched by
   * row and column, so passing a fresh array of equal annotations keeps
   * their elements; an annotation is re-rendered only when its `data`
   * changes (by `===`).
   */
  id?: string | number;
}

/** Instance counter, for ids unique within a page. */
let gridIds = 0;

/** A merged range in table coordinates (0-based rows and columns of `Table.rows`). */
type Merge = { r0: number; r1: number; c0: number; c1: number };

/** Everything about one cell of the diff, as passed to event callbacks and returned by `cellInfo()`. */
export interface CellInfo {
  /** Aligned row index. */
  row: number;
  /** The aligned column. */
  column: AlignedColumn;
  /** Index into `diff.columns`. */
  columnIndex: number;
  /** Status of the row. */
  status: RowStatus;
  /** True when this cell changed (in a modified row). */
  changed: boolean;
  /** True when the row moved (rows matched by key; see `TableDiff.moved`). */
  moved: boolean;
  /** Row index in the old table, or -1. */
  oldRow: number;
  /** Row index in the new table, or -1. */
  newRow: number;
  /** Old value; `undefined` when the row or column doesn't exist in the old table. */
  oldValue: CellValue | undefined;
  /** New value; `undefined` when the row or column doesn't exist in the new table. */
  newValue: CellValue | undefined;
  /** Old formula text, `null` if none, `undefined` if the cell doesn't exist on that side. */
  oldFormula?: string | null;
  /** New formula text, `null` if none, `undefined` if the cell doesn't exist on that side. */
  newFormula?: string | null;
  /** For changed cells: `input`, `formula` or `computed` (from the trace when attached). */
  kind?: ChangeKind;
  /** The cell's trace (causes, impacts), when a `trace` is attached and the cell changed. */
  trace?: CellTrace;
  /** The cell's review decision, when `decisions` has one for it (or its row or column). */
  decision?: Decision;
}

/**
 * How a changed cell shows its change:
 * - `stacked`: old value struck through above the new one (modified rows are taller).
 * - `new`: just the new value on a tinted background.
 * - `delta`: the new value, with the absolute and % change beneath it
 *   (or `was …` for text). Modified rows are taller.
 * - `inline`: `old → new` on one line.
 */
/**
 * A rectangle of selected cells, from `DiffGrid.getRange()` and
 * `onRangeChange`. Rows and columns are as displayed: rows hidden in a
 * collapsed region and hidden columns aren't in it.
 */
export interface GridRange {
  /** The active cell: where the selection started. */
  anchor: { row: number; col: number };
  /** The opposite corner. Equal to `anchor` for one cell. */
  focus: { row: number; col: number };
  /** Aligned rows in the range, top to bottom. */
  rows: number[];
  /** Column indices (into `diff.columns`) in the range, left to right. */
  columns: number[];
  /**
   * The range in A1 notation on each side (`"B2:D5"`), when its cells are
   * one block there. Undefined for a side the range has no cells on, or
   * where its cells aren't one block (moved rows can scatter them).
   */
  a1: { old?: string; new?: string };
}

export type CellDisplay = 'stacked' | 'new' | 'delta' | 'inline';


/** Options for {@link DiffGrid}. All are optional and can be changed later with `setOptions`. */
export interface DiffGridOptions<A = unknown> {
  /** `unified` or `split`. Default `unified`. */
  layout?: DiffLayout;
  /**
   * Mirror the visible cells as hidden ARIA grid rows and cells, so screen
   * readers can read them and arrow keys announce the selected cell.
   * Default true; false skips the DOM work (the canvas is unaffected).
   */
  accessibility?: boolean;
  /** The grid's accessible name. Default: a summary such as "Spreadsheet diff: 3 rows modified, 1 added, …". */
  ariaLabel?: string;
  /** Collapse runs of unchanged rows into separators that expand on click. Default true. */
  collapseUnchanged?: boolean;
  /** Unchanged rows kept visible around each change when collapsing. Default 2. */
  context?: number;
  /** Hide columns with no changed cells. Key, added, removed and renamed columns stay, and so does the row-label column (the first mostly-text one) when rows were matched by order. Default false. */
  hideUnchangedColumns?: boolean;
  /**
   * Columns pinned to the left while scrolling horizontally. `'keys'` (default)
   * moves the diff's key columns to the front and pins them; a number pins
   * that many leading columns.
   */
  frozenColumns?: number | 'keys';
  /**
   * Added to a 0-based row index to get the displayed row number. Defaults
   * to the table's `origin`, so numbers match the spreadsheet (2 for the
   * first data row under a header).
   */
  rowNumberOffset?: number;
  /**
   * What the row headers (the gutter) show:
   * - `'labels'` (default): what the sheet calls each row, inferred from
   *   the layout: its label, under its section headings
   *   (`India › Revenue`). Section headings are bold; long labels drop
   *   their outer sections first. A row with no label shows its number.
   * - `'numbers'`: spreadsheet row numbers (old and new, in `unified`).
   * - `'both'`: the label, then the row number.
   * - `'none'`: only the change bar and sign.
   *
   * See [How cells are named](https://github.com/samueldbarton/spreadiff/blob/main/docs/spreadiff-lite.md#how-cells-are-named).
   */
  rowHeaders?: 'labels' | 'numbers' | 'both' | 'none';
  /**
   * What the column headers show:
   * - `'labels'` (default): each column's heading, the header-row name or
   *   the heading rows above the data (`Forecast › Q2`, `2025`), as they
   *   stand at the top row in view. A column with neither shows its letter.
   * - `'letters'`: spreadsheet letters (`A`, `B`, …).
   * - `'both'`: the letter, then the heading.
   */
  columnHeaders?: 'labels' | 'letters' | 'both';
  /** Width of labelled row headers, in CSS pixels. Default: fits the labels, between 96 and 260. */
  rowHeaderWidth?: number;
  /**
   * When the headers show labels, fold away what they already show, unless
   * it changed: columns that only hold row labels (a model's label column,
   * a records table's id and name) and rows that only hold column headings
   * (a model's row of years). Default true.
   */
  foldLabels?: boolean;
  /** `system` (default) follows the page's `prefers-color-scheme`; `light` / `dark` force one. */
  theme?: 'system' | 'light' | 'dark';
  /** Change overview strip on the right edge; click it to jump. Default true. */
  minimap?: boolean;
  /** Popover with old → new, character diff, Δ and trace when hovering a changed cell. Default true. */
  tooltips?: boolean;
  /** Fixed widths in CSS pixels, by column name. Other columns size to their content. */
  columnWidths?: Record<string, number>;
  /** Notes to show below rows. Keep the array referentially stable; a new array re-renders all of them. */
  annotations?: CellAnnotation<A>[];
  /**
   * Render an annotation's content. The element is placed in the light DOM
   * (projected into the grid through a `<slot>`), so your page CSS styles
   * it, and spans the grid's width below its row. Its height is measured and
   * tracked as it resizes.
   */
  renderAnnotation?: (annotation: CellAnnotation<A>) => HTMLElement | null | undefined;
  /** Called on pointer down on a cell (after selecting it). */
  onCellClick?: (cell: CellInfo, event: MouseEvent) => void;
  /** Called when the pointer enters a cell, with `null` when it leaves the cells. */
  onCellHover?: (cell: CellInfo | null) => void;
  /** Called whenever the selected cell changes: click, keyboard, `nextChange()`, `select()`. */
  onSelectionChange?: (cell: CellInfo | null) => void;
  /**
   * Called whenever the selected range changes: drag across cells,
   * shift-click, shift + arrow keys, `selectRange()`. A single selected cell
   * is a one-cell range; `null` when nothing is selected.
   */
  onRangeChange?: (range: GridRange | null) => void;
  /** How changed cells show old vs. new. Default `stacked`. */
  cellDisplay?: CellDisplay;
  /**
   * Show values as the sheet formats them (`12.3%`, `$1,234`, `1-Jan-24`),
   * from the file's number formats. Default true; false shows the stored
   * values. Changes follow the cell's format too (`+1.2 pp` for a
   * percentage). Copying a cell always copies the exact value.
   */
  numberFormats?: boolean;
  /**
   * Color downstream changes (formula results that changed because
   * something upstream did) amber, and mark formula edits with `ƒ`. Works
   * without a trace. Default true.
   */
  highlightComputed?: boolean;
  /** Attach a workbook trace: tooltips list causes and impacts, and selection can focus them. */
  trace?: TraceContext;
  /**
   * Review decisions to show (see `decideRange`): a rejected change is
   * dimmed and marked ↺, an accepted one ✓. Undecided changes look as usual.
   */
  decisions?: SheetDecisions;
  /**
   * When a traced cell is selected: outline its causes (solid), computed
   * parents (dotted) and impacts (dashed) on this sheet, and dim the rest.
   * Default true.
   */
  traceFocus?: boolean;
}

/** Arguments to `DiffGrid.render`. */
export type DiffGridRenderProps =
  | {
      /** The diff to show. */
      diff: TableDiff;
      /** Element to render into. The grid fills it, so give it a size (e.g. a height). */
      container: HTMLElement;
    }
  | {
      /** A sheet from `compareWorkbooks` / `diffFiles`: shows its diff, with its trace attached. */
      sheet: SheetView;
      /** Element to render into. The grid fills it, so give it a size (e.g. a height). */
      container: HTMLElement;
    };

const HEADER_HEIGHT = 36;
const SEPARATOR_HEIGHT = 30;
const MINIMAP_WIDTH = 12;
const SPLIT_DIVIDER = 1;
const CELL_PAD = 10;
const MIN_COL = 56;
const MAX_COL = 360;
const BAR_WIDTH = 3;
/** Room for the +, −, • or ↕ at the right of a labelled gutter. */
const SIGN_SLOT = 18;
/** Between a row's sections and its label. */
const LABEL_SEP = ' › ';
/** Tables with more cells than this (both sides) infer labels after the first frame. */
const LABELS_SYNC_CELLS = 600_000;

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
 * A canvas-rendered, virtualized spreadsheet diff. Only the visible cells
 * are drawn, so size doesn't matter; the grid fills its container.
 *
 * ```ts
 * import { DiffGrid, diffTables, parseCsv } from 'spreadiff';
 * declare const before: string, after: string, container: HTMLElement;
 *
 * const grid = new DiffGrid({ layout: 'split' });
 * grid.render({ diff: diffTables(parseCsv(before), parseCsv(after)), container });
 * grid.nextChange(); // select and scroll to the first change
 * // later:
 * grid.cleanUp();
 * ```
 *
 * Keyboard (when the grid has focus): `n`/`j` next change, `p`/`k`
 * previous, arrow keys move the selection, ⌘/Ctrl+C copies the selected
 * value, Esc clears the selection. Tab moves focus out of the grid.
 */
export class DiffGrid<A = unknown> {
  /**
   * The element the grid renders into its container (it holds the shadow root). Style its size from outside if needed.
   */
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
  /** Merged ranges by table cell (row * width + col), per side; built on first use for each diff. */
  private merges: { diff: TableDiff; old: Map<number, Merge>; new: Map<number, Merge> } | null = null;
  /** The accessibility mirror (see `syncA11y`). */
  private a11y: HTMLElement | null = null;
  private a11yKey = '';
  private a11yDiff: TableDiff | null = null;
  private a11yOptions: DiffGridOptions<A> | null = null;
  private a11ySel = '';
  private a11yScroll = '';
  private a11yTimer = 0;
  private readonly a11yId = `sd${(gridIds = (gridIds + 1) % 1e9)}`;
  private expanded = new Set<string>();
  private annotationEls = new Map<string, { a: CellAnnotation<A>; el: HTMLElement; height: number; ro: ResizeObserver }>();
  private annotationRender: DiffGridOptions<A>['renderAnnotation'] = undefined;
  private cellMarks = new Set<string>();
  private width = 0;
  private height = 0;
  private dpr = 1;
  private frame = 0;
  private hover: Hit | null = null;
  private selected: { row: number; col: number } | null = null;
  /** The far corner of a range selection; null for a single cell. */
  private extent: { row: number; col: number } | null = null;
  private rangeKey = '';
  private tooltipTimer = 0;
  private resizing: { vcol: number; startX: number; startW: number } | null = null;
  private lastPointer: { x: number; y: number } | null = null;
  /** Cells related to the selection through the trace: `row:col` → role. */
  private focus: Map<string, 'cause' | 'impact' | 'parent'> | null = null;
  private focusKey = '';
  private themeFrame = 0;
  private textCache = new Map<string, string>();
  /** Inferred labels for the current diff (see `ensureLabels`), and the columns the row headers show. */
  private labels: { diff: TableDiff; old: SheetLabels; new: SheetLabels; absorbed: Set<number>; width: number } | null = null;
  private labelsPending: TableDiff | null = null;
  /** Rows whose label didn't fit this frame (`pane:row`), for the full label on hover. */
  private truncatedLabels = new Set<string>();
  /** Table rows at the top of the view, per side: column headings are read there. */
  private headerTop = { old: 0, new: 0 };
  private minimapCache: { key: string; colors: (string | null)[]; alphas: Float32Array } | null = null;
  private cleanups: (() => void)[] = [];
  /** `options.trace` with its sheet resolved against the diff shown (see `resolveTrace`). */
  private tc: (TraceContext & { sheet: number }) | null = null;

  constructor(options: DiffGridOptions<A> = {}) {
    this.options = options;
    this.host = document.createElement('div');
    this.host.className = 'spreadiff';
    this.host.style.cssText = 'display:block;position:relative;width:100%;height:100%;';
    this.shadow = this.host.attachShadow({ mode: 'open' });
    // An adopted stylesheet isn't inline CSS, so a Content-Security-Policy
    // without 'unsafe-inline' allows it. Older browsers get a <style>.
    const adopt = 'adoptedStyleSheets' in ShadowRoot.prototype && 'replaceSync' in CSSStyleSheet.prototype;
    let style: HTMLStyleElement | null = null;
    if (adopt) {
      const sheet = new CSSStyleSheet();
      sheet.replaceSync(STYLE);
      this.shadow.adoptedStyleSheets = [sheet];
    } else {
      style = document.createElement('style');
      style.textContent = STYLE;
    }
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
    this.tooltip.setAttribute('aria-hidden', 'true');
    this.tooltip.id = `${this.a11yId}-tip`;
    if (style) this.shadow.append(style);
    this.shadow.append(this.scroller, this.canvas, this.minimap, this.overlay, this.tooltip);
    this.ctx = this.canvas.getContext('2d', { alpha: false })!;
    this.mctx = this.minimap.getContext('2d')!;
    this.applyThemeAttr();
  }

  // ---- Public API ------------------------------------------------------

  /** Mount into `container` (first call, or a new container) and show `diff`. */
  render(props: DiffGridRenderProps): void {
    const { container } = props;
    let diff: TableDiff;
    if ('sheet' in props && props.sheet) {
      diff = props.sheet.diff;
      if (props.sheet.trace !== this.options.trace) {
        this.options = { ...this.options, trace: props.sheet.trace };
        if (this.diff === diff) this.resolveTrace();
      }
    } else diff = (props as { diff: TableDiff }).diff;
    if (this.container !== container) {
      this.detach();
      this.container = container;
      container.appendChild(this.host);
      this.attach();
    }
    this.setDiff(diff);
  }

  /**
   * Show a different diff. Resets scroll, selection, expanded regions and resized columns. A no-op for the same diff.
   */
  setDiff(diff: TableDiff): void {
    if (this.diff === diff) return;
    this.diff = diff;
    this.resolveTrace();
    this.expanded.clear();
    this.selected = null;
    this.extent = null;
    this.rangeKey = '';
    this.hover = null;
    this.widthOverrides.clear();
    this.labels = null;
    this.labelsPending = null;
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
    if (has('trace')) this.resolveTrace();
    if (has('theme')) {
      this.applyThemeAttr();
      this.refreshTheme();
    }
    if (!this.diff) return;
    if (has('numberFormats')) this.textCache.clear();
    if (has('columnWidths', 'cellDisplay', 'numberFormats')) this.measureColumns();
    if (has('annotations', 'renderAnnotation')) this.syncAnnotations();
    if (has('trace', 'traceFocus')) this.updateFocus();
    if (has('layout', 'collapseUnchanged', 'context', 'hideUnchangedColumns', 'frozenColumns', 'columnWidths', 'annotations', 'renderAnnotation', 'rowNumberOffset', 'minimap', 'cellDisplay', 'rowHeaders', 'columnHeaders', 'rowHeaderWidth', 'foldLabels', 'numberFormats')) {
      this.relayout();
    } else {
      this.schedule();
    }
  }

  /** A table cell's text as its sheet shows it, or plain with `numberFormats: false`. */
  private shown(table: Table, row: number, col: number): string {
    if (row < 0 || col < 0) return '';
    if (this.options.numberFormats === false || !table.numberFormats) return formatCell(table.rows[row]?.[col] ?? null);
    return displayCell(table, row, col);
  }

  /** A decision set for this cell, its row or its column; undefined when none was. */
  private decisionAt(r: number, ci: number): Decision | undefined {
    const dec = this.options.decisions;
    const d = this.diff;
    if (!dec || !d) return undefined;
    const st = d.status[r];
    if (st === ROW_STATUS.added || st === ROW_STATUS.removed) return dec.rows.get(r);
    if (d.columns[ci].status !== 'equal' && dec.columns.has(ci)) return dec.columns.get(ci);
    return st === ROW_STATUS.modified && d.changedCells.get(r)?.[ci] ? dec.cells.get(r * d.columns.length + ci) : undefined;
  }

  /** The number format code of an aligned cell (its new side, else old), for formatting its change. */
  private codeOf(r: number, ci: number): string | undefined {
    const d = this.diff!;
    if (this.options.numberFormats === false) return undefined;
    const c = d.columns[ci];
    const ni = d.newRows[r];
    if (ni >= 0 && c.newIndex >= 0) return formatCodeAt(d.newTable, ni, c.newIndex);
    const oi = d.oldRows[r];
    return oi >= 0 && c.oldIndex >= 0 ? formatCodeAt(d.oldTable, oi, c.oldIndex) : undefined;
  }

  /**
   * Which sheet of `trace.workbook` this grid shows: the one whose `diff` it
   * is. That beats `trace.sheet`, which can be stale or wrong (React applies
   * a new diff before new options); the index is only used for a diff that
   * isn't one of the workbook's sheets. With neither, nothing is traced.
   */
  private resolveTrace(): void {
    const tc = this.options.trace;
    if (!tc || !this.diff) {
      this.tc = null;
      return;
    }
    const found = tc.workbook.sheets.findIndex((s) => s.diff === this.diff);
    const sheet = found >= 0 ? found : tc.sheet;
    if (sheet === undefined) {
      // Usually transient (a new diff set before its new trace); a real mismatch would show nothing traced.
      this.tc = null;
      return;
    }
    this.tc = { ...tc, sheet };
  }

  /** The current options. */
  getOptions(): Readonly<DiffGridOptions<A>> {
    return this.options;
  }

  /**
   * Select the first changed cell of the next block of changed rows and
   * scroll to it. Returns false when there is no further change.
   */
  nextChange(): boolean {
    return this.jumpChange(1);
  }

  /** Like `nextChange`, backwards. Returns false at the first change. */
  prevChange(): boolean {
    return this.jumpChange(-1);
  }

  /** Scroll an aligned row into view, expanding its collapsed region if needed. Out-of-range rows are ignored. */
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

  /** Collapse every expanded region again. */
  collapseAll(): void {
    this.expanded.clear();
    if (this.diff) this.relayout();
  }

  /** Select a cell by aligned row and column index. Returns false (and changes nothing) when out of range. */
  select(row: number | null, columnIndex = 0): boolean {
    if (row !== null && !this.isCell(row, columnIndex)) return false;
    this.selected = row === null ? null : { row, col: columnIndex };
    this.extent = null;
    this.options.onSelectionChange?.(row === null ? null : this.cellInfo(row, columnIndex));
    this.emitRange();
    if (row !== null) this.scrollToRow(row, 'nearest');
    this.schedule();
    return true;
  }

  /**
   * Select a range: from the active cell (`row`, `col`) to the opposite
   * corner. Returns false (and changes nothing) when either is out of range.
   */
  selectRange(row: number, col: number, focusRow: number, focusCol: number): boolean {
    if (!this.isCell(row, col) || !this.isCell(focusRow, focusCol)) return false;
    const was = this.selected;
    this.selected = { row, col };
    this.extent = focusRow === row && focusCol === col ? null : { row: focusRow, col: focusCol };
    if (was?.row !== row || was?.col !== col) this.options.onSelectionChange?.(this.cellInfo(row, col));
    this.emitRange();
    this.scrollToRow(focusRow, 'nearest');
    this.schedule();
    return true;
  }

  /** The selected range (a single cell is a one-cell range), or null. */
  getRange(): GridRange | null {
    const sel = this.selected;
    const d = this.diff;
    if (!sel || !d) return null;
    const focus = this.extent ?? sel;
    const { da, df, va, vf } = this.rangeBounds(sel, focus);
    const rows: number[] = [];
    let columns: number[];
    if (da < 0 || df < 0 || va < 0 || vf < 0) {
      // Nothing of it shown (all collapsed or hidden): just the active cell.
      rows.push(sel.row);
      columns = [sel.col];
    } else {
      for (let i = Math.min(da, df); i <= Math.max(da, df); i++) if (this.rows.kind[i] === ROW_KIND.data) rows.push(this.rows.ref[i]);
      columns = Array.from(this.cols.cols.slice(Math.min(va, vf), Math.max(va, vf) + 1));
    }
    const a1 = (side: 'old' | 'new'): string | undefined => {
      const t = side === 'old' ? d.oldTable : d.newTable;
      const map = side === 'old' ? d.oldRows : d.newRows;
      const rs = rows.map((r) => map[r]).filter((r) => r >= 0);
      const cs = columns.map((c) => (side === 'old' ? d.columns[c].oldIndex : d.columns[c].newIndex)).filter((c) => c >= 0);
      if (!rs.length || !cs.length) return undefined;
      const block = (xs: number[]) => xs.every((x, i) => i === 0 || x === xs[i - 1] + 1);
      if (!block(rs) || !block(cs)) return undefined;
      const o = t.origin ?? { row: 1, col: 0 };
      const at = (r: number, c: number) => `${columnLabel(c + o.col)}${r + o.row + 1}`;
      const first = at(rs[0], cs[0]);
      const last = at(rs[rs.length - 1], cs[cs.length - 1]);
      return first === last ? first : `${first}:${last}`;
    };
    return { anchor: { ...sel }, focus: { ...focus }, rows, columns, a1: { old: a1('old'), new: a1('new') } };
  }

  /**
   * Display rows and visible columns of a range's corners. A corner in a
   * collapsed region, or a hidden or folded column (a label column the row
   * headers show), counts from the nearest shown one toward the other corner.
   * -1 when nothing between them is shown.
   */
  private rangeBounds(sel: { row: number; col: number }, focus: { row: number; col: number }): { da: number; df: number; va: number; vf: number } {
    const shownFrom = (row: number, toward: number) => {
      const step = toward >= row ? 1 : -1;
      for (let r = row; step > 0 ? r <= toward : r >= toward; ) {
        if (this.rows.displayOf[r] >= 0) return this.rows.displayOf[r];
        // Jump over the collapsed region holding r, rather than row by row.
        const g = this.rows.regions.find((x) => r >= x.start && r < x.end);
        r = g ? (step > 0 ? g.end : g.start - 1) : r + step;
      }
      return -1;
    };
    const da = shownFrom(sel.row, focus.row);
    const df = shownFrom(focus.row, sel.row);
    const visibleFrom = (col: number, toward: number) => {
      const step = toward >= col ? 1 : -1;
      for (let c = col; step > 0 ? c <= toward : c >= toward; c += step) {
        const v = this.cols.cols.indexOf(c);
        if (v >= 0) return v;
      }
      return -1;
    };
    const va = visibleFrom(sel.col, focus.col);
    const vf = visibleFrom(focus.col, sel.col);
    return { da, df, va, vf };
  }

  /** Tell `onRangeChange` about a new range (once per distinct range). */
  private emitRange(): void {
    const sel = this.selected;
    const f = this.extent ?? sel;
    const key = sel && f ? `${sel.row}:${sel.col}:${f.row}:${f.col}` : '';
    if (key === this.rangeKey) return;
    this.rangeKey = key;
    this.options.onRangeChange?.(this.getRange());
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

  /**
   * Remove the grid from the page and release its observers and listeners. The instance can be `render`ed again.
   */
  cleanUp(): void {
    this.detach();
    clearTimeout(this.a11yTimer);
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
    const tc = this.tc;
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
    for (const info of this.annotationEls.values()) {
      const list = annotationsAfter.get(info.a.row) ?? [];
      list.push(info.height);
      annotationsAfter.set(info.a.row, list);
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
      filter: this.foldedRows(),
    });
    this.minimapCache = null;
    this.relayoutColumns();
  }

  /** Recompute column geometry only (resizing a column must not walk every row). */
  private relayoutColumns(): void {
    const diff = this.diff!;
    const keys = new Set(diff.keyColumns);
    const label = this.options.hideUnchangedColumns ? rowLabelColumn(diff) : -1; // without keys, keep what names the rows
    // Columns the row headers already show (labels mode), unless they changed.
    const mode = this.rowHeaderMode();
    const absorbed = (mode === 'labels' || mode === 'both') && this.options.foldLabels !== false ? this.labels?.absorbed : undefined;
    const visible: number[] = [];
    diff.columns.forEach((c, i) => {
      if (absorbed?.has(i)) return;
      if (!this.options.hideUnchangedColumns || c.status !== 'equal' || c.changes > 0 || keys.has(c.name) || i === label) {
        visible.push(i);
      }
    });
    if (!visible.length && diff.columns.length) visible.push(0); // never an empty grid
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
    // Digits at the gutter font (~7px each) plus the padding `text` keeps on both sides.
    const numW = Math.max(28, Math.ceil(digits * 7 + CELL_PAD * 2 + 2));
    const mode = this.rowHeaderMode();
    // Labels (and 'none') keep a sign slot on the right of the gutter, in both layouts.
    const labelW = mode === 'labels' || mode === 'both' ? (this.options.rowHeaderWidth ?? this.labels?.width ?? 120) : 0;
    const extra = mode === 'both' ? numW : 0;
    if (this.layout === 'split') {
      const pw = Math.floor((w - SPLIT_DIVIDER) / 2);
      const gutter = mode === 'numbers' ? numW + BAR_WIDTH + 4 : labelW + extra + BAR_WIDTH + SIGN_SLOT;
      return [
        { x: 0, width: pw, gutter, side: 'old' },
        { x: pw + SPLIT_DIVIDER, width: w - pw - SPLIT_DIVIDER, gutter, side: 'new' },
      ];
    }
    const gutter = mode === 'numbers' ? numW * 2 + BAR_WIDTH + 16 : labelW + extra + BAR_WIDTH + SIGN_SLOT;
    return [{ x: 0, width: w, gutter, side: 'unified' }];
  }

  /** The row-header mode in effect: labels fall back to numbers until they're inferred. */
  private rowHeaderMode(): 'labels' | 'numbers' | 'both' | 'none' {
    const m = this.options.rowHeaders ?? 'labels';
    return (m === 'labels' || m === 'both') && !this.ensureLabels() ? 'numbers' : m;
  }

  /** Whether labels are needed at all (row or column headers in a labels mode). */
  private wantsLabels(): boolean {
    const r = this.options.rowHeaders ?? 'labels';
    const c = this.options.columnHeaders ?? 'labels';
    return r === 'labels' || r === 'both' || c !== 'letters';
  }

  /**
   * Infer row and column labels for the current diff. Small tables are done
   * now; large ones after the first frame (about 50 ms per 100,000 rows), so
   * showing a big diff never waits on it. Returns whether labels are ready.
   */
  private ensureLabels(): boolean {
    const diff = this.diff;
    if (!diff || !this.wantsLabels()) return false;
    if (this.labels?.diff === diff) return true;
    const cells = (diff.oldTable.rows.length + diff.newTable.rows.length) * Math.max(1, diff.columns.length);
    if (cells <= LABELS_SYNC_CELLS) {
      this.computeLabels(diff);
      return true;
    }
    if (this.labelsPending !== diff) {
      this.labelsPending = diff;
      setTimeout(() => {
        if (this.diff !== diff || this.labelsPending !== diff) return;
        this.labelsPending = null;
        this.computeLabels(diff);
        this.measureColumns();
        this.relayout();
      }, 0);
    }
    return false;
  }

  private computeLabels(diff: TableDiff): void {
    const opts = { keyColumns: diff.keyColumns };
    const old = sheetLabels(diff.oldTable, opts);
    const neu = sheetLabels(diff.newTable, opts);
    const absorbed = new Set<number>();
    const oldAbs = new Set(old.labelColumns());
    const newAbs = new Set(neu.labelColumns());
    diff.columns.forEach((c, i) => {
      if (c.status === 'equal' && c.changes === 0 && newAbs.has(c.newIndex) && oldAbs.has(c.oldIndex)) absorbed.add(i);
    });
    // Width: fit the sampled labels, within bounds; longer ones truncate (outer sections first).
    const t = this.theme;
    const ctx = this.ctx;
    let widest = 0;
    for (const r of this.sampleRows(400)) {
      const side = diff.newRows[r] >= 0 ? 'new' : 'old';
      const ti = side === 'new' ? diff.newRows[r] : diff.oldRows[r];
      const l = (side === 'new' ? neu : old).row(ti);
      if (!l.row.length && !l.sections.length) continue;
      ctx.font = `${l.heading ? 600 : 400} ${t.fontSize - 1}px ${t.fontFamily}`;
      const main = ctx.measureText(l.row.join(LABEL_SEP)).width;
      ctx.font = `${t.fontSize - 1}px ${t.fontFamily}`;
      const sec = l.sections.length ? ctx.measureText(l.sections.join(LABEL_SEP) + LABEL_SEP).width : 0;
      widest = Math.max(widest, main + sec);
    }
    const width = Math.round(Math.min(260, Math.max(96, widest + 18)));
    this.labels = { diff, old, new: neu, absorbed, width };
  }

  /**
   * Rows to keep when column headers show labels: unchanged rows that only
   * hold column headings (the header shows them) are folded away.
   */
  private foldedRows(): ((r: number) => boolean) | undefined {
    const diff = this.diff!;
    if (this.options.foldLabels === false || (this.options.columnHeaders ?? 'labels') === 'letters' || !this.ensureLabels()) return undefined;
    const L = this.labels!;
    return (r) => {
      if (diff.status[r] !== ROW_STATUS.equal) return true;
      const oi = diff.oldRows[r];
      const ni = diff.newRows[r];
      return !(ni >= 0 && L.new.headingRow(ni) && oi >= 0 && L.old.headingRow(oi));
    };
  }

  /** The labels of an aligned row, on a side (`unified`: the new row, else the old). */
  private rowLabelOf(r: number, side: 'old' | 'new' | 'unified'): RowLabel | null {
    const diff = this.diff!;
    const L = this.labels;
    if (!L || L.diff !== diff) return null;
    const s = side === 'unified' ? (diff.newRows[r] >= 0 ? 'new' : 'old') : side;
    const ti = s === 'new' ? diff.newRows[r] : diff.oldRows[r];
    if (ti < 0) return null;
    const l = L[s].row(ti);
    return l.row.length || l.sections.length ? l : null;
  }

  /** The heading of an aligned column at a table row, on a side: inferred parts, outermost first; empty when none. */
  private columnHeadingOf(ci: number, side: 'old' | 'new', tableRow: number): string[] {
    const diff = this.diff!;
    const L = this.labels;
    const col = diff.columns[ci];
    const s = side === 'new' && col.newIndex < 0 ? 'old' : side === 'old' && col.oldIndex < 0 ? 'new' : side;
    const idx = s === 'new' ? col.newIndex : col.oldIndex;
    if (!L || L.diff !== diff || idx < 0) return [];
    return L[s].column(tableRow, idx);
  }

  /** Column heading parts for the header, at the top row in view; the column's name when it has a real one. */
  private columnHeading(ci: number, side: 'old' | 'new'): string[] {
    const col = this.diff!.columns[ci];
    const h = this.columnHeadingOf(ci, side, side === 'old' ? this.headerTop.old : this.headerTop.new);
    if (h.length) return h;
    return /^[A-Z]{1,3}$/.test(col.name) ? [] : [col.name];
  }

  /** `C7`: a cell's address in the new table (the old one for removed cells). */
  private addressOf(info: CellInfo): string {
    const diff = this.diff!;
    const side = info.newRow >= 0 && info.column.newIndex >= 0 ? 'new' : 'old';
    const row = side === 'new' ? info.newRow : info.oldRow;
    const t = side === 'new' ? diff.newTable : diff.oldTable;
    return `${this.columnLetter(info.column, side)}${row + (t.origin?.row ?? 1) + 1}`;
  }

  /** A row label as one line: `India › Revenue`. */
  private rowLabelText(l: RowLabel): string {
    return [...l.sections, ...l.row].join(LABEL_SEP);
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
    const headings = (this.options.columnHeaders ?? 'labels') !== 'letters' && this.ensureLabels();
    this.widths = diff.columns.map((c, i) => {
      const fixed = this.options.columnWidths?.[c.name];
      ctx.font = `600 ${t.fontSize - 1}px ${t.fontFamily}`;
      // The widest of the name and the inferred heading (at the top: "Forecast › Q2").
      const heading = headings ? this.columnHeadingOf(i, c.newIndex >= 0 ? 'new' : 'old', 0).join(LABEL_SEP) : '';
      const title = heading.length > c.name.length ? heading : c.name;
      let header = ctx.measureText(c.status === 'renamed' ? `${c.oldName} → ${c.name}` : c.status === 'equal' ? title : `+ ${title}`).width;
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
          const s = this.shown(table, ri, idx);
          ws.push(s.length > 80 ? 80 * 7 : ctx.measureText(s).width);
        }
        // Changed cells in inline / delta displays need room for more than the value.
        if (diff.changedCells.get(r)?.[i] && c.oldIndex >= 0 && c.newIndex >= 0 && (this.display === 'inline' || this.display === 'delta')) {
          const ov = diff.oldTable.rows[diff.oldRows[r]][c.oldIndex] ?? null;
          const nv = diff.newTable.rows[diff.newRows[r]][c.newIndex] ?? null;
          const ot = this.shown(diff.oldTable, diff.oldRows[r], c.oldIndex);
          const nt = this.shown(diff.newTable, diff.newRows[r], c.newIndex);
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
                ws.push(ctx.measureText(formatDelta(dl, this.codeOf(r, i)).label).width);
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
    const renderChanged = render !== this.annotationRender;
    this.annotationRender = render;
    // Match by id, else by row + column (+ position among equal ones), not by
    // object identity: a fresh array of equal annotations keeps its elements.
    const seen = new Map<string, number>();
    const keyed = list.map((a) => {
      const base = a.id !== undefined ? `id:${a.id}` : `${a.row}\u0000${a.column ?? ''}`;
      const n = seen.get(base) ?? 0;
      seen.set(base, n + 1);
      return { a, key: `${base}\u0000${n}` };
    });
    const keep = new Set(keyed.map((k) => k.key));
    let moved = false;
    for (const [key, info] of this.annotationEls) {
      if (!keep.has(key) || !render) {
        info.ro.disconnect();
        info.el.remove();
        this.annotationEls.delete(key);
        moved = true;
      }
    }
    this.cellMarks.clear();
    for (const { a, key } of keyed) {
      if (a.column) this.cellMarks.add(`${a.row}:${a.column}`);
      if (!render) continue;
      const existing = this.annotationEls.get(key);
      if (existing) {
        if (existing.a.row !== a.row) moved = true;
        if (existing.a.data !== a.data || renderChanged) {
          const content = render(a);
          if (content) existing.el.replaceChildren(content);
        }
        existing.a = a;
        continue;
      }
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
      const info = { a, el: wrap, height: Math.max(1, wrap.offsetHeight || 40), ro };
      this.annotationEls.set(key, info);
      ro.observe(wrap);
    }
    if (moved && this.diff) this.relayout();
  }

  private positionAnnotations(): void {
    if (!this.annotationEls.size || !this.rows) return;
    const top = this.scroller.scrollTop;
    const byRow = new Map<number, HTMLElement[]>();
    for (const info of this.annotationEls.values()) {
      const list = byRow.get(info.a.row) ?? [];
      list.push(info.el);
      byRow.set(info.a.row, list);
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

  /** Draw synchronously instead of on the next animation frame. Rarely needed (tests, benchmarks). */
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
    this.syncA11y();
  }

  // ---- Accessibility ------------------------------------------------------

  /**
   * A visually hidden DOM copy of the visible window (ARIA grid rows and
   * cells), so screen readers can read what the canvas draws. The selected
   * cell is the grid's `aria-activedescendant`, so arrow keys announce it.
   * Rebuilt only when the window, the selection or the diff changes.
   */
  private syncA11y(now = false): void {
    if (this.options.accessibility === false) {
      this.a11y?.replaceChildren();
      return;
    }
    // Per frame, only compare scroll position and selection: rebuilding while
    // scrolling waits until it settles; a new selection or diff applies at once.
    const sel0 = this.selected;
    const selNow = sel0 ? `${sel0.row}:${sel0.col}` : '';
    const scrollNow = `${this.scroller.scrollTop}:${this.scroller.scrollLeft}:${this.rows.count}:${this.width}:${this.height}`;
    if (!now && this.a11y && this.a11yDiff === this.diff && selNow === this.a11ySel) {
      if (scrollNow === this.a11yScroll) return;
      this.a11yScroll = scrollNow;
      clearTimeout(this.a11yTimer);
      this.a11yTimer = window.setTimeout(() => this.syncA11y(true), 150);
      return;
    }
    this.a11yScroll = scrollNow;
    const diff = this.diff!;
    const rows = this.rows;
    const cols = this.cols;
    const scrollY = this.scroller.scrollTop;
    const p = this.panes()[this.panes().length - 1];
    const first = Math.max(0, rowAt(rows, scrollY));
    let last = first;
    while (last + 1 < rows.count && rows.offsets[last + 1] < scrollY + this.height - HEADER_HEIGHT) last++;
    const view = p.width - p.gutter - cols.frozenWidth;
    const fromCol = Math.max(cols.frozen, colAt(cols, cols.frozenWidth + this.scroller.scrollLeft));
    const toCol = colAt(cols, cols.frozenWidth + this.scroller.scrollLeft + view);
    const visible: number[] = [];
    for (let vi = 0; vi < cols.frozen; vi++) visible.push(vi);
    for (let vi = fromCol; vi <= toCol && vi < cols.cols.length; vi++) visible.push(vi);
    const sel = this.selected;
    const selKey = sel ? `${sel.row}:${sel.col}` : '';
    const key = `${first}:${last}:${visible.join(',')}:${selKey}:${rows.count}:${this.layout}`;
    if (key === this.a11yKey && this.a11yDiff === diff && this.a11yOptions === this.options) return;
    this.a11yKey = key;
    this.a11yDiff = diff;
    // New options (decisions, number formats) can change what cells say.
    this.a11yOptions = this.options;
    this.a11ySel = selKey;
    if (!this.a11y) {
      this.a11y = el('div', 'sd-a11y');
      this.a11y.setAttribute('role', 'rowgroup');
      this.scroller.appendChild(this.a11y);
    }
    const s = diff.stats;
    const n = (k: number, one: string, many = `${one}s`) => `${k.toLocaleString()} ${k === 1 ? one : many}`;
    const parts = [s.rowsModified && `${n(s.rowsModified, 'row')} modified`, s.rowsAdded && `${n(s.rowsAdded, 'row')} added`, s.rowsRemoved && `${n(s.rowsRemoved, 'row')} removed`, s.rowsMoved && `${n(s.rowsMoved, 'row')} moved`, s.cellsChanged && `${n(s.cellsChanged, 'cell')} changed`].filter(Boolean);
    this.scroller.setAttribute('aria-label', this.options.ariaLabel ?? `Spreadsheet diff: ${parts.length ? parts.join(', ') : 'no changes'}`);
    this.scroller.setAttribute('aria-rowcount', String(rows.count + 1));
    this.scroller.setAttribute('aria-colcount', String(cols.cols.length + 1));
    const id = (d: number, vi: number) => `${this.a11yId}-${d}-${vi}`;
    const cell = (role: string, text: string, colIndex: number, cellId?: string) => {
      const c = el('div', 'sd-a11y-cell');
      c.setAttribute('role', role);
      c.setAttribute('aria-colindex', String(colIndex));
      if (cellId) c.id = cellId;
      c.textContent = text;
      return c;
    };
    const out: HTMLElement[] = [];
    const header = el('div', 'sd-a11y-row');
    header.setAttribute('role', 'row');
    header.setAttribute('aria-rowindex', '1');
    header.append(cell('columnheader', 'Row', 1));
    for (const vi of visible) {
      const c = diff.columns[cols.cols[vi]];
      const heading = (this.options.columnHeaders ?? 'labels') === 'letters' ? [] : this.columnHeading(cols.cols[vi], 'new');
      const title = heading.length && c.status !== 'renamed' ? heading.join(', ') : c.name;
      const name = c.status === 'renamed' ? `${c.name}, renamed from ${c.oldName}` : c.status === 'added' ? `${title}, added column` : c.status === 'removed' ? `${title}, removed column` : title;
      header.append(cell('columnheader', name, vi + 2));
    }
    out.push(header);
    let active = '';
    for (let d = first; d <= last; d++) {
      const kind = rows.kind[d];
      if (kind === ROW_KIND.annotation) continue;
      const row = el('div', 'sd-a11y-row');
      row.setAttribute('role', 'row');
      row.setAttribute('aria-rowindex', String(d + 2));
      if (kind === ROW_KIND.separator) {
        const region = rows.regions[rows.ref[d]];
        const n = region.end - region.start;
        row.append(cell('gridcell', `${n.toLocaleString()} unchanged ${n === 1 ? 'row' : 'rows'}, collapsed`, 1));
        out.push(row);
        continue;
      }
      const r = rows.ref[d];
      row.append(cell('rowheader', this.describeRow(r), 1));
      for (const vi of visible) {
        const ci = cols.cols[vi];
        const cid = id(d, vi);
        row.append(cell('gridcell', this.describeCell(r, ci), vi + 2, cid));
        if (sel && sel.row === r && sel.col === ci) active = cid;
      }
      if (sel && sel.row === r) row.setAttribute('aria-selected', 'true');
      out.push(row);
    }
    this.a11y.replaceChildren(...out);
    if (active) this.scroller.setAttribute('aria-activedescendant', active);
    else this.scroller.removeAttribute('aria-activedescendant');
  }

  /** "Row 12, modified, moved from row 4" for the row header a screen reader announces. */
  private describeRow(r: number): string {
    const diff = this.diff!;
    const oi = diff.oldRows[r];
    const ni = diff.newRows[r];
    const n = ni >= 0 ? ni + this.rowOffset('new') : oi + this.rowOffset('old');
    const status = ROW_STATUS_NAMES[diff.status[r]];
    const label = this.rowHeaderMode() === 'numbers' ? null : this.rowLabelOf(r, 'unified');
    let text = `${label ? `${this.rowLabelText(label)}, row ${n}` : `Row ${n}`}${status === 'equal' ? '' : `, ${status}`}`;
    if (diff.moved?.[r]) text += `, moved from row ${oi + this.rowOffset('old')}`;
    return text;
  }

  /** A cell as words: its value, and for a change the old value, the delta and why it changed. */
  private describeCell(r: number, ci: number): string {
    const info = this.cellInfo(r, ci);
    const dd = this.diff!;
    const ov = this.shown(dd.oldTable, info.oldRow, info.column.oldIndex);
    const nv = this.shown(dd.newTable, info.newRow, info.column.newIndex);
    const v = (x: string) => x || 'empty';
    if (info.status === 'added') return info.newValue === undefined ? 'not in the new table' : v(nv);
    if (info.status === 'removed') return v(ov);
    if (info.column.status === 'added') return v(nv);
    if (info.column.status === 'removed') return `${v(ov)}, column removed`;
    if (!info.changed) return v(nv);
    const d = numericDelta(info.oldValue ?? null, info.newValue ?? null);
    const delta = d && d.delta !== 0 ? `, ${d.delta > 0 ? 'up' : 'down'} ${displayMagnitude(Math.abs(d.delta), this.codeOf(r, ci))}${d.ratio === null ? '' : ` (${Math.abs(d.ratio * 100).toFixed(1)}%)`}` : '';
    const why = info.kind === 'computed' ? ', recalculated result' : info.kind === 'formula' ? ', formula edited' : ', edited';
    const formula = info.oldFormula !== info.newFormula && info.newFormula ? `, formula ${info.newFormula}` : '';
    const decided = info.decision ? `, ${info.decision === 'reject' ? 'rejected' : 'accepted'}` : '';
    return `${v(nv)}, was ${v(ov)}${delta}${why}${formula}${decided}`;
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

    // A range: a tint over its cells and an outline around it.
    if (this.selected && this.extent) {
      const { da, df, va, vf } = this.rangeBounds(this.selected, this.extent);
      if (da >= 0 && df >= 0 && va >= 0 && vf >= 0) {
        const [d0, d1] = da < df ? [da, df] : [df, da];
        const [v0, v1] = va < vf ? [va, vf] : [vf, va];
        const y0 = bodyTop + rows.offsets[d0] - scrollY;
        const y1 = bodyTop + rows.offsets[d1 + 1] - scrollY;
        ctx.save();
        ctx.beginPath();
        ctx.rect(dataX, bodyTop, right - dataX, this.height - bodyTop);
        ctx.clip();
        // Frozen and scrolling columns sit apart on screen: tint each column, outline the span.
        ctx.globalAlpha = 0.12;
        ctx.fillStyle = t.selection;
        for (let vi = v0; vi <= v1; vi++) ctx.fillRect(colX(vi), y0, cols.widths[vi], y1 - y0);
        ctx.globalAlpha = 1;
        ctx.strokeStyle = t.selection;
        ctx.lineWidth = 1;
        const x0 = colX(v0);
        ctx.strokeRect(x0 + 0.5, y0 + 0.5, colX(v1) + cols.widths[v1] - x0 - 1, y1 - y0 - 1);
        ctx.restore();
      }
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
        ctx.strokeStyle = role === 'impact' ? t.computedFg : t.modifiedFg;
        ctx.setLineDash(role === 'impact' ? [4, 3] : role === 'parent' ? [1.5, 2.5] : []);
        roundRect(ctx, colX(vi) + 1.5, y + 1.5, cols.widths[vi] - 3, h - 3, 3);
        ctx.stroke();
      }
      ctx.setLineDash([]);
      ctx.restore();
    }

    // Column headings follow the top row in view (a model may repeat its year row per section).
    const top = { old: -1, new: -1 };
    for (let d = first, k = 0; d < rows.count && k < 80 && (top.old < 0 || top.new < 0); d++, k++) {
      if (rows.kind[d] !== ROW_KIND.data) continue;
      const r = rows.ref[d];
      if (top.old < 0 && diff.oldRows[r] >= 0) top.old = diff.oldRows[r];
      if (top.new < 0 && diff.newRows[r] >= 0) top.new = diff.newRows[r];
    }
    this.headerTop = { old: Math.max(0, top.old), new: Math.max(0, top.new) };
    this.drawHeader(p, firstScrolled, lastScrolled, colX);
    ctx.restore();
  }

  /** The side whose cells a pane shows for a row (unified: new, unless the row was removed). */
  private sideOf(p: PaneGeometry, r: number): 'old' | 'new' {
    return p.side === 'old' ? 'old' : p.side === 'new' ? 'new' : this.diff!.newRows[r] >= 0 ? 'new' : 'old';
  }

  /** The horizontal merge holding an unchanged cell on `side`, if any. */
  private mergeAt(side: 'old' | 'new', r: number, ci: number): Merge | undefined {
    const diff = this.diff!;
    if (!diff.oldTable.merges && !diff.newTable.merges) return undefined;
    if (diff.changedCells.get(r)?.[ci]) return undefined; // a change keeps its own box
    if (this.merges?.diff !== diff) {
      const index = (t: Table) => {
        const m = new Map<number, Merge>();
        const o = t.origin ?? { row: 1, col: 0 };
        let budget = 200_000; // cells indexed: enough for real sheets, bounded for giant merges
        for (const g of t.merges ?? []) {
          if (g.c1 === g.c0) continue; // vertical merges cross row alignment; drawn cell by cell
          const merge = { r0: g.r0 - o.row, r1: g.r1 - o.row, c0: g.c0 - o.col, c1: g.c1 - o.col };
          for (let row = Math.max(0, merge.r0); row <= Math.min(t.rows.length - 1, merge.r1) && budget > 0; row++) {
            for (let c = Math.max(0, merge.c0); c <= Math.min(t.columns.length - 1, merge.c1); c++, budget--) m.set(row * t.columns.length + c, merge);
          }
        }
        return m;
      };
      this.merges = { diff, old: index(diff.oldTable), new: index(diff.newTable) };
    }
    const t = side === 'old' ? diff.oldTable : diff.newTable;
    const row = side === 'old' ? diff.oldRows[r] : diff.newRows[r];
    const c = side === 'old' ? diff.columns[ci].oldIndex : diff.columns[ci].newIndex;
    return row < 0 || c < 0 ? undefined : this.merges[side].get(row * t.columns.length + c);
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
    // Merged with the next visible column (same merge, both unchanged): no gridline between them.
    const side = this.sideOf(p, r);
    const merge = missing ? undefined : this.mergeAt(side, r, ci);
    const joinsNext = !!merge && vi + 1 < this.cols.cols.length && this.mergeAt(side, r, this.cols.cols[vi + 1]) === merge;
    if (!joinsNext) vline(ctx, right, y, y + h, t.border);
    if (missing) return;
    if (merge && vi > 0 && this.mergeAt(side, r, this.cols.cols[vi - 1]) === merge) return; // drawn by the merge's first visible cell

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

    let oldText = this.shown(diff.oldTable, oi, col.oldIndex);
    let newText = this.shown(diff.newTable, ni, col.newIndex);
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
      // A merge's text spans its visible columns.
      let span = w;
      for (let k = vi + 1; merge && k < this.cols.cols.length && this.mergeAt(side, r, this.cols.cols[k]) === merge; k++) span += this.cols.widths[k];
      const shown = v === undefined ? '' : p.side === 'old' || (p.side === 'unified' && newV === undefined) ? oldText : newText;
      this.text(shown, x, y + h / 2 + 0.5, span, span > w ? 'left' : align, font, color, p.side === 'unified' && status === ROW_STATUS.removed);
    }
    if (this.focus && !(this.selected?.row === r && this.selected.col === ci) && !this.focus.has(`${r}:${ci}`)) {
      ctx.fillStyle = t.dim;
      ctx.fillRect(x, y, w, h);
    }
    const decision = this.options.decisions && this.decisionAt(r, ci);
    if (decision) {
      if (decision === 'reject') {
        ctx.fillStyle = t.dim;
        ctx.fillRect(x, y, w, h);
      }
      const mark = decision === 'reject' ? '↺' : '✓';
      ctx.font = `600 ${t.fontSize - 2}px ${t.fontFamily}`;
      ctx.textAlign = 'right';
      ctx.textBaseline = 'top';
      ctx.fillStyle = decision === 'reject' ? t.removedFg : t.addedFg;
      ctx.fillText(mark, x + w - 3, y + 2);
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
    const tc = this.tc;
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
    const tc = this.tc!;
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
      const code = this.options.numberFormats === false ? undefined : formatCodeAt(d.newTable, ni, c.newIndex);
      const f = formatDelta(dl, code);
      return { label, delta: f.label, sign: f.up ? '+' : '-' };
    }
    return { label, delta: `${this.shown(d.oldTable, oi, c.oldIndex) || '∅'} → ${this.shown(d.newTable, ni, c.newIndex) || '∅'}`, sign: '' };
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
    const deltaText = delta && delta.delta !== 0 ? formatDelta(delta, this.codeOf(r, ci)) : null;

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
        this.text(this.fitDelta(deltaText, w / 2, tiny), x, y + h / 2 + 0.5, w, align === 'right' ? 'left' : 'right', tiny, deltaText.up ? t.addedFg : t.removedFg);
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
      if (deltaText) this.text(this.fitDelta(deltaText, w, tiny), x, bottom, w, align, tiny, deltaText.up ? t.addedFg : t.removedFg);
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
    // Formula edits get an ƒ in the corner. Computed cells are told apart by
    // their amber tint (and the tooltip), with no extra mark.
    if (kind === 'formula' && p.side !== 'old') {
      ctx.font = `italic 600 ${t.fontSize - 3}px ${t.fontFamily}`;
      ctx.fillStyle = t.formula;
      ctx.textAlign = 'left';
      ctx.fillText('ƒ', x + 3, y + 8);
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
    const mode = this.rowHeaderMode();
    if (mode !== 'numbers') {
      this.drawRowHeader(p, r, status, missing, cy, mode, numColor, selected || hovered, d);
      vline(this.ctx, p.x + p.gutter, y, y + h, t.border);
      return;
    }
    if (p.side === 'unified') {
      const numW = (p.gutter - BAR_WIDTH - 16) / 2;
      if (oi >= 0) this.text(String(oi + offOld), p.x + BAR_WIDTH, cy, numW, 'right', mono, numColor);
      if (ni >= 0) this.text(String(ni + offNew), p.x + BAR_WIDTH + numW, cy, numW, 'right', mono, numColor);
      const sign = status === ROW_STATUS.added ? '+' : status === ROW_STATUS.removed ? '−' : '';
      if (diff.moved?.[r]) {
        // Moved (and maybe modified: the bar shows that): an up-down arrow in the sign slot.
        ctx.font = `600 ${t.fontSize}px ${t.monoFontFamily}`;
        ctx.fillStyle = t.modifiedFg;
        ctx.textAlign = 'center';
        ctx.fillText('↕', p.x + p.gutter - 9, cy);
      } else if (status === ROW_STATUS.modified) {
        ctx.fillStyle = t.modifiedFg;
        ctx.beginPath();
        ctx.arc(p.x + p.gutter - 9, cy - 0.5, 2.5, 0, Math.PI * 2);
        ctx.fill();
      } else if (sign) {
        ctx.font = `600 ${t.fontSize}px ${t.monoFontFamily}`;
        ctx.fillStyle = status === ROW_STATUS.added ? t.addedFg : t.removedFg;
        ctx.textAlign = 'center';
        ctx.fillText(sign, p.x + p.gutter - 9, cy);
      }
    } else if (!missing) {
      const idx = p.side === 'old' ? oi : ni;
      const off = p.side === 'old' ? offOld : offNew;
      if (idx >= 0) this.text(String(idx + off), p.x + BAR_WIDTH, cy, p.gutter - BAR_WIDTH, 'right', mono, numColor);
      if (diff.moved?.[r] && p.side === 'new') this.text('↕', p.x + BAR_WIDTH + 2, cy, 12, 'left', `600 ${t.fontSize - 1}px ${t.monoFontFamily}`, t.modifiedFg);
    }
    vline(this.ctx, p.x + p.gutter, y, y + h, t.border);
  }

  /** A labelled (or bare) row header: label, optional number, and the sign slot. */
  private drawRowHeader(p: PaneGeometry, r: number, status: number, missing: boolean, cy: number, mode: 'labels' | 'both' | 'none', numColor: string, active: boolean, d: number): void {
    const ctx = this.ctx;
    const t = this.theme;
    const diff = this.diff!;
    const signX = p.x + p.gutter - SIGN_SLOT / 2;
    // Sign slot: ↕ moved, • modified, + / − added or removed (the unified gutter, or each side's own rows).
    if (!missing) {
      const added = status === ROW_STATUS.added;
      const removed = status === ROW_STATUS.removed;
      if (diff.moved?.[r] && p.side !== 'old') {
        ctx.font = `600 ${t.fontSize}px ${t.monoFontFamily}`;
        ctx.fillStyle = t.modifiedFg;
        ctx.textAlign = 'center';
        ctx.fillText('↕', signX, cy);
      } else if (status === ROW_STATUS.modified && p.side === 'unified') {
        ctx.fillStyle = t.modifiedFg;
        ctx.beginPath();
        ctx.arc(signX, cy - 0.5, 2.5, 0, Math.PI * 2);
        ctx.fill();
      } else if (added || removed) {
        ctx.font = `600 ${t.fontSize}px ${t.monoFontFamily}`;
        ctx.fillStyle = added ? t.addedFg : t.removedFg;
        ctx.textAlign = 'center';
        ctx.fillText(added ? '+' : '−', signX, cy);
      }
    }
    if (missing || mode === 'none') return;
    const side = p.side === 'unified' ? (diff.newRows[r] >= 0 ? 'new' : 'old') : p.side;
    const ti = side === 'new' ? diff.newRows[r] : diff.oldRows[r];
    const number = ti >= 0 ? String(ti + this.rowOffset(side)) : '';
    const mono = `${t.fontSize - 2}px ${t.monoFontFamily}`;
    const numW = mode === 'both' ? p.gutter - BAR_WIDTH - SIGN_SLOT - (this.options.rowHeaderWidth ?? this.labels?.width ?? 120) : 0;
    const x0 = p.x + BAR_WIDTH + 8;
    const x1 = p.x + p.gutter - SIGN_SLOT - numW - 4;
    if (mode === 'both' && number) this.text(number, x1, cy, numW + CELL_PAD, 'right', mono, t.fgFaint);
    const label = this.rowLabelOf(r, p.side);
    const key = `${p.side}:${r}`;
    if (!label) {
      // No label for this row: its number, where the label would be.
      if (mode === 'labels' && number) this.text(number, x0 - CELL_PAD, cy, x1 - x0 + CELL_PAD * 2, 'left', mono, numColor);
      this.truncatedLabels.delete(key);
      return;
    }
    const main = label.row.join(LABEL_SEP) || label.sections[label.sections.length - 1] || '';
    const sections = label.row.length ? label.sections : label.sections.slice(0, -1);
    const strong = label.heading || active;
    const mainFont = `${label.heading ? 600 : 450} ${t.fontSize - 1}px ${t.fontFamily}`;
    const secFont = `${t.fontSize - 1}px ${t.fontFamily}`;
    const mainColor = status === ROW_STATUS.added ? t.addedFg : status === ROW_STATUS.removed ? t.removedFg : strong ? t.fg : t.fgMuted;
    const room = x1 - x0;
    ctx.font = mainFont;
    const mw = ctx.measureText(main).width;
    // Drop outer sections until the label fits; mark the cut with "…".
    ctx.font = secFont;
    let secs = sections;
    let prefix = secs.length ? secs.join(LABEL_SEP) + LABEL_SEP : '';
    let cut = false;
    while (secs.length && ctx.measureText(prefix).width + mw > room) {
      secs = secs.slice(1);
      cut = true;
      prefix = `…${LABEL_SEP}${secs.length ? secs.join(LABEL_SEP) + LABEL_SEP : ''}`;
    }
    let pw = prefix ? ctx.measureText(prefix).width : 0;
    if (pw + mw > room) {
      prefix = '';
      pw = 0;
    }
    if (prefix) {
      ctx.fillStyle = t.fgFaint;
      ctx.textAlign = 'left';
      ctx.fillText(prefix, x0, cy);
    }
    const fits = pw + mw <= room;
    this.text(main, x0 + pw - CELL_PAD, cy, room - pw + CELL_PAD * 2, 'left', mainFont, mainColor, status === ROW_STATUS.removed && p.side === 'unified');
    if (cut || !fits) this.truncatedLabels.add(key);
    else this.truncatedLabels.delete(key);
    void d;
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
    const mode = this.rowHeaderMode();
    if (mode !== 'numbers') {
      // The columns the labels stand for ("region"), or nothing; OLD / NEW in split.
      if (mode !== 'none') {
        const names = [...(this.labels?.absorbed ?? [])].map((ci) => this.diff!.columns[ci].name).filter((n) => !/^[A-Z]{1,3}$/.test(n));
        const title = names.length ? names.join(' · ') : '';
        this.text(title, p.x + BAR_WIDTH + 8 - CELL_PAD, ly, p.gutter - BAR_WIDTH - SIGN_SLOT - (p.side === 'unified' ? 0 : 26) + CELL_PAD, 'left', `600 ${t.fontSize - 1}px ${t.fontFamily}`, t.fgMuted);
      }
      if (p.side !== 'unified') label(p.side === 'old' ? 'OLD' : 'NEW', p.x + p.gutter - 4, p.side === 'old' ? t.removedFg : t.addedFg);
    } else if (p.side !== 'unified') {
      label(p.side === 'old' ? 'OLD' : 'NEW', p.x + p.gutter - CELL_PAD, p.side === 'old' ? t.removedFg : t.addedFg);
    } else {
      const numW = (p.gutter - BAR_WIDTH - 16) / 2;
      label('OLD', p.x + BAR_WIDTH + numW - CELL_PAD, t.removedFg);
      label('NEW', p.x + BAR_WIDTH + numW * 2 - CELL_PAD, t.addedFg);
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
    const side = p.side === 'old' ? 'old' : 'new';
    const letter = this.columnLetter(col, side);
    const mode = this.options.columnHeaders ?? 'labels';
    const heading = mode === 'letters' ? [] : this.columnHeading(ci, side);
    const bold = `600 ${t.fontSize - 1}px ${t.fontFamily}`;
    const showLetter = mode !== 'labels' || !heading.length;
    let tx = x + CELL_PAD;
    if (showLetter) {
      // The letter alone is the heading when there is no other ('letters', or no label found).
      const lone = mode === 'letters' || !heading.length;
      const font = lone ? `600 ${t.fontSize - 2}px ${t.monoFontFamily}` : `500 ${t.fontSize - 3}px ${t.monoFontFamily}`;
      ctx.font = font;
      ctx.fillStyle = lone ? t.fgMuted : t.fgFaint;
      ctx.textAlign = 'left';
      ctx.fillText(letter, tx, cy);
      tx += ctx.measureText(letter).width + 6;
    }
    let avail = x + w - tx - CELL_PAD;
    if (col.changes > 0) avail -= 28;
    const accent = col.status === 'added' ? t.addedFg : col.status === 'removed' ? t.removedFg : t.fg;
    if (!heading.length && col.status !== 'renamed') {
      // Nothing more to say ('letters', or a lettered column with no heading found).
    } else if (col.status !== 'renamed' && heading.length) {
      // Outer headings muted, the innermost bold: "Forecast › Q2".
      const inner = heading[heading.length - 1];
      const prefix = heading.length > 1 ? heading.slice(0, -1).join(LABEL_SEP) + LABEL_SEP : '';
      const sign = col.status === 'added' && p.side === 'unified' ? '+ ' : col.status === 'removed' && p.side === 'unified' ? '− ' : '';
      ctx.font = `${t.fontSize - 1}px ${t.fontFamily}`;
      const pwid = prefix ? ctx.measureText(prefix).width : 0;
      ctx.font = bold;
      const iw = ctx.measureText(sign + inner).width;
      const usePrefix = prefix && pwid + iw <= avail;
      if (usePrefix) {
        ctx.font = `${t.fontSize - 1}px ${t.fontFamily}`;
        ctx.fillStyle = t.fgFaint;
        ctx.textAlign = 'left';
        ctx.fillText(prefix, tx, cy);
      }
      const ix = usePrefix ? tx + pwid : tx;
      this.text(sign + inner, ix - CELL_PAD, cy, x + w - ix - (col.changes ? 28 : 0) + CELL_PAD, 'left', bold, accent, col.status === 'removed' && p.side === 'unified');
    } else if (col.status === 'renamed' && p.side !== 'new') {
      const oldName = col.oldName ?? '';
      if (p.side === 'old') {
        this.text(oldName, tx - CELL_PAD, cy, avail + CELL_PAD * 2, 'left', bold, t.fg);
      } else {
        ctx.font = bold;
        const ow = Math.min(ctx.measureText(oldName).width, avail * 0.45);
        this.text(oldName, tx - CELL_PAD, cy, ow + CELL_PAD * 2, 'left', bold, t.fgMuted, true);
        tx += ow + 4;
        this.text('→', tx - CELL_PAD, cy, 14 + CELL_PAD * 2, 'left', bold, t.modifiedFg);
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
      ctx.fillStyle = mixAlpha(t.modified, t.dark ? 0.16 : 0.1);
      roundRect(ctx, bx, cy - 7.5, bw, 15, 7.5);
      ctx.fill();
      ctx.fillStyle = t.modifiedFg;
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
      // Density shading: each pixel row gets the share of its rows that
      // changed (as opacity) and the most severe kind among them (as color).
      // Without it, thousands of changes in a long sheet paint the strip solid.
      const n = Math.ceil(trackH);
      const rowsW = new Float64Array(n);
      const changedW = new Float64Array(n);
      const ranks = new Uint8Array(n);
      const rank = (s: number) => (s === ROW_STATUS.removed ? 3 : s === ROW_STATUS.added ? 2 : s === ROW_STATUS.modified ? 1 : 0);
      for (let d = 0; d < this.rows.count; d++) {
        if (this.rows.kind[d] !== ROW_KIND.data) continue;
        const r = this.rows.ref[d];
        const k = rank(this.diff!.status[r]) || (this.diff!.moved?.[r] ? 1 : 0);
        const y0 = Math.floor((this.rows.offsets[d] / total) * trackH);
        const y1 = Math.max(y0 + 1, Math.ceil((this.rows.offsets[d + 1] / total) * trackH));
        const w = 1 / (y1 - y0);
        for (let y = y0; y < y1 && y < n; y++) {
          rowsW[y] += w;
          if (!k) continue;
          changedW[y] += w;
          if (k > ranks[y]) ranks[y] = k;
        }
      }
      const colors: (string | null)[] = new Array(n).fill(null);
      const alphas = new Float32Array(n);
      for (let y = 0; y < n; y++) {
        if (!ranks[y]) continue;
        const k = ranks[y];
        colors[y] = k === 3 ? t.removed : k === 2 ? t.added : t.modified;
        // 35% for a lone change, solid where every row changed; 10 steps so runs merge.
        alphas[y] = Math.round((0.35 + 0.65 * Math.min(1, changedW[y] / rowsW[y])) * 10) / 10;
      }
      this.minimapCache = { key, colors, alphas };
    }
    const { colors, alphas } = this.minimapCache;
    let y = 0;
    while (y < colors.length) {
      const c = colors[y];
      const al = alphas[y];
      let e = y + 1;
      while (e < colors.length && colors[e] === c && alphas[e] === al) e++;
      if (c) {
        ctx.globalAlpha = al;
        ctx.fillStyle = c;
        ctx.fillRect(3, trackTop + y, mw - 5, e - y);
      }
      y = e;
    }
    ctx.globalAlpha = 1;
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
      if ((info?.changed || info?.moved) && this.options.tooltips !== false) {
        this.tooltipTimer = window.setTimeout(() => this.showTooltip(info, hit), 260);
      } else if (hit.area === 'gutter' && this.options.tooltips !== false) {
        // A label cut short: the whole of it, and the row's numbers.
        const p = this.panes()[hit.pane];
        const r = this.rows.ref[hit.display];
        if (this.truncatedLabels.has(`${p.side}:${r}`)) this.tooltipTimer = window.setTimeout(() => this.showRowTooltip(r, p.side, hit), 260);
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
      if (e.shiftKey && this.selected) {
        this.extent = info.row === this.selected.row && info.columnIndex === this.selected.col ? null : { row: info.row, col: info.columnIndex };
      } else {
        this.selected = { row: info.row, col: info.columnIndex };
        this.extent = null;
        this.options.onSelectionChange?.(info);
      }
      this.emitRange();
      this.options.onCellClick?.(info, e);
      this.schedule();
      // Drag to extend the range.
      if (e.button !== 0) return;
      this.scroller.setPointerCapture?.(e.pointerId);
      const move = (m: PointerEvent) => {
        const c = this.cellAt(this.hitTest(m.clientX, m.clientY));
        if (!c || !this.selected) return;
        const same = c.row === this.selected.row && c.columnIndex === this.selected.col;
        this.extent = same ? null : { row: c.row, col: c.columnIndex };
        this.emitRange();
        this.schedule();
      };
      const up = () => {
        this.scroller.removeEventListener('pointermove', move);
        for (const t of ['pointerup', 'pointercancel', 'lostpointercapture'] as const) this.scroller.removeEventListener(t, up);
      };
      this.scroller.addEventListener('pointermove', move);
      for (const t of ['pointerup', 'pointercancel', 'lostpointercapture'] as const) this.scroller.addEventListener(t, up);
    }
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (e.metaKey || e.ctrlKey) {
      if (e.key === 'c' && this.selected) {
        // Exact values, not display-rounded ones; a range as tab-separated rows.
        const text = (row: number, col: number) => {
          const info = this.cellInfo(row, col);
          const v = info.newValue !== undefined ? info.newValue : info.oldValue;
          return typeof v === 'number' ? String(v) : formatCell(v ?? null);
        };
        const range = this.getRange()!;
        void navigator.clipboard?.writeText(range.rows.map((r) => range.columns.map((c) => text(r, c)).join('\t')).join('\n'));
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
      const extend = e.shiftKey && this.selected !== null;
      const from = extend ? (this.extent ?? this.selected) : this.selected;
      let d = from ? this.rows.displayOf[from.row] : -1;
      const col = from?.col ?? this.cols.cols[0];
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
      const to = { row: this.rows.ref[d], col: this.cols.cols[vi] };
      if (extend) {
        const a = this.selected!;
        this.extent = to.row === a.row && to.col === a.col ? null : to;
      } else {
        this.selected = to;
        this.extent = null;
        this.options.onSelectionChange?.(this.cellInfo(to.row, to.col));
      }
      this.emitRange();
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
    this.extent = null;
    this.options.onSelectionChange?.(this.cellInfo(r, col));
    this.emitRange();
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

  /** Everything about one cell of the diff. Throws a `RangeError` for a row or column that doesn't exist. */
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
      moved: diff.moved?.[row] === 1,
      oldRow: oi,
      newRow: ni,
      oldValue: oi >= 0 && column.oldIndex >= 0 ? diff.oldTable.rows[oi][column.oldIndex] ?? null : undefined,
      newValue: ni >= 0 && column.newIndex >= 0 ? diff.newTable.rows[ni][column.newIndex] ?? null : undefined,
      oldFormula: oi >= 0 && column.oldIndex >= 0 ? diff.oldTable.formulas?.[oi]?.[column.oldIndex] : undefined,
      newFormula: ni >= 0 && column.newIndex >= 0 ? diff.newTable.formulas?.[ni]?.[column.newIndex] : undefined,
      kind: diff.changedCells.get(row)?.[columnIndex] === 1 ? this.kindOf(row, columnIndex) : undefined,
      trace: this.tc?.trace.get({ sheet: this.tc.sheet, row, col: columnIndex }),
      decision: this.decisionAt(row, columnIndex),
    };
  }

  private showTooltip(info: CellInfo, hit: Hit): void {
    const d = this.diff!;
    const before = this.shown(d.oldTable, info.oldRow, info.column.oldIndex);
    const after = this.shown(d.newTable, info.newRow, info.column.newIndex);
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
    const rowNumber = info.newRow >= 0 ? `row ${info.newRow + this.rowOffset('new')}` : `row ${info.oldRow + this.rowOffset('old')}`;
    // With labels: "India › Revenue · C7" under the column heading; else the column name and row number.
    const rl = this.rowHeaderMode() === 'numbers' ? null : this.rowLabelOf(info.row, 'unified');
    const heading = (this.options.columnHeaders ?? 'labels') === 'letters' ? [] : this.columnHeading(info.columnIndex, info.newRow >= 0 ? 'new' : 'old');
    const colTitle = heading.length && info.column.status !== 'renamed' ? heading.join(LABEL_SEP) : info.column.name;
    const where = rl ? `${esc(this.rowLabelText(rl))} · ${esc(this.addressOf(info))}` : rowNumber;
    let html = `<div class="sd-tt-head"><span class="sd-tt-col">${esc(colTitle)}</span><span>${where}</span></div>`;
    if (info.moved) html += `<div class="sd-tt-moved">↕ Moved from row ${info.oldRow + this.rowOffset('old')}</div>`;
    if (!info.changed) {
      this.placeTooltip(html, hit);
      return;
    }
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
      const abs = displayMagnitude(Math.abs(delta.delta), this.codeOf(info.row, info.columnIndex));
      const pct = delta.ratio === null ? '' : ` (${sign}${Math.abs(delta.ratio * 100).toFixed(Math.abs(delta.ratio) < 0.1 ? 2 : 1)}%)`;
      html += `<div class="sd-tt-delta">Δ <b data-sign="${sign}">${sign}${abs}${pct}</b></div>`;
    }
    const tc = this.tc;
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
        // Each cause, its own change, and (when recalculated) how much of this cell's change it accounts for.
        const contributions = kind === 'computed' ? t.contributions : undefined;
        for (let i = 0; i < Math.min(list.length, 6); i++) {
          const d = this.describeRef(list[i]);
          const c = contributions?.[i];
          // A difference of two doubles: round away the binary noise (56082.4700000002).
          const share = typeof c === 'number' && c !== 0 ? `<span class="sd-tt-share" data-sign="${c > 0 ? '+' : '-'}">${c > 0 ? '+' : '−'}${esc(formatCell(Number(Math.abs(c).toPrecision(12))))}</span>` : contributions ? '<span></span>' : '';
          trace += `<div class="sd-tt-cause${contributions ? ' sd-tt-has-share' : ''}"><span>${kind === 'computed' ? '←' : '→'}</span><span class="sd-tt-ref">${esc(d.label)}</span><span class="sd-tt-d" data-sign="${d.sign}">${esc(d.delta)}</span>${share}</div>`;
        }
        if (count > 6) trace += `<div class="sd-tt-trace-title">+ ${count - 6} more</div>`;
        if (kind === 'computed' && count) {
          trace +=
            t.method === 'recalculated'
              ? `<div class="sd-tt-note">${contributions?.some((x) => typeof x === 'number') ? 'Right column: this cell\'s change from each edit, found by undoing it and recalculating.' : 'Found by undoing each edit and recalculating.'}</div>`
              : '<div class="sd-tt-note">From formula references; some listed edits may not have mattered.</div>';
        }
        if (t.stale) {
          trace += `<div class="sd-tt-stale">The file saved <b>${esc(formatCell(t.stale.cached) || 'nothing')}</b> here without recalculating; recalculating its inputs gives <b>${esc(formatCell(t.stale.recalculated))}</b>.</div>`;
        }
      } else {
        trace += '</div>';
      }
      html += trace + '</div>';
    }
    this.placeTooltip(html, hit);
  }

  /** Fill the tooltip and place it by the hovered cell. */
  private placeTooltip(html: string, hit: Hit): void {
    this.tooltip.innerHTML = html;
    // Position under the cell, flipping above near the bottom edge.
    const panes = this.panes();
    const p = panes[hit.pane];
    const cols = this.cols;
    const vx = hit.vcol < 0 ? p.x + BAR_WIDTH : hit.vcol < cols.frozen ? p.x + p.gutter + cols.offsets[hit.vcol] : p.x + p.gutter + cols.offsets[hit.vcol] - this.scroller.scrollLeft;
    const y = HEADER_HEIGHT + this.rows.offsets[hit.display + 1] - this.scroller.scrollTop;
    const tw = this.tooltip.offsetWidth;
    const th = this.tooltip.offsetHeight;
    const left = Math.max(8, Math.min(vx, this.width - this.minimapWidth - tw - 8));
    const rowH = this.rows.offsets[hit.display + 1] - this.rows.offsets[hit.display];
    // A tall tooltip (many causes) may fit neither side: take the roomier one, kept inside the grid.
    const below = y + 6;
    const above = y - rowH - th - 6;
    const top =
      below + th + 8 <= this.height ? below
      : above >= 4 ? above
      : this.height - y > y - rowH ? Math.max(4, Math.min(below, this.height - th - 4))
      : 4;
    this.tooltip.style.left = `${left}px`;
    this.tooltip.style.top = `${top}px`;
    this.tooltip.setAttribute('data-open', '');
    this.tooltip.removeAttribute('aria-hidden');
    this.scroller.setAttribute('aria-describedby', this.tooltip.id);
  }

  private showRowTooltip(r: number, side: 'old' | 'new' | 'unified', hit: Hit): void {
    const label = this.rowLabelOf(r, side);
    if (!label) return;
    const diff = this.diff!;
    const esc = (s: string) => s.replace(/[&<>]/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : '&gt;'));
    const nums = [diff.oldRows[r] >= 0 ? `old row ${diff.oldRows[r] + this.rowOffset('old')}` : '', diff.newRows[r] >= 0 ? `new row ${diff.newRows[r] + this.rowOffset('new')}` : ''].filter(Boolean).join(' · ');
    this.placeTooltip(`<div class="sd-tt-head"><span class="sd-tt-col">${esc(this.rowLabelText(label))}</span><span>${nums}</span></div>`, hit);
  }

  private hideTooltip(): void {
    clearTimeout(this.tooltipTimer);
    this.tooltip.removeAttribute('data-open');
    this.tooltip.setAttribute('aria-hidden', 'true');
    this.scroller.removeAttribute('aria-describedby');
  }
}

// ---- helpers ---------------------------------------------------------------

const NUMERIC = /^\s*[-+]?(\d[\d,]*\.?\d*|\.\d+)([eE][-+]?\d+)?%?\s*$/;

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  e.className = cls;
  return e;
}

/**
 * `+5.23 · +8.0%` (or `+5.23` when the old value was 0). With the cell's
 * number format, the absolute part is in it (`+$1,200`), and a percentage
 * cell's change is in points (`+1.2 pp`), so it isn't read as the ratio.
 */
function formatDelta(d: { delta: number; ratio: number | null }, code?: string): { label: string; up: boolean; abs: string; pct?: string } {
  const up = d.delta > 0;
  const sign = up ? '+' : '−';
  const abs = `${sign}${displayMagnitude(Math.abs(d.delta), code)}`;
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
