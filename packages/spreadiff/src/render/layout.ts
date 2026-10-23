import { ROW_STATUS, type TableDiff } from '../core/types.js';

/** A collapsed run of unchanged aligned rows `[start, end)`. */
export interface CollapsedRegion {
  start: number;
  end: number;
}

export const ROW_KIND = { data: 0, separator: 1, annotation: 2 } as const;

/**
 * The vertical layout: an ordered list of display rows (data rows, collapsed
 * separators, annotation rows) with prefix-summed offsets for O(log n)
 * hit-testing and windowing.
 */
export interface RowLayout {
  count: number;
  kind: Uint8Array;
  /** Aligned row index for data rows; region index for separators; annotation index for annotations. */
  ref: Int32Array;
  /** Top offset of each display row; `offsets[count]` is the total height. */
  offsets: Float64Array;
  regions: CollapsedRegion[];
  /** Display index of each aligned row, -1 while collapsed. */
  displayOf: Int32Array;
}

export interface LayoutInput {
  diff: TableDiff;
  collapse: boolean;
  context: number;
  /** Regions the user expanded, keyed by `start:end` of the original region. */
  expanded: Set<string>;
  rowHeight: number;
  /** Height for modified rows (taller in the unified layout to fit old/new). */
  modifiedHeight: number;
  separatorHeight: number;
  /** aligned row → annotation heights to place after that row. */
  annotationsAfter: Map<number, number[]>;
  /** Filter rows (e.g. only show changes). Return false to hide an aligned row. */
  filter?: (alignedRow: number) => boolean;
}

/** Minimum collapsed run worth a separator (smaller runs just render). */
const MIN_COLLAPSE = 4;

export function computeRowLayout(input: LayoutInput): RowLayout {
  const { diff, collapse, context } = input;
  const n = diff.rowCount;
  const status = diff.status;
  const visible = new Uint8Array(n).fill(1);
  if (input.filter) for (let r = 0; r < n; r++) visible[r] = input.filter(r) ? 1 : 0;

  // Mark rows within `context` of any change as kept.
  const keep = new Uint8Array(n);
  if (collapse) {
    let lastChange = -Infinity;
    for (let r = 0; r < n; r++) {
      if (status[r] !== ROW_STATUS.equal) lastChange = r;
      if (r - lastChange <= context) keep[r] = 1;
    }
    let nextChange = Infinity;
    for (let r = n - 1; r >= 0; r--) {
      if (status[r] !== ROW_STATUS.equal) nextChange = r;
      if (nextChange - r <= context) keep[r] = 1;
    }
  } else {
    keep.fill(1);
  }

  const regions: CollapsedRegion[] = [];
  const kinds: number[] = [];
  const refs: number[] = [];
  const heights: number[] = [];
  const displayOf = new Int32Array(n).fill(-1);
  const pushRow = (r: number) => {
    displayOf[r] = kinds.length;
    kinds.push(ROW_KIND.data);
    refs.push(r);
    heights.push(status[r] === ROW_STATUS.modified ? input.modifiedHeight : input.rowHeight);
    const anns = input.annotationsAfter.get(r);
    if (anns) {
      for (const h of anns) {
        kinds.push(ROW_KIND.annotation);
        refs.push(r);
        heights.push(h);
      }
    }
  };

  let r = 0;
  while (r < n) {
    if (!visible[r]) {
      r++;
      continue;
    }
    if (keep[r]) {
      pushRow(r++);
      continue;
    }
    let end = r;
    while (end < n && !keep[end]) end++;
    const key = `${r}:${end}`;
    if (end - r < MIN_COLLAPSE || input.expanded.has(key)) {
      for (let i = r; i < end; i++) if (visible[i]) pushRow(i);
    } else {
      regions.push({ start: r, end });
      kinds.push(ROW_KIND.separator);
      refs.push(regions.length - 1);
      heights.push(input.separatorHeight);
    }
    r = end;
  }

  const count = kinds.length;
  const offsets = new Float64Array(count + 1);
  for (let i = 0; i < count; i++) offsets[i + 1] = offsets[i] + heights[i];
  return {
    count,
    kind: Uint8Array.from(kinds),
    ref: Int32Array.from(refs),
    offsets,
    regions,
    displayOf,
  };
}

/** Index of the display row containing `y` (clamped). */
export function rowAt(layout: RowLayout, y: number): number {
  const o = layout.offsets;
  let lo = 0;
  let hi = layout.count - 1;
  if (hi < 0) return -1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (o[mid] <= y) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** Display index of the first separator or changed row after/before `from`. */
export function findChange(
  layout: RowLayout,
  diff: TableDiff,
  from: number,
  dir: 1 | -1,
): number {
  // Jump to the start of the next change *block*, not every changed row.
  const isChange = (d: number) =>
    layout.kind[d] === ROW_KIND.data && diff.status[layout.ref[d]] !== ROW_STATUS.equal;
  let d = from + dir;
  if (dir === 1) {
    while (d < layout.count && isChange(d) && isChange(d - 1)) d++;
    while (d < layout.count && !isChange(d)) d++;
    return d < layout.count ? d : -1;
  }
  while (d >= 0 && !isChange(d)) d--;
  while (d > 0 && isChange(d - 1)) d--;
  return d >= 0 ? d : -1;
}

/** Visible columns (indices into `diff.columns`) with x offsets. */
export interface ColumnLayout {
  cols: Int32Array;
  widths: Float64Array;
  /** Left offset of each visible column in content space; last = total width. */
  offsets: Float64Array;
  frozen: number;
  frozenWidth: number;
}

export function computeColumnLayout(
  visibleCols: number[],
  widthOf: (col: number) => number,
  frozen: number,
): ColumnLayout {
  const cols = Int32Array.from(visibleCols);
  const widths = Float64Array.from(visibleCols, widthOf);
  const offsets = new Float64Array(cols.length + 1);
  for (let i = 0; i < cols.length; i++) offsets[i + 1] = offsets[i] + widths[i];
  const f = Math.min(frozen, cols.length);
  return { cols, widths, offsets, frozen: f, frozenWidth: offsets[f] };
}
