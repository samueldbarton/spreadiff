import type { TableDiff } from './types.js';

/**
 * The column that names each row when rows were matched by order (no key
 * columns): the first column on both sides whose values are mostly text,
 * like a model's "line" column. -1 if there's none. Renderers keep it when
 * hiding unchanged columns, so each changed row still says which row it is.
 */
export function rowLabelColumn(diff: TableDiff): number {
  if (diff.keyColumns.length) return -1;
  const t = diff.newTable;
  const step = Math.max(1, Math.floor(t.rows.length / 200)); // a sample of rows is enough
  for (let ci = 0; ci < diff.columns.length; ci++) {
    const c = diff.columns[ci];
    if (c.oldIndex < 0 || c.newIndex < 0) continue;
    let text = 0;
    let filled = 0;
    for (let r = 0; r < t.rows.length; r += step) {
      const v = t.rows[r][c.newIndex];
      if (v === null || v === undefined || v === '') continue;
      filled++;
      if (typeof v === 'string') text++;
    }
    if (filled && text / filled >= 0.8) return ci;
  }
  return -1;
}
