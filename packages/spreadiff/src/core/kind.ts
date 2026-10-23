import type { ChangeKind } from './trace.js';
import type { TableDiff } from './types.js';

/**
 * Classify a changed cell without a full trace: an edited input, an edited
 * formula, or a computed (downstream) result — same formula, new value.
 */
export function changeKind(diff: TableDiff, row: number, col: number): ChangeKind {
  if (diff.formulaChanged?.get(row)?.[col]) return 'formula';
  const ni = diff.newRows[row];
  const oi = diff.oldRows[row];
  const c = diff.columns[col];
  const nf = ni >= 0 && c.newIndex >= 0 ? diff.newTable.formulas?.[ni]?.[c.newIndex] : null;
  const of = oi >= 0 && c.oldIndex >= 0 ? diff.oldTable.formulas?.[oi]?.[c.oldIndex] : null;
  return nf && of ? 'computed' : 'input';
}
