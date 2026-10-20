import { formatCell } from './text.js';
import { ROW_STATUS, type TableDiff } from './types.js';

/**
 * daff-style "highlighter" CSV: a leading action column (`+++`, `---`, `->`)
 * and `old->new` in changed cells. Readable in any spreadsheet app.
 */
export function toUnifiedCsv(diff: TableDiff): string {
  const esc = (s: string) => (/[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const lines: string[] = [];
  const header = ['@@', ...diff.columns.map((c) =>
    c.status === 'added' ? `+++ ${c.name}` : c.status === 'removed' ? `--- ${c.name}` : c.status === 'renamed' ? `${c.oldName}->${c.name}` : c.name,
  )];
  lines.push(header.map(esc).join(','));
  for (let r = 0; r < diff.rowCount; r++) {
    const st = diff.status[r];
    const o = diff.oldRows[r] >= 0 ? diff.oldTable.rows[diff.oldRows[r]] : null;
    const n = diff.newRows[r] >= 0 ? diff.newTable.rows[diff.newRows[r]] : null;
    const mask = diff.changedCells.get(r);
    const action = st === ROW_STATUS.added ? '+++' : st === ROW_STATUS.removed ? '---' : st === ROW_STATUS.modified ? '->' : '';
    const cells = diff.columns.map((c, i) => {
      const ov = o && c.oldIndex >= 0 ? formatCell(o[c.oldIndex] ?? null) : '';
      const nv = n && c.newIndex >= 0 ? formatCell(n[c.newIndex] ?? null) : '';
      if (mask?.[i]) return `${ov}->${nv}`;
      return n ? nv : ov;
    });
    lines.push([action, ...cells].map(esc).join(','));
  }
  return lines.join('\n') + '\n';
}
