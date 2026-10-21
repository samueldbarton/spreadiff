import { formatCell } from './text.js';
import { ROW_STATUS, type TableDiff } from './types.js';

/**
 * daff-style "highlighter" CSV: a leading action column (`+++`, `---`, `->`)
 * and `old->new` in changed cells. Readable in any spreadsheet app.
 */
export function toUnifiedCsv(diff: TableDiff): string {
  const esc = (s: string) => (/[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  // Cells from untrusted files must not become live formulas when the export
  // is opened in a spreadsheet app (CSV injection). Plain numbers pass.
  const safe = (s: string) => (/^[=+\-@\t\r]/.test(s) && !/^[-+]?\d/.test(s) ? `'${s}` : s);
  const lines: string[] = [];
  const header = ['@@', ...diff.columns.map((c) =>
    c.status === 'added' ? `+++ ${c.name}` : c.status === 'removed' ? `--- ${c.name}` : c.status === 'renamed' ? `${c.oldName}->${c.name}` : c.name,
  )];
  lines.push(header.map((h, i) => (i === 0 ? h : esc(safe(h)))).join(','));
  for (let r = 0; r < diff.rowCount; r++) {
    const st = diff.status[r];
    const o = diff.oldRows[r] >= 0 ? diff.oldTable.rows[diff.oldRows[r]] : null;
    const n = diff.newRows[r] >= 0 ? diff.newTable.rows[diff.newRows[r]] : null;
    const mask = diff.changedCells.get(r);
    const action = st === ROW_STATUS.added ? '+++' : st === ROW_STATUS.removed ? '---' : st === ROW_STATUS.modified ? '->' : '';
    const cells = diff.columns.map((c, i) => {
      const ov = o && c.oldIndex >= 0 ? formatCell(o[c.oldIndex] ?? null) : '';
      const nv = n && c.newIndex >= 0 ? formatCell(n[c.newIndex] ?? null) : '';
      if (mask?.[i]) return safe(`${ov}->${nv}`);
      return safe(n ? nv : ov);
    });
    lines.push([action, ...cells].map(esc).join(','));
  }
  return lines.join('\n') + '\n';
}
