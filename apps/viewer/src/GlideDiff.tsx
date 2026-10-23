import { DataEditor, GridCellKind, type GridCell, type GridColumn, type Item, type Theme } from '@glideapps/glide-data-grid';
import '@glideapps/glide-data-grid/dist/index.css';
import { useCallback, useMemo } from 'react';
import { formatCell, ROW_STATUS, type TableDiff } from 'spreadiff';

/**
 * The same diff rendered with Glide Data Grid, the most widely used canvas
 * grid for React. It's here as a baseline for the benchmark, and to show
 * what you get from a general-purpose grid with diff colors bolted on.
 */
export function GlideDiff({ diff, dark }: { diff: TableDiff; dark: boolean }) {
  const palette = dark
    ? { added: '#16301f', removed: '#3a1c1e', addedCell: '#1f4a2d', removedCell: '#5a2528', modified: '#11233a', fg: '#ededef', bg: '#0b0b0c', header: '#151517', border: '#26262a' }
    : { added: '#e7f8ed', removed: '#ffe9eb', addedCell: '#c3efd2', removedCell: '#ffcdd2', modified: '#eef6ff', fg: '#0b0b0c', bg: '#ffffff', header: '#f7f7f8', border: '#e6e6e8' };

  const theme: Partial<Theme> = useMemo(
    () => ({
      bgCell: palette.bg,
      bgHeader: palette.header,
      bgHeaderHasFocus: palette.header,
      textDark: palette.fg,
      textHeader: palette.fg,
      borderColor: palette.border,
      fontFamily: 'Inter, system-ui, sans-serif',
      baseFontStyle: '13px',
      headerFontStyle: '600 12px',
    }),
    [dark],
  );

  const columns: GridColumn[] = useMemo(
    () => [
      { title: '', id: '__status', width: 36 },
      ...diff.columns.map((c) => ({
        title: c.status === 'renamed' ? `${c.oldName} → ${c.name}` : c.status === 'added' ? `+ ${c.name}` : c.status === 'removed' ? `− ${c.name}` : c.name,
        id: c.name,
        width: 140,
      })),
    ],
    [diff],
  );

  const getCellContent = useCallback(
    ([col, row]: Item): GridCell => {
      const st = diff.status[row];
      const rowBg = st === ROW_STATUS.added ? palette.added : st === ROW_STATUS.removed ? palette.removed : st === ROW_STATUS.modified ? palette.modified : undefined;
      if (col === 0) {
        const sign = st === ROW_STATUS.added ? '+' : st === ROW_STATUS.removed ? '−' : st === ROW_STATUS.modified ? '~' : '';
        return { kind: GridCellKind.Text, data: sign, displayData: sign, allowOverlay: false, themeOverride: rowBg ? { bgCell: rowBg } : undefined };
      }
      const ci = col - 1;
      const c = diff.columns[ci];
      const oi = diff.oldRows[row];
      const ni = diff.newRows[row];
      const ov = oi >= 0 && c.oldIndex >= 0 ? formatCell(diff.oldTable.rows[oi][c.oldIndex] ?? null) : '';
      const nv = ni >= 0 && c.newIndex >= 0 ? formatCell(diff.newTable.rows[ni][c.newIndex] ?? null) : '';
      const changed = diff.changedCells.get(row)?.[ci] === 1;
      const text = changed ? `${ov} → ${nv}` : ni >= 0 ? nv : ov;
      return {
        kind: GridCellKind.Text,
        data: text,
        displayData: text,
        allowOverlay: false,
        themeOverride: changed ? { bgCell: palette.addedCell } : rowBg ? { bgCell: rowBg } : undefined,
      };
    },
    [diff, dark],
  );

  return (
    <DataEditor
      columns={columns}
      rows={diff.rowCount}
      getCellContent={getCellContent}
      theme={theme}
      rowMarkers="number"
      smoothScrollX
      smoothScrollY
      freezeColumns={1}
      width="100%"
      height="100%"
    />
  );
}
