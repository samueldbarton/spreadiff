/**
 * Every word the grid and `SpreadsheetDiff` show or announce, so an app can
 * translate them. Pass a partial set as the `strings` option; the rest stay
 * English. Messages with a number take it twice: `count` formatted for the
 * `locale` option ("1,234"), and `n` itself, to choose a plural form.
 *
 * Values from the files (cell values, sheet and column names, formulas)
 * aren't translated: they're shown as the sheet has them.
 */
export interface DiffStrings {
  // ---- The grid: what a screen reader hears ------------------------------
  /** The grid's accessible name before a diff is shown. */
  gridLabel: string;
  /** The grid's accessible name: a summary of the diff. Each `parts` entry comes from the `summary*` messages. */
  gridSummary: (parts: string[]) => string;
  /** Summary part: "3 rows modified". */
  summaryRowsModified: (count: string, n: number) => string;
  /** Summary part: "1 row added". */
  summaryRowsAdded: (count: string, n: number) => string;
  /** Summary part: "2 rows removed". */
  summaryRowsRemoved: (count: string, n: number) => string;
  /** Summary part: "1 row moved". */
  summaryRowsMoved: (count: string, n: number) => string;
  /** Summary part: "7 cells changed". */
  summaryCellsChanged: (count: string, n: number) => string;
  /** Heading of the row-header column. */
  rowHeaderColumn: string;
  /** A renamed column's header: "Price, renamed from Cost". */
  columnRenamed: (name: string, oldName: string) => string;
  /** An added column's header: "Region, added column". */
  columnAdded: (name: string) => string;
  /** A removed column's header: "Notes, removed column". */
  columnRemoved: (name: string) => string;
  /** A folded run of unchanged rows, as announced. */
  unchangedRowsCollapsed: (count: string, n: number) => string;
  /** A row header: `label` when the row has an inferred label, else null. */
  rowHeader: (row: number, label: string | null) => string;
  /** Status after a row header: "modified". */
  rowModified: string;
  /** Status after a row header: "added". */
  rowAdded: string;
  /** Status after a row header: "removed". */
  rowRemoved: string;
  /** After a moved row's header: "moved from row 4". */
  movedFromRow: (row: number) => string;
  /** An empty cell. */
  empty: string;
  /** A cell of an added row whose column the new table doesn't have. */
  notInNewTable: string;
  /** A cell in a removed column. */
  cellColumnRemoved: string;
  /** Before the old value of a changed cell (also drawn under it in the `delta` display). */
  was: (value: string) => string;
  /** An increase: "up 2.5 (25.0%)". */
  up: (amount: string) => string;
  /** A decrease: "down 2.5 (20.0%)". */
  down: (amount: string) => string;
  /** Why a cell changed: someone edited it. */
  kindInput: string;
  /** Why a cell changed: its formula was edited. */
  kindFormula: string;
  /** Why a cell changed: it's a formula result that moved. */
  kindComputed: string;
  /** The new formula of a cell whose formula changed. */
  formulaIs: (formula: string) => string;
  /** A cell whose change was accepted in review. */
  accepted: string;
  /** A cell whose change was rejected in review. */
  rejected: string;

  // ---- The grid: what it draws -------------------------------------------
  /** The separator over a folded run of unchanged rows. */
  unchangedRows: (count: string, n: number) => string;

  // ---- The grid: the popover ----------------------------------------------
  /** Where a cell is, when rows have no labels. */
  tooltipRow: (row: number) => string;
  /** A moved row: "Moved from row 4". */
  tooltipMovedFromRow: (row: number) => string;
  /** A row label's popover: "old row 4". */
  tooltipOldRow: (row: number) => string;
  /** A row label's popover: "new row 5". */
  tooltipNewRow: (row: number) => string;
  /** The popover's label for an input edit. */
  tooltipKindInput: string;
  /** The popover's label for a formula edit. */
  tooltipKindFormula: string;
  /** The popover's label for a downstream change. */
  tooltipKindComputed: string;
  /** Before a downstream change's causes: "changed because of 3 edits:". */
  changedBecause: (count: string, n: number) => string;
  /** A downstream change from a volatile function, with no edit behind it. */
  causeVolatile: string;
  /** A downstream change behind a dynamic reference the trace can't follow. */
  causeDynamic: string;
  /** A downstream change whose cause wasn't found. */
  causeNotFound: string;
  /** Before an edit's impacts: "flows into 4 cells:". */
  flowsInto: (count: string, n: number) => string;
  /** An edit that moved nothing downstream. */
  nothingDownstream: string;
  /** Causes or impacts past the six listed: "+ 3 more". */
  more: (count: string, n: number) => string;
  /** Under causes with contributions: how they were found. */
  noteContributions: string;
  /** Under causes found by recalculating, without contributions. */
  noteRecalculated: string;
  /** Under causes found from references alone. */
  noteReferences: string;
  /** A result saved without recalculating. The two values are shown in bold. */
  stale: (saved: string, recalculated: string) => string;
  /** What `stale` shows for a cell saved with nothing in it. */
  staleNothing: string;

  // ---- SpreadsheetDiff ---------------------------------------------------
  /** The sheet tabs' accessible name. */
  sheetsLabel: string;
  /** The outputs panel's accessible name. */
  outputsLabel: string;
  /** Shown while the files are read and compared. */
  comparing: string;
  /** The count on a sheet's tab, for added and removed sheets. */
  sheetNew: string;
  /** The count on a removed sheet's tab. */
  sheetRemoved: string;
  /** An unchanged sheet's tab tooltip. */
  sheetUnchangedTitle: (sheet: string) => string;
  /** A changed sheet's tab tooltip. */
  sheetTitle: (sheet: string, cellsChanged: number, rowsAdded: number, rowsRemoved: number) => string;
  /** The review bar's count: "2 rejected · 1 accepted · 9 changes". */
  tally: (rejected: number, accepted: number, changes: number) => string;
  /** The Accept button. */
  accept: string;
  /** The Accept button's tooltip. */
  acceptTitle: string;
  /** The Reject button. */
  reject: string;
  /** The Reject button's tooltip. */
  rejectTitle: string;
  /** The Undo button, which clears decisions. */
  undo: string;
  /** The Undo button's tooltip. */
  undoTitle: string;
  /** The download button, for an .xlsx. */
  downloadXlsx: string;
  /** The download button, for a CSV. */
  downloadCsv: string;
  /** The download button's tooltip. */
  downloadTitle: string;
  /** Added to the downloaded file's name, before its extension. */
  reviewedSuffix: string;
  /** The outputs panel heading: "Changed outputs · 12". */
  outputsHeading: (count: string, n: number) => string;
  /** The outputs panel heading when no output changed. */
  noOutputs: string;
  /** The button that lists more outputs: "Show 100 more". */
  showMore: (count: string, n: number) => string;
  /** The waterfall's first bar. */
  waterfallBefore: string;
  /** The waterfall's last bar. */
  waterfallAfter: string;
  /** The waterfall bar for causes not listed. */
  waterfallOther: string;
  /** The waterfall bar for what edits do together beyond their own contributions. */
  waterfallInteraction: string;
  /** In place of a waterfall, for a value that isn't a number. */
  waterfallNotNumber: string;
  /** Summary of changed names, charts and pivot tables. */
  objectsSummary: (count: string, n: number) => string;
  /** A changed defined name. */
  objectName: string;
  /** A changed chart. */
  objectChart: string;
  /** A changed pivot table. */
  objectPivot: string;
}

const s = (n: number, one: string, many = `${one}s`) => (n === 1 ? one : many);

/** The English strings: the defaults. */
export const ENGLISH_STRINGS: DiffStrings = {
  gridLabel: 'Spreadsheet diff',
  gridSummary: (parts) => `Spreadsheet diff: ${parts.length ? parts.join(', ') : 'no changes'}`,
  summaryRowsModified: (c, n) => `${c} ${s(n, 'row')} modified`,
  summaryRowsAdded: (c, n) => `${c} ${s(n, 'row')} added`,
  summaryRowsRemoved: (c, n) => `${c} ${s(n, 'row')} removed`,
  summaryRowsMoved: (c, n) => `${c} ${s(n, 'row')} moved`,
  summaryCellsChanged: (c, n) => `${c} ${s(n, 'cell')} changed`,
  rowHeaderColumn: 'Row',
  columnRenamed: (name, old) => `${name}, renamed from ${old}`,
  columnAdded: (name) => `${name}, added column`,
  columnRemoved: (name) => `${name}, removed column`,
  unchangedRowsCollapsed: (c, n) => `${c} unchanged ${s(n, 'row')}, collapsed`,
  rowHeader: (row, label) => (label ? `${label}, row ${row}` : `Row ${row}`),
  rowModified: 'modified',
  rowAdded: 'added',
  rowRemoved: 'removed',
  movedFromRow: (row) => `moved from row ${row}`,
  empty: 'empty',
  notInNewTable: 'not in the new table',
  cellColumnRemoved: 'column removed',
  was: (v) => `was ${v}`,
  up: (a) => `up ${a}`,
  down: (a) => `down ${a}`,
  kindInput: 'edited',
  kindFormula: 'formula edited',
  kindComputed: 'recalculated result',
  formulaIs: (f) => `formula ${f}`,
  accepted: 'accepted',
  rejected: 'rejected',
  unchangedRows: (c, n) => `${c} unchanged ${s(n, 'row')}`,
  tooltipRow: (row) => `row ${row}`,
  tooltipMovedFromRow: (row) => `Moved from row ${row}`,
  tooltipOldRow: (row) => `old row ${row}`,
  tooltipNewRow: (row) => `new row ${row}`,
  tooltipKindInput: 'input edit',
  tooltipKindFormula: 'formula edit',
  tooltipKindComputed: 'downstream',
  changedBecause: (c, n) => `changed because of ${c} ${s(n, 'edit')}:`,
  causeVolatile: 'volatile function (NOW, RAND, …)',
  causeDynamic: 'dynamic reference (INDIRECT, OFFSET, …)',
  causeNotFound: 'cause not found',
  flowsInto: (c, n) => `flows into ${c} ${s(n, 'cell')}:`,
  nothingDownstream: 'nothing downstream changed',
  more: (c) => `+ ${c} more`,
  noteContributions: "Right column: this cell's change from each edit, found by undoing it and recalculating.",
  noteRecalculated: 'Found by undoing each edit and recalculating.',
  noteReferences: 'From formula references; some listed edits may not have mattered.',
  stale: (saved, recalculated) => `The file saved ${saved} here without recalculating; recalculating its inputs gives ${recalculated}.`,
  staleNothing: 'nothing',
  sheetsLabel: 'Sheets',
  outputsLabel: 'Changed outputs',
  comparing: 'Comparing…',
  sheetNew: 'new',
  sheetRemoved: 'removed',
  sheetUnchangedTitle: (sheet) => `${sheet}: no changes`,
  sheetTitle: (sheet, cells, added, removed) => `${sheet}: ${cells} cells changed, ${added} rows added, ${removed} removed`,
  tally: (rejected, accepted, changes) => `${rejected} rejected · ${accepted} accepted · ${changes} ${s(changes, 'change')}`,
  accept: 'Accept',
  acceptTitle: 'Accept the selected changes (a)',
  reject: 'Reject',
  rejectTitle: 'Reject the selected changes (r)',
  undo: 'Undo',
  undoTitle: 'Clear decisions on the selection (u)',
  downloadXlsx: 'Download .xlsx',
  downloadCsv: 'Download .csv',
  downloadTitle: 'The new file with rejected changes undone',
  reviewedSuffix: ' (reviewed)',
  outputsHeading: (c) => `Changed outputs · ${c}`,
  noOutputs: 'No outputs changed',
  showMore: (c) => `Show ${c} more`,
  waterfallBefore: 'Before',
  waterfallAfter: 'After',
  waterfallOther: 'Other causes, together',
  waterfallInteraction: 'Interaction of the edits',
  waterfallNotNumber: 'Not a number: no waterfall.',
  objectsSummary: (c, n) => `${c} ${s(n, 'change')} to names, charts or pivot tables`,
  objectName: 'Name',
  objectChart: 'Chart',
  objectPivot: 'Pivot table',
};

/** The strings in effect, and counts formatted for the locale. */
export interface Words {
  t: DiffStrings;
  /** A count as text in the locale: 1234 → "1,234". */
  count: (n: number) => string;
}

const formats = new Map<string, Intl.NumberFormat>();

/** Merge `strings` over the English defaults; `locale` formats counts (default: the browser's). */
export function words(strings: Partial<DiffStrings> | undefined, locale: string | undefined): Words {
  const key = locale ?? '';
  let f = formats.get(key);
  if (!f) formats.set(key, (f = new Intl.NumberFormat(locale)));
  const nf = f;
  return { t: strings ? { ...ENGLISH_STRINGS, ...strings } : ENGLISH_STRINGS, count: (n) => nf.format(n) };
}
