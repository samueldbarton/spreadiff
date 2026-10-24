import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { diffWorkbookObjects, diffWorkbooks, parseXlsx, type Workbook } from '../src/core';

const poi = (name: string) => fileURLToPath(new URL(`fixtures/poi/${name}`, import.meta.url));

/** The file with one part's text rewritten. */
function edit(bytes: Uint8Array, part: string, f: (xml: string) => string): Uint8Array {
  const files = unzipSync(bytes);
  files[part] = strToU8(f(strFromU8(files[part])));
  return zipSync(files);
}

describe('workbook objects', () => {
  it('reports names, charts and pivots that changed', () => {
    const base = (charts: Workbook['charts'], pivots: Workbook['pivots'], names: Record<string, string>): Workbook => ({ name: 'w', sheets: [{ name: 'S', columns: ['A'], rows: [[1]], names: { Local: '$A$1' } }], charts, pivots, names });
    const a = base(
      [{ sheet: 'S', title: 'Revenue', types: ['bar'], series: [{ name: 'FY24', values: 'S!$B$2:$B$13' }] }, { sheet: 'S', title: 'Chart 2', types: ['line'], series: [] }],
      [{ sheet: 'P', name: 'PivotTable1', location: 'A3:C9', source: 'Data!A1:D100', values: ['Sum of Revenue'] }],
      { Rate: 'S!$B$1', Old: 'S!$C$1' },
    );
    const b = base(
      [{ sheet: 'S', title: 'Revenue', types: ['bar', 'line'], series: [{ name: 'FY24', values: 'S!$B$2:$B$14' }, { name: 'FY25', values: 'S!$C$2:$C$14' }] }],
      [{ sheet: 'P', name: 'PivotTable1', location: 'A3:C9', source: 'Data!A1:D120', values: ['Sum of Revenue', 'Count of Units'] }],
      { Rate: 'S!$B$2', New: 'S!$D$1' },
    );
    const changes = diffWorkbookObjects(a, b);
    expect(changes).toEqual([
      { kind: 'name', change: 'modified', name: 'Rate', details: ['S!$B$1 → S!$B$2'] },
      { kind: 'name', change: 'added', name: 'New', details: ['= S!$D$1'] },
      { kind: 'name', change: 'removed', name: 'Old', details: ['was S!$C$1'] },
      { kind: 'chart', change: 'modified', name: 'S › Revenue', details: ['type: bar → bar + line', 'series 1 (FY24) values: S!$B$2:$B$13 → S!$B$2:$B$14', 'series 2 (FY25) added: S!$C$2:$C$14'] },
      { kind: 'chart', change: 'removed', name: 'S › Chart 2', details: [''] },
      { kind: 'pivot', change: 'modified', name: 'P › PivotTable1', details: ['source: Data!A1:D100 → Data!A1:D120', 'values: Sum of Revenue → Sum of Revenue, Count of Units'] },
    ]);
    expect(diffWorkbookObjects(a, a)).toEqual([]);
  });

  it('reads them from real Excel files, and diffWorkbooks reports them', () => {
    const charts = new Uint8Array(readFileSync(poi('WithTwoCharts.xlsx')));
    const wa = parseXlsx(charts);
    expect(wa.charts).toEqual([
      { sheet: 'Sheet2', title: 'Chart 1', types: ['line'], series: [{ name: '1st Column', values: 'Sheet1!$A$1:$A$6' }, { name: '2nd Column', values: 'Sheet1!$B$1:$B$6' }] },
      { sheet: 'Sheet3', title: 'Chart 1', types: ['area'], series: [{ name: 'Squares', values: 'Sheet1!$E$1:$E$6' }, { name: 'Base Numbers', values: 'Sheet1!$D$1:$D$6' }] },
    ]);
    const wb = parseXlsx(edit(charts, 'xl/charts/chart1.xml', (x) => x.replace('Sheet1!$A$1:$A$6', 'Sheet1!$A$1:$A$7')));
    expect(diffWorkbooks(wa, wb).objects).toEqual([{ kind: 'chart', change: 'modified', name: 'Sheet2 › Chart 1', details: ['series 1 (1st Column) values: Sheet1!$A$1:$A$6 → Sheet1!$A$1:$A$7'] }]);
    expect(diffWorkbooks(wa, wa).objects).toBeUndefined();

    const pivots = new Uint8Array(readFileSync(poi('ExcelPivotTableSample.xlsx')));
    const pa = parseXlsx(pivots);
    expect(pa.pivots?.map((p) => [p.sheet, p.name, p.location, p.source, p.values.join()])).toEqual([
      ['Tabelle2', 'PivotTable1', 'A3:D7', 'Tabelle1!A1:C4', 'Summe von Field2'],
      ['Tabelle3', 'PivotTable1', 'A3:B6', 'Tabelle1', 'Summe von Field2'],
    ]);
    const pb = parseXlsx(edit(pivots, 'xl/pivotCache/pivotCacheDefinition1.xml', (x) => x.replace('ref="A1:C4"', 'ref="A1:C9"')));
    expect(diffWorkbooks(pa, pb).objects).toEqual([{ kind: 'pivot', change: 'modified', name: 'Tabelle2 › PivotTable1', details: ['source: Tabelle1!A1:C4 → Tabelle1!A1:C9'] }]);
  });

  it('finds parts whose names differ in case from their links', () => {
    const wb = parseXlsx(new Uint8Array(readFileSync(poi('123233_charts.xlsx'))));
    expect(wb.charts?.length).toBe(4);
  });
});
