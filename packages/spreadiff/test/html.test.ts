import { describe, expect, it } from 'vitest';
import { diffTables, parseCsv, renderHTML, renderText, HTML_CSS } from '../src/html';
import type { Table } from '../src/core';

const csv = (s: string) => parseCsv(s.trim() + '\n');
const a = csv('id,name,price\n1,Lamp,10\n2,Chair,20\n3,Desk,30\n4,Mug,4\n5,Rug,50\n6,Vase,6\n7,Tray,7');
const b = csv('id,name,price\n1,Lamp,10\n2,Armchair,25\n3,Desk,30\n4,Mug,4\n5,Rug,50\n6,Vase,6\n8,Clock,12');

describe('html', () => {
  it('split layout: aligned rows, line numbers, hunks, intra-cell marks', () => {
    const html = renderHTML(diffTables(a, b), { layout: 'split', context: 1 });
    expect(html).toContain('sdl-split');
    expect(html).toContain('<mark>Armc</mark>hair'); // only the differing characters (C → Armc)
    expect(html).toMatch(/⋯ 2 unchanged rows/);
    expect(html).toContain('sdl-removed'); // id 7
    expect(html).toContain('sdl-added'); // id 8
    // Removed row: right side is a placeholder; added row: left side is.
    expect(html.match(/class="sdl-gap" colspan="3"/g)).toHaveLength(2);
  });

  it('unified stacked renders a − line then a + line for each modified row', () => {
    const html = renderHTML(diffTables(a, b), { layout: 'unified' });
    const minus = html.indexOf('sdl-minus');
    const plus = html.indexOf('sdl-plus');
    expect(minus).toBeGreaterThan(0);
    expect(plus).toBeGreaterThan(minus);
  });

  it('delta and inline displays', () => {
    const delta = renderHTML(diffTables(a, b), { layout: 'unified', display: 'delta' });
    expect(delta).toContain('+5 · +25.0%');
    const inline = renderHTML(diffTables(a, b), { layout: 'unified', display: 'inline' });
    expect(inline).toContain('<s>20</s><span class="sdl-arrow">→</span>25');
  });

  it('marks computed (downstream) cells', () => {
    const mk = (p: number): Table => ({
      name: 's', columns: ['item', 'price', 'qty', 'total'], origin: { row: 1, col: 0 },
      rows: [['a', p, 2, p * 2]], formulas: [[null, null, null, '=B2*C2']],
    });
    const html = renderHTML(diffTables(mk(1), mk(3)), { layout: 'split' });
    expect(html).toContain('data-kind="computed"');
    expect(html).toContain('data-kind="input"');
  });

  it('escapes cell content', () => {
    const html = renderHTML(diffTables(csv('k,v\n1,a'), csv('k,v\n1,<img src=x onerror=alert(1)>')));
    expect(html).not.toContain('<img');
    expect(html).toContain('&lt;img');
  });

  it('renders aligned plain text', () => {
    expect(renderText(diffTables(a, b), { context: 0 })).toBe(
      [
        '   id  name      price',
        '⋯  1 unchanged row',
        '-  2   Chair     20',
        '+  2   Armchair  25',
        '⋯  4 unchanged rows',
        '-  7   Tray      7',
        '+  8   Clock     12',
        '',
      ].join('\n'),
    );
  });

  it('ships its CSS', () => expect(HTML_CSS).toContain('.sdl-split'));
});

describe('changedColumnsOnly keeps what names the rows', () => {
  it('keeps the row-label column when rows were matched by order', async () => {
    const { diffTables, parseCsv, renderText } = await import('../src/html');
    const d = diffTables(parseCsv('line,a,b\nRevenue,1,2\nCosts,3,4'), parseCsv('line,a,b\nRevenue,1,9\nCosts,3,4'), { keys: 'none' });
    expect(renderText(d, { changedColumnsOnly: true, context: 0 }).split('\n')[0].trim().split(/\s+/)).toEqual(['line', 'b']);
  });
  it('keeps only keys when there are keys, and no label column for numbers', async () => {
    const { diffTables, parseCsv, renderText } = await import('../src/html');
    const keyed = diffTables(parseCsv('id,name,b\n1,x,2\n2,y,4'), parseCsv('id,name,b\n1,x,9\n2,y,4'), { keys: ['id'] });
    expect(renderText(keyed, { changedColumnsOnly: true, context: 0 }).split('\n')[0].trim().split(/\s+/)).toEqual(['id', 'b']);
    const numeric = diffTables(parseCsv('a,b\n1,2\n3,4'), parseCsv('a,b\n1,9\n3,4'), { keys: 'none' });
    expect(renderText(numeric, { changedColumnsOnly: true, context: 0 }).split('\n')[0].trim().split(/\s+/)).toEqual(['b']);
  });
});

describe('html: row and column labels', () => {
  const model = (g: number) => [',FY2024,FY2025', 'India,,', '  Revenue,8300,' + 8300 * (1 + g), '  Costs,4150,4565'].join('\n');
  const d = diffTables(parseCsv(model(0.1), { header: false }), parseCsv(model(0.12), { header: false }), { keys: 'none' });
  it('labels rows and columns, and folds the label column and the year row', () => {
    const html = renderHTML(d, { layout: 'unified', context: 9 });
    expect(html).toContain('<span class="sdl-sec">India › </span><span class="sdl-lab">Revenue</span>');
    expect(html).toContain('>FY2025</th>');
    expect(html).not.toContain('>A</span>'); // the label column is in the row headers now
    expect(html).not.toMatch(/<td[^>]*>FY2024<\/td>/); // and the year row in the column headers
  });
  it('shows numbers and letters when asked', () => {
    const html = renderHTML(d, { layout: 'unified', context: 9, rowHeaders: 'numbers', columnHeaders: 'letters' });
    expect(html).toContain('<th class="sdl-ln">old</th><th class="sdl-ln">new</th>');
    expect(html).toContain('<span class="sdl-letter">C</span></th>');
    expect(html).not.toContain('sdl-label');
    const both = renderHTML(d, { layout: 'split', context: 9, rowHeaders: 'both' });
    expect(both).toMatch(/<span class="sdl-lab">Revenue<\/span><span class="sdl-n">3<\/span>/);
  });
});
