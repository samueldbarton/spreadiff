import { describe, expect, it } from 'vitest';
import { diffTables, parseCsv, renderHTML, renderText, LITE_CSS } from '../src/lite';
import type { Table } from '../src/core';

const csv = (s: string) => parseCsv(s.trim() + '\n');
const a = csv('id,name,price\n1,Lamp,10\n2,Chair,20\n3,Desk,30\n4,Mug,4\n5,Rug,50\n6,Vase,6\n7,Tray,7');
const b = csv('id,name,price\n1,Lamp,10\n2,Armchair,25\n3,Desk,30\n4,Mug,4\n5,Rug,50\n6,Vase,6\n8,Clock,12');

describe('lite', () => {
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

  it('ships its CSS', () => expect(LITE_CSS).toContain('.sdl-split'));
});

describe('changedColumnsOnly keeps what names the rows', () => {
  it('keeps the row-label column when rows were matched by order', async () => {
    const { diffTables, parseCsv, renderText } = await import('../src/lite');
    const d = diffTables(parseCsv('line,a,b\nRevenue,1,2\nCosts,3,4'), parseCsv('line,a,b\nRevenue,1,9\nCosts,3,4'), { keys: 'none' });
    expect(renderText(d, { changedColumnsOnly: true, context: 0 }).split('\n')[0].trim().split(/\s+/)).toEqual(['line', 'b']);
  });
  it('keeps only keys when there are keys, and no label column for numbers', async () => {
    const { diffTables, parseCsv, renderText } = await import('../src/lite');
    const keyed = diffTables(parseCsv('id,name,b\n1,x,2\n2,y,4'), parseCsv('id,name,b\n1,x,9\n2,y,4'), { keys: ['id'] });
    expect(renderText(keyed, { changedColumnsOnly: true, context: 0 }).split('\n')[0].trim().split(/\s+/)).toEqual(['id', 'b']);
    const numeric = diffTables(parseCsv('a,b\n1,2\n3,4'), parseCsv('a,b\n1,9\n3,4'), { keys: 'none' });
    expect(renderText(numeric, { changedColumnsOnly: true, context: 0 }).split('\n')[0].trim().split(/\s+/)).toEqual(['b']);
  });
});
