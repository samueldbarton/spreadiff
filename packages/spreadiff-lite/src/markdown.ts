import { formatCell, type CellValue } from 'spreadiff/core';
import type { CellChange, ChangeReport, StructureChange } from './report.js';

/** Options for {@link toMarkdown}. */
export interface MarkdownOptions {
  /** Max rows per table; the rest are counted in a line under it. Default 50. `Infinity` for all. */
  maxRows?: number;
  /** Max causes listed per downstream result. Default 3. */
  maxCauses?: number;
  /** Heading level of the title (`##` = 2). Sections are one level below. Default 2. */
  headingLevel?: number;
  /** Significant digits for values and changes. Default 15: exact, as the file stores them. Contributions show at most 4. */
  digits?: number;
  /** List intermediate results (downstream cells that other changed formulas read) as well as final outputs. Default true. */
  intermediate?: boolean;
}

const PIPE = /\|/g;
const cellText = (s: string) => s.replace(/\r?\n/g, ' ').replace(PIPE, '\\|');

/**
 * The report as Markdown: a one-line summary, then tables of edits and
 * downstream results (each with what drove it), shape changes and notes.
 * Sections with nothing in them are left out.
 *
 * ```ts
 * import { compare, toMarkdown } from 'spreadiff-lite';
 * declare const report: ReturnType<typeof compare>;
 * const md = toMarkdown(report, { maxRows: 20 });
 * ```
 */
export function toMarkdown(report: ChangeReport, options: MarkdownOptions = {}): string {
  const maxRows = options.maxRows ?? 50;
  const maxCauses = options.maxCauses ?? 3;
  const level = Math.min(5, Math.max(1, options.headingLevel ?? 2));
  const digits = options.digits ?? 15;
  const h = (n: number, text: string) => `${'#'.repeat(level + n)} ${text}`;
  const num = (v: number) => {
    if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
    return String(Number(v.toPrecision(Math.min(15, Math.max(1, digits)))));
  };
  const val = (v: CellValue) => {
    if (v === null || v === '') return '*(empty)*';
    return cellText(typeof v === 'number' ? num(v) : formatCell(v));
  };
  const signed = (v: number) => (v > 0 ? `+${num(v)}` : v < 0 ? `−${num(-v)}` : '0');
  const pct = (r: number) => {
    const p = r * 100;
    const s = Math.abs(p) >= 10 ? p.toFixed(1) : Math.abs(p) >= 1 ? p.toFixed(2) : Number(p.toPrecision(2)).toString();
    return `${p > 0 ? '+' : p < 0 ? '−' : ''}${s.replace('-', '')}%`;
  };
  // Contributions are an attribution, not a stored value: 4 significant digits read better.
  const share = (v: number) => {
    const r = Number(v.toPrecision(Math.min(4, digits)));
    return r > 0 ? `+${r}` : r < 0 ? `−${-r}` : '0';
  };
  const change = (c: CellChange) => (c.delta === undefined ? '' : c.ratio === undefined ? signed(c.delta) : `${signed(c.delta)} (${pct(c.ratio)})`);
  const code = (f: string | null) => (f ? `\`${cellText(f).replace(/`/g, "'")}\`` : '*(value)*');
  const what = (c: CellChange) => (c.label ? cellText(c.label) : '');

  const { totals: t } = report;
  const out: string[] = [h(0, `${cellText(report.old)} → ${cellText(report.new)}`), ''];

  // ---- Summary line -------------------------------------------------------
  const parts: string[] = [];
  const plural = (n: number, one: string, many = `${one}s`) => `${n.toLocaleString('en-US')} ${n === 1 ? one : many}`;
  if (t.inputs) parts.push(plural(t.inputs, 'input'));
  if (t.formulas) parts.push(plural(t.formulas, 'formula'));
  if (t.external) parts.push(plural(t.external, 'outside-data value'));
  const edits = report.edits.length;
  const shape = report.structure.length;
  const objects = report.objects.length;
  if (!edits && !t.downstream && !shape && !objects) {
    out.push('No changes.');
    return out.join('\n') + '\n';
  }
  let summary = edits ? `**${plural(edits, 'edit')}** (${parts.join(', ')})` : '**No cell edits**';
  if (t.downstream) {
    summary += ` moved **${plural(t.downstream, 'result')} downstream**`;
    if (t.outputs && t.outputs < t.downstream) summary += ` (${plural(t.outputs, 'final output')})`;
  } else if (edits) summary += '; nothing downstream moved';
  if (shape) summary += `${t.downstream || edits ? '; ' : ', '}${plural(shape, 'change')} to rows, columns or sheets`;
  if (objects) summary += `${t.downstream || edits || shape ? '; ' : ', '}${plural(objects, 'change')} to names, charts or pivot tables`;
  summary += '.';
  if (t.downstream) {
    summary +=
      t.recalculated === t.downstream
        ? ' Every cause was found by recalculating.'
        : t.recalculated
          ? ` Causes of ${plural(t.recalculated, 'result')} were found by recalculating; the rest by following formula references.`
          : ' Causes were found by following formula references (the formulas on the path couldn\'t be recalculated).';
  }
  out.push(summary, '');

  const table = (head: string[], rows: string[][], total: number, more: string) => {
    out.push(`| ${head.join(' | ')} |`, `| ${head.map(() => '---').join(' | ')} |`);
    for (const r of rows) out.push(`| ${r.join(' | ')} |`);
    if (total > rows.length) out.push('', `…and ${plural(total - rows.length, more)}.`);
    out.push('');
  };

  // ---- Edits ----------------------------------------------------------------
  if (edits) {
    out.push(h(1, 'Edits'), '');
    const rows = report.edits.slice(0, maxRows).map((c) => {
      const before = c.kind === 'formula' ? code(c.formula?.before ?? null) : val(c.before);
      const after = c.kind === 'formula' ? code(c.formula?.after ?? null) : val(c.after);
      const note = c.kind === 'external' ? ` *(${cellText(c.external!.provider)} ${cellText(c.external!.function)})*` : c.kind === 'input' && c.formula?.before ? ' *(was a formula)*' : '';
      const result = c.kind === 'formula' && (c.before !== null || c.after !== null) ? `${val(c.before)} → ${val(c.after)}` : change(c);
      return [`\`${cellText(c.id)}\``, what(c) + note, before, after, result, c.impact ? String(c.impact) : ''];
    });
    table(['Cell', 'What', 'Before', 'After', 'Change', 'Moves'], rows, edits, 'more edit');
  }

  // ---- Downstream ---------------------------------------------------------
  const downstreamTable = (title: string, list: CellChange[], more: string) => {
    if (!list.length) return;
    out.push(h(1, title), '');
    const rows = list.slice(0, maxRows).map((c) => {
      const causes = (c.causes ?? []).slice(0, maxCauses).map((x) => {
        const name = x.label || x.id;
        return x.contribution === null ? cellText(name) : `${cellText(name)} ${share(x.contribution)}`;
      });
      const extra = (c.causeCount ?? 0) - causes.length;
      // "; " between causes: " · " already separates the parts of a label.
      let driven = causes.join('; ') + (extra > 0 ? `; ${plural(extra, 'more', 'more')}` : '');
      if (!c.causeCount) driven = c.unexplained === 'volatile' ? '*volatile function*' : c.unexplained === 'dynamic' ? '*dynamic reference*' : '*no edit found*';
      else if (c.method === 'references') driven += ' *(by references)*';
      const after = val(c.after) + (c.stale ? ` *(file saved ${val(c.stale.saved)})*` : '');
      return [`\`${cellText(c.id)}\``, what(c), val(c.before), after, change(c), driven];
    });
    table(['Cell', 'What', 'Before', 'After', 'Change', 'Driven by'], rows, list.length, more);
  };
  const split = t.outputs > 0 && t.outputs < t.downstream;
  if (!split) downstreamTable('Downstream', report.downstream, 'more result');
  else {
    downstreamTable('Final outputs', report.downstream.filter((c) => c.output), 'more output');
    if (options.intermediate !== false) downstreamTable('Intermediate results', report.downstream.filter((c) => !c.output), 'more intermediate result');
  }

  // ---- Structure ------------------------------------------------------------
  if (shape) {
    out.push(h(1, 'Rows, columns and sheets'), '');
    for (const s of report.structure.slice(0, maxRows)) out.push(`- ${structureLine(s)}`);
    if (shape > maxRows) out.push(`- …and ${plural(shape - maxRows, 'more')}.`);
    out.push('');
  }

  // ---- Names, charts and pivot tables ---------------------------------------
  if (objects) {
    out.push(h(1, 'Names, charts and pivot tables'), '');
    const noun = { name: 'Name', chart: 'Chart', pivot: 'Pivot table' };
    for (const o of report.objects.slice(0, maxRows)) {
      const details = o.details.filter(Boolean).map((d) => cellText(d)).join('; ');
      out.push(`- ${noun[o.kind]} \`${cellText(o.name)}\` ${o.change}${details ? `: ${details}` : ''}`);
    }
    if (objects > maxRows) out.push(`- …and ${plural(objects - maxRows, 'more')}.`);
    out.push('');
  }

  // ---- Notes ----------------------------------------------------------------
  if (report.notes.length) {
    out.push(h(1, 'Notes'), '');
    for (const n of report.notes) out.push(`- ${n}`);
    out.push('');
  }
  return out.join('\n').replace(/\n+$/, '\n');
}

function structureLine(s: StructureChange): string {
  const sheet = `**${cellText(s.sheet)}**`;
  const moves = s.impact ? `; moves ${s.impact} result${s.impact === 1 ? '' : 's'}` : '';
  const rows = () => {
    const r = s.rows ?? [];
    const span = r.length > 1 && r[r.length - 1] - r[0] === r.length - 1 ? `rows ${r[0]}–${r[r.length - 1]}` : r.length === 1 ? `row ${r[0]}` : `rows ${r.slice(0, 8).join(', ')}${r.length > 8 ? ', …' : ''}`;
    const labels = (s.labels ?? []).filter(Boolean);
    const named = labels.length ? ` (${labels.slice(0, 5).map(cellText).join(', ')}${labels.length > 5 ? ', …' : ''})` : '';
    return span + named;
  };
  switch (s.kind) {
    case 'sheet-added':
      return `${sheet}: sheet added`;
    case 'sheet-removed':
      return `${sheet}: sheet removed`;
    case 'rows-added':
      return `${sheet}: ${rows()} added${moves}`;
    case 'rows-removed':
      return `${sheet}: ${rows()} removed${moves}`;
    case 'rows-moved':
      return `${sheet}: ${rows()} moved`;
    case 'column-added':
      return `${sheet}: column *${cellText(s.column ?? '')}* added${moves}`;
    case 'column-removed':
      return `${sheet}: column *${cellText(s.column ?? '')}* removed${moves}`;
    case 'column-renamed':
      return `${sheet}: column renamed *${cellText(s.column ?? '')}*`;
  }
}
