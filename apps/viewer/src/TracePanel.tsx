import { useMemo, useState } from 'react';
import { columnLabel, formatCell, numericDelta, ROW_STATUS, type CellRef, type CellTrace, type WorkbookDiff, type WorkbookTrace } from 'spreadiff';

export type TraceView = 'outputs' | 'impact' | 'why';

interface Described {
  ref: CellRef;
  sheet: string;
  addr: string;
  name: string;
  before: string;
  after: string;
  delta: { text: string; up: boolean } | null;
  kind: CellTrace['kind'] | 'equal';
}

/** Everything a row in the panel needs about one cell. */
export function describe(wb: WorkbookDiff, trace: WorkbookTrace, ref: CellRef): Described {
  const sd = wb.sheets[ref.sheet];
  const d = sd.diff;
  const c = d.columns[ref.col];
  const ni = d.newRows[ref.row];
  const oi = d.oldRows[ref.row];
  const useNew = ni >= 0 && c.newIndex >= 0;
  const t = useNew ? d.newTable : d.oldTable;
  const ri = useNew ? ni : oi;
  const ci = useNew ? c.newIndex : c.oldIndex;
  const o = t.origin ?? { row: 1, col: 0 };
  const ov = oi >= 0 && c.oldIndex >= 0 ? d.oldTable.rows[oi][c.oldIndex] ?? null : null;
  const nv = ni >= 0 && c.newIndex >= 0 ? d.newTable.rows[ni][c.newIndex] ?? null : null;
  const label = t.rows[ri]?.[0];
  const dl = numericDelta(ov, nv);
  let delta: Described['delta'] = null;
  if (dl && dl.delta !== 0) {
    const up = dl.delta > 0;
    const s = up ? '+' : '−';
    const pct = dl.ratio === null ? '' : ` (${s}${Math.abs(dl.ratio * 100).toFixed(Math.abs(dl.ratio) < 0.1 ? 2 : 1)}%)`;
    delta = { text: `${s}${formatCell(Math.abs(dl.delta))}${pct}`, up };
  }
  const st = d.status[ref.row];
  return {
    ref,
    sheet: sd.name,
    addr: `${columnLabel(ci + o.col)}${ri + o.row + 1}`,
    name: `${label !== null && label !== undefined ? formatCell(label) + ' · ' : ''}${c.name}`,
    before: st === ROW_STATUS.added ? '' : formatCell(ov),
    after: st === ROW_STATUS.removed ? '' : formatCell(nv),
    delta,
    kind: trace.get(ref)?.kind ?? (st === ROW_STATUS.added ? 'added' : st === ROW_STATUS.removed ? 'removed' : 'equal'),
  };
}

const KIND_LABEL: Record<string, string> = {
  input: 'input',
  formula: 'formula',
  computed: 'downstream',
  added: 'row added',
  removed: 'row removed',
  equal: '',
};

function Kind({ kind }: { kind: string }) {
  return KIND_LABEL[kind] ? <span className={`kchip k-${kind}`}>{KIND_LABEL[kind]}</span> : null;
}

function Value({ d, showOld = true }: { d: Described; showOld?: boolean }) {
  return (
    <span className="tval">
      {showOld && d.before !== '' && d.kind !== 'added' && (
        <>
          <s>{d.before}</s>
          <span className="tarrow">→</span>
        </>
      )}
      <b>{d.after || '∅'}</b>
      {d.delta && <span className={`tdelta ${d.delta.up ? 'up' : 'down'}`}>{d.delta.text}</span>}
    </span>
  );
}

function Addr({ d, current, onPick }: { d: Described; current: number; onPick: (r: CellRef) => void }) {
  return (
    <button className="taddr" onClick={() => onPick(d.ref)} title="Show in grid">
      <span>
        {d.ref.sheet !== current && <span className="tsheet">{d.sheet}!</span>}
        {d.addr}
      </span>
      <span className="tname">{d.name}</span>
    </button>
  );
}

interface Props {
  wb: WorkbookDiff;
  trace: WorkbookTrace;
  sheet: number;
  selected: CellRef | null;
  view: TraceView;
  onView: (v: TraceView) => void;
  onPick: (ref: CellRef) => void;
}

export function TracePanel({ wb, trace, sheet, selected, view, onView, onPick }: Props) {
  // Root edits (inputs, formula edits, structural), biggest blast radius first.
  const roots = useMemo(() => {
    const out: { ref: CellRef; t: CellTrace }[] = [];
    wb.sheets.forEach((_, si) => {
      const cols = wb.sheets[si].diff.columns.length;
      for (const [k, t] of trace.sheet(si)) {
        if (t.kind !== 'computed') out.push({ ref: { sheet: si, row: Math.floor(k / cols), col: k % cols }, t });
      }
    });
    return out.sort((a, b) => b.t.impactCount - a.t.impactCount);
  }, [wb, trace]);

  // End results: changed computed cells that no other changed cell reads.
  const outputs = useMemo(() => {
    const computed: CellRef[] = [];
    wb.sheets.forEach((s, si) => {
      const cols = s.diff.columns.length;
      for (const [k, t] of trace.sheet(si)) if (t.kind === 'computed') computed.push({ sheet: si, row: Math.floor(k / cols), col: k % cols });
    });
    if (computed.length > 5000) return null;
    const feeds = new Set<string>();
    for (const r of computed) for (const p of trace.parents(r)) feeds.add(`${p.sheet}:${p.row}:${p.col}`);
    return computed.filter((r) => !feeds.has(`${r.sheet}:${r.row}:${r.col}`));
  }, [wb, trace]);

  return (
    <aside className="trace-panel">
      <div className="tp-tabs" role="tablist">
        {(['outputs', 'impact', 'why'] as const).map((v) => (
          <button key={v} role="tab" aria-selected={view === v} onClick={() => onView(v)}>
            {v === 'outputs' ? 'Outputs' : v === 'impact' ? 'Impact' : 'Why?'}
          </button>
        ))}
      </div>
      <div className="tp-body">
        {view === 'outputs' && <Outputs wb={wb} trace={trace} outputs={outputs} sheet={sheet} onPick={onPick} />}
        {view === 'impact' && <Impact wb={wb} trace={trace} roots={roots} sheet={sheet} onPick={onPick} />}
        {view === 'why' && <Why wb={wb} trace={trace} selected={selected} sheet={sheet} onPick={onPick} />}
      </div>
      <p className="tp-foot">
        {trace.stats.inputs} input edits · {trace.stats.formulas} formula edits · {trace.stats.computed} downstream
        {trace.stats.unexplained ? ` · ${trace.stats.unexplained} unexplained` : ''}
      </p>
    </aside>
  );
}

/** "What happened to the bottom line?" — end results, each with its drivers. */
function Outputs({ wb, trace, outputs, sheet, onPick }: { wb: WorkbookDiff; trace: WorkbookTrace; outputs: CellRef[] | null; sheet: number; onPick: (r: CellRef) => void }) {
  if (!outputs) return <p className="tp-empty">Too many downstream changes to summarize.</p>;
  if (!outputs.length) return <p className="tp-empty">No formula results changed.</p>;
  return (
    <ul className="tp-list">
      {outputs.map((r) => {
        const d = describe(wb, trace, r);
        const t = trace.get(r)!;
        return (
          <li key={`${r.sheet}:${r.row}:${r.col}`} className="tp-card">
            <div className="tp-line">
              <Addr d={d} current={sheet} onPick={onPick} />
            </div>
            <div className="tp-big">
              <Value d={d} />
            </div>
            <div className="tp-drivers">
              <span className="tp-muted">driven by</span>
              {t.causes.map((c) => {
                const cd = describe(wb, trace, c);
                return (
                  <button key={`${c.sheet}:${c.row}:${c.col}`} className={`driver k-${cd.kind}`} onClick={() => onPick(c)} title={`${cd.sheet}!${cd.addr} ${cd.before} → ${cd.after}`}>
                    {cd.name}
                    {cd.delta && <span className={`tdelta ${cd.delta.up ? 'up' : 'down'}`}>{cd.delta.text.replace(/ \(.*\)/, '')}</span>}
                    {!cd.delta && cd.kind === 'formula' && <span className="tdelta">ƒ</span>}
                    {!cd.delta && (cd.kind === 'added' || cd.kind === 'removed') && <span className="tdelta">{cd.kind === 'added' ? '+row' : '−row'}</span>}
                  </button>
                );
              })}
              {t.causeCount > t.causes.length && <span className="tp-muted">+{t.causeCount - t.causes.length}</span>}
              {t.unexplained && <span className="tp-muted">{t.unexplained === 'volatile' ? 'volatile function' : t.unexplained === 'dynamic' ? 'dynamic reference' : 'unknown cause'}</span>}
            </div>
          </li>
        );
      })}
    </ul>
  );
}

/** "What did each edit move?" — root edits, expandable to their downstream cells. */
function Impact({ wb, trace, roots, sheet, onPick }: { wb: WorkbookDiff; trace: WorkbookTrace; roots: { ref: CellRef; t: CellTrace }[]; sheet: number; onPick: (r: CellRef) => void }) {
  const [open, setOpen] = useState<string | null>(null);
  if (!roots.length) return <p className="tp-empty">No edits.</p>;
  return (
    <ul className="tp-list">
      {roots.slice(0, 300).map(({ ref, t }) => {
        const d = describe(wb, trace, ref);
        const key = `${ref.sheet}:${ref.row}:${ref.col}`;
        return (
          <li key={key} className="tp-card">
            <div className="tp-line">
              <Kind kind={t.kind} />
              <Addr d={d} current={sheet} onPick={onPick} />
            </div>
            <div className="tp-big">
              <Value d={d} />
            </div>
            <button className="tp-impacts" onClick={() => setOpen(open === key ? null : key)} disabled={!t.impactCount}>
              {t.impactCount ? `${open === key ? '▾' : '▸'} moved ${t.impactCount} downstream cell${t.impactCount === 1 ? '' : 's'}` : 'nothing downstream changed'}
            </button>
            {open === key && (
              <ul className="tp-sub">
                {t.impacts.map((i) => {
                  const id = describe(wb, trace, i);
                  return (
                    <li key={`${i.sheet}:${i.row}:${i.col}`}>
                      <Addr d={id} current={sheet} onPick={onPick} />
                      <Value d={id} showOld={false} />
                    </li>
                  );
                })}
              </ul>
            )}
          </li>
        );
      })}
    </ul>
  );
}

/** "Why did this cell change?" — the cause tree of the selected cell. */
function Why({ wb, trace, selected, sheet, onPick }: { wb: WorkbookDiff; trace: WorkbookTrace; selected: CellRef | null; sheet: number; onPick: (r: CellRef) => void }) {
  if (!selected) return <p className="tp-empty">Select a changed cell to see why it changed.</p>;
  const t = trace.get(selected);
  if (!t) return <p className="tp-empty">This cell didn't change.</p>;
  return (
    <div className="tp-tree">
      <Node wb={wb} trace={trace} r={selected} sheet={sheet} onPick={onPick} depth={0} seen={new Set()} />
    </div>
  );
}

function Node({ wb, trace, r, sheet, onPick, depth, seen }: { wb: WorkbookDiff; trace: WorkbookTrace; r: CellRef; sheet: number; onPick: (r: CellRef) => void; depth: number; seen: Set<string> }) {
  const d = describe(wb, trace, r);
  const key = `${r.sheet}:${r.row}:${r.col}`;
  const t = trace.get(r);
  // First two levels open; deeper ones on demand, so the root causes stay in view.
  const [open, setOpen] = useState(depth < 2);
  const kids = t?.kind === 'computed' && depth < 12 && !seen.has(key) ? trace.parents(r) : [];
  const next = new Set(seen).add(key);
  return (
    <div className="tnode">
      <div className="tnode-row">
        {kids.length > 0 ? (
          <button className="ttoggle" onClick={() => setOpen(!open)} aria-expanded={open}>{open ? '▾' : '▸'}</button>
        ) : (
          <span className="ttoggle" />
        )}
        <Kind kind={d.kind} />
        <Addr d={d} current={sheet} onPick={onPick} />
        <Value d={d} />
        {!open && kids.length > 0 && <span className="tp-muted">{kids.length} input{kids.length === 1 ? '' : 's'}</span>}
      </div>
      {open && kids.length > 0 && (
        <div className="tnode-kids">
          {kids.slice(0, 40).map((k) => (
            <Node key={`${k.sheet}:${k.row}:${k.col}`} wb={wb} trace={trace} r={k} sheet={sheet} onPick={onPick} depth={depth + 1} seen={next} />
          ))}
        </div>
      )}
    </div>
  );
}
