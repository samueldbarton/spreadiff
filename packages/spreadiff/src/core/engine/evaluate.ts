/**
 * A formula evaluator, used to check cached values and to find which cells a
 * result really depends on. It reads cell values through a host (so callers
 * can overlay recalculated values) and records every cell it actually reads:
 * a lookup reads its key column and one result cell, an IF only the branch
 * it takes. Anything it can't model throws {@link Unsupported}, and the
 * caller falls back to static analysis.
 */
import { lookupName, structuredRange } from '../formula.js';
import type { TableDefinition } from '../types.js';
import { FUNCTIONS } from './functions.js';
import { lex, parseFormula, parseTokens, relativeKey, Unsupported, type Node, type RefRange } from './parse.js';
import { compare, ERR, isArray, isError, isRange, toNumber, toText, type ArrayVal, type RangeVal, type Scalar, type Value, XlError } from './values.js';

export { Unsupported } from './parse.js';

/** What the engine needs from a workbook. Coordinates are 0-based sheet coordinates. */
export interface EngineHost {
  /** A cell's value (dates as serial numbers), or null when empty. */
  get(sheet: string, row: number, col: number): Scalar;
  /** The used extent of a sheet (bounds for `A:A`), or null if there's no such sheet. */
  extent(sheet: string): { rows: number; cols: number } | null;
  /** Sheet names in workbook order (for 3-D references and `sheet` lookups). */
  sheets: string[];
  /** Workbook-scoped defined names. */
  names?: Record<string, string>;
  /** The names a formula on `sheet` sees, sheet-scoped ones included (see `scopedNames`). Defaults to `names`. */
  namesFor?(sheet: string): Record<string, string> | undefined;
  /** Excel tables, for structured references. */
  tables?: TableDefinition[];
  /** The workbook uses the 1904 date system. */
  date1904?: boolean;
  /** Array-formula blocks covering a cell, so each cell takes its element of the result. */
  arrayBlock?(sheet: string, row: number, col: number): { r0: number; c0: number } | null;
  /** All values of a rectangle, row-major (faster than `get` per cell). */
  getRange?(sheet: string, r0: number, c0: number, r1: number, c1: number): Scalar[];
  /**
   * Bumped whenever a value the host returns changes; lets the engine reuse
   * range reads (a lookup table read by thousands of VLOOKUPs). Omit it and
   * nothing is reused.
   */
  version?: number;
  /**
   * The cells changed since `version` (oldest first), so cached ranges can
   * be patched instead of re-read; null if the log doesn't go back that far.
   */
  changesSince?(version: number): { sheet: string; row: number; col: number }[] | null;
}

/** A rectangle the evaluation read. */
export interface Read {
  sheet: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

/** Context passed to function implementations. */
export interface Ctx {
  sheet: string;
  row: number;
  col: number;
  host: EngineHost;
  /** The engine, for shared caches. */
  engine: Engine;
  /** Evaluate an argument. */
  ev(n: Node): Value;
  /** Reduce to one value (implicit intersection for ranges); records the read. */
  scalar(v: Value): Scalar;
  /** Every value of a range or array, row-major; records the read. */
  values(v: Value): Scalar[];
  /** Rows × cols of a range or array (scalars are 1 × 1). */
  dims(v: Value): { rows: number; cols: number };
  /** One element of a range or array by 0-based position; records only that cell. */
  at(v: Value, r: number, c: number): Scalar;
  /** Record that a rectangle was read. */
  read(r: Read): void;
  /** A range as an array (records the read). */
  array(v: Value): ArrayVal;
}

const BINARY_OPS = new Set(['+', '-', '*', '/', '^', '&', '=', '<>', '<', '>', '<=', '>=']);

/** A formula ready to evaluate at one cell: a tree shared with its copies, and the shift to this cell. */
export interface Compiled {
  node: Node;
  dRow: number;
  dCol: number;
  /** The formula relative to its cell; copies of a formula share it. */
  shape: string;
}

interface CachedRange {
  sheet: string;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  values: Scalar[];
  version: number;
}

export class Engine {
  private parsed = new Map<string, Node | Unsupported>();
  /** Range values, patched as the host's values change. */
  private ranges = new Map<string, CachedRange>();
  /** Exact-match indexes over cached ranges (first position of each key). */
  readonly indexes = new WeakMap<Scalar[], Map<string, number>>();
  /** Trees by relative key: a formula copied down a column is parsed once. */
  private byKey = new Map<string, { node: Node; row: number; col: number } | Unsupported>();
  /** Per cell and text, the compiled tree and its shift. */
  private compiled = new Map<string, Compiled | Unsupported>();

  constructor(readonly host: EngineHost) {}

  /** Values of a range, row-major, reused across evaluations and patched when cells change. */
  rangeValues(sheet: string, r0: number, c0: number, r1: number, c1: number): Scalar[] {
    const host = this.host;
    const cells = (r1 - r0 + 1) * (c1 - c0 + 1);
    const read = (): Scalar[] => {
      if (host.getRange) return host.getRange(sheet, r0, c0, r1, c1);
      const out: Scalar[] = new Array(cells);
      let i = 0;
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) out[i++] = host.get(sheet, r, c);
      return out;
    };
    const version = host.version;
    if (version === undefined || cells < 16) return read();
    const k = `${sheet}\u0000${r0},${c0},${r1},${c1}`;
    const hit = this.ranges.get(k);
    if (hit) {
      if (hit.version === version) return hit.values;
      const changes = host.changesSince?.(hit.version);
      if (changes && changes.length < cells / 4) {
        const w = c1 - c0 + 1;
        let patched = false;
        for (const ch of changes) {
          if (ch.sheet !== sheet || ch.row < r0 || ch.row > r1 || ch.col < c0 || ch.col > c1) continue;
          hit.values[(ch.row - r0) * w + (ch.col - c0)] = host.get(sheet, ch.row, ch.col);
          patched = true;
        }
        if (patched) this.indexes.delete(hit.values);
        hit.version = version;
        return hit.values;
      }
    }
    const values = read();
    // Reading may have recalculated something (lazy evaluation): only keep a stable read.
    if (host.version === version) this.ranges.set(k, { sheet, r0, c0, r1, c1, values, version });
    return values;
  }

  /** Parse an absolute formula (names, INDIRECT text) with a cache; throws Unsupported. */
  parse(formula: string): Node {
    let p = this.parsed.get(formula);
    if (p === undefined) {
      try {
        p = parseFormula(formula);
      } catch (e) {
        p = e instanceof Unsupported ? e : new Unsupported(String(e));
      }
      this.parsed.set(formula, p);
    }
    if (p instanceof Unsupported) throw p;
    return p;
  }

  /** The relative key of a formula (shared by its copies), or null if it doesn't parse. */
  shapeOf(formula: string, row: number, col: number): string | null {
    try {
      return relativeKey(lex(formula), row, col);
    } catch {
      return null;
    }
  }

  /** The tree for a cell's formula, shared with copies of it; throws Unsupported. */
  compile(formula: string, row: number, col: number): Compiled {
    const ck = `${row},${col}\u0000${formula}`;
    let c = this.compiled.get(ck);
    if (c === undefined) {
      try {
        const toks = lex(formula);
        const key = relativeKey(toks, row, col);
        let hit = this.byKey.get(key);
        if (hit === undefined) {
          try {
            hit = { node: parseTokens(toks), row, col };
          } catch (e) {
            hit = e instanceof Unsupported ? e : new Unsupported(String(e));
          }
          this.byKey.set(key, hit);
        }
        c = hit instanceof Unsupported ? hit : { node: hit.node, dRow: row - hit.row, dCol: col - hit.col, shape: key };
      } catch (e) {
        c = e instanceof Unsupported ? e : new Unsupported(String(e));
      }
      this.compiled.set(ck, c);
    }
    if (c instanceof Unsupported) throw c;
    return c;
  }

  /**
   * Evaluate the formula of the cell at (sheet, row, col): its value and,
   * when `record`, the rectangles read. Throws Unsupported for anything not
   * modelled. `formula` may be text or the result of {@link compile}.
   */
  evaluate(
    formula: string | Compiled,
    sheet: string,
    row: number,
    col: number,
    record = true,
  ): { value: Scalar; reads: Read[] } {
    const reads: Read[] = [];
    const { node, dRow, dCol } = typeof formula === 'string' ? this.compile(formula, row, col) : formula;
    const ctx = new EvalContext(this, sheet, row, col, record ? reads : null, 0, dRow, dCol);
    const v = ctx.ev(node);
    let out: Scalar;
    if (isArray(v) || isRange(v)) {
      const block = this.host.arrayBlock?.(sheet, row, col);
      if (block) {
        const d = ctx.dims(v);
        const r = row - block.r0;
        const c = col - block.c0;
        // An array result fills its block; a smaller result is #N/A outside it (1-wide results repeat).
        const rr = d.rows === 1 ? 0 : r;
        const cc = d.cols === 1 ? 0 : c;
        out = rr < d.rows && cc < d.cols ? ctx.at(v, rr, cc) : ERR.na;
      } else out = ctx.scalar(v);
    } else out = v;
    // A formula pointing at an empty cell shows 0.
    return { value: out === null ? 0 : out, reads };
  }
}

/** One evaluation's state. A class, so evaluating a cell allocates one small object. */
class EvalContext implements Ctx {
  readonly host: EngineHost;
  constructor(
    readonly engine: Engine,
    readonly sheet: string,
    readonly row: number,
    readonly col: number,
    private readonly reads: Read[] | null,
    private readonly depth: number,
    private readonly dRow: number,
    private readonly dCol: number,
  ) {
    this.host = engine.host;
  }

  read(r: Read): void {
    this.reads?.push(r);
  }

  private rangeOf(ref: RefRange): RangeVal | XlError {
    const s = ref.sheet ?? this.sheet;
    const ext = this.host.extent(s);
    if (!ext) return ERR.ref;
    if (ref.sheetTo !== undefined) throw new Unsupported('3-D reference outside an aggregate');
    // A tree shared with a copied formula: shift the relative bounds to this cell.
    const rel = ref.rel ?? 0;
    let r0 = ref.r0 < 0 ? 0 : ref.r0 + (rel & 1 ? this.dRow : 0);
    let r1 = ref.r1 < 0 ? Math.max(0, ext.rows - 1) : ref.r1 + (rel & 4 ? this.dRow : 0);
    let c0 = ref.c0 < 0 ? 0 : ref.c0 + (rel & 2 ? this.dCol : 0);
    let c1 = ref.c1 < 0 ? Math.max(0, ext.cols - 1) : ref.c1 + (rel & 8 ? this.dCol : 0);
    if (r0 < 0 || c0 < 0) return ERR.ref;
    if (r0 > r1) [r0, r1] = [r1, r0];
    if (c0 > c1) [c0, c1] = [c1, c0];
    return { kind: 'range', sheet: s, r0, c0, r1, c1 };
  }

  /** A 3-D reference, as a single array of every sheet's cells (aggregates only). */
  private threeD(ref: RefRange): ArrayVal {
    const sheets = this.host.sheets;
    const a = sheets.findIndex((s) => s.toUpperCase() === ref.sheet!.toUpperCase());
    const b = sheets.findIndex((s) => s.toUpperCase() === ref.sheetTo!.toUpperCase());
    if (a < 0 || b < 0) throw new Unsupported('3-D reference to unknown sheet');
    const data: Scalar[] = [];
    for (let k = Math.min(a, b); k <= Math.max(a, b); k++) {
      const rg = this.rangeOf({ ...ref, sheet: sheets[k], sheetTo: undefined }); // keeps the relative shift
      if (isError(rg)) throw new Unsupported('3-D reference');
      data.push(...this.values(rg));
    }
    return { kind: 'array', rows: data.length, cols: 1, data };
  }

  ev(n: Node): Value {
    switch (n.t) {
      case 'num':
      case 'str':
      case 'bool':
        return n.v;
      case 'err':
        return new XlError(n.v);
      case 'missing':
        return null;
      case 'ref':
        return n.ref.sheetTo !== undefined ? this.threeD(n.ref) : this.rangeOf(n.ref);
      case 'name': {
        const target = lookupName(this.host.namesFor ? this.host.namesFor(this.sheet) : this.host.names, n.name);
        if (target === undefined) return ERR.name;
        if (this.depth > 8) throw new Unsupported('name nesting');
        // Names are workbook formulas (Inputs!$B$2, or a constant/expression), absolute.
        const sub = new EvalContext(this.engine, this.sheet, this.row, this.col, this.reads, this.depth + 1, 0, 0);
        return sub.ev(this.engine.parse(target));
      }
      case 'struct': {
        const r = structuredRange(n.table, n.spec, { tables: this.host.tables, row: this.row, sheets: this.host.sheets });
        if (!r) throw new Unsupported('structured reference');
        return this.rangeOf(r);
      }
      case 'array': {
        const rows = n.rows.length;
        const cols = n.rows[0].length;
        const data: Scalar[] = [];
        for (const r of n.rows) {
          if (r.length !== cols) return ERR.value;
          for (const x of r) data.push(this.scalar(this.ev(x)));
        }
        return { kind: 'array', rows, cols, data };
      }
      case 'unary': {
        const v = this.ev(n.arg);
        const op = n.op;
        const f = (x: Scalar): Scalar => {
          const num = toNumber(x);
          if (isError(num)) return num;
          return op === '-' ? -num : op === '%' ? num / 100 : num;
        };
        if (isRange(v) || isArray(v)) {
          const a = this.array(v);
          return { kind: 'array', rows: a.rows, cols: a.cols, data: a.data.map(f) };
        }
        return f(v);
      }
      case 'binary': {
        if (n.op === ':') {
          const a = this.ev(n.l);
          const b = this.ev(n.r);
          if (!isRange(a) || !isRange(b) || a.sheet !== b.sheet) return ERR.value;
          return { kind: 'range', sheet: a.sheet, r0: Math.min(a.r0, b.r0), c0: Math.min(a.c0, b.c0), r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1) };
        }
        if (!BINARY_OPS.has(n.op)) throw new Unsupported(`operator ${n.op}`);
        const a = this.ev(n.l);
        const b = this.ev(n.r);
        if (isRange(a) || isArray(a) || isRange(b) || isArray(b)) return this.elementwise(n.op, a, b);
        return scalarOp(n.op, a as Scalar, b as Scalar);
      }
      case 'call': {
        const fn = FUNCTIONS[n.name];
        if (!fn) throw new Unsupported(`function ${n.name}`);
        return fn(n.args, this);
      }
    }
  }

  scalar(v: Value): Scalar {
    if (isRange(v)) {
      if (v.r0 === v.r1 && v.c0 === v.c1) {
        this.reads?.push({ sheet: v.sheet, r0: v.r0, c0: v.c0, r1: v.r0, c1: v.c0 });
        return this.host.get(v.sheet, v.r0, v.c0);
      }
      // Implicit intersection with the formula's row or column.
      if (v.sheet === this.sheet && v.c0 === v.c1 && this.row >= v.r0 && this.row <= v.r1) return this.at(v, this.row - v.r0, 0);
      if (v.sheet === this.sheet && v.r0 === v.r1 && this.col >= v.c0 && this.col <= v.c1) return this.at(v, 0, this.col - v.c0);
      return ERR.value;
    }
    if (isArray(v)) return v.data[0] ?? null;
    return v;
  }

  values(v: Value): Scalar[] {
    if (isRange(v)) {
      this.reads?.push({ sheet: v.sheet, r0: v.r0, c0: v.c0, r1: v.r1, c1: v.c1 });
      return this.engine.rangeValues(v.sheet, v.r0, v.c0, v.r1, v.c1);
    }
    if (isArray(v)) return v.data;
    return [v];
  }

  dims(v: Value): { rows: number; cols: number } {
    if (isRange(v)) return { rows: v.r1 - v.r0 + 1, cols: v.c1 - v.c0 + 1 };
    if (isArray(v)) return { rows: v.rows, cols: v.cols };
    return { rows: 1, cols: 1 };
  }

  at(v: Value, r: number, c: number): Scalar {
    if (isRange(v)) {
      const rr = v.r0 + r;
      const cc = v.c0 + c;
      if (rr > v.r1 || cc > v.c1) return ERR.ref;
      this.reads?.push({ sheet: v.sheet, r0: rr, c0: cc, r1: rr, c1: cc });
      return this.host.get(v.sheet, rr, cc);
    }
    if (isArray(v)) return r < v.rows && c < v.cols ? v.data[r * v.cols + c] : ERR.na;
    return r === 0 && c === 0 ? v : ERR.na;
  }

  array(v: Value): ArrayVal {
    if (isArray(v)) return v;
    if (isRange(v)) {
      const d = this.dims(v);
      if (d.rows * d.cols > 5e6) throw new Unsupported('range too large');
      return { kind: 'array', rows: d.rows, cols: d.cols, data: this.values(v) };
    }
    return { kind: 'array', rows: 1, cols: 1, data: [v] };
  }

  /** Array arithmetic: broadcast 1-row / 1-column / scalar operands, #N/A outside the overlap. */
  private elementwise(op: string, a: Value, b: Value): ArrayVal {
    const x = this.array(a);
    const y = this.array(b);
    const rows = Math.max(x.rows, y.rows);
    const cols = Math.max(x.cols, y.cols);
    if (rows * cols > 5e6) throw new Unsupported('array too large');
    const pick = (m: ArrayVal, r: number, c: number): Scalar => {
      const rr = m.rows === 1 ? 0 : r;
      const cc = m.cols === 1 ? 0 : c;
      return rr < m.rows && cc < m.cols ? m.data[rr * m.cols + cc] : ERR.na;
    };
    const data: Scalar[] = new Array(rows * cols);
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) data[r * cols + c] = scalarOp(op, pick(x, r, c), pick(y, r, c));
    return { kind: 'array', rows, cols, data };
  }
}

export function scalarOp(op: string, a: Scalar, b: Scalar): Scalar {
  switch (op) {
    case '&': {
      const x = toText(a);
      if (isError(x)) return x;
      const y = toText(b);
      if (isError(y)) return y;
      return x + y;
    }
    case '=':
    case '<>':
    case '<':
    case '>':
    case '<=':
    case '>=': {
      if (isError(a)) return a;
      if (isError(b)) return b;
      const c = compare(a, b);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
  }
  const x = toNumber(a);
  if (isError(x)) return x;
  const y = toNumber(b);
  if (isError(y)) return y;
  switch (op) {
    case '+':
      return x + y;
    case '-':
      return x - y;
    case '*':
      return x * y;
    case '/':
      return y === 0 ? ERR.div0 : x / y;
    case '^': {
      if (x === 0 && y < 0) return ERR.div0;
      const r = Math.pow(x, y);
      return Number.isFinite(r) ? r : ERR.num;
    }
  }
  throw new Unsupported(`operator ${op}`);
}
