// Built by scripts/build-action.mjs from action/src/main.ts. Do not edit.

// action/src/main.ts
import { execFileSync } from "node:child_process";
import { appendFileSync, readFileSync } from "node:fs";

// packages/spreadiff/src/core/types.ts
var ROW_STATUS = { equal: 0, added: 1, removed: 2, modified: 3 };

// packages/spreadiff/src/core/formula.ts
var DYNAMIC_FUNCS = /* @__PURE__ */ new Set(["INDIRECT", "OFFSET", "TABLE"]);
var VOLATILE_FUNCS = /* @__PURE__ */ new Set(["NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "CELL", "INFO"]);
var MAX_ROW = 1048575;
var MAX_COL = 16383;
function colIndex(letters) {
  let c = 0;
  for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
  return c - 1;
}
var IDENT_CHAR = /[A-Za-z0-9_.\\À-￿]/;
var STRING = /"(?:[^"]|"")*"/y;
var QUOTED_SHEET = /'((?:[^']|'')+)'!/y;
var BARE_SHEET = /([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!/y;
var BARE_3D = /([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*):([A-Za-z_À-￿][A-Za-z0-9_.À-￿]*)!/y;
var EXTERNAL = /\[[^\]]*\]/y;
var ERROR = /#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)/y;
var CELL_RANGE = /\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/y;
var COL_RANGE = /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/y;
var ROW_RANGE = /\$?(\d+):\$?(\d+)/y;
var IDENT = /[A-Za-z_\\À-￿][A-Za-z0-9_.À-￿]*/y;
var NUMBER = /\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?/y;
function at(re, s, i) {
  re.lastIndex = i;
  return re.exec(s);
}
function tokenizeFormula(formula) {
  const f = formula.startsWith("=") ? formula.slice(1) : formula;
  const tokens = [];
  const names = [];
  const structured = [];
  let dynamic = false;
  let volatile = false;
  let text3 = "";
  const flush = () => {
    if (text3) tokens.push({ type: "text", text: text3 });
    text3 = "";
  };
  const boundaryBefore = (i2) => i2 === 0 || !IDENT_CHAR.test(f[i2 - 1]);
  const boundaryAfter = (j) => j >= f.length || !IDENT_CHAR.test(f[j]) && f[j] !== "(" && f[j] !== "!" && f[j] !== "[";
  const tryRef = (i2, sheet) => {
    let m = at(CELL_RANGE, f, i2);
    if (m && boundaryAfter(i2 + m[0].length)) {
      const r0 = Number(m[2]) - 1;
      const c0 = colIndex(m[1]);
      const r1 = m[4] ? Number(m[4]) - 1 : r0;
      const c1 = m[3] ? colIndex(m[3]) : c0;
      if (r0 <= MAX_ROW && r1 <= MAX_ROW && c0 <= MAX_COL && c1 <= MAX_COL) {
        return { end: i2 + m[0].length, range: { sheet, r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) } };
      }
    }
    m = at(COL_RANGE, f, i2);
    if (m && boundaryAfter(i2 + m[0].length)) {
      const a = colIndex(m[1]);
      const b = colIndex(m[2]);
      return { end: i2 + m[0].length, range: { sheet, r0: -1, r1: -1, c0: Math.min(a, b), c1: Math.max(a, b) } };
    }
    m = at(ROW_RANGE, f, i2);
    if (m && boundaryAfter(i2 + m[0].length)) {
      const a = Number(m[1]) - 1;
      const b = Number(m[2]) - 1;
      return { end: i2 + m[0].length, range: { sheet, r0: Math.min(a, b), r1: Math.max(a, b), c0: -1, c1: -1 } };
    }
    return null;
  };
  let i = 0;
  while (i < f.length) {
    let m;
    if (m = at(STRING, f, i)) {
      text3 += m[0];
      i += m[0].length;
      continue;
    }
    if (m = at(EXTERNAL, f, i)) {
      dynamic = true;
      text3 += m[0];
      i += m[0].length;
      continue;
    }
    if (m = at(ERROR, f, i)) {
      if (m[0] === "#REF!") dynamic = true;
      text3 += m[0];
      i += m[0].length;
      continue;
    }
    if (boundaryBefore(i)) {
      const q3 = at(BARE_3D, f, i);
      const q = q3 ?? at(QUOTED_SHEET, f, i) ?? at(BARE_SHEET, f, i);
      if (q) {
        const spec = q3 ? `${q3[1]}:${q3[2]}` : q[1].replace(/''/g, "'");
        const colon = spec.indexOf(":");
        const sheet = colon >= 0 ? spec.slice(0, colon) : spec;
        const ref2 = tryRef(i + q[0].length, sheet);
        if (ref2 && colon >= 0) ref2.range.sheetTo = spec.slice(colon + 1);
        if (ref2) {
          flush();
          tokens.push({ type: "ref", text: f.slice(i, ref2.end), range: ref2.range });
          i = ref2.end;
          continue;
        }
        const nm = at(IDENT, f, i + q[0].length);
        if (nm && f[i + q[0].length + nm[0].length] !== "(" && colon < 0) {
          const qualified = `${sheet}!${nm[0]}`;
          names.push(qualified);
          flush();
          tokens.push({ type: "text", text: `\0name:${qualified.toUpperCase()}\0` });
          i += q[0].length + nm[0].length;
          continue;
        }
        dynamic = true;
        text3 += q[0];
        i += q[0].length;
        continue;
      }
      const ref = tryRef(i, null);
      if (ref) {
        flush();
        tokens.push({ type: "ref", text: f.slice(i, ref.end), range: ref.range });
        i = ref.end;
        continue;
      }
      if (m = at(IDENT, f, i)) {
        const word = m[0];
        const end = i + word.length;
        const upper = word.toUpperCase();
        if ((upper === "TRUE" || upper === "FALSE") && f[end] === "(" && f[end + 1] === ")") {
          text3 += upper;
          i = end + 2;
          continue;
        }
        if (f[end] === "(") {
          const name = upper.replace(/^_XLFN\./, "");
          if (DYNAMIC_FUNCS.has(name)) dynamic = true;
          if (VOLATILE_FUNCS.has(name)) volatile = true;
          text3 += upper;
        } else if (f[end] === "[") {
          let j = end;
          let depth = 0;
          for (; j < f.length; j++) {
            const ch = f[j];
            if (ch === "'") j++;
            else if (ch === "[") depth++;
            else if (ch === "]" && --depth === 0) break;
          }
          if (depth !== 0) {
            dynamic = true;
            text3 += word;
          } else {
            structured.push({ table: word, spec: f.slice(end, j + 1) });
            text3 += f.slice(i, j + 1);
            i = j + 1;
            continue;
          }
        } else if (upper === "TRUE" || upper === "FALSE") {
          text3 += upper;
        } else {
          names.push(word);
          flush();
          tokens.push({ type: "text", text: `\0name:${upper}\0` });
        }
        i = end;
        continue;
      }
      if (m = at(NUMBER, f, i)) {
        text3 += m[0];
        i += m[0].length;
        continue;
      }
    }
    if (f[i] !== " " && f[i] !== "\n" && f[i] !== "\r" && f[i] !== "	") text3 += f[i];
    i++;
  }
  flush();
  return { tokens, dynamic, volatile, names, structured };
}
function formulaRefs(formula, names, context) {
  return refsOf(formula, names, context, 0);
}
function refsOf(formula, names, context, depth) {
  const t = tokenizeFormula(formula);
  const ranges = [];
  let dynamic = t.dynamic;
  let volatile = t.volatile;
  for (const tok of t.tokens) {
    if (tok.type !== "ref") continue;
    const rg = tok.range;
    if (rg.sheetTo === void 0) {
      ranges.push(rg);
      continue;
    }
    const order = context?.sheets?.map((x) => x.toUpperCase()) ?? [];
    const a = order.indexOf(rg.sheet.toUpperCase());
    const b = order.indexOf(rg.sheetTo.toUpperCase());
    if (a < 0 || b < 0) {
      dynamic = true;
      continue;
    }
    for (let k = Math.min(a, b); k <= Math.max(a, b); k++) ranges.push({ sheet: context.sheets[k], r0: rg.r0, c0: rg.c0, r1: rg.r1, c1: rg.c1 });
  }
  for (const n of t.names) {
    const target = lookupName(names, n);
    if (target === void 0 || depth > 8) {
      dynamic = true;
      continue;
    }
    const sub = refsOf(target, names, context, depth + 1);
    ranges.push(...sub.ranges);
    dynamic ||= sub.dynamic;
    volatile ||= sub.volatile;
  }
  for (const s of t.structured) {
    const r = structuredRange(s.table, s.spec, context);
    if (r) ranges.push(r);
    else dynamic = true;
  }
  return { ranges, dynamic, volatile };
}
function structuredRange(table, spec, context) {
  const def = context?.tables?.find((t) => t.name.toUpperCase() === table.toUpperCase());
  if (!def) return null;
  const inner = spec.slice(1, -1).trim();
  const items = [];
  if (inner.startsWith("[")) {
    let depth = 0;
    let cur = "";
    for (let i = 0; i < inner.length; i++) {
      const ch = inner[i];
      if (ch === "'") {
        cur += inner[++i] ?? "";
        continue;
      }
      if (ch === "[") {
        if (depth++ === 0) continue;
      } else if (ch === "]") {
        if (--depth === 0) {
          items.push(cur);
          cur = "";
          continue;
        }
      } else if (depth === 0) {
        if (ch === ":") items.push(":");
        continue;
      }
      cur += ch;
    }
  } else if (inner.startsWith("@")) {
    items.push("@");
    if (inner.length > 1) items.push(inner.slice(1).replace(/'(.)/g, "$1"));
  } else if (inner) {
    items.push(inner.replace(/'(.)/g, "$1"));
  }
  const { range, headerRows, totalsRows } = def;
  const dataTop = range.r0 + headerRows;
  const dataBottom = range.r1 - totalsRows;
  let r0 = dataTop;
  let r1 = dataBottom;
  const special = items.filter((x) => x.startsWith("#") || x === "@").map((x) => x.toUpperCase());
  if (special.length) {
    const has = (k) => special.includes(k);
    if (has("#ALL")) {
      r0 = range.r0;
      r1 = range.r1;
    } else if (has("#THIS ROW") || has("@")) {
      if (context?.row === void 0) return null;
      r0 = r1 = context.row;
    } else {
      r0 = has("#HEADERS") ? range.r0 : has("#DATA") ? dataTop : dataBottom + 1;
      r1 = has("#TOTALS") ? range.r1 : has("#DATA") ? dataBottom : range.r0 + headerRows - 1;
    }
  }
  const cols = items.filter((x) => !x.startsWith("#") && x !== "@");
  let c0 = range.c0;
  let c1 = range.c1;
  if (cols.length) {
    const idx = (name) => def.columns.findIndex((c) => c.toUpperCase() === name.replace(/^@/, "").toUpperCase());
    const names = cols.filter((x) => x !== ":");
    const a = idx(names[0]);
    const b = idx(names[names.length - 1]);
    if (a < 0 || b < 0) return null;
    c0 = range.c0 + Math.min(a, b);
    c1 = range.c0 + Math.max(a, b);
  }
  if (r0 > r1) return null;
  return { sheet: def.sheet, r0, r1, c0, c1 };
}
function scopedNames(workbookNames, sheets, sheet) {
  if (!sheets.some((s) => s.names)) return workbookNames;
  const out = { ...workbookNames };
  for (const s of sheets) {
    if (!s.names) continue;
    for (const [k, v] of Object.entries(s.names)) out[`${s.name}!${k}`] = qualify(v, s.name);
  }
  const own = sheets.find((s) => s.name === sheet)?.names;
  if (own) for (const [k, v] of Object.entries(own)) out[k] = qualify(v, sheet);
  return out;
}
function qualify(formula, sheet) {
  const prefix = /^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`;
  return formula.replace(/("[^"]*")|((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(!])/g, (all, str2, q, ref) => str2 || q ? all : `${prefix}!${ref}`);
}
function lookupName(names, name) {
  if (!names) return void 0;
  if (name in names) return names[name];
  const upper = name.toUpperCase();
  for (const k of Object.keys(names)) if (k.toUpperCase() === upper) return names[k];
  return void 0;
}
function canonicalFormula(formula, map) {
  const t = tokenizeFormula(formula);
  let out = "";
  for (const tok of t.tokens) {
    if (tok.type !== "ref") out += tok.text;
    else if (tok.range.sheetTo === void 0) out += `${map(tok.range)}`;
    else {
      const r = tok.range;
      out += `3d:${r.sheet.toUpperCase()}:${r.sheetTo.toUpperCase()}:${r.r0},${r.c0},${r.r1},${r.c1}`;
    }
  }
  return out;
}

// packages/spreadiff/src/core/sequence.ts
var WORK_BUDGET = 4e7;
function alignSequences(a, b) {
  const out = [];
  diffRange(a, b, 0, a.length, 0, b.length, out);
  return out;
}
function diffRange(a, b, aLo, aHi, bLo, bHi, out) {
  while (aLo < aHi && bLo < bHi && a[aLo] === b[bLo]) {
    out.push(aLo++, bLo++);
  }
  let aEnd = aHi;
  let bEnd = bHi;
  while (aEnd > aLo && bEnd > bLo && a[aEnd - 1] === b[bEnd - 1]) {
    aEnd--;
    bEnd--;
  }
  if (aLo < aEnd && bLo < bEnd) {
    const n = aEnd - aLo;
    const m = bEnd - bLo;
    const cap = Math.max(64, Math.floor(WORK_BUDGET / (n + m)));
    const split = bisect(a, b, aLo, n, bLo, m, cap);
    if (split === NO_COMMON) {
    } else if (split === OVER_BUDGET) {
      patience(a, b, aLo, aEnd, bLo, bEnd, out);
    } else {
      const x = split[0];
      const y = split[1];
      diffRange(a, b, aLo, aLo + x, bLo, bLo + y, out);
      diffRange(a, b, aLo + x, aEnd, bLo + y, bEnd, out);
    }
  }
  for (let i = 0; i < aHi - aEnd; i++) out.push(aEnd + i, bEnd + i);
}
var NO_COMMON = 0;
var OVER_BUDGET = 1;
function bisect(a, b, aLo, n, bLo, m, cap) {
  const maxD = Math.ceil((n + m) / 2);
  const dLimit = Math.min(maxD, cap);
  const vOffset = dLimit + 1;
  const vLength = 2 * dLimit + 2;
  const v1 = new Int32Array(vLength).fill(-1);
  const v2 = new Int32Array(vLength).fill(-1);
  v1[vOffset + 1] = 0;
  v2[vOffset + 1] = 0;
  const delta = n - m;
  const front = (delta & 1) !== 0;
  let k1start = 0;
  let k1end = 0;
  let k2start = 0;
  let k2end = 0;
  for (let d = 0; d < dLimit; d++) {
    for (let k1 = -d + k1start; k1 <= d - k1end; k1 += 2) {
      const k1Offset = vOffset + k1;
      let x1 = k1 === -d || k1 !== d && v1[k1Offset - 1] < v1[k1Offset + 1] ? v1[k1Offset + 1] : v1[k1Offset - 1] + 1;
      let y1 = x1 - k1;
      while (x1 < n && y1 < m && a[aLo + x1] === b[bLo + y1]) {
        x1++;
        y1++;
      }
      v1[k1Offset] = x1;
      if (x1 > n) k1end += 2;
      else if (y1 > m) k1start += 2;
      else if (front) {
        const k2Offset = vOffset + delta - k1;
        if (k2Offset >= 0 && k2Offset < vLength && v2[k2Offset] !== -1) {
          if (x1 >= n - v2[k2Offset]) return [x1, y1];
        }
      }
    }
    for (let k2 = -d + k2start; k2 <= d - k2end; k2 += 2) {
      const k2Offset = vOffset + k2;
      let x2 = k2 === -d || k2 !== d && v2[k2Offset - 1] < v2[k2Offset + 1] ? v2[k2Offset + 1] : v2[k2Offset - 1] + 1;
      let y2 = x2 - k2;
      while (x2 < n && y2 < m && a[aLo + n - x2 - 1] === b[bLo + m - y2 - 1]) {
        x2++;
        y2++;
      }
      v2[k2Offset] = x2;
      if (x2 > n) k2end += 2;
      else if (y2 > m) k2start += 2;
      else if (!front) {
        const k1Offset = vOffset + delta - k2;
        if (k1Offset >= 0 && k1Offset < vLength && v1[k1Offset] !== -1) {
          const x1 = v1[k1Offset];
          const y1 = x1 - (k1Offset - vOffset);
          if (x1 >= n - x2) return [x1, y1];
        }
      }
    }
  }
  return dLimit === maxD ? NO_COMMON : OVER_BUDGET;
}
function patience(a, b, aLo, aHi, bLo, bHi, out) {
  const seen = /* @__PURE__ */ new Map();
  for (let i = aLo; i < aHi; i++) {
    const e = seen.get(a[i]);
    if (e) e[0]++;
    else seen.set(a[i], [1, i, 0, -1]);
  }
  for (let j = bLo; j < bHi; j++) {
    const e = seen.get(b[j]);
    if (e) {
      e[2]++;
      e[3] = j;
    }
  }
  const candA = [];
  const candB = [];
  for (let i = aLo; i < aHi; i++) {
    const e = seen.get(a[i]);
    if (e[0] === 1 && e[2] === 1) {
      candA.push(i);
      candB.push(e[3]);
    }
  }
  if (candA.length === 0) return;
  const lis = longestIncreasing(candB);
  let pa = aLo;
  let pb = bLo;
  for (const idx of lis) {
    const ia = candA[idx];
    const ib = candB[idx];
    if (pa < ia && pb < ib) diffGapNoPatience(a, b, pa, ia, pb, ib, out);
    out.push(ia, ib);
    pa = ia + 1;
    pb = ib + 1;
  }
  if (pa < aHi && pb < bHi) diffGapNoPatience(a, b, pa, aHi, pb, bHi, out);
}
function diffGapNoPatience(a, b, aLo, aHi, bLo, bHi, out) {
  const n = aHi - aLo;
  const m = bHi - bLo;
  if ((n + m) * Math.min(n, m) <= WORK_BUDGET) diffRange(a, b, aLo, aHi, bLo, bHi, out);
}
function longestIncreasing(xs) {
  const tails = [];
  const prev = new Int32Array(xs.length);
  for (let i = 0; i < xs.length; i++) {
    let lo = 0;
    let hi = tails.length;
    while (lo < hi) {
      const mid = lo + hi >> 1;
      if (xs[tails[mid]] < xs[i]) lo = mid + 1;
      else hi = mid;
    }
    prev[i] = lo > 0 ? tails[lo - 1] : -1;
    tails[lo] = i;
  }
  const out = new Array(tails.length);
  let k = tails.length ? tails[tails.length - 1] : -1;
  for (let i = tails.length - 1; i >= 0; i--) {
    out[i] = k;
    k = prev[k];
  }
  return out;
}

// packages/spreadiff/src/core/xml.ts
function eachElement(xml, tag, fn) {
  const open = `<${tag}`;
  const close = `</${tag}>`;
  let i = xml.indexOf(open);
  while (i >= 0) {
    const next = xml.charCodeAt(i + open.length);
    if (next !== 62 && next !== 47 && next !== 32 && next !== 9 && next !== 10 && next !== 13) {
      i = xml.indexOf(open, i + 1);
      continue;
    }
    const gt = xml.indexOf(">", i);
    if (gt < 0) return;
    if (xml.charCodeAt(gt - 1) === 47) {
      fn(xml.slice(i + open.length, gt - 1), null);
      i = xml.indexOf(open, gt + 1);
      continue;
    }
    const end = xml.indexOf(close, gt + 1);
    if (end < 0) return;
    fn(xml.slice(i + open.length, gt), xml.slice(gt + 1, end));
    i = xml.indexOf(open, end + close.length);
  }
}
function textRuns(body) {
  let s = "";
  const clean = body.indexOf("<rPh") < 0 ? body : removeElements(body, "rPh");
  eachElement(clean, "t", (_, b) => {
    if (b) s += b;
  });
  return decode(s);
}
function removeElements(xml, tag) {
  let out = "";
  let pos = 0;
  const open = `<${tag}`;
  const close = `</${tag}>`;
  for (; ; ) {
    const i = xml.indexOf(open, pos);
    if (i < 0) break;
    const end = xml.indexOf(close, i);
    if (end < 0) break;
    out += xml.slice(pos, i);
    pos = end + close.length;
  }
  return out + xml.slice(pos);
}
function attr(attrs, name) {
  let i = attrs.indexOf(name);
  while (i >= 0) {
    const before = attrs.charCodeAt(i - 1);
    let j = i + name.length;
    if (before === 32 || before === 9 || before === 10 || before === 13) {
      while (attrs.charCodeAt(j) === 32) j++;
      if (attrs.charCodeAt(j) === 61) {
        j++;
        while (attrs.charCodeAt(j) === 32) j++;
        const q = attrs[j];
        if (q === '"' || q === "'") {
          const end = attrs.indexOf(q, j + 1);
          if (end >= 0) return attrs.slice(j + 1, end);
        }
      }
    }
    i = attrs.indexOf(name, i + 1);
  }
  return void 0;
}
var ENTITY = /&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g;
function decode(s) {
  if (s.indexOf("\r") >= 0) s = s.replace(/\r\n?/g, "\n");
  if (s.indexOf("_x") >= 0) s = s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
  if (s.indexOf("&") < 0) return s;
  return s.replace(ENTITY, (all, e) => {
    switch (e) {
      case "amp":
        return "&";
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      case "apos":
        return "'";
      default: {
        const cp = e[1] === "x" ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return cp >= 0 && cp <= 1114111 ? String.fromCodePoint(cp) : all;
      }
    }
  });
}
function decodeRef(ref) {
  let col2 = 0;
  let i = 0;
  for (; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c < 65 || c > 90) break;
    col2 = col2 * 26 + (c - 64);
  }
  return [Number(ref.slice(i)) - 1, col2 - 1];
}
function resolvePath(from, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = from.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
}

// packages/spreadiff/src/core/objects.ts
var strip = (xml) => xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, "<$1");
function refOrValue(body) {
  if (!body) return void 0;
  let out;
  eachElement(body, "f", (_, b) => out ??= b ? decode(b) : void 0);
  if (out === void 0) eachElement(body, "v", (_, b) => out ??= b ? decode(b) : void 0);
  return out;
}
function parseChart(xml, sheet, n) {
  const x = strip(xml);
  const types = [];
  for (const m of x.matchAll(/<(\w+?(?:3D)?)Chart[\s>]/g)) if (!types.includes(m[1])) types.push(m[1]);
  const head = x.slice(0, Math.max(0, x.indexOf("<plotArea")) || x.length);
  let title = "";
  eachElement(head, "title", (_, body) => {
    if (title || !body) return;
    const f = refOrValue(/<tx>([\s\S]*?)<\/tx>/.exec(body)?.[1] ?? null);
    let runs = "";
    eachElement(body, "t", (_a2, b) => runs += b ? decode(b) : "");
    title = runs || f || "";
  });
  const series = [];
  eachElement(x, "ser", (_, body) => {
    if (!body) return;
    const part = (tag) => {
      let v;
      eachElement(body, tag, (_a2, b) => v ??= refOrValue(b));
      return v;
    };
    const s = {};
    const name = part("tx");
    const categories = part("cat") ?? part("xVal");
    const values = part("val") ?? part("yVal");
    if (name !== void 0) s.name = name;
    if (categories !== void 0) s.categories = categories;
    if (values !== void 0) s.values = values;
    series.push(s);
  });
  return { sheet, title: title || `Chart ${n}`, types, series };
}
function parsePivot(xml, cacheXml, sheet) {
  const x = strip(xml);
  const c = strip(cacheXml);
  let name = "";
  eachElement(x, "pivotTableDefinition", (a) => name ||= decode(attr(a, "name") ?? ""));
  let location = "";
  eachElement(x, "location", (a) => location ||= attr(a, "ref") ?? "");
  const values = [];
  eachElement(x, "dataField", (a) => values.push(decode(attr(a, "name") ?? "")));
  let source = "external";
  eachElement(c, "worksheetSource", (a) => {
    const ref = attr(a, "ref");
    const sh = attr(a, "sheet");
    const nm = attr(a, "name");
    source = nm ? decode(nm) : `${sh ? decode(sh) + "!" : ""}${ref ?? ""}`;
  });
  return { sheet, name, location, source, values };
}
function parseObjects(text3, sheets) {
  const charts = [];
  const pivots = [];
  const relsOf2 = (path) => {
    const out = [];
    eachElement(strip(text3(path.replace(/([^/]+)$/, "_rels/$1.rels"))), "Relationship", (a) => {
      const target = attr(a, "Target");
      const type = attr(a, "Type") ?? "";
      if (target && attr(a, "TargetMode") !== "External") out.push({ type: type.slice(type.lastIndexOf("/") + 1), target: resolvePath(path, target) });
    });
    return out;
  };
  for (const { name, path } of sheets) {
    let n = 0;
    for (const r of relsOf2(path)) {
      if (r.type === "drawing") {
        for (const d of relsOf2(r.target)) if (d.type === "chart") charts.push(parseChart(text3(d.target), name, ++n));
      } else if (r.type === "pivotTable") {
        const cache2 = relsOf2(r.target).find((p) => p.type === "pivotCacheDefinition");
        pivots.push(parsePivot(text3(r.target), cache2 ? text3(cache2.target) : "", name));
      }
    }
  }
  return { charts, pivots };
}
function chartDetails(a, b) {
  const out = [];
  if (a.types.join() !== b.types.join()) out.push(`type: ${a.types.join(" + ") || "?"} \u2192 ${b.types.join(" + ") || "?"}`);
  const n = Math.max(a.series.length, b.series.length);
  for (let i = 0; i < n; i++) {
    const x = a.series[i];
    const y = b.series[i];
    const label = `series ${i + 1}${(y ?? x)?.name ? ` (${(y ?? x).name})` : ""}`;
    if (!x) out.push(`${label} added: ${y.values ?? ""}`);
    else if (!y) out.push(`${label} removed`);
    else
      for (const k of ["name", "categories", "values"]) {
        if (x[k] !== y[k]) out.push(`${label} ${k}: ${x[k] ?? "\u2014"} \u2192 ${y[k] ?? "\u2014"}`);
      }
  }
  return out;
}
function diffWorkbookObjects(a, b) {
  const out = [];
  const namesOf = (wb) => {
    const m = new Map(Object.entries(wb.names ?? {}));
    for (const t of wb.sheets) for (const [k, v] of Object.entries(t.names ?? {})) m.set(`${t.name}!${k}`, v);
    return m;
  };
  const na = namesOf(a);
  const nb = namesOf(b);
  for (const [k, v] of nb) {
    const old = na.get(k);
    if (old === void 0) out.push({ kind: "name", change: "added", name: k, details: [`= ${v}`] });
    else if (old !== v) out.push({ kind: "name", change: "modified", name: k, details: [`${old} \u2192 ${v}`] });
  }
  for (const [k, v] of na) if (!nb.has(k)) out.push({ kind: "name", change: "removed", name: k, details: [`was ${v}`] });
  const pair = (xs, ys, key, fallback) => {
    const pairs = [];
    const used = /* @__PURE__ */ new Set();
    const keyed = /* @__PURE__ */ new Map();
    ys.forEach((y, i) => keyed.set(key(y), [...keyed.get(key(y)) ?? [], i]));
    const unmatched = [];
    for (const x of xs) {
      const i = keyed.get(key(x))?.find((j) => !used.has(j));
      if (i === void 0) unmatched.push(x);
      else {
        used.add(i);
        pairs.push([x, ys[i]]);
      }
    }
    const rest = ys.map((y, i) => ({ y, i })).filter(({ i }) => !used.has(i));
    for (const x of unmatched) {
      const k = rest.findIndex(({ y }) => fallback(y, 0) === fallback(x, 0));
      if (k >= 0) {
        pairs.push([x, rest[k].y]);
        rest.splice(k, 1);
      } else pairs.push([x, void 0]);
    }
    for (const { y } of rest) pairs.push([void 0, y]);
    return pairs;
  };
  const ca = a.charts ?? [];
  const cb = b.charts ?? [];
  for (const [x, y] of pair(ca, cb, (c) => `${c.sheet}\0${c.title}`, (c) => c.sheet)) {
    const t = y ?? x;
    const name = `${t.sheet} \u203A ${t.title}`;
    if (!x) out.push({ kind: "chart", change: "added", name, details: [t.series.map((s) => s.values).filter(Boolean).join(", ")] });
    else if (!y) out.push({ kind: "chart", change: "removed", name, details: [x.series.map((s) => s.values).filter(Boolean).join(", ")] });
    else {
      const details = chartDetails(x, y);
      if (x.title !== y.title) details.unshift(`title: ${x.title} \u2192 ${y.title}`);
      if (details.length) out.push({ kind: "chart", change: "modified", name, details });
    }
  }
  const pa = a.pivots ?? [];
  const pb = b.pivots ?? [];
  for (const [x, y] of pair(pa, pb, (p) => `${p.sheet}\0${p.name}`, (p) => p.sheet)) {
    const t = y ?? x;
    const name = `${t.sheet} \u203A ${t.name}`;
    if (!x) out.push({ kind: "pivot", change: "added", name, details: [`of ${t.source}`] });
    else if (!y) out.push({ kind: "pivot", change: "removed", name, details: [`was of ${x.source}`] });
    else {
      const details = [];
      if (x.name !== y.name) details.push(`name: ${x.name} \u2192 ${y.name}`);
      if (x.source !== y.source) details.push(`source: ${x.source} \u2192 ${y.source}`);
      if (x.location !== y.location) details.push(`location: ${x.location} \u2192 ${y.location}`);
      if (x.values.join("\0") !== y.values.join("\0")) details.push(`values: ${x.values.join(", ") || "\u2014"} \u2192 ${y.values.join(", ") || "\u2014"}`);
      if (details.length) out.push({ kind: "pivot", change: "modified", name, details });
    }
  }
  return out;
}

// packages/spreadiff/src/core/diff.ts
var DEFAULTS = {
  keys: "auto",
  numericTolerance: 0,
  relativeTolerance: 0,
  precision: 15,
  trimWhitespace: true,
  ignoreCase: false,
  compareFormulas: true,
  rowSimilarity: 0.5,
  detectRenames: true
};
var PAIRING_WINDOW = 24;
function makeNormalizer(o) {
  return (v) => {
    if (v === null || v === void 0) return "";
    if (typeof v === "number") return String(v);
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    let s = v;
    if (o.trimWhitespace) s = s.trim();
    if (o.ignoreCase) s = s.toLowerCase();
    return s;
  };
}
function toNumber(v) {
  if (typeof v === "number") return v;
  if (typeof v === "string" && v.trim() !== "") return Number(v);
  return NaN;
}
function makeComparer(o) {
  const norm = makeNormalizer(o);
  const tol = o.numericTolerance;
  const rel = o.relativeTolerance;
  const digits = o.precision;
  return (a, b) => {
    if (a === b) return true;
    if (digits < 17 && typeof a === "number" && typeof b === "number" && sameDigits(a, b, digits)) return true;
    if (tol > 0 || rel > 0) {
      const x = toNumber(a);
      const y = toNumber(b);
      if (!Number.isNaN(x) && !Number.isNaN(y)) {
        const d = Math.abs(x - y);
        return d <= tol || d <= rel * Math.max(Math.abs(x), Math.abs(y));
      }
    }
    return norm(a) === norm(b);
  };
}
function relativeFormulaHash(f, row, col2) {
  h1 = 2166136261;
  h2 = 2654435769;
  const n = f.length;
  let i = 0;
  while (i < n) {
    const ch = f.charCodeAt(i);
    if (ch === 34) {
      const close = f.indexOf('"', i + 1);
      const end2 = close < 0 ? i + 1 : close + 1;
      for (; i < end2; i++) mix(f.charCodeAt(i));
      continue;
    }
    const cAbs = ch === 36;
    const j = cAbs ? i + 1 : i;
    let run2 = 0;
    while (j + run2 < n && isUpper(f.charCodeAt(j + run2))) run2++;
    let k = j + run2;
    const next = k < n ? f.charCodeAt(k) : 0;
    if (run2 === 0) {
      mix(ch);
      i++;
      continue;
    }
    if (next !== 36 && !isDigit(next)) {
      for (; i < k; i++) mix(f.charCodeAt(i));
      continue;
    }
    if (run2 > 3) {
      for (const stop = k - 3; i < stop; i++) mix(f.charCodeAt(i));
      continue;
    }
    const rAbs = next === 36;
    if (rAbs) k++;
    let end = k;
    let digits = 0;
    while (end < n && isDigit(f.charCodeAt(end))) digits = digits * 10 + f.charCodeAt(end++) - 48;
    if (end === k) {
      mix(ch);
      i++;
      continue;
    }
    let c = 0;
    for (let x = j; x < j + run2; x++) c = c * 26 + (f.charCodeAt(x) - 64);
    mix(65536 | (cAbs ? 1 : 0) | (rAbs ? 2 : 0));
    mix(cAbs ? c : c - 1 - col2);
    mix(rAbs ? digits : digits - 1 - row);
    if (rAbs) mix(end - k);
    i = end;
  }
  return (h1 >>> 0) * 2097152 + (h2 >>> 11);
}
var h1 = 0;
var h2 = 0;
function mix(x) {
  h1 = Math.imul(h1 ^ x, 16777619);
  h2 = Math.imul(h2 ^ x, 1540483477);
  h2 ^= h2 >>> 15;
}
var isUpper = (c) => c >= 65 && c <= 90;
var isDigit = (c) => c >= 48 && c <= 57;
function sameDigits(a, b, digits) {
  if (!Number.isFinite(a) || !Number.isFinite(b)) return false;
  return Math.abs(a - b) <= Math.max(Math.abs(a), Math.abs(b)) * 10 ** (1 - digits);
}
function alignTables(oldTable, newTable, o) {
  const norm = makeNormalizer(o);
  const equal = makeComparer(o);
  let columns = alignColumns(oldTable.columns, newTable.columns);
  const shared = columns.filter((c) => c.oldIndex >= 0 && c.newIndex >= 0);
  const resolved = resolveKeys(oldTable, newTable, shared, o, norm);
  const keyColumns = resolved.keys;
  const pairs = [];
  if (keyColumns.length) alignByKey(oldTable, newTable, columns, resolved, norm, pairs);
  else alignBySequence(oldTable, newTable, shared, norm, equal, o.rowSimilarity, pairs);
  if (o.detectRenames) columns = detectRenames(columns, oldTable, newTable, pairs, equal);
  return { oldTable, newTable, columns, pairs, keyColumns, o };
}
function formulaKeyer(sheets, self) {
  const axis = (map, survives) => {
    const n = map.length;
    const next = new Int32Array(n + 1).fill(-1);
    const prev = new Int32Array(n).fill(-1);
    for (let i = n - 1; i >= 0; i--) next[i] = survives(i) ? i : next[i + 1];
    for (let i = 0, last = -1; i < n; i++) prev[i] = last = survives(i) ? i : last;
    return { map, next, prev };
  };
  const maps = { old: [], new: [] };
  const byName = { old: /* @__PURE__ */ new Map(), new: /* @__PURE__ */ new Map() };
  sheets.forEach((a, idx) => {
    for (const side of ["old", "new"]) {
      const table = side === "old" ? a.oldTable : a.newTable;
      const own = side === "old" ? 0 : 1;
      const rows = new Int32Array(table.rows.length).fill(-1);
      const both = new Uint8Array(table.rows.length);
      for (let k = 0; k < a.pairs.length; k += 2) {
        const r = a.pairs[k + own];
        if (r < 0) continue;
        rows[r] = k / 2;
        both[r] = a.pairs[k + 1 - own] >= 0 ? 1 : 0;
      }
      const cols = new Int32Array(table.columns.length).fill(-1);
      const colBoth = new Uint8Array(table.columns.length);
      a.columns.forEach((c, ci) => {
        const j = side === "old" ? c.oldIndex : c.newIndex;
        if (j < 0) return;
        cols[j] = ci;
        colBoth[j] = c.oldIndex >= 0 && c.newIndex >= 0 ? 1 : 0;
      });
      maps[side][idx] = {
        rows: axis(rows, (i) => both[i] === 1),
        cols: axis(cols, (j) => colBoth[j] === 1),
        origin: table.origin ?? { row: 1, col: 0 }
      };
      byName[side].set(table.name.toUpperCase(), idx);
    }
  });
  const span = (v0, v1, origin, ax, tag) => {
    if (v0 < 0) return "*";
    const n = ax.map.length;
    const one = (i) => i < 0 ? `${tag}h${i}` : i >= n ? `${tag}e${i - n}` : `${tag}${ax.map[i]}`;
    let i0 = v0 - origin;
    let i1 = v1 - origin;
    if (i1 > i0) {
      const a = i0 >= 0 && i0 < n ? ax.next[i0] : -1;
      const b = i1 >= 0 && i1 < n ? ax.prev[i1] : -1;
      if (a >= 0 && a <= i1) i0 = a;
      if (b >= 0 && b >= i0) i1 = b;
    }
    return `${one(i0)}:${one(i1)}`;
  };
  return (side, formula) => canonicalFormula(formula, (rg) => {
    const idx = rg.sheet === null ? self : byName[side].get(rg.sheet.toUpperCase());
    if (idx === void 0) return `?${rg.sheet.toUpperCase()}!${rg.r0},${rg.c0}:${rg.r1},${rg.c1}`;
    const m = maps[side][idx];
    return `${idx}!${span(rg.r0, rg.r1, m.origin.row, m.rows, "r")}|${span(rg.c0, rg.c1, m.origin.col, m.cols, "c")}`;
  });
}
function compareAligned(a, keyer) {
  const { oldTable, newTable, columns, pairs, keyColumns, o } = a;
  const equal = makeComparer(o);
  const rowCount = pairs.length / 2;
  const oldRows = new Int32Array(rowCount);
  const newRows = new Int32Array(rowCount);
  const status = new Uint8Array(rowCount);
  const changedCells = /* @__PURE__ */ new Map();
  const formulaChanged = /* @__PURE__ */ new Map();
  const cmpCols = columns.map((c, i) => ({ c, i })).filter(({ c }) => c.oldIndex >= 0 && c.newIndex >= 0);
  const of = oldTable.formulas;
  const nf = newTable.formulas;
  const checkFormulas = o.compareFormulas && !!of && !!nf;
  const stats = {
    rowsAdded: 0,
    rowsRemoved: 0,
    rowsModified: 0,
    rowsEqual: 0,
    rowsMoved: 0,
    cellsChanged: 0,
    columnsAdded: 0,
    columnsRemoved: 0,
    columnsRenamed: 0
  };
  for (let r = 0; r < rowCount; r++) {
    const oi = pairs[2 * r];
    const ni = pairs[2 * r + 1];
    oldRows[r] = oi;
    newRows[r] = ni;
    if (oi < 0) {
      status[r] = ROW_STATUS.added;
      stats.rowsAdded++;
      continue;
    }
    if (ni < 0) {
      status[r] = ROW_STATUS.removed;
      stats.rowsRemoved++;
      continue;
    }
    const x = oldTable.rows[oi];
    const y = newTable.rows[ni];
    let mask = null;
    for (const { c, i } of cmpCols) {
      const sameValue = equal(x[c.oldIndex] ?? null, y[c.newIndex] ?? null);
      let formulaDiffers = false;
      if (checkFormulas) {
        const fo = of[oi]?.[c.oldIndex] ?? null;
        const fn = nf[ni]?.[c.newIndex] ?? null;
        if (fo !== fn) formulaDiffers = fo === null || fn === null || keyer("old", fo) !== keyer("new", fn);
      }
      if (!sameValue || formulaDiffers) {
        mask ??= new Uint8Array(columns.length);
        mask[i] = 1;
        c.changes++;
        stats.cellsChanged++;
        if (formulaDiffers) {
          let fm = formulaChanged.get(r);
          if (!fm) formulaChanged.set(r, fm = new Uint8Array(columns.length));
          fm[i] = 1;
        }
      }
    }
    if (mask) {
      status[r] = ROW_STATUS.modified;
      changedCells.set(r, mask);
      stats.rowsModified++;
    } else {
      status[r] = ROW_STATUS.equal;
      stats.rowsEqual++;
    }
  }
  const moved = new Uint8Array(rowCount);
  const matchedAt = [];
  const oldOrder = [];
  for (let r = 0; r < rowCount; r++) {
    if (oldRows[r] >= 0 && newRows[r] >= 0) {
      matchedAt.push(r);
      oldOrder.push(oldRows[r]);
    }
  }
  if (keyColumns.length && oldOrder.some((v, i) => i > 0 && v < oldOrder[i - 1])) {
    const keep = new Uint8Array(oldOrder.length);
    for (const i of longestIncreasing(oldOrder)) keep[i] = 1;
    for (let i = 0; i < oldOrder.length; i++) {
      if (!keep[i]) {
        moved[matchedAt[i]] = 1;
        stats.rowsMoved++;
      }
    }
  }
  for (const c of columns) {
    if (c.status === "added") stats.columnsAdded++;
    else if (c.status === "removed") stats.columnsRemoved++;
    else if (c.status === "renamed") stats.columnsRenamed++;
  }
  return { oldTable, newTable, columns, rowCount, oldRows, newRows, status, moved, changedCells, formulaChanged, keyColumns, stats, options: o };
}
function diffWorkbooks(oldBook, newBook, options = {}) {
  const o = { ...DEFAULTS, ...options };
  const empty = (name) => ({ name, columns: [], rows: [] });
  const oldByName = new Map(oldBook.sheets.map((s) => [s.name, s]));
  const newNames = new Set(newBook.sheets.map((s) => s.name));
  const positional = oldBook.sheets.length === 1 && newBook.sheets.length === 1;
  const plan = [];
  for (const ns of newBook.sheets) {
    const os = positional ? oldBook.sheets[0] : oldByName.get(ns.name);
    plan.push({ name: ns.name, status: os ? null : "added", old: os ?? empty(ns.name), new: ns });
  }
  if (!positional) {
    oldBook.sheets.forEach((os, i) => {
      if (newNames.has(os.name)) return;
      plan.splice(Math.min(i, plan.length), 0, { name: os.name, status: "removed", old: os, new: empty(os.name) });
    });
  }
  const aligned = plan.map((p) => alignTables(p.old, p.new, o));
  const sheets = plan.map((p, i) => {
    const d = compareAligned(aligned[i], formulaKeyer(aligned, i));
    return { name: p.name, status: p.status ?? (hasChanges(d) ? "modified" : "equal"), diff: d };
  });
  const objects = diffWorkbookObjects(oldBook, newBook);
  return {
    sheets,
    names: oldBook.names || newBook.names ? { old: oldBook.names, new: newBook.names } : void 0,
    tables: oldBook.tables || newBook.tables ? { old: oldBook.tables, new: newBook.tables } : void 0,
    date1904: oldBook.date1904 || newBook.date1904 ? { old: oldBook.date1904, new: newBook.date1904 } : void 0,
    computedValues: oldBook.computedValues || newBook.computedValues ? { old: oldBook.computedValues, new: newBook.computedValues } : void 0,
    objects: objects.length ? objects : void 0
  };
}
function hasChanges(d) {
  const s = d.stats;
  return s.rowsAdded + s.rowsRemoved + s.rowsModified + (s.rowsMoved ?? 0) + s.columnsAdded + s.columnsRemoved + s.columnsRenamed > 0;
}
function alignColumns(oldCols, newCols) {
  const oldIndex = /* @__PURE__ */ new Map();
  oldCols.forEach((n, i) => {
    const q = oldIndex.get(n);
    if (q) q.push(i);
    else oldIndex.set(n, [i]);
  });
  const used = /* @__PURE__ */ new Set();
  const out = newCols.map((name, ni) => {
    const oi = oldIndex.get(name)?.shift();
    if (oi !== void 0) {
      used.add(oi);
      return { oldIndex: oi, newIndex: ni, name, status: "equal", changes: 0 };
    }
    return { oldIndex: -1, newIndex: ni, name, status: "added", changes: 0 };
  });
  for (let oi = 0; oi < oldCols.length; oi++) {
    if (used.has(oi)) continue;
    let at3 = 0;
    for (let p = oi - 1; p >= 0; p--) {
      const pos = out.findIndex((c) => c.oldIndex === p);
      if (pos >= 0) {
        at3 = pos + 1;
        break;
      }
    }
    while (at3 < out.length && out[at3].status === "removed" && out[at3].oldIndex < oi) at3++;
    out.splice(at3, 0, { oldIndex: oi, newIndex: -1, name: oldCols[oi], status: "removed", changes: 0 });
  }
  return out;
}
var isKeyName = (name) => /^(id|key|uuid|guid|sku|code)$/i.test(name) || /[ _.-](id|key|code)$/i.test(name) || /[a-z0-9](Id|ID)$/.test(name);
function keyValue(v, norm) {
  if (typeof v === "number") return v;
  const s = norm(v);
  const c = s.charCodeAt(0);
  if (c >= 48 && c <= 57 || c === 45 || c === 73 || c === 78) {
    const n = Number(s);
    if (String(n) === s) return n;
  }
  return s;
}
function uniqueKeys(t, col2, norm) {
  const keys = new Array(t.rows.length);
  const set = /* @__PURE__ */ new Set();
  for (let r = 0; r < t.rows.length; r++) {
    const k = keyValue(t.rows[r][col2] ?? null, norm);
    if (k === "" || set.has(k)) return null;
    set.add(k);
    keys[r] = k;
  }
  return { keys, set };
}
function resolveKeys(oldTable, newTable, shared, o, norm) {
  if (o.keys === "none") return { keys: [] };
  if (Array.isArray(o.keys)) {
    const names = new Set(shared.map((c) => c.name));
    const missing2 = o.keys.filter((k) => !names.has(k));
    if (missing2.length) throw new Error(`Key column(s) not in both tables: ${missing2.join(", ")}`);
    return { keys: o.keys };
  }
  if (oldTable.rows.length < 2 || newTable.rows.length < 2) return { keys: [] };
  const candidates = shared.filter((c) => isKeyName(c.name)).sort((x, y) => x.newIndex - y.newIndex);
  for (const c of candidates.slice(0, 4)) {
    const a = uniqueKeys(oldTable, c.oldIndex, norm);
    if (!a) continue;
    const b = uniqueKeys(newTable, c.newIndex, norm);
    if (!b) continue;
    let hits = 0;
    for (const k of b.keys) if (a.set.has(k)) hits++;
    if (hits >= 0.5 * Math.min(oldTable.rows.length, newTable.rows.length)) {
      return { keys: [c.name], values: { old: a.keys, new: b.keys } };
    }
  }
  return { keys: [] };
}
function alignByKey(oldTable, newTable, columns, resolved, norm, pairs) {
  const keyCols = resolved.keys.map((k) => columns.find((c) => c.name === k));
  const keysOf = (t, side) => {
    if (keyCols.length === 1) {
      const col2 = keyCols[0][side];
      return t.rows.map((r) => keyValue(r[col2] ?? null, norm));
    }
    return t.rows.map((r) => {
      let k = "";
      for (const c of keyCols) {
        const v = norm(r[c[side]] ?? null);
        k += v.length + ":" + v;
      }
      return k;
    });
  };
  const oldKeys = resolved.values?.old ?? keysOf(oldTable, "oldIndex");
  const newKeys = resolved.values?.new ?? keysOf(newTable, "newIndex");
  const first = /* @__PURE__ */ new Map();
  const dups = /* @__PURE__ */ new Map();
  oldKeys.forEach((k, i) => {
    const f = first.get(k);
    if (f === void 0) first.set(k, i);
    else {
      const q = dups.get(k);
      if (q) q.push(i);
      else dups.set(k, [f, i]);
    }
  });
  const matchedOld = new Int32Array(newTable.rows.length).fill(-1);
  const oldUsed = new Uint8Array(oldTable.rows.length);
  const cursor = /* @__PURE__ */ new Map();
  for (let ni = 0; ni < newKeys.length; ni++) {
    const k = newKeys[ni];
    const q = dups.size ? dups.get(k) : void 0;
    let oi = -1;
    if (q) {
      const at3 = cursor.get(k) ?? 0;
      if (at3 < q.length) {
        oi = q[at3];
        cursor.set(k, at3 + 1);
      }
    } else {
      const f = first.get(k);
      if (f !== void 0 && !oldUsed[f]) oi = f;
    }
    if (oi >= 0) {
      matchedOld[ni] = oi;
      oldUsed[oi] = 1;
    }
  }
  let nextOld = 0;
  const flushRemovedBefore = (limit) => {
    for (; nextOld < limit; nextOld++) if (!oldUsed[nextOld]) pairs.push(nextOld, -1);
  };
  const upcoming = new Int32Array(newTable.rows.length + 1);
  upcoming[newTable.rows.length] = oldTable.rows.length;
  for (let ni = newTable.rows.length - 1; ni >= 0; ni--) {
    upcoming[ni] = matchedOld[ni] >= 0 ? matchedOld[ni] : upcoming[ni + 1];
  }
  for (let ni = 0; ni < newTable.rows.length; ni++) {
    const oi = matchedOld[ni];
    if (oi >= nextOld || oi < 0) flushRemovedBefore(upcoming[ni]);
    pairs.push(oi, ni);
  }
  flushRemovedBefore(oldTable.rows.length);
}
function alignBySequence(oldTable, newTable, shared, norm, equal, threshold, pairs) {
  const nv = shared.length;
  const width = oldTable.formulas || newTable.formulas ? nv * 2 : nv;
  const oldKeys = new Array(oldTable.rows.length * width);
  const newKeys = new Array(newTable.rows.length * width);
  const formulaKey = (t, r, k, side) => {
    const c = shared[k][side];
    const f = t.formulas?.[r]?.[c];
    if (!f) return -1;
    const origin = t.origin ?? { row: 1, col: 0 };
    return relativeFormulaHash(f, r + origin.row, c + origin.col);
  };
  const buckets = /* @__PURE__ */ new Map();
  const repAt = [];
  const f64 = new Float64Array(1);
  const u32 = new Uint32Array(f64.buffer);
  const intern = (table, side, keys) => {
    const rows = table.rows;
    const ids = new Int32Array(rows.length);
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      const at3 = r * width;
      let h = 2166136261;
      for (let c = 0; c < width; c++) {
        const k = c < nv ? keyValue(row[shared[c][side]] ?? null, norm) : formulaKey(table, r, c - nv, side);
        keys[at3 + c] = k;
        if (typeof k === "number") {
          f64[0] = k === 0 ? 0 : k !== k ? NaN : k;
          h = Math.imul(h ^ u32[0], 16777619);
          h = Math.imul(h ^ u32[1], 16777619);
        } else {
          h = Math.imul(h ^ 1540483477, 16777619);
          for (let i = 0; i < k.length; i++) h = Math.imul(h ^ k.charCodeAt(i), 16777619);
          h = Math.imul(h ^ k.length, 16777619);
        }
      }
      let bucket = buckets.get(h);
      let id = -1;
      if (bucket) {
        for (const cand of bucket) {
          const o = repAt[cand];
          const repKeys = o >= 0 ? oldKeys : newKeys;
          const repOff = o >= 0 ? o : ~o;
          let same2 = true;
          for (let c = 0; c < width && same2; c++) {
            const x = repKeys[repOff + c];
            const y = keys[at3 + c];
            same2 = x === y || x !== x && y !== y;
          }
          if (same2) {
            id = cand;
            break;
          }
        }
      } else buckets.set(h, bucket = []);
      if (id < 0) {
        id = repAt.length;
        repAt.push(side === "oldIndex" ? at3 : ~at3);
        bucket.push(id);
      }
      ids[r] = id;
    }
    return ids;
  };
  const a = intern(oldTable, "oldIndex", oldKeys);
  const b = intern(newTable, "newIndex", newKeys);
  buckets.clear();
  const matches = alignSequences(a, b);
  const of = oldTable.formulas;
  const nf = newTable.formulas;
  const similarity = (oi, ni) => {
    const x = oldTable.rows[oi];
    const y = newTable.rows[ni];
    let same2 = 0;
    let total = 0;
    for (const c of shared) {
      if (of?.[oi]?.[c.oldIndex] && nf?.[ni]?.[c.newIndex]) continue;
      total++;
      if (equal(x[c.oldIndex] ?? null, y[c.newIndex] ?? null)) same2++;
    }
    if (total === 0) {
      for (const c of shared) {
        total++;
        if (equal(x[c.oldIndex] ?? null, y[c.newIndex] ?? null)) same2++;
      }
    }
    return total ? same2 / total : 0;
  };
  const modelSheet = !!(of?.length || nf?.length);
  const sharesFormulas = (oi, ni) => {
    for (const c of shared) if (of?.[oi]?.[c.oldIndex] && nf?.[ni]?.[c.newIndex]) return true;
    return false;
  };
  const gap = (o0, o1, n0, n1) => {
    let nj = n0;
    const floor = o1 - o0 === 1 && n1 - n0 === 1 ? Math.min(threshold, Number.MIN_VALUE) : threshold;
    const sameSize = o1 - o0 === n1 - n0;
    for (let oi = o0; oi < o1; oi++) {
      const j = n0 + (oi - o0);
      const inPlace = sameSize && j >= nj;
      if (inPlace && (similarity(oi, j) >= threshold || sharesFormulas(oi, j))) {
        for (; nj < j; nj++) pairs.push(-1, nj);
        pairs.push(oi, j);
        nj = j + 1;
        continue;
      }
      let best = -1;
      let bestScore = floor;
      const end = Math.min(n1, nj + PAIRING_WINDOW);
      for (let j2 = nj; j2 < end; j2++) {
        const s = similarity(oi, j2);
        if (s >= bestScore && (best < 0 || s > bestScore)) {
          best = j2;
          bestScore = s;
          if (s === 1) break;
        }
      }
      if (best < 0 && inPlace && modelSheet) best = j;
      if (best < 0) {
        pairs.push(oi, -1);
        continue;
      }
      for (; nj < best; nj++) pairs.push(-1, nj);
      pairs.push(oi, best);
      nj = best + 1;
    }
    for (; nj < n1; nj++) pairs.push(-1, nj);
  };
  let po = 0;
  let pn = 0;
  for (let k = 0; k < matches.length; k += 2) {
    const oi = matches[k];
    const ni = matches[k + 1];
    if (po < oi || pn < ni) gap(po, oi, pn, ni);
    pairs.push(oi, ni);
    po = oi + 1;
    pn = ni + 1;
  }
  gap(po, oldTable.rows.length, pn, newTable.rows.length);
}
function detectRenames(columns, oldTable, newTable, pairs, equal) {
  const removed = columns.filter((c) => c.status === "removed");
  const added = columns.filter((c) => c.status === "added");
  if (!removed.length || !added.length) return columns;
  const matched = [];
  for (let k = 0; k < pairs.length; k += 2) if (pairs[k] >= 0 && pairs[k + 1] >= 0) matched.push(k);
  if (matched.length === 0) return columns;
  const step = Math.max(1, Math.floor(matched.length / 2e3));
  const isBlank2 = (v) => v === null || v === "";
  const score = (rc, ac) => {
    let same2 = 0;
    let total = 0;
    for (let s = 0; s < matched.length; s += step) {
      const k = matched[s];
      const x = oldTable.rows[pairs[k]][rc.oldIndex] ?? null;
      const y = newTable.rows[pairs[k + 1]][ac.newIndex] ?? null;
      if (isBlank2(x) && isBlank2(y)) continue;
      total++;
      if (equal(x, y)) same2++;
    }
    return total >= Math.min(3, matched.length) ? same2 / total : 0;
  };
  const fingerprint = (table, col2, side) => {
    let h = "";
    for (let s = 0, n = 0; s < matched.length && n < 64; s += step, n++) {
      h += String(table.rows[pairs[matched[s] + side]][col2] ?? "") + "\0";
    }
    return h;
  };
  const byPrint = /* @__PURE__ */ new Map();
  for (const ac of added) {
    const f = fingerprint(newTable, ac.newIndex, 1);
    const list = byPrint.get(f);
    if (list) list.push(ac);
    else byPrint.set(f, [ac]);
  }
  const exhaustive = removed.length * added.length <= 4096;
  const merged = /* @__PURE__ */ new Set();
  removed.forEach((rc, i) => {
    let best = null;
    let bestScore = 0.8;
    const consider = (ac) => {
      if (!ac || merged.has(ac)) return;
      const sc2 = score(rc, ac);
      if (sc2 >= bestScore && (!best || sc2 > bestScore)) {
        best = ac;
        bestScore = sc2;
      }
    };
    consider(added[i]);
    if (bestScore < 1) for (const ac of byPrint.get(fingerprint(oldTable, rc.oldIndex, 0)) ?? []) consider(ac);
    if (bestScore < 1 && exhaustive) for (const ac of added) consider(ac);
    if (best) {
      const b = best;
      b.oldIndex = rc.oldIndex;
      b.oldName = rc.name;
      b.status = "renamed";
      merged.add(b);
    }
  });
  const mergedOld = new Set([...merged].map((m) => m.oldIndex));
  return columns.filter((c) => !(c.status === "removed" && mergedOld.has(c.oldIndex)));
}

// packages/spreadiff/src/core/csv.ts
function columnLabel(index) {
  let s = "";
  let n = index + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function detectDelimiter(text3) {
  const sample = text3.slice(0, 64 * 1024);
  const lines = sample.split(/\r?\n/, 20).filter((l) => l.length > 0);
  let best = ",";
  let bestScore = -1;
  for (const d of [",", "	", ";", "|"]) {
    const counts = lines.map((l) => l.split(d).length - 1);
    if (counts.length === 0 || counts[0] === 0) continue;
    const consistent = counts.filter((c) => c === counts[0]).length;
    const score = consistent * 1e3 + counts[0];
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}
function parseCsv(text3, options = {}) {
  if (text3.charCodeAt(0) === 65279) text3 = text3.slice(1);
  const hint = /^sep=(.)\r?\n/.exec(text3.slice(0, 8));
  if (hint) text3 = text3.slice(hint[0].length);
  const delim = (options.delimiter ?? hint?.[1] ?? detectDelimiter(text3)).charCodeAt(0);
  const QUOTE = 34;
  const CR = 13;
  const LF = 10;
  const rows = [];
  let row = [];
  let i = 0;
  const len = text3.length;
  while (i < len) {
    let field;
    if (text3.charCodeAt(i) === QUOTE) {
      i++;
      let start = i;
      let parts = null;
      for (; ; ) {
        const q = text3.indexOf('"', i);
        if (q === -1) {
          i = len;
          field = (parts ? parts.join("") : "") + text3.slice(start);
          break;
        }
        if (text3.charCodeAt(q + 1) === QUOTE) {
          (parts ??= []).push(text3.slice(start, q + 1));
          i = q + 2;
          start = i;
          continue;
        }
        field = (parts ? parts.join("") : "") + text3.slice(start, q);
        i = q + 1;
        break;
      }
      const tail = i;
      while (i < len) {
        const c2 = text3.charCodeAt(i);
        if (c2 === delim || c2 === LF || c2 === CR) break;
        i++;
      }
      if (i > tail) field += text3.slice(tail, i);
    } else {
      const start = i;
      while (i < len) {
        const c2 = text3.charCodeAt(i);
        if (c2 === delim || c2 === LF || c2 === CR) break;
        i++;
      }
      field = text3.slice(start, i);
    }
    row.push(field);
    const c = text3.charCodeAt(i);
    if (c === delim) {
      i++;
      if (i === len) row.push("");
    } else {
      if (c === CR) i++;
      if (text3.charCodeAt(i) === LF) i++;
      rows.push(row);
      row = [];
    }
  }
  if (row.length) rows.push(row);
  return toTable(rows, options);
}
var NUMERIC = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;
function isCanonicalNumber(f) {
  const len = f.length;
  const neg = f.charCodeAt(0) === 45;
  let i = neg ? 1 : 0;
  if (i >= len) return false;
  let c = f.charCodeAt(i);
  if (c < 48 || c > 57) return false;
  let digits = 0;
  if (c === 48) i++;
  else {
    while (i < len && (c = f.charCodeAt(i)) >= 48 && c <= 57) {
      i++;
      digits++;
    }
  }
  if (i === len) return digits <= 15 && !(neg && digits === 0);
  if (f.charCodeAt(i) !== 46 || ++i === len) return false;
  const intZero = digits === 0;
  let leadingZeros = 0;
  let last = 0;
  for (; i < len; i++) {
    c = f.charCodeAt(i);
    if (c < 48 || c > 57) return false;
    if (intZero && digits === 0 && c === 48) leadingZeros++;
    else digits++;
    last = c;
  }
  return last !== 48 && digits <= 15 && leadingZeros < 6;
}
function toTable(raw, options) {
  while (raw.length && raw[raw.length - 1].every((f) => f === "")) raw.pop();
  const width = raw.reduce((w, r) => Math.max(w, r.length), 0);
  const header = options.header ?? true;
  const columns = header ? uniqueNames(Array.from({ length: width }, (_, c) => raw[0]?.[c] ?? "")) : Array.from({ length: width }, (_, c) => columnLabel(c));
  const body = header ? raw.slice(1) : raw;
  const numbers2 = options.numbers ?? "lossless";
  const rows = new Array(body.length);
  for (let r = 0; r < body.length; r++) {
    const src = body[r];
    const out = new Array(width);
    for (let c = 0; c < width; c++) {
      const f = src[c];
      if (f === void 0 || f === "") out[c] = null;
      else if (numbers2 === "lossless") out[c] = isCanonicalNumber(f) ? Number(f) : f;
      else if (numbers2 && NUMERIC.test(f)) out[c] = Number(f);
      else out[c] = f;
    }
    rows[r] = out;
  }
  return { name: options.name ?? "Sheet1", columns, rows, origin: { row: header ? 1 : 0, col: 0 } };
}
function uniqueNames(names) {
  const bases = names.map((raw, i) => raw.trim() === "" ? columnLabel(i) : raw.trim());
  const taken = /* @__PURE__ */ new Set();
  const firstSeen = /* @__PURE__ */ new Set();
  const isDup = bases.map((b) => firstSeen.has(b) ? true : (firstSeen.add(b), false));
  bases.forEach((b, i) => {
    if (!isDup[i]) taken.add(b);
  });
  return bases.map((b, i) => {
    if (!isDup[i]) return b;
    let n = 2;
    while (taken.has(`${b} (${n})`)) n++;
    const name = `${b} (${n})`;
    taken.add(name);
    return name;
  });
}

// node_modules/fflate/esm/index.mjs
import { createRequire } from "module";
var require2 = createRequire("/");
var _a;
var Worker;
var isMarkedAsUntransferable;
try {
  _a = require2("worker_threads"), Worker = _a.Worker, isMarkedAsUntransferable = _a.isMarkedAsUntransferable;
} catch (e) {
}
var u8 = Uint8Array;
var u16 = Uint16Array;
var i32 = Int32Array;
var fleb = new u8([
  0,
  0,
  0,
  0,
  0,
  0,
  0,
  0,
  1,
  1,
  1,
  1,
  2,
  2,
  2,
  2,
  3,
  3,
  3,
  3,
  4,
  4,
  4,
  4,
  5,
  5,
  5,
  5,
  0,
  /* unused */
  0,
  0,
  /* impossible */
  0
]);
var fdeb = new u8([
  0,
  0,
  0,
  0,
  1,
  1,
  2,
  2,
  3,
  3,
  4,
  4,
  5,
  5,
  6,
  6,
  7,
  7,
  8,
  8,
  9,
  9,
  10,
  10,
  11,
  11,
  12,
  12,
  13,
  13,
  /* unused */
  0,
  0
]);
var clim = new u8([16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15]);
var freb = function(eb, start) {
  var b = new u16(31);
  for (var i = 0; i < 31; ++i) {
    b[i] = start += 1 << eb[i - 1];
  }
  var r = new i32(b[30]);
  for (var i = 1; i < 30; ++i) {
    for (var j = b[i]; j < b[i + 1]; ++j) {
      r[j] = j - b[i] << 5 | i;
    }
  }
  return { b, r };
};
var _a = freb(fleb, 2);
var fl = _a.b;
var revfl = _a.r;
fl[28] = 258, revfl[258] = 28;
var _b = freb(fdeb, 0);
var fd = _b.b;
var revfd = _b.r;
var rev = new u16(32768);
for (i = 0; i < 32768; ++i) {
  x = (i & 43690) >> 1 | (i & 21845) << 1;
  x = (x & 52428) >> 2 | (x & 13107) << 2;
  x = (x & 61680) >> 4 | (x & 3855) << 4;
  rev[i] = ((x & 65280) >> 8 | (x & 255) << 8) >> 1;
}
var x;
var i;
var hMap = (function(cd, mb, r) {
  var s = cd.length;
  var i = 0;
  var l = new u16(mb);
  for (; i < s; ++i) {
    if (cd[i])
      ++l[cd[i] - 1];
  }
  var le = new u16(mb);
  for (i = 1; i < mb; ++i) {
    le[i] = le[i - 1] + l[i - 1] << 1;
  }
  var co;
  if (r) {
    co = new u16(1 << mb);
    var rvb = 15 - mb;
    for (i = 0; i < s; ++i) {
      if (cd[i]) {
        var sv = i << 4 | cd[i];
        var r_1 = mb - cd[i];
        var v = le[cd[i] - 1]++ << r_1;
        for (var m = v | (1 << r_1) - 1; v <= m; ++v) {
          co[rev[v] >> rvb] = sv;
        }
      }
    }
  } else {
    co = new u16(s);
    for (i = 0; i < s; ++i) {
      if (cd[i]) {
        co[i] = rev[le[cd[i] - 1]++] >> 15 - cd[i];
      }
    }
  }
  return co;
});
var flt = new u8(288);
for (i = 0; i < 144; ++i)
  flt[i] = 8;
var i;
for (i = 144; i < 256; ++i)
  flt[i] = 9;
var i;
for (i = 256; i < 280; ++i)
  flt[i] = 7;
var i;
for (i = 280; i < 288; ++i)
  flt[i] = 8;
var i;
var fdt = new u8(32);
for (i = 0; i < 32; ++i)
  fdt[i] = 5;
var i;
var flrm = /* @__PURE__ */ hMap(flt, 9, 1);
var fdrm = /* @__PURE__ */ hMap(fdt, 5, 1);
var max = function(a) {
  var m = a[0];
  for (var i = 1; i < a.length; ++i) {
    if (a[i] > m)
      m = a[i];
  }
  return m;
};
var bits = function(d, p, m) {
  var o = p / 8 | 0;
  return (d[o] | d[o + 1] << 8) >> (p & 7) & m;
};
var bits16 = function(d, p) {
  var o = p / 8 | 0;
  return (d[o] | d[o + 1] << 8 | d[o + 2] << 16) >> (p & 7);
};
var shft = function(p) {
  return (p + 7) / 8 | 0;
};
var slc = function(v, s, e) {
  if (s == null || s < 0)
    s = 0;
  if (e == null || e > v.length)
    e = v.length;
  return new u8(v.subarray(s, e));
};
var ec = [
  "unexpected EOF",
  "invalid block type",
  "invalid length/literal",
  "invalid distance",
  "stream finished",
  "no stream handler",
  ,
  // determined by compression function
  "no callback",
  "invalid UTF-8 data",
  "extra field too long",
  "date not in range 1980-2099",
  "filename too long",
  "stream finishing",
  "invalid zip data"
  // determined by unknown compression method
];
var err = function(ind, msg, nt) {
  var e = new Error(msg || ec[ind]);
  e.code = ind;
  if (Error.captureStackTrace)
    Error.captureStackTrace(e, err);
  if (!nt)
    throw e;
  return e;
};
var inflt = function(dat, st, buf, dict) {
  var sl = dat.length, dl = dict ? dict.length : 0;
  if (!sl || st.f && !st.l)
    return buf || new u8(0);
  var noBuf = !buf;
  var resize = noBuf || st.i != 2;
  var noSt = st.i;
  if (noBuf)
    buf = new u8(sl * 3);
  var cbuf = function(l2) {
    var bl = buf.length;
    if (l2 > bl) {
      var nbuf = new u8(Math.max(bl * 2, l2));
      nbuf.set(buf);
      buf = nbuf;
    }
  };
  var final = st.f || 0, pos = st.p || 0, bt = st.b || 0, lm = st.l, dm = st.d, lbt = st.m, dbt = st.n;
  var tbts = sl * 8;
  do {
    if (!lm) {
      final = bits(dat, pos, 1);
      var type = bits(dat, pos + 1, 3);
      pos += 3;
      if (!type) {
        var s = shft(pos) + 4, l = dat[s - 4] | dat[s - 3] << 8, t = s + l;
        if (t > sl) {
          if (noSt)
            err(0);
          break;
        }
        if (resize)
          cbuf(bt + l);
        buf.set(dat.subarray(s, t), bt);
        st.b = bt += l, st.p = pos = t * 8, st.f = final;
        continue;
      } else if (type == 1)
        lm = flrm, dm = fdrm, lbt = 9, dbt = 5;
      else if (type == 2) {
        var hLit = bits(dat, pos, 31) + 257, hcLen = bits(dat, pos + 10, 15) + 4;
        var tl = hLit + bits(dat, pos + 5, 31) + 1;
        pos += 14;
        var ldt = new u8(tl);
        var clt = new u8(19);
        for (var i = 0; i < hcLen; ++i) {
          clt[clim[i]] = bits(dat, pos + i * 3, 7);
        }
        pos += hcLen * 3;
        var clb = max(clt), clbmsk = (1 << clb) - 1;
        var clm = hMap(clt, clb, 1);
        for (var i = 0; i < tl; ) {
          var r = clm[bits(dat, pos, clbmsk)];
          pos += r & 15;
          var s = r >> 4;
          if (s < 16) {
            ldt[i++] = s;
          } else {
            var c = 0, n = 0;
            if (s == 16)
              n = 3 + bits(dat, pos, 3), pos += 2, c = ldt[i - 1];
            else if (s == 17)
              n = 3 + bits(dat, pos, 7), pos += 3;
            else if (s == 18)
              n = 11 + bits(dat, pos, 127), pos += 7;
            while (n--)
              ldt[i++] = c;
          }
        }
        var lt = ldt.subarray(0, hLit), dt = ldt.subarray(hLit);
        lbt = max(lt);
        dbt = max(dt);
        lm = hMap(lt, lbt, 1);
        dm = hMap(dt, dbt, 1);
      } else
        err(1);
      if (pos > tbts) {
        if (noSt)
          err(0);
        break;
      }
    }
    if (resize)
      cbuf(bt + 131072);
    var lms = (1 << lbt) - 1, dms = (1 << dbt) - 1;
    var lpos = pos;
    for (; ; lpos = pos) {
      var c = lm[bits16(dat, pos) & lms], sym = c >> 4;
      pos += c & 15;
      if (pos > tbts) {
        if (noSt)
          err(0);
        break;
      }
      if (!c)
        err(2);
      if (sym < 256)
        buf[bt++] = sym;
      else if (sym == 256) {
        lpos = pos, lm = null;
        break;
      } else {
        var add = sym - 254;
        if (sym > 264) {
          var i = sym - 257, b = fleb[i];
          add = bits(dat, pos, (1 << b) - 1) + fl[i];
          pos += b;
        }
        var d = dm[bits16(dat, pos) & dms], dsym = d >> 4;
        if (!d)
          err(3);
        pos += d & 15;
        var dt = fd[dsym];
        if (dsym > 3) {
          var b = fdeb[dsym];
          dt += bits16(dat, pos) & (1 << b) - 1, pos += b;
        }
        if (pos > tbts) {
          if (noSt)
            err(0);
          break;
        }
        if (resize)
          cbuf(bt + 131072);
        var end = bt + add;
        if (bt < dt) {
          var shift = dl - dt, dend = Math.min(dt, end);
          if (shift + bt < 0)
            err(3);
          for (; bt < dend; ++bt)
            buf[bt] = dict[shift + bt];
        }
        for (; bt < end; ++bt)
          buf[bt] = buf[bt - dt];
      }
    }
    st.l = lm, st.p = lpos, st.b = bt, st.f = final;
    if (lm)
      final = 1, st.m = lbt, st.d = dm, st.n = dbt;
  } while (!final);
  return bt != buf.length && noBuf ? slc(buf, 0, bt) : buf.subarray(0, bt);
};
var et = /* @__PURE__ */ new u8(0);
var b2 = function(d, b) {
  return d[b] | d[b + 1] << 8;
};
var b4 = function(d, b) {
  return (d[b] | d[b + 1] << 8 | d[b + 2] << 16 | d[b + 3] << 24) >>> 0;
};
var b8 = function(d, b) {
  return b4(d, b) + b4(d, b + 4) * 4294967296;
};
function inflateSync(data, opts) {
  return inflt(data, { i: 2 }, opts && opts.out, opts && opts.dictionary);
}
var td = typeof TextDecoder != "undefined" && /* @__PURE__ */ new TextDecoder();
var tds = 0;
try {
  td.decode(et, { stream: true });
  tds = 1;
} catch (e) {
}
var dutf8 = function(d) {
  for (var r = "", i = 0; ; ) {
    var c = d[i++];
    var eb = (c > 127) + (c > 223) + (c > 239);
    if (i + eb > d.length)
      return { s: r, r: slc(d, i - 1) };
    if (!eb)
      r += String.fromCharCode(c);
    else if (eb == 3) {
      c = ((c & 15) << 18 | (d[i++] & 63) << 12 | (d[i++] & 63) << 6 | d[i++] & 63) - 65536, r += String.fromCharCode(55296 | c >> 10, 56320 | c & 1023);
    } else if (eb & 1)
      r += String.fromCharCode((c & 31) << 6 | d[i++] & 63);
    else
      r += String.fromCharCode((c & 15) << 12 | (d[i++] & 63) << 6 | d[i++] & 63);
  }
};
function strFromU8(dat, latin1) {
  if (latin1) {
    var r = "";
    for (var i = 0; i < dat.length; i += 16384)
      r += String.fromCharCode.apply(null, dat.subarray(i, i + 16384));
    return r;
  } else if (td) {
    return td.decode(dat);
  } else {
    var _a2 = dutf8(dat), s = _a2.s, r = _a2.r;
    if (r.length)
      err(8);
    return s;
  }
}
var slzh = function(d, b) {
  return b + 30 + b2(d, b + 26) + b2(d, b + 28);
};
var zh = function(d, b, z) {
  var fnl = b2(d, b + 28), efl = b2(d, b + 30), fn = strFromU8(d.subarray(b + 46, b + 46 + fnl), !(b2(d, b + 8) & 2048)), es = b + 46 + fnl;
  var _a2 = z64hs(d, es, efl, z, b4(d, b + 20), b4(d, b + 24), b4(d, b + 42)), sc2 = _a2[0], su = _a2[1], off = _a2[2];
  return [b2(d, b + 10), sc2, su, fn, es + efl + b2(d, b + 32), off];
};
var z64hs = function(d, b, l, z, sc2, su, off) {
  var nsc = sc2 == 4294967295, nsu = su == 4294967295, noff = off == 4294967295, e = b + l;
  var nf = nsc + nsu + noff;
  if (z && nf) {
    for (; b + 4 < e; b += 4 + b2(d, b + 2)) {
      if (b2(d, b) == 1) {
        return [
          nsc ? b8(d, b + 4 + 8 * nsu) : sc2,
          nsu ? b8(d, b + 4) : su,
          noff ? b8(d, b + 4 + 8 * (nsu + nsc)) : off,
          1
        ];
      }
    }
    if (z < 2)
      err(13);
  }
  return [sc2, su, off, 0];
};
function unzipSync(data, opts) {
  var files = {};
  var e = data.length - 22;
  for (; b4(data, e) != 101010256; --e) {
    if (!e || data.length - e > 65558)
      err(13);
  }
  ;
  var c = b2(data, e + 8);
  if (!c)
    return {};
  var o = b4(data, e + 16);
  var z = b4(data, e - 20) == 117853008;
  if (z) {
    var ze = b4(data, e - 12);
    z = b4(data, ze) == 101075792;
    if (z) {
      c = b4(data, ze + 32);
      o = b4(data, ze + 48);
    }
  }
  var fltr = opts && opts.filter;
  for (var i = 0; i < c; ++i) {
    var _a2 = zh(data, o, z), c_2 = _a2[0], sc2 = _a2[1], su = _a2[2], fn = _a2[3], no = _a2[4], off = _a2[5], b = slzh(data, off);
    o = no;
    if (!fltr || fltr({
      name: fn,
      size: sc2,
      originalSize: su,
      compression: c_2
    })) {
      if (!c_2)
        files[fn] = slc(data, b, b + sc2);
      else if (c_2 == 8)
        files[fn] = inflateSync(data.subarray(b, b + sc2), { out: new u8(su) });
      else
        err(14, "unknown compression type " + c_2);
    }
  }
  return files;
}

// packages/spreadiff/src/core/serial.ts
var DAY = 864e5;
var EPOCH = Date.UTC(1899, 11, 30);
function serialFromParts(y, m, d, date19042 = false) {
  const ms = Date.UTC(y, m - 1, d);
  let s = Math.round((ms - EPOCH) / DAY);
  if (date19042) return s - 1462;
  if (s < 61) s -= 1;
  return s;
}
function partsFromSerial(serial, date19042 = false) {
  let s = Math.floor(serial);
  if (date19042) s += 1462;
  else if (s < 61) {
    if (s === 60) return { y: 1900, m: 2, d: 29, dow: 3 };
    if (s === 0) return { y: 1900, m: 1, d: 0, dow: 6 };
    if (s < 1) s += 1;
    else {
      const dt2 = new Date(EPOCH + (s + 1) * DAY);
      return { y: dt2.getUTCFullYear(), m: dt2.getUTCMonth() + 1, d: dt2.getUTCDate(), dow: (s - 1) % 7 };
    }
  }
  const dt = new Date(EPOCH + s * DAY);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: dt.getUTCDay() };
}
var ISO = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}):(\d{2}(?:\.\d+)?))?$/;
var TIME = /^(\d{2}):(\d{2}):(\d{2}(?:\.\d+)?)$/;
function isoToSerial(s, date19042 = false) {
  let m = ISO.exec(s);
  if (m) {
    const day = serialFromParts(Number(m[1]), Number(m[2]), Number(m[3]), date19042);
    const t = m[4] ? (Number(m[4]) * 3600 + Number(m[5]) * 60 + Number(m[6])) / 86400 : 0;
    return day + t;
  }
  m = TIME.exec(s);
  if (m) return (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3])) / 86400;
  return null;
}
function serialToIso(serial, date19042) {
  if (!Number.isFinite(serial) || serial < 0 || serial > 2958465) return null;
  const s = date19042 ? serial + 1462 : serial < 60 ? serial + 1 : serial;
  const ms = Math.round((s - 25569) * 864e5);
  const iso = new Date(ms).toISOString();
  if (Number.isInteger(serial)) return iso.slice(0, 10);
  const frac = iso.slice(19, 23) === ".000" ? "" : iso.slice(19, 23).replace(/0+$/, "");
  if (serial < 1) return iso.slice(11, 19) + frac;
  return `${iso.slice(0, 10)} ${iso.slice(11, 19)}${frac}`;
}

// packages/spreadiff/src/core/numfmt.ts
var BUILTIN_FORMATS = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: "$#,##0_);($#,##0)",
  6: "$#,##0_);[Red]($#,##0)",
  7: "$#,##0.00_);($#,##0.00)",
  8: "$#,##0.00_);[Red]($#,##0.00)",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0_);(#,##0)",
  38: "#,##0_);[Red](#,##0)",
  39: "#,##0.00_);(#,##0.00)",
  40: "#,##0.00_);[Red](#,##0.00)",
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mm:ss.0",
  48: "##0.0E+0",
  49: "@",
  // Built-in dates and times of East Asian Excel (these are Japanese; Chinese and Korean differ only in text).
  27: 'yyyy"\u5E74"m"\u6708"',
  28: 'm"\u6708"d"\u65E5"',
  29: 'm"\u6708"d"\u65E5"',
  30: "m/d/yy",
  31: 'yyyy"\u5E74"m"\u6708"d"\u65E5"',
  32: 'h"\u6642"mm"\u5206"',
  33: 'h"\u6642"mm"\u5206"ss"\u79D2"',
  34: 'yyyy"\u5E74"m"\u6708"',
  35: 'm"\u6708"d"\u65E5"',
  36: 'yyyy"\u5E74"m"\u6708"',
  50: 'yyyy"\u5E74"m"\u6708"',
  51: 'm"\u6708"d"\u65E5"',
  52: 'yyyy"\u5E74"m"\u6708"',
  53: 'm"\u6708"d"\u65E5"',
  54: 'm"\u6708"d"\u65E5"',
  55: 'yyyy"\u5E74"m"\u6708"',
  56: 'm"\u6708"d"\u65E5"',
  57: 'yyyy"\u5E74"m"\u6708"',
  58: 'm"\u6708"d"\u65E5"'
};
var MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
var DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
function splitSections(code) {
  const out = [];
  let cur = "";
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const j = end < 0 ? code.length : end;
      cur += code.slice(i, j + 1);
      i = j;
    } else if (ch === "\\" || ch === "_" || ch === "*") {
      cur += code.slice(i, i + 2);
      i++;
    } else if (ch === "[") {
      const end = code.indexOf("]", i);
      const j = end < 0 ? code.length : end;
      cur += code.slice(i, j + 1);
      i = j;
    } else if (ch === ";") {
      out.push(cur);
      cur = "";
    } else cur += ch;
  }
  out.push(cur);
  return out;
}
var bare = (s) => s.replace(/"[^"]*"|\\.|_.|\*./g, "");
function sectionIsDate(s) {
  const b = bare(s).replace(/\[(?![hms]+\])[^\]]*\]/gi, "");
  return /[yd]|[hs]|\[[hms]+\]|(?<![#0?.,])m|am\/pm|a\/p/i.test(b) && !/[#0?]/.test(b.replace(/\.0+|0+(?=\])/g, ""));
}
function isDateFormat(code) {
  return sectionIsDate(splitSections(code)[0]);
}
function parseSection(s) {
  let cond;
  const currency = [];
  const body = s.replace(/"[^"]*"|\\.|\[([^\]]*)\]/g, (t, inner) => {
    if (inner === void 0) return t;
    const c = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:[eE][-+]?\d+)?)$/.exec(inner);
    if (c) {
      cond = { op: c[1], n: Number(c[2]) };
      return "";
    }
    if (inner[0] === "$") {
      const sym = inner.slice(1).split("-")[0];
      if (sym) currency.push(sym);
      return sym ? `"${sym.replace(/"/g, "")}"` : "";
    }
    if (/^[hms]+$/i.test(inner)) return t;
    return "";
  });
  return { body, cond, currency };
}
var test = (c, v) => {
  if (!c) return true;
  switch (c.op) {
    case "<":
      return v < c.n;
    case "<=":
      return v <= c.n;
    case ">":
      return v > c.n;
    case ">=":
      return v >= c.n;
    case "=":
      return v === c.n;
    default:
      return v !== c.n;
  }
};
function formatNumberCode(v, code, date19042 = false, general = generalText) {
  if (!Number.isFinite(v)) return null;
  const raw = splitSections(code);
  const secs = raw.map(parseSection);
  const numeric = secs.slice(0, 3);
  let sec;
  let value = v;
  let minus = false;
  if (numeric.some((s) => s.cond)) {
    const [a, b, c] = numeric;
    if (test(a.cond ?? { op: ">=", n: 0 }, v)) sec = a;
    else if (b && test(b.cond ?? { op: "<", n: 0 }, v)) sec = b;
    else sec = c ?? b ?? a;
    if (v < 0) {
      if (sec === a || numeric.length === 1) minus = true;
      value = -v;
    }
  } else if (v < 0 && numeric.length >= 2) {
    sec = numeric[1];
    value = -v;
  } else if (v === 0 && numeric.length >= 3) sec = numeric[2];
  else {
    sec = numeric[0];
    if (v < 0) {
      minus = true;
      value = -v;
    }
  }
  const body = sec.body;
  if (body === "" && raw.length > 1) return "";
  if (/^\s*general\s*$/i.test(body)) return (minus ? "-" : "") + general(value);
  if (/general/i.test(bare(body))) return (minus ? "-" : "") + body.replace(/"([^"]*)"|\\(.)|_.|\*.|general/gi, (t, q, e) => q !== void 0 ? q : e !== void 0 ? e : t[0] === "_" ? " " : t[0] === "*" ? "" : general(value));
  if (sectionIsDate(body)) return minus ? null : formatDate(value, body, date19042);
  return formatDigits(value, body, minus);
}
function formatTextCode(v, code) {
  const s = splitSections(code);
  const sec = s.length > 3 ? s[3] : s.length === 1 && /@/.test(bare(s[0])) ? s[0] : null;
  if (sec === null) return null;
  return parseSection(sec).body.replace(/"([^"]*)"|\\(.)|_.|\*.|@/g, (t, q, e) => q !== void 0 ? q : e !== void 0 ? e : t[0] === "_" ? " " : t[0] === "*" ? "" : v);
}
function generalText(n) {
  if (n === 0) return "0";
  const a = Math.abs(n);
  if (a >= 1e-9 && a < 1e11) {
    const int = Math.floor(a).toString().length;
    const decimals = Math.max(0, Math.min(10 - int, 10));
    const s = Number(n.toFixed(decimals)).toString();
    if (!/e/i.test(s) && s.replace(/^-/, "").length <= 11) return s;
  }
  const [m, e] = n.toExponential(5).split("e");
  const mant = m.replace(/\.?0+$/, "");
  const exp = Number(e);
  return `${mant}E${exp < 0 ? "-" : "+"}${String(Math.abs(exp)).padStart(2, "0")}`;
}
function roundHalfAway(x, digits) {
  const f = 10 ** digits;
  const y = Number((Math.abs(x) * f).toPrecision(15));
  return Math.sign(x) * Math.round(y) / f;
}
function tokenize(body) {
  const out = [];
  const lit = (s) => {
    const last = out[out.length - 1];
    if (last?.k === "lit") last.s += s;
    else out.push({ k: "lit", s });
  };
  let seenDot = false;
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '"') {
      const end = body.indexOf('"', i + 1);
      const j = end < 0 ? body.length : end;
      lit(body.slice(i + 1, j));
      i = j;
    } else if (ch === "\\") {
      lit(body[i + 1] ?? "");
      i++;
    } else if (ch === "_") {
      lit(" ");
      i++;
    } else if (ch === "*") i++;
    else if (ch === "0" || ch === "#" || ch === "?") out.push({ k: "dig", c: ch });
    else if (ch === "." && !seenDot) {
      seenDot = true;
      out.push({ k: "dot" });
    } else if (ch === ",") out.push({ k: "comma" });
    else if (ch === "%") out.push({ k: "pct" });
    else if ((ch === "E" || ch === "e") && (body[i + 1] === "+" || body[i + 1] === "-")) {
      out.push({ k: "exp", sign: body[i + 1] });
      i++;
    } else if (ch === "/") out.push({ k: "slash" });
    else if (ch === "@") out.push({ k: "at" });
    else if (ch === "[") return null;
    else lit(ch);
  }
  return out;
}
function fillInt(digits, ph) {
  const out = new Array(ph.length).fill("");
  let d = digits.length - 1;
  for (let i = ph.length - 1; i >= 0; i--) {
    if (d >= 0) out[i] = digits[d--];
    else out[i] = ph[i] === "0" ? "0" : ph[i] === "?" ? " " : "";
  }
  if (d >= 0 && ph.length) out[0] = digits.slice(0, d + 1) + out[0];
  return out;
}
function fillDec(digits, ph) {
  const out = ph.map((_, i) => digits[i] ?? "0");
  for (let i = ph.length - 1; i >= 0 && out[i] === "0" && ph[i] !== "0"; i--) out[i] = ph[i] === "?" ? " " : "";
  return out;
}
function formatDigits(x, body, minus) {
  const toks = tokenize(body);
  if (!toks) return null;
  if (toks.some((t) => t.k === "at")) return null;
  const sign = minus ? "-" : "";
  const digIdx = toks.map((t, i) => t.k === "dig" ? i : -1).filter((i) => i >= 0);
  if (!digIdx.length) {
    return sign + toks.map((t) => t.k === "lit" ? t.s : t.k === "pct" ? "%" : t.k === "comma" ? "," : t.k === "dot" ? "." : t.k === "slash" ? "/" : "").join("");
  }
  const slash = toks.findIndex((t) => t.k === "slash");
  if (slash >= 0 && toks.slice(0, slash).some((t) => t.k === "dig") && toks.slice(slash + 1).some((t) => t.k === "dig" || t.k === "lit" && /\d/.test(t.s))) {
    return formatFraction(x, toks, slash, sign);
  }
  const pct = toks.filter((t) => t.k === "pct").length;
  x *= 100 ** pct;
  const expAt = toks.findIndex((t) => t.k === "exp");
  const dotAt = toks.findIndex((t) => t.k === "dot");
  const mantEnd = expAt >= 0 ? expAt : toks.length;
  const intEnd = dotAt >= 0 && dotAt < mantEnd ? dotAt : mantEnd;
  const firstDig = digIdx[0];
  const lastIntDig = [...digIdx].reverse().find((i) => i < intEnd) ?? -1;
  let group = false;
  for (let i = firstDig + 1; i < lastIntDig; i++) if (toks[i].k === "comma") group = true;
  const lastDig = [...digIdx].reverse().find((i) => i < mantEnd) ?? -1;
  let scale = 0;
  for (let i = lastDig + 1; i < mantEnd && (toks[i].k === "comma" || toks[i].k === "dot" && i === intEnd); i++) if (toks[i].k === "comma") scale++;
  x /= 1e3 ** scale;
  const intPh = toks.slice(0, intEnd).filter((t) => t.k === "dig").map((t) => t.c);
  const decPh = toks.slice(intEnd, mantEnd).filter((t) => t.k === "dig").map((t) => t.c);
  let exp = 0;
  let expPh = [];
  if (expAt >= 0) {
    expPh = toks.slice(expAt + 1).filter((t) => t.k === "dig").map((t) => t.c);
    if (x !== 0) {
      const lg = Math.floor(Math.log10(x));
      const step = intPh.includes("#") && intPh.length > 1 ? intPh.length : 0;
      exp = step ? Math.floor(lg / step) * step : lg - Math.max(0, intPh.filter((c) => c === "0").length - 1);
      x /= 10 ** exp;
      const width = step || Math.max(1, intPh.filter((c) => c === "0").length);
      if (roundHalfAway(x, decPh.length) >= 10 ** width) {
        exp += step || 1;
        x /= 10 ** (step || 1);
      }
    }
  }
  const r = roundHalfAway(x, decPh.length);
  const fixed = r.toFixed(decPh.length);
  const [ip, dp = ""] = fixed.split(".");
  const intDigits = ip === "0" ? "" : ip;
  let ints = fillInt(intDigits, intPh);
  if (group) {
    const field = ints.join("");
    const lead = /^ */.exec(field)[0];
    const grouped = lead + field.slice(lead.length).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    ints = [grouped, ...new Array(intPh.length - 1).fill("")];
  }
  const decs = fillDec(dp, decPh);
  const expDigits = expAt >= 0 ? fillInt(String(Math.abs(exp)), expPh.length ? expPh : ["0"]) : [];
  let out = "";
  let ii = 0;
  let di = 0;
  let ei = 0;
  let dotShown = false;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.k === "lit") out += t.s;
    else if (t.k === "pct") out += "%";
    else if (t.k === "slash") out += "/";
    else if (t.k === "comma") continue;
    else if (t.k === "dot") {
      if (!dotShown) out += ".";
      dotShown = true;
    } else if (t.k === "exp") out += "E" + (exp < 0 ? "-" : t.sign === "+" ? "+" : "");
    else if (i < intEnd) out += ints[ii++] ?? "";
    else if (i < mantEnd) out += decs[di++] ?? "";
    else out += expDigits[ei++] ?? "";
  }
  return sign + out;
}
function formatFraction(x, toks, slash, sign) {
  const before = toks.slice(0, slash);
  const after = toks.slice(slash + 1);
  let numStart = before.length;
  while (numStart > 0 && before[numStart - 1].k === "dig") numStart--;
  const wholeToks = before.slice(0, numStart);
  const numToks = before.slice(numStart);
  const hasWhole = wholeToks.some((t) => t.k === "dig");
  const fixedDen = after.find((t) => t.k === "lit" && /^\d+/.test(t.s));
  const denPh = after.filter((t) => t.k === "dig").length;
  const whole = hasWhole ? Math.floor(x) : 0;
  const frac = x - whole;
  let n;
  let d;
  if (fixedDen) {
    d = Number(/^\d+/.exec(fixedDen.s)[0]);
    n = Math.round(frac * d);
  } else {
    [n, d] = bestFraction(frac, 10 ** Math.max(1, denPh) - 1);
  }
  let w = whole;
  if (n === d && hasWhole) {
    w++;
    n = 0;
  }
  const pad = (s, ph, left) => {
    const width = ph.filter((t) => t.k === "dig").length;
    return left ? s.padStart(width, " ") : s.padEnd(width, " ");
  };
  const wholeText = hasWhole ? w === 0 && n !== 0 ? "" : String(w) : "";
  const lits = (ts) => ts.filter((t) => t.k === "lit").map((t) => t.s).join("");
  if (hasWhole && n === 0) {
    const width = numToks.length + 1 + (fixedDen ? fixedDen.s.length : denPh);
    return sign + wholeText + lits(wholeToks).replace(/^\s*/, (s) => s) + " ".repeat(width);
  }
  const num3 = String(hasWhole ? n : Math.round(n + whole * d));
  const den = fixedDen ? fixedDen.s : pad(String(d), after, false);
  return sign + (wholeText ? wholeText + lits(wholeToks) : hasWhole ? " ".repeat(wholeText.length) + lits(wholeToks).replace(/./g, " ") : "") + pad(num3, numToks, true) + "/" + den;
}
function bestFraction(x, max2) {
  let best = [Math.round(x), 1];
  let err2 = Math.abs(x - best[0]);
  for (let d = 1; d <= max2 && err2 > 1e-12; d++) {
    const n = Math.round(x * d);
    const e = Math.abs(x - n / d);
    if (e < err2 - 1e-15) {
      best = [n, d];
      err2 = e;
    }
  }
  return best;
}
function formatDate(v, fmt, d1904) {
  if (v < 0 || v >= (d1904 ? 2957004 : 2958466)) return null;
  const fd2 = Math.min(3, Math.max(0, ...[...fmt.matchAll(/s+\.(0+)/gi)].map((m) => m[1].length)));
  const scale = 10 ** fd2;
  const totalUnits = Math.round(v * 86400 * scale);
  const dayCarry = Math.floor(totalUnits / (86400 * scale));
  const p = partsFromSerial(dayCarry, d1904);
  const units = totalUnits - dayCarry * 86400 * scale;
  const secs = Math.floor(units / scale);
  const frac = fd2 ? String(units % scale).padStart(fd2, "0") : "";
  const hh = Math.floor(secs / 3600);
  const mm = Math.floor(secs % 3600 / 60);
  const ss = secs % 60;
  const elapsedSecs = Math.floor(totalUnits / scale);
  const ampm = /AM\/PM|A\/P/i.test(bare(fmt));
  const tokens = fmt.match(/"[^"]*"|\\.|_.|\*.|\[[hms]+\]|AM\/PM|A\/P|y+|m+|d+|h+|s+|[aA]|[^ymdhsaA"\\_*[]+|\[/gi) ?? [];
  const isUnit = (x) => /^(\[?[ymdhs])/i.test(x);
  let out = "";
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k];
    const l = t.toLowerCase();
    let prev = "";
    for (let j = k - 1; j >= 0; j--) if (isUnit(tokens[j])) {
      prev = tokens[j].toLowerCase().replace("[", "");
      break;
    }
    let nextTok = "";
    for (let j = k + 1; j < tokens.length; j++) if (isUnit(tokens[j])) {
      nextTok = tokens[j].toLowerCase().replace("[", "");
      break;
    }
    if (t.startsWith('"')) out += t.slice(1, -1);
    else if (t.startsWith("\\")) out += t.slice(1);
    else if (t.startsWith("_")) out += " ";
    else if (t.startsWith("*")) continue;
    else if (/^\[h+\]$/i.test(t)) out += String(Math.floor(elapsedSecs / 3600)).padStart(t.length - 2, "0");
    else if (/^\[m+\]$/i.test(t)) out += String(Math.floor(elapsedSecs / 60)).padStart(t.length - 2, "0");
    else if (/^\[s+\]$/i.test(t)) out += String(elapsedSecs).padStart(t.length - 2, "0");
    else if (/^am\/pm$/i.test(t)) out += hh < 12 ? "AM" : "PM";
    else if (/^a\/p$/i.test(t)) out += hh < 12 ? t[0] : t[2];
    else if (l[0] === "y") out += l.length <= 2 ? String(p.y % 100).padStart(2, "0") : String(p.y);
    else if (l[0] === "d") out += l.length === 1 ? p.d : l.length === 2 ? String(p.d).padStart(2, "0") : l.length === 3 ? DAYS[p.dow].slice(0, 3) : DAYS[p.dow];
    else if (l[0] === "h") {
      const h = ampm ? hh % 12 || 12 : hh;
      out += l.length === 1 ? h : String(h).padStart(2, "0");
    } else if (l[0] === "s") out += l.length === 1 ? ss : String(ss).padStart(2, "0");
    else if (fd2 && /^\.0+/.test(t) && tokens[k - 1]?.toLowerCase().startsWith("s")) {
      const zeros = /^\.(0+)/.exec(t)[1].length;
      out += "." + frac.slice(0, zeros).padEnd(zeros, "0") + t.slice(1 + zeros);
    } else if (l[0] === "m") {
      if (prev.startsWith("h") || nextTok.startsWith("s")) out += l.length === 1 ? mm : String(mm).padStart(2, "0");
      else out += l.length === 1 ? p.m : l.length === 2 ? String(p.m).padStart(2, "0") : l.length === 3 ? MONTHS[p.m - 1].slice(0, 3) : l.length === 5 ? MONTHS[p.m - 1][0] : MONTHS[p.m - 1];
    } else out += t;
  }
  return out;
}

// packages/spreadiff/src/core/engine/parse.ts
var Unsupported = class extends Error {
};
var ERROR2 = /#(?:REF!|NAME\?|VALUE!|DIV\/0!|N\/A|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)/y;
var CELL = /(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?/y;
var COLS = /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/y;
var ROWS = /\$?(\d+):\$?(\d+)/y;
var QUOTED_SHEET2 = /'((?:[^']|'')+)'!/y;
var SHEET = /([A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*)(?::([A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*))?!/y;
var IDENT2 = /[A-Za-z_\\\u00C0-\uFFFF][A-Za-z0-9_.?\\\u00C0-\uFFFF]*/y;
var NUMBER2 = /(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y;
var MAX_ROW2 = 1048575;
var MAX_COL2 = 16383;
function at2(re, s, i) {
  re.lastIndex = i;
  return re.exec(s);
}
function col(letters) {
  let c = 0;
  for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
  return c - 1;
}
var IDENT_CHAR2 = /[A-Za-z0-9_.\u00C0-\uFFFF]/;
var identChar = (ch) => ch !== void 0 && IDENT_CHAR2.test(ch);
function lex(src) {
  const s = src.charCodeAt(0) === 61 ? src.slice(1) : src;
  const out = [];
  let i = 0;
  const refAt = (j, sheet, sheetTo) => {
    const okEnd = (e) => !identChar(s[e]) && s[e] !== "(" && s[e] !== "!";
    let m = at2(CELL, s, j);
    if (m && okEnd(j + m[0].length)) {
      const a = { r: Number(m[4]) - 1, c: col(m[2]), rr: !m[3], cr: !m[1] };
      const b = m[8] ? { r: Number(m[8]) - 1, c: col(m[6]), rr: !m[7], cr: !m[5] } : a;
      if (a.r <= MAX_ROW2 && b.r <= MAX_ROW2 && a.c <= MAX_COL2 && b.c <= MAX_COL2) {
        const [top, bottom] = a.r <= b.r ? [a, b] : [b, a];
        const [left, right] = a.c <= b.c ? [a, b] : [b, a];
        const rel = (top.rr ? 1 : 0) | (left.cr ? 2 : 0) | (bottom.rr ? 4 : 0) | (right.cr ? 8 : 0);
        return { end: j + m[0].length, ref: { sheet, sheetTo, r0: top.r, c0: left.c, r1: bottom.r, c1: right.c, rel } };
      }
    }
    m = at2(COLS, s, j);
    if (m && okEnd(j + m[0].length)) {
      const a = col(m[1]);
      const b = col(m[2]);
      return { end: j + m[0].length, ref: { sheet, sheetTo, r0: -1, r1: -1, c0: Math.min(a, b), c1: Math.max(a, b) } };
    }
    m = at2(ROWS, s, j);
    if (m && okEnd(j + m[0].length)) {
      const a = Number(m[1]) - 1;
      const b = Number(m[2]) - 1;
      return { end: j + m[0].length, ref: { sheet, sheetTo, r0: Math.min(a, b), r1: Math.max(a, b), c0: -1, c1: -1 } };
    }
    return null;
  };
  while (i < s.length) {
    const ch = s[i];
    if (ch === " " || ch === "\n" || ch === "\r" || ch === "	") {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let v = "";
      for (; ; ) {
        if (j >= s.length) throw new Unsupported("unterminated string");
        if (s[j] === '"') {
          if (s[j + 1] === '"') {
            v += '"';
            j += 2;
            continue;
          }
          break;
        }
        v += s[j++];
      }
      out.push({ k: "str", v });
      i = j + 1;
      continue;
    }
    if (ch === "#") {
      const m = at2(ERROR2, s, i);
      if (!m) throw new Unsupported("bad error literal");
      out.push({ k: "err", v: m[0] });
      i += m[0].length;
      continue;
    }
    if (ch === "[") throw new Unsupported("external or unqualified structured reference");
    if (!(i > 0 && identChar(s[i - 1]))) {
      const q = at2(QUOTED_SHEET2, s, i) ?? at2(SHEET, s, i);
      if (q) {
        let sheet = q[1].replace(/''/g, "'");
        let sheetTo = q[2];
        if (q[0].startsWith("'") && sheet.includes(":")) [sheet, sheetTo] = sheet.split(":");
        const r2 = refAt(i + q[0].length, sheet, sheetTo);
        if (!r2) {
          const id2 = !sheetTo ? at2(IDENT2, s, i + q[0].length) : null;
          if (!id2 || s[i + q[0].length + id2[0].length] === "(") throw new Unsupported("sheet-qualified name");
          out.push({ k: "ident", v: `${sheet}!${id2[0]}` });
          i += q[0].length + id2[0].length;
          continue;
        }
        out.push({ k: "ref", ref: r2.ref });
        i = r2.end;
        continue;
      }
      const r = refAt(i, null);
      if (r) {
        out.push({ k: "ref", ref: r.ref });
        i = r.end;
        continue;
      }
      const n = at2(NUMBER2, s, i);
      if (n) {
        out.push({ k: "num", v: Number(n[0]) });
        i += n[0].length;
        continue;
      }
      const id = at2(IDENT2, s, i);
      if (id) {
        const word = id[0];
        const end = i + word.length;
        if (s[end] === "(") {
          out.push({ k: "func", v: word.toUpperCase().replace(/^_XLFN\.(?:_XLWS\.)?/, "").replace(/^_XLL\./, "") });
          i = end + 1;
          continue;
        }
        if (s[end] === "[") {
          let j = end;
          let depth = 0;
          for (; j < s.length; j++) {
            if (s[j] === "'") j++;
            else if (s[j] === "[") depth++;
            else if (s[j] === "]" && --depth === 0) break;
          }
          if (depth) throw new Unsupported("bad structured reference");
          out.push({ k: "struct", table: word, spec: s.slice(end, j + 1) });
          i = j + 1;
          continue;
        }
        out.push({ k: "ident", v: word });
        i = end;
        continue;
      }
    }
    const two = s.slice(i, i + 2);
    if (two === "<=" || two === ">=" || two === "<>") {
      out.push({ k: "op", v: two });
      i += 2;
      continue;
    }
    if ("+-*/^&=<>%:".includes(ch)) {
      out.push({ k: "op", v: ch });
      i++;
      continue;
    }
    if (ch === "(" || ch === ")" || ch === "," || ch === ";" || ch === "{" || ch === "}") {
      out.push({ k: ch });
      i++;
      continue;
    }
    throw new Unsupported(`unexpected ${ch}`);
  }
  out.push({ k: "eof" });
  return out;
}
function relativeKey(toks, row, col2) {
  let k = "";
  for (const t of toks) {
    switch (t.k) {
      case "ref": {
        const r = t.ref;
        const rel = r.rel ?? 0;
        const f = (v, bit, base) => rel & bit ? `~${v - base}` : `${v}`;
        k += `${r.sheet ?? ""}:${r.sheetTo ?? ""}:${f(r.r0, 1, row)},${f(r.c0, 2, col2)},${f(r.r1, 4, row)},${f(r.c1, 8, col2)}`;
        break;
      }
      case "num":
        k += `${t.v}`;
        break;
      case "str":
        k += `${t.v.replace(/\u0003/g, "")}`;
        break;
      case "err":
      case "ident":
      case "func":
      case "op":
        k += `${t.k}${t.v}`;
        break;
      case "struct":
        k += `${t.table}${t.spec}`;
        break;
      default:
        k += t.k;
    }
  }
  return k;
}
var BINARY = { "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5 };
function parseFormula(formula) {
  return parseTokens(lex(formula));
}
function parseTokens(toks) {
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (k) => {
    const t = next();
    if (t.k !== k) throw new Unsupported(`expected ${k}`);
    return t;
  };
  const primary = () => {
    const t = next();
    switch (t.k) {
      case "num":
        return { t: "num", v: t.v };
      case "str":
        return { t: "str", v: t.v };
      case "err":
        return { t: "err", v: t.v };
      case "ref":
        return { t: "ref", ref: t.ref };
      case "struct":
        return { t: "struct", table: t.table, spec: t.spec };
      case "ident": {
        const u = t.v.toUpperCase();
        if (u === "TRUE" || u === "FALSE") return { t: "bool", v: u === "TRUE" };
        return { t: "name", name: t.v };
      }
      case "func": {
        const args = [];
        if (peek().k === ")") {
          next();
          return { t: "call", name: t.v, args };
        }
        for (; ; ) {
          if (peek().k === "," || peek().k === ")") args.push({ t: "missing" });
          else args.push(expr(0));
          const sep = next();
          if (sep.k === ")") break;
          if (sep.k !== ",") throw new Unsupported("bad argument list");
        }
        return { t: "call", name: t.v, args };
      }
      case "(": {
        const e = expr(0);
        expect(")");
        return e;
      }
      case "{": {
        const rows = [[]];
        for (; ; ) {
          rows[rows.length - 1].push(unary());
          const sep = next();
          if (sep.k === "}") break;
          if (sep.k === ";") rows.push([]);
          else if (sep.k !== ",") throw new Unsupported("bad array constant");
        }
        return { t: "array", rows };
      }
      case "op":
        if (t.v === "-" || t.v === "+") return { t: "unary", op: t.v, arg: unary() };
        throw new Unsupported(`unexpected ${t.v}`);
      default:
        throw new Unsupported(`unexpected ${t.k}`);
    }
  };
  const rangeOp = () => {
    let l = primary();
    while (peek().k === "op" && peek().v === ":") {
      next();
      const r = primary();
      l = { t: "binary", op: ":", l, r };
    }
    return l;
  };
  const postfix = () => {
    let n = rangeOp();
    while (peek().k === "op" && peek().v === "%") {
      next();
      n = { t: "unary", op: "%", arg: n };
    }
    return n;
  };
  const unary = () => {
    const t = peek();
    if (t.k === "op" && (t.v === "-" || t.v === "+")) {
      next();
      return { t: "unary", op: t.v, arg: unary() };
    }
    return postfix();
  };
  const expr = (min) => {
    let l = unary();
    for (; ; ) {
      const t = peek();
      if (t.k !== "op") break;
      const prec = BINARY[t.v];
      if (prec === void 0 || prec < min) break;
      next();
      const r = expr(prec + 1);
      l = { t: "binary", op: t.v, l, r };
    }
    return l;
  };
  const tree = expr(0);
  if (peek().k !== "eof") throw new Unsupported("trailing input");
  return tree;
}

// packages/spreadiff/src/core/engine/values.ts
var XlError = class {
  constructor(code) {
    this.code = code;
  }
  toString() {
    return this.code;
  }
};
var ERR = {
  div0: new XlError("#DIV/0!"),
  na: new XlError("#N/A"),
  value: new XlError("#VALUE!"),
  ref: new XlError("#REF!"),
  num: new XlError("#NUM!"),
  name: new XlError("#NAME?"),
  null: new XlError("#NULL!")
};
var BY_CODE = new Map(Object.values(ERR).map((e) => [e.code, e]));
var errorOf = (code) => BY_CODE.get(code) ?? new XlError(code);
var isError = (v) => v instanceof XlError;
var isRange = (v) => typeof v === "object" && v !== null && v.kind === "range";
var isArray = (v) => typeof v === "object" && v !== null && v.kind === "array";
function textToNumber(s) {
  let t = s.trim();
  if (t === "") return null;
  let neg = false;
  if (t.startsWith("(") && t.endsWith(")")) {
    neg = true;
    t = t.slice(1, -1).trim();
  }
  const pct = t.endsWith("%");
  if (pct) t = t.slice(0, -1).trim();
  const m = /^([+-]?)\$?\s*((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?$/.exec(t);
  if (!m) return null;
  if (neg && m[1]) return null;
  let n = Number(`${m[1]}${m[2].replace(/,/g, "")}${m[3] !== void 0 ? `e${m[3]}` : ""}`);
  if (!Number.isFinite(n)) return null;
  if (neg) n = -n;
  return pct ? n / 100 : n;
}
var MONTHS2 = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
var monthOf = (w) => {
  const i = MONTHS2.indexOf(w.slice(0, 3).toLowerCase());
  return i >= 0 && (w.length === 3 || "january february march april may june july august september october november december".split(" ")[i].startsWith(w.toLowerCase())) ? i + 1 : 0;
};
var TIME_TEXT = /^(\d{1,4}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp][Mm])?$/;
function dateTextToSerial(s) {
  const t = s.trim().replace(/\s+/g, " ");
  if (!t || !/\d/.test(t)) return null;
  const time = (x) => {
    const m2 = TIME_TEXT.exec(x);
    if (!m2) return null;
    let h = Number(m2[1]);
    const min = Number(m2[2]);
    const sec = m2[3] ? Number(m2[3]) : 0;
    if (m2[4]) {
      if (h < 1 || h > 12) return null;
      h = h % 12 + (m2[4].toLowerCase() === "pm" ? 12 : 0);
    }
    if (min > 59 || sec >= 60) return null;
    return (h * 3600 + min * 60 + sec) / 86400;
  };
  const year = (y) => y.length <= 2 ? (Number(y) < 30 ? 2e3 : 1900) + Number(y) : Number(y);
  const date = (y, mo, d) => {
    if (!(mo >= 1 && mo <= 12 && d >= 1 && y >= 1900 && y <= 9999)) return null;
    const last = new Date(Date.UTC(y, mo, 0)).getUTCDate() + (y === 1900 && mo === 2 ? 1 : 0);
    if (d > last) return null;
    return y === 1900 && mo === 2 && d === 29 ? 60 : serialFromParts(y, mo, d);
  };
  const onlyTime = time(t);
  if (onlyTime !== null) return onlyTime;
  if (/^\d{1,4}\/\d{1,2}\/\d{1,4}(?:\s|$)/.test(t)) throw new Unsupported("locale-dependent date text");
  let m;
  let day = null;
  let rest = "";
  if (m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](.+))?$/.exec(t)) {
    day = date(Number(m[1]), Number(m[2]), Number(m[3]));
    rest = m[4] ?? "";
  } else if (m = /^(\d{1,2})[- ]([A-Za-z]+)[- ](\d{2,4})(?: (.+))?$/.exec(t)) {
    day = monthOf(m[2]) ? date(year(m[3]), monthOf(m[2]), Number(m[1])) : null;
    rest = m[4] ?? "";
  } else if (m = /^([A-Za-z]+) (\d{1,2}),? (\d{4})(?: (.+))?$/.exec(t)) {
    day = monthOf(m[1]) ? date(Number(m[3]), monthOf(m[1]), Number(m[2])) : null;
    rest = m[4] ?? "";
  } else return null;
  if (day === null) return null;
  if (!rest) return day;
  const tm = time(rest);
  return tm === null ? null : day + tm;
}
function toNumber2(v) {
  if (typeof v === "number") return v;
  if (v === null) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (isError(v)) return v;
  const n = textToNumber(v) ?? dateTextToSerial(v);
  return n === null ? ERR.value : n;
}
function toBool(v) {
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v === null) return false;
  if (isError(v)) return v;
  const u = v.toUpperCase();
  if (u === "TRUE") return true;
  if (u === "FALSE") return false;
  return ERR.value;
}
function numberToText(n) {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const s = Math.abs(n) >= 1e15 ? Number(n.toPrecision(15)).toExponential() : String(Number(n.toPrecision(15)));
  if (!s.includes("e")) return s;
  const [mant, exp] = s.split("e");
  return `${mant}E${exp.startsWith("-") ? "-" : "+"}${exp.replace(/^[+-]/, "").padStart(2, "0")}`;
}
function toText(v) {
  if (typeof v === "string") return v;
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (isError(v)) return v;
  return numberToText(v);
}
var PLAIN = /^[a-z0-9 ]*$/;
var collator = null;
function compareText(a, b) {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  if (x === y) return 0;
  if (PLAIN.test(x) && PLAIN.test(y)) return x < y ? -1 : 1;
  collator ??= new Intl.Collator("en", { sensitivity: "accent", usage: "sort" });
  const c = collator.compare(x, y);
  return c < 0 ? -1 : c > 0 ? 1 : x < y ? -1 : 1;
}
var rank = (v) => typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2;
function compare(a, b) {
  if (a === null && b === null) return 0;
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra < rb ? -1 : 1;
  if (typeof a === "number") return a < b ? -1 : a > b ? 1 : 0;
  if (typeof a === "string") return compareText(a, b);
  return a === b ? 0 : a ? 1 : -1;
}
function sameResult(a, b, date19042 = false) {
  if (a === b) return true;
  if (typeof a === "string" && typeof b === "number") a = isoToSerial(a, date19042) ?? a;
  else if (typeof a === "number" && typeof b === "string") b = isoToSerial(b, date19042) ?? b;
  if (typeof a === "number" && typeof b === "number") {
    const d = Math.abs(a - b);
    return d <= 1e-9 * Math.max(Math.abs(a), Math.abs(b)) || d <= 1e-14;
  }
  if (isError(a) && isError(b)) return a.code === b.code;
  if (a === "" && b === null || a === null && b === "") return true;
  return false;
}

// packages/spreadiff/src/core/engine/functions.ts
var arity = (args, min, max2 = min) => {
  if (args.length < min || args.length > max2) throw new Unsupported("argument count");
};
var sc = (ctx, n) => n === void 0 ? null : ctx.scalar(ctx.ev(n));
var missing = (n) => n === void 0 || n.t === "missing";
function num(ctx, n, dflt) {
  if (missing(n)) {
    if (dflt === void 0) return ERR.value;
    return dflt;
  }
  return toNumber2(sc(ctx, n));
}
function bool(ctx, n, dflt) {
  if (missing(n)) return dflt;
  return toBool(sc(ctx, n));
}
function text(ctx, n) {
  return toText(sc(ctx, n));
}
var firstError = (...xs) => xs.find(isError);
function numbers(ctx, args, opts = {}) {
  const out = [];
  for (const a of args) {
    if (a.t === "missing") continue;
    const v = ctx.ev(a);
    if (isRange(v) || isArray(v)) {
      for (const x of ctx.values(v)) {
        if (typeof x === "number") out.push(x);
        else if (isError(x)) return x;
      }
    } else {
      if (isError(v)) return v;
      if (v === null) {
        if (opts.countText) continue;
        out.push(0);
        continue;
      }
      const n = toNumber2(v);
      if (isError(n)) return n;
      out.push(n);
    }
  }
  return out;
}
var agg = (f) => (args, ctx) => {
  const xs = numbers(ctx, args);
  return isError(xs) ? xs : f(xs);
};
var sum = (xs) => {
  let s = 0;
  for (const x of xs) s += x;
  return s;
};
var mean = (xs) => xs.length ? sum(xs) / xs.length : ERR.div0;
var variance = (xs, sample) => {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return ERR.div0;
  const m = sum(xs) / n;
  let s = 0;
  for (const x of xs) s += (x - m) * (x - m);
  return s / (sample ? n - 1 : n);
};
var sqrtOr = (v) => isError(v) ? v : Math.sqrt(v);
function roundHalfAway2(x, digits) {
  const f = 10 ** digits;
  const y = Number((Math.abs(x) * f).toPrecision(15));
  return Math.sign(x) * Math.round(y) / f;
}
var math1 = (f) => (args, ctx) => {
  arity(args, 1);
  const x = num(ctx, args[0]);
  if (isError(x)) return x;
  const r = f(x);
  return Number.isFinite(r) ? r : ERR.num;
};
function wildcard(pattern) {
  let re = "";
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "~" && i + 1 < pattern.length) re += escapeRe(pattern[++i]);
    else if (ch === "*") re += "[\\s\\S]*";
    else if (ch === "?") re += "[\\s\\S]";
    else re += escapeRe(ch);
  }
  return new RegExp(`^${re}$`, "i");
}
var escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
var same = (a, b) => typeof a === "string" && typeof b === "string" ? a === b || a.toLowerCase() === b.toLowerCase() : compare(a, b) === 0;
function criterion(c) {
  if (typeof c === "number") return (v) => typeof v === "number" ? v === c : typeof v === "string" && textToNumber(v) === c;
  if (typeof c === "boolean") return (v) => v === c;
  if (c === null) return (v) => v === null || v === "";
  if (isError(c)) return (v) => isError(v) && v.code === c.code;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c);
  const op = m[1] ?? "";
  const rhs = m[2];
  const n = textToNumber(rhs) ?? (/\d/.test(rhs) && !/[*?~]/.test(rhs) ? dateTextToSerial(rhs) : null);
  if (op === "" || op === "=") {
    if (rhs === "") return op === "=" ? (v) => v === null || v === "" : (v) => v === null || v === "";
    if (n !== null) return (v) => typeof v === "number" ? v === n : typeof v === "string" && textToNumber(v) === n;
    const lower = rhs.toLowerCase();
    if (/[*?~]/.test(rhs)) {
      const re = wildcard(rhs);
      return (v) => typeof v === "string" && re.test(v);
    }
    if (lower === "true" || lower === "false") return (v) => v === (lower === "true");
    return (v) => typeof v === "string" && v.toLowerCase() === lower;
  }
  if (op === "<>") {
    if (rhs === "") return (v) => v !== null && v !== "";
    if (n !== null) return (v) => !(typeof v === "number" && v === n);
    if (/[*?~]/.test(rhs)) {
      const re = wildcard(rhs);
      return (v) => !(typeof v === "string" && re.test(v));
    }
    const lower = rhs.toLowerCase();
    return (v) => !(typeof v === "string" && v.toLowerCase() === lower);
  }
  const test2 = (cmp) => op === "<" ? cmp < 0 : op === ">" ? cmp > 0 : op === "<=" ? cmp <= 0 : cmp >= 0;
  if (n !== null) return (v) => typeof v === "number" && test2(compare(v, n));
  return (v) => typeof v === "string" && test2(compare(v, rhs));
}
function matchingIndexes(ctx, pairs) {
  let dims = null;
  let idx = null;
  for (const [rn, cn] of pairs) {
    const rv = ctx.ev(rn);
    const d = ctx.dims(rv);
    if (dims && (d.rows !== dims.rows || d.cols !== dims.cols)) return ERR.value;
    dims = d;
    const vals = ctx.values(rv);
    const test2 = criterion(sc(ctx, cn));
    const next = [];
    if (idx === null) {
      for (let i = 0; i < vals.length; i++) if (test2(vals[i])) next.push(i);
    } else {
      for (const i of idx) if (test2(vals[i])) next.push(i);
    }
    idx = next;
  }
  return { dims, idx };
}
function pickAt(ctx, target, dims, idx) {
  return idx.map((i) => ctx.at(target, Math.floor(i / dims.cols), i % dims.cols));
}
function ifsAggregate(kind) {
  return (args, ctx) => {
    const hasTarget = kind !== "count";
    if (args.length < (hasTarget ? 3 : 2) || (args.length - (hasTarget ? 1 : 0)) % 2) throw new Unsupported("argument count");
    const pairs = [];
    for (let i = hasTarget ? 1 : 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const m = matchingIndexes(ctx, pairs);
    if (isError(m)) return m;
    if (kind === "count") return m.idx.length;
    const target = ctx.ev(args[0]);
    const td2 = ctx.dims(target);
    if (td2.rows !== m.dims.rows || td2.cols !== m.dims.cols) return ERR.value;
    const xs = [];
    for (const v of pickAt(ctx, target, m.dims, m.idx)) {
      if (isError(v)) return v;
      if (typeof v === "number") xs.push(v);
    }
    if (kind === "sum") return sum(xs);
    if (kind === "avg") return mean(xs);
    if (!xs.length) return 0;
    return kind === "max" ? Math.max(...xs) : Math.min(...xs);
  };
}
function ifAggregate(kind) {
  return (args, ctx) => {
    arity(args, 2, 3);
    const m = matchingIndexes(ctx, [[args[0], args[1]]]);
    if (isError(m)) return m;
    let target = ctx.ev(missing(args[2]) ? args[0] : args[2]);
    if (isRange(target)) target = { ...target, r1: target.r0 + m.dims.rows - 1, c1: target.c0 + m.dims.cols - 1 };
    const xs = [];
    for (const v of pickAt(ctx, target, m.dims, m.idx)) {
      if (isError(v)) return v;
      if (typeof v === "number") xs.push(v);
    }
    return kind === "sum" ? sum(xs) : mean(xs);
  };
}
var indexKey = (v) => typeof v === "number" ? `n${v}` : typeof v === "string" ? `s${v.toLowerCase()}` : typeof v === "boolean" ? `b${v}` : null;
function exactIndex(vals, key, wild, ctx) {
  if (typeof key === "string" && wild && /[*?~]/.test(key)) {
    const re = wildcard(key);
    return vals.findIndex((v) => typeof v === "string" && re.test(v));
  }
  if (ctx && vals.length > 32) {
    let idx = ctx.engine.indexes.get(vals);
    if (!idx) {
      idx = /* @__PURE__ */ new Map();
      for (let i = 0; i < vals.length; i++) {
        const k2 = indexKey(vals[i]);
        if (k2 !== null && !idx.has(k2)) idx.set(k2, i);
      }
      ctx.engine.indexes.set(vals, idx);
    }
    const k = indexKey(key);
    return k === null ? -1 : idx.get(k) ?? -1;
  }
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i];
    if (v === null) continue;
    if (typeof v === typeof key && same(v, key)) return i;
  }
  return -1;
}
function approxIndex(vals, key, descending = false) {
  let lo = 0;
  let hi = vals.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = lo + hi >> 1;
    const v = vals[mid];
    const c = v === null || typeof v !== typeof key ? typeof v === "number" ? -1 : 1 : compare(v, key);
    const ok = descending ? c >= 0 : c <= 0;
    if (ok) {
      if (v !== null && typeof v === typeof key) found = mid;
      lo = mid + 1;
    } else hi = mid - 1;
  }
  return found;
}
function vhlookup(horizontal) {
  return (args, ctx) => {
    arity(args, 3, 4);
    const key = sc(ctx, args[0]);
    if (isError(key)) return key;
    const table = ctx.ev(args[1]);
    const index = num(ctx, args[2]);
    if (isError(index)) return index;
    const approx = bool(ctx, args[3], true);
    if (isError(approx)) return approx;
    const d = ctx.dims(table);
    const k = Math.trunc(index);
    if (k < 1) return ERR.value;
    if (k > (horizontal ? d.rows : d.cols)) return ERR.ref;
    const keys = horizontal ? sliceRow(ctx, table, 0) : sliceCol(ctx, table, 0);
    const i = approx ? approxIndex(keys, key) : exactIndex(keys, key, true, ctx);
    if (i < 0) return ERR.na;
    return horizontal ? ctx.at(table, k - 1, i) : ctx.at(table, i, k - 1);
  };
}
function sliceCol(ctx, v, c) {
  if (isRange(v)) return ctx.values({ ...v, c0: v.c0 + c, c1: v.c0 + c });
  const d = ctx.dims(v);
  const out = [];
  for (let r = 0; r < d.rows; r++) out.push(ctx.at(v, r, c));
  return out;
}
function sliceRow(ctx, v, r) {
  if (isRange(v)) return ctx.values({ ...v, r0: v.r0 + r, r1: v.r0 + r });
  const d = ctx.dims(v);
  const out = [];
  for (let c = 0; c < d.cols; c++) out.push(ctx.at(v, r, c));
  return out;
}
function vector(ctx, v) {
  const d = ctx.dims(v);
  if (d.rows !== 1 && d.cols !== 1) return ERR.na;
  return ctx.values(v);
}
function matchFn(args, ctx) {
  arity(args, 2, 3);
  const key = sc(ctx, args[0]);
  if (isError(key)) return key;
  const vals = vector(ctx, ctx.ev(args[1]));
  if (isError(vals)) return vals;
  const type = num(ctx, args[2], 1);
  if (isError(type)) return type;
  const i = type === 0 ? exactIndex(vals, key, true, ctx) : approxIndex(vals, key, type < 0);
  return i < 0 ? ERR.na : i + 1;
}
function indexFn(args, ctx) {
  arity(args, 1, 3);
  const src = ctx.ev(args[0]);
  if (isError(src)) return src;
  const d = ctx.dims(src);
  let r = num(ctx, args[1], 0);
  if (isError(r)) return r;
  let c = num(ctx, args[2], 0);
  if (isError(c)) return c;
  r = Math.trunc(r);
  c = Math.trunc(c);
  if (missing(args[2]) && d.rows === 1 && d.cols > 1) {
    c = r;
    r = 1;
  }
  if (r < 0 || c < 0 || r > d.rows || c > d.cols) return ERR.ref;
  if (isRange(src)) {
    const r0 = r === 0 ? src.r0 : src.r0 + r - 1;
    const r1 = r === 0 ? src.r1 : r0;
    const c0 = c === 0 ? src.c0 : src.c0 + c - 1;
    const c1 = c === 0 ? src.c1 : c0;
    if (r === 0 && c === 0 && d.cols > 1 && d.rows > 1) return src;
    return { kind: "range", sheet: src.sheet, r0, c0, r1, c1 };
  }
  if (r === 0 && c === 0) return src;
  if (r === 0 || c === 0) {
    const a = ctx.array(src);
    if (r === 0) return { kind: "array", rows: a.rows, cols: 1, data: Array.from({ length: a.rows }, (_, i) => a.data[i * a.cols + c - 1]) };
    return { kind: "array", rows: 1, cols: a.cols, data: a.data.slice((r - 1) * a.cols, r * a.cols) };
  }
  return ctx.at(src, r - 1, c - 1);
}
function xlookup(args, ctx) {
  arity(args, 3, 6);
  const key = sc(ctx, args[0]);
  if (isError(key)) return key;
  const look = ctx.ev(args[1]);
  const ret = ctx.ev(args[2]);
  const mode = num(ctx, args[4], 0);
  const search = num(ctx, args[5], 1);
  if (isError(mode) || isError(search)) return ERR.value;
  if (search !== 1 && search !== -1) throw new Unsupported("XLOOKUP binary search");
  const vals = vector(ctx, look);
  if (isError(vals)) return vals;
  const order = vals.map((_, i) => i);
  if (search === -1) order.reverse();
  let hit = -1;
  if (mode === 0 || mode === 2) {
    const re = mode === 2 && typeof key === "string" ? wildcard(key) : null;
    hit = order.find((i) => re ? typeof vals[i] === "string" && re.test(vals[i]) : vals[i] !== null && typeof vals[i] === typeof key && same(vals[i], key)) ?? -1;
  } else if (mode === -1 || mode === 1) {
    let best = -1;
    for (const i of order) {
      const v = vals[i];
      if (v === null || typeof v !== typeof key) continue;
      const c = compare(v, key);
      if (c === 0) {
        best = i;
        break;
      }
      if (mode === -1 ? c < 0 : c > 0) {
        if (best < 0 || (mode === -1 ? compare(v, vals[best]) > 0 : compare(v, vals[best]) < 0)) best = i;
      }
    }
    hit = best;
  } else throw new Unsupported("XLOOKUP match mode");
  if (hit < 0) return missing(args[3]) ? ERR.na : ctx.ev(args[3]);
  const ld = ctx.dims(look);
  const rd = ctx.dims(ret);
  if (ld.cols === 1) {
    if (isRange(ret)) return { kind: "range", sheet: ret.sheet, r0: ret.r0 + hit, r1: ret.r0 + hit, c0: ret.c0, c1: ret.c1 };
    return rd.cols === 1 ? ctx.at(ret, hit, 0) : ERR.value;
  }
  if (isRange(ret)) return { kind: "range", sheet: ret.sheet, r0: ret.r0, r1: ret.r1, c0: ret.c0 + hit, c1: ret.c0 + hit };
  return rd.rows === 1 ? ctx.at(ret, 0, hit) : ERR.value;
}
function lookupFn(args, ctx) {
  arity(args, 2, 3);
  const key = sc(ctx, args[0]);
  if (isError(key)) return key;
  const look = ctx.ev(args[1]);
  const vals = vector(ctx, look);
  if (isError(vals)) throw new Unsupported("LOOKUP array form");
  const i = approxIndex(vals, key);
  if (i < 0) return ERR.na;
  const res = missing(args[2]) ? look : ctx.ev(args[2]);
  const d = ctx.dims(res);
  return d.cols === 1 ? ctx.at(res, i, 0) : ctx.at(res, 0, i);
}
function refArg(ctx, n) {
  if (missing(n)) return { kind: "range", sheet: ctx.sheet, r0: ctx.row, c0: ctx.col, r1: ctx.row, c1: ctx.col };
  const v = ctx.ev(n);
  if (isError(v)) return v;
  if (!isRange(v)) return ERR.value;
  return v;
}
function offsetFn(args, ctx) {
  arity(args, 3, 5);
  const base = refArg(ctx, args[0]);
  if (isError(base)) return base;
  const dr = num(ctx, args[1]);
  const dc = num(ctx, args[2]);
  const h = num(ctx, args[3], base.r1 - base.r0 + 1);
  const w = num(ctx, args[4], base.c1 - base.c0 + 1);
  const e = firstError(dr, dc, h, w);
  if (e) return e;
  const r0 = base.r0 + Math.trunc(dr);
  const c0 = base.c0 + Math.trunc(dc);
  const hh = Math.trunc(h);
  const ww = Math.trunc(w);
  if (r0 < 0 || c0 < 0 || hh < 1 || ww < 1) return ERR.ref;
  return { kind: "range", sheet: base.sheet, r0, c0, r1: r0 + hh - 1, c1: c0 + ww - 1 };
}
function indirectFn(args, ctx) {
  arity(args, 1, 2);
  const t = text(ctx, args[0]);
  if (isError(t)) return t;
  const a1 = bool(ctx, args[1], true);
  if (a1 !== true) throw new Unsupported("INDIRECT R1C1");
  let node;
  try {
    node = parseFormula(t);
  } catch {
    if (lookupName(ctx.host.namesFor ? ctx.host.namesFor(ctx.sheet) : ctx.host.names, t) !== void 0) node = { t: "name", name: t };
    else return ERR.ref;
  }
  if (node.t !== "ref" && node.t !== "name" && !(node.t === "binary" && node.op === ":")) return ERR.ref;
  const v = ctx.ev(node);
  return isRange(v) ? v : ERR.ref;
}
var date1904 = (ctx) => ctx.host.date1904 ?? false;
function datePart(part) {
  return (args, ctx) => {
    arity(args, part === "dow" ? 1 : 1, part === "dow" ? 2 : 1);
    const s = num(ctx, args[0]);
    if (isError(s)) return s;
    if (s < 0) return ERR.num;
    const p = partsFromSerial(s, date1904(ctx));
    if (part !== "dow") return p[part];
    const type = num(ctx, args[1], 1);
    if (isError(type)) return type;
    const dow = partsFromSerial(s, date1904(ctx)).dow;
    if (type === 1) return dow + 1;
    if (type === 2) return (dow + 6) % 7 + 1;
    if (type === 3) return (dow + 6) % 7;
    throw new Unsupported("WEEKDAY type");
  };
}
function edate(endOfMonth) {
  return (args, ctx) => {
    arity(args, 2);
    const s = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const e = firstError(s, m);
    if (e) return e;
    const p = partsFromSerial(s, date1904(ctx));
    const months = Math.trunc(m);
    if (endOfMonth) return serialFromParts(p.y, p.m + months + 1, 0, date1904(ctx));
    const last = partsFromSerial(serialFromParts(p.y, p.m + months + 1, 0, date1904(ctx)), date1904(ctx)).d;
    return serialFromParts(p.y, p.m + months, Math.min(p.d, last), date1904(ctx));
  };
}
function formatValue(v, format, d1904 = false) {
  const s = formatNumberCode(v, format, d1904);
  if (s === null) throw new Unsupported("number format");
  return s;
}
var formatText = (v, format) => formatTextCode(v, format) ?? v;
function pmt(rate, nper, pv, fv, type) {
  if (rate === 0) return -(pv + fv) / nper;
  const f = (1 + rate) ** nper;
  return -(fv + pv * f) * rate / ((1 + rate * type) * (f - 1));
}
function fvOf(rate, nper, p, pv, type) {
  if (rate === 0) return -(pv + p * nper);
  const f = (1 + rate) ** nper;
  return -(pv * f + p * (1 + rate * type) * (f - 1) / rate);
}
function ipmt(rate, per, nper, pv, fv, type) {
  if (type === 1 && per === 1) return 0;
  const i = fvOf(rate, per - 1, pmt(rate, nper, pv, fv, type), pv, type) * rate;
  return type === 1 ? i / (1 + rate) : i;
}
var fin = (min, max2, f) => (args, ctx) => {
  arity(args, min, max2);
  const xs = [];
  for (let i = 0; i < max2; i++) {
    const v = num(ctx, args[i], 0);
    if (isError(v)) return v;
    xs.push(v);
  }
  const r = f(xs);
  return Number.isFinite(r) ? r : ERR.num;
};
function newton(f, guess) {
  let r = guess;
  for (let k = 0; k < 100; k++) {
    const y = f(r);
    const h = 1e-7 * Math.max(1, Math.abs(r));
    const dy = (f(r + h) - y) / h;
    if (!Number.isFinite(dy) || dy === 0) return ERR.num;
    const next = r - y / dy;
    if (Math.abs(next - r) < 1e-12) return next;
    r = next;
  }
  return ERR.num;
}
var passthrough = (f, nums = 0) => (args, ctx) => {
  arity(args, 1, 1 + nums);
  const s = text(ctx, args[0]);
  if (isError(s)) return s;
  const rest = [];
  for (let i = 1; i < args.length; i++) {
    const n = num(ctx, args[i]);
    if (isError(n)) return n;
    rest.push(n);
  }
  return f(s, ...rest);
};
var isType = (test2) => (args, ctx) => {
  arity(args, 1);
  return test2(sc(ctx, args[0]));
};
var FUNCTIONS = {
  // Logic (lazy branches)
  IF: (args, ctx) => {
    arity(args, 1, 3);
    const c = toBool(sc(ctx, args[0]));
    if (isError(c)) return c;
    if (c) return missing(args[1]) ? args.length > 1 ? 0 : true : ctx.ev(args[1]);
    return args.length < 3 ? false : missing(args[2]) ? 0 : ctx.ev(args[2]);
  },
  IFS: (args, ctx) => {
    if (args.length < 2 || args.length % 2) throw new Unsupported("argument count");
    for (let i = 0; i < args.length; i += 2) {
      const c = toBool(sc(ctx, args[i]));
      if (isError(c)) return c;
      if (c) return ctx.ev(args[i + 1]);
    }
    return ERR.na;
  },
  IFERROR: (args, ctx) => {
    arity(args, 2);
    const v = ctx.ev(args[0]);
    const s = isRange(v) || isArray(v) ? ctx.scalar(v) : v;
    return isError(s) ? ctx.ev(args[1]) : v;
  },
  IFNA: (args, ctx) => {
    arity(args, 2);
    const v = ctx.ev(args[0]);
    const s = isRange(v) || isArray(v) ? ctx.scalar(v) : v;
    return isError(s) && s.code === "#N/A" ? ctx.ev(args[1]) : v;
  },
  CHOOSE: (args, ctx) => {
    if (args.length < 2) throw new Unsupported("argument count");
    const i = num(ctx, args[0]);
    if (isError(i)) return i;
    const k = Math.trunc(i);
    return k < 1 || k >= args.length ? ERR.value : ctx.ev(args[k]);
  },
  SWITCH: (args, ctx) => {
    if (args.length < 3) throw new Unsupported("argument count");
    const v = sc(ctx, args[0]);
    if (isError(v)) return v;
    let i = 1;
    for (; i + 1 < args.length; i += 2) {
      const c = sc(ctx, args[i]);
      if (!isError(c) && typeof v === typeof c && same(v, c)) return ctx.ev(args[i + 1]);
    }
    return i < args.length ? ctx.ev(args[i]) : ERR.na;
  },
  AND: (args, ctx) => logic(args, ctx, (xs) => xs.every(Boolean)),
  OR: (args, ctx) => logic(args, ctx, (xs) => xs.some(Boolean)),
  XOR: (args, ctx) => logic(args, ctx, (xs) => xs.filter(Boolean).length % 2 === 1),
  NOT: (args, ctx) => {
    arity(args, 1);
    const b = toBool(sc(ctx, args[0]));
    return isError(b) ? b : !b;
  },
  TRUE: () => true,
  FALSE: () => false,
  // Aggregates
  SUM: agg(sum),
  AVERAGE: agg(mean),
  MIN: agg((xs) => xs.length ? Math.min(...xs) : 0),
  MAX: agg((xs) => xs.length ? Math.max(...xs) : 0),
  PRODUCT: agg((xs) => xs.reduce((a, b) => a * b, 1)),
  SUMSQ: agg((xs) => sum(xs.map((x) => x * x))),
  MEDIAN: agg((xs) => {
    if (!xs.length) return ERR.num;
    const s = [...xs].sort((a, b) => a - b);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }),
  STDEV: agg((xs) => sqrtOr(variance(xs, true))),
  "STDEV.S": agg((xs) => sqrtOr(variance(xs, true))),
  STDEVP: agg((xs) => sqrtOr(variance(xs, false))),
  "STDEV.P": agg((xs) => sqrtOr(variance(xs, false))),
  VAR: agg((xs) => variance(xs, true)),
  "VAR.S": agg((xs) => variance(xs, true)),
  VARP: agg((xs) => variance(xs, false)),
  "VAR.P": agg((xs) => variance(xs, false)),
  COUNT: (args, ctx) => {
    let n = 0;
    for (const a of args) {
      const v = ctx.ev(a);
      if (isRange(v) || isArray(v)) for (const x of ctx.values(v)) n += typeof x === "number" ? 1 : 0;
      else if (typeof v === "number" || typeof v === "boolean" || typeof v === "string" && textToNumber(v) !== null) n++;
    }
    return n;
  },
  COUNTA: (args, ctx) => {
    let n = 0;
    for (const a of args) {
      if (a.t === "missing") continue;
      const v = ctx.ev(a);
      if (isRange(v) || isArray(v)) for (const x of ctx.values(v)) n += x === null ? 0 : 1;
      else n++;
    }
    return n;
  },
  COUNTBLANK: (args, ctx) => {
    arity(args, 1);
    return ctx.values(ctx.ev(args[0])).filter((x) => x === null || x === "").length;
  },
  LARGE: (args, ctx) => nth(args, ctx, true),
  SMALL: (args, ctx) => nth(args, ctx, false),
  RANK: (args, ctx) => rank2(args, ctx),
  "RANK.EQ": (args, ctx) => rank2(args, ctx),
  SUMPRODUCT: (args, ctx) => {
    if (!args.length) throw new Unsupported("argument count");
    const arrays = args.map((a) => ctx.array(ctx.ev(a)));
    const { rows, cols } = arrays[0];
    if (arrays.some((a) => a.rows !== rows || a.cols !== cols)) return ERR.value;
    let s = 0;
    for (let i = 0; i < rows * cols; i++) {
      let p = 1;
      for (const a of arrays) {
        const v = a.data[i];
        if (isError(v)) return v;
        p *= typeof v === "number" ? v : 0;
      }
      s += p;
    }
    return s;
  },
  SUBTOTAL: (args, ctx) => {
    if (args.length < 2) throw new Unsupported("argument count");
    const code = num(ctx, args[0]);
    if (isError(code)) return code;
    const fn = { 1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV", 8: "STDEVP", 9: "SUM", 10: "VAR", 11: "VARP" }[code % 100];
    if (!fn || !(code >= 1 && code < 12) && !(code >= 101 && code < 112)) return ERR.value;
    const refs = args.slice(1).map((a) => ctx.ev(a));
    if (refs.some((v) => !isRange(v))) return FUNCTIONS[fn](args.slice(1), ctx);
    const host = ctx.host;
    const skipHidden = code > 100 ? 1 : 2;
    const kept = [];
    for (const v of refs) {
      const values = ctx.values(v);
      const w = v.c1 - v.c0 + 1;
      for (let i = 0; i < values.length; i++) {
        const r = v.r0 + Math.floor(i / w);
        const h = host.rowHidden?.(v.sheet, r) ?? 0;
        if (h && h >= skipHidden) continue;
        const f = host.formulaAt?.(v.sheet, r, v.c0 + i % w);
        if (f && /\b(?:SUBTOTAL|AGGREGATE)\s*\(/i.test(f)) continue;
        kept.push(values[i]);
      }
    }
    if (fn === "COUNTA") return kept.filter((x) => x !== null).length;
    if (fn === "COUNT") return kept.filter((x) => typeof x === "number").length;
    const err2 = kept.find(isError);
    if (err2) return err2;
    const xs = kept.filter((x) => typeof x === "number");
    switch (fn) {
      case "SUM":
        return sum(xs);
      case "AVERAGE":
        return mean(xs);
      case "MAX":
        return xs.length ? Math.max(...xs) : 0;
      case "MIN":
        return xs.length ? Math.min(...xs) : 0;
      case "PRODUCT":
        return xs.length ? xs.reduce((a, b) => a * b, 1) : 0;
      case "STDEV":
        return sqrtOr(variance(xs, true));
      case "STDEVP":
        return sqrtOr(variance(xs, false));
      case "VAR":
        return variance(xs, true);
      default:
        return variance(xs, false);
    }
  },
  SUMIF: ifAggregate("sum"),
  AVERAGEIF: ifAggregate("avg"),
  COUNTIF: (args, ctx) => {
    arity(args, 2);
    return FUNCTIONS.COUNTIFS(args, ctx);
  },
  SUMIFS: ifsAggregate("sum"),
  COUNTIFS: ifsAggregate("count"),
  AVERAGEIFS: ifsAggregate("avg"),
  MAXIFS: ifsAggregate("max"),
  MINIFS: ifsAggregate("min"),
  // Math
  ABS: math1(Math.abs),
  INT: math1(Math.floor),
  SQRT: (args, ctx) => {
    arity(args, 1);
    const x = num(ctx, args[0]);
    return isError(x) ? x : x < 0 ? ERR.num : Math.sqrt(x);
  },
  EXP: math1(Math.exp),
  LN: (args, ctx) => {
    arity(args, 1);
    const x = num(ctx, args[0]);
    return isError(x) ? x : x <= 0 ? ERR.num : Math.log(x);
  },
  LOG10: (args, ctx) => {
    arity(args, 1);
    const x = num(ctx, args[0]);
    return isError(x) ? x : x <= 0 ? ERR.num : Math.log10(x);
  },
  LOG: (args, ctx) => {
    arity(args, 1, 2);
    const x = num(ctx, args[0]);
    const b = num(ctx, args[1], 10);
    const e = firstError(x, b);
    if (e) return e;
    return x <= 0 || b <= 0 || b === 1 ? ERR.num : Math.log(x) / Math.log(b);
  },
  SIGN: math1(Math.sign),
  PI: () => Math.PI,
  SIN: math1(Math.sin),
  COS: math1(Math.cos),
  TAN: math1(Math.tan),
  ASIN: math1(Math.asin),
  ACOS: math1(Math.acos),
  ATAN: math1(Math.atan),
  SINH: math1(Math.sinh),
  COSH: math1(Math.cosh),
  TANH: math1(Math.tanh),
  ATAN2: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const y = num(ctx, args[1]);
    const e = firstError(x, y);
    if (e) return e;
    return x === 0 && y === 0 ? ERR.div0 : Math.atan2(y, x);
  },
  RADIANS: math1((x) => x * Math.PI / 180),
  DEGREES: math1((x) => x * 180 / Math.PI),
  POWER: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const y = num(ctx, args[1]);
    const e = firstError(x, y);
    if (e) return e;
    const r = x ** y;
    return Number.isFinite(r) ? r : ERR.num;
  },
  MOD: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const y = num(ctx, args[1]);
    const e = firstError(x, y);
    if (e) return e;
    if (y === 0) return ERR.div0;
    return x - y * Math.floor(x / y);
  },
  ROUND: round((x, d) => roundHalfAway2(x, d)),
  ROUNDUP: round((x, d) => {
    const f = 10 ** d;
    return Math.sign(x) * Math.ceil(Number((Math.abs(x) * f).toPrecision(15))) / f;
  }),
  ROUNDDOWN: round((x, d) => {
    const f = 10 ** d;
    return Math.sign(x) * Math.floor(Number((Math.abs(x) * f).toPrecision(15))) / f;
  }),
  TRUNC: (args, ctx) => {
    arity(args, 1, 2);
    const x = num(ctx, args[0]);
    const d = num(ctx, args[1], 0);
    const e = firstError(x, d);
    if (e) return e;
    const f = 10 ** Math.trunc(d);
    return Math.trunc(x * f) / f;
  },
  MROUND: (args, ctx) => {
    arity(args, 2);
    const x = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const e = firstError(x, m);
    if (e) return e;
    if (m === 0) return 0;
    if (Math.sign(x) * Math.sign(m) < 0) return ERR.num;
    return roundHalfAway2(x / m, 0) * m;
  },
  CEILING: (args, ctx) => ceilFloor(args, ctx, true, false),
  "CEILING.MATH": (args, ctx) => ceilFloor(args, ctx, true, true),
  FLOOR: (args, ctx) => ceilFloor(args, ctx, false, false),
  "FLOOR.MATH": (args, ctx) => ceilFloor(args, ctx, false, true),
  EVEN: math1((x) => Math.sign(x) * Math.ceil(Math.abs(x) / 2) * 2),
  ODD: math1((x) => {
    const n = Math.ceil(Math.abs(x));
    return Math.sign(x || 1) * (n % 2 ? n : n + 1);
  }),
  FACT: math1((x) => {
    if (x < 0) return NaN;
    let r = 1;
    for (let i = 2; i <= Math.floor(x); i++) r *= i;
    return r;
  }),
  // Lookup and reference
  VLOOKUP: vhlookup(false),
  HLOOKUP: vhlookup(true),
  MATCH: matchFn,
  INDEX: indexFn,
  XLOOKUP: xlookup,
  LOOKUP: lookupFn,
  OFFSET: offsetFn,
  INDIRECT: indirectFn,
  ROW: (args, ctx) => {
    arity(args, 0, 1);
    const r = refArg(ctx, args[0]);
    return isError(r) ? r : r.r0 + 1;
  },
  COLUMN: (args, ctx) => {
    arity(args, 0, 1);
    const r = refArg(ctx, args[0]);
    return isError(r) ? r : r.c0 + 1;
  },
  ROWS: (args, ctx) => {
    arity(args, 1);
    const v = ctx.ev(args[0]);
    return isError(v) ? v : ctx.dims(v).rows;
  },
  COLUMNS: (args, ctx) => {
    arity(args, 1);
    const v = ctx.ev(args[0]);
    return isError(v) ? v : ctx.dims(v).cols;
  },
  // Arrays and matrices
  TRANSPOSE: (args, ctx) => {
    arity(args, 1);
    const a = ctx.array(ctx.ev(args[0]));
    const data = new Array(a.rows * a.cols);
    for (let r = 0; r < a.rows; r++) for (let c = 0; c < a.cols; c++) data[c * a.rows + r] = a.data[r * a.cols + c];
    return { kind: "array", rows: a.cols, cols: a.rows, data };
  },
  MMULT: (args, ctx) => {
    arity(args, 2);
    const a = matrix(ctx, args[0]);
    const b = matrix(ctx, args[1]);
    if (isError(a)) return a;
    if (isError(b)) return b;
    if (a.cols !== b.rows) return ERR.value;
    const data = [];
    for (let r = 0; r < a.rows; r++) {
      for (let c = 0; c < b.cols; c++) {
        let s = 0;
        for (let k = 0; k < a.cols; k++) s += a.m[r * a.cols + k] * b.m[k * b.cols + c];
        data.push(s);
      }
    }
    return { kind: "array", rows: a.rows, cols: b.cols, data };
  },
  MDETERM: (args, ctx) => {
    arity(args, 1);
    const a = matrix(ctx, args[0]);
    if (isError(a)) return a;
    if (a.rows !== a.cols) return ERR.value;
    return gauss(a.m, a.rows, false).det;
  },
  MINVERSE: (args, ctx) => {
    arity(args, 1);
    const a = matrix(ctx, args[0]);
    if (isError(a)) return a;
    if (a.rows !== a.cols) return ERR.value;
    const g = gauss(a.m, a.rows, true);
    if (g.det === 0) return ERR.num;
    return { kind: "array", rows: a.rows, cols: a.cols, data: g.inverse };
  },
  HEX2DEC: (args, ctx) => baseToDec(args, ctx, 16),
  BIN2DEC: (args, ctx) => baseToDec(args, ctx, 2),
  OCT2DEC: (args, ctx) => baseToDec(args, ctx, 8),
  // Information
  ISBLANK: isType((v) => v === null),
  ISNUMBER: isType((v) => typeof v === "number"),
  ISTEXT: isType((v) => typeof v === "string"),
  ISNONTEXT: isType((v) => typeof v !== "string"),
  ISLOGICAL: isType((v) => typeof v === "boolean"),
  ISERROR: isType((v) => isError(v)),
  ISERR: isType((v) => isError(v) && v.code !== "#N/A"),
  ISNA: isType((v) => isError(v) && v.code === "#N/A"),
  ISEVEN: (args, ctx) => {
    const x = num(ctx, args[0]);
    return isError(x) ? x : Math.trunc(x) % 2 === 0;
  },
  ISODD: (args, ctx) => {
    const x = num(ctx, args[0]);
    return isError(x) ? x : Math.abs(Math.trunc(x) % 2) === 1;
  },
  NA: () => ERR.na,
  N: (args, ctx) => {
    arity(args, 1);
    const v = sc(ctx, args[0]);
    return typeof v === "number" ? v : typeof v === "boolean" ? v ? 1 : 0 : isError(v) ? v : 0;
  },
  T: (args, ctx) => {
    arity(args, 1);
    const v = sc(ctx, args[0]);
    return typeof v === "string" ? v : isError(v) ? v : "";
  },
  // Text
  CONCATENATE: (args, ctx) => {
    let s = "";
    for (const a of args) {
      const t = text(ctx, a);
      if (isError(t)) return t;
      s += t;
    }
    return s;
  },
  CONCAT: (args, ctx) => {
    let s = "";
    for (const a of args) {
      const v = ctx.ev(a);
      for (const x of isRange(v) || isArray(v) ? ctx.values(v) : [v]) {
        const t = toText(x);
        if (isError(t)) return t;
        s += t;
      }
    }
    return s;
  },
  TEXTJOIN: (args, ctx) => {
    if (args.length < 3) throw new Unsupported("argument count");
    const delim = text(ctx, args[0]);
    const skip = bool(ctx, args[1], true);
    if (isError(delim)) return delim;
    if (isError(skip)) return skip;
    const parts = [];
    for (const a of args.slice(2)) {
      const v = ctx.ev(a);
      for (const x of isRange(v) || isArray(v) ? ctx.values(v) : [v]) {
        const t = toText(x);
        if (isError(t)) return t;
        if (skip && t === "") continue;
        parts.push(t);
      }
    }
    return parts.join(delim);
  },
  LEN: passthrough((s) => s.length),
  UPPER: passthrough((s) => s.toUpperCase()),
  LOWER: passthrough((s) => s.toLowerCase()),
  PROPER: passthrough((s) => s.toLowerCase().replace(/(^|[^A-Za-z\u00C0-\uFFFF])([a-z\u00E0-\uFFFF])/g, (_, a, b) => a + b.toUpperCase())),
  TRIM: passthrough((s) => s.replace(/ +/g, " ").trim()),
  CLEAN: passthrough((s) => s.replace(/[\x00-\x1f]/g, "")),
  LEFT: (args, ctx) => {
    arity(args, 1, 2);
    const s = text(ctx, args[0]);
    const n = num(ctx, args[1], 1);
    const e = firstError(s, n);
    if (e) return e;
    return n < 0 ? ERR.value : s.slice(0, Math.trunc(n));
  },
  RIGHT: (args, ctx) => {
    arity(args, 1, 2);
    const s = text(ctx, args[0]);
    const n = num(ctx, args[1], 1);
    const e = firstError(s, n);
    if (e) return e;
    const k = Math.trunc(n);
    return k < 0 ? ERR.value : k === 0 ? "" : s.slice(-k);
  },
  MID: (args, ctx) => {
    arity(args, 3);
    const s = text(ctx, args[0]);
    const a = num(ctx, args[1]);
    const n = num(ctx, args[2]);
    const e = firstError(s, a, n);
    if (e) return e;
    if (a < 1 || n < 0) return ERR.value;
    const st = Math.trunc(a) - 1;
    return s.slice(st, st + Math.trunc(n));
  },
  REPT: (args, ctx) => {
    arity(args, 2);
    const s = text(ctx, args[0]);
    const n = num(ctx, args[1]);
    const e = firstError(s, n);
    if (e) return e;
    return n < 0 ? ERR.value : s.repeat(Math.trunc(n));
  },
  EXACT: (args, ctx) => {
    arity(args, 2);
    const a = text(ctx, args[0]);
    const b = text(ctx, args[1]);
    const e = firstError(a, b);
    return e ?? a === b;
  },
  SUBSTITUTE: (args, ctx) => {
    arity(args, 3, 4);
    const s = text(ctx, args[0]);
    const from = text(ctx, args[1]);
    const to = text(ctx, args[2]);
    const e = firstError(s, from, to);
    if (e) return e;
    if (from === "") return s;
    if (missing(args[3])) return s.split(from).join(to);
    const k = num(ctx, args[3]);
    if (isError(k)) return k;
    let seen = 0;
    let idx = -1;
    let pos = 0;
    while ((idx = s.indexOf(from, pos)) >= 0) {
      if (++seen === Math.trunc(k)) return s.slice(0, idx) + to + s.slice(idx + from.length);
      pos = idx + from.length;
    }
    return s;
  },
  REPLACE: (args, ctx) => {
    arity(args, 4);
    const s = text(ctx, args[0]);
    const a = num(ctx, args[1]);
    const n = num(ctx, args[2]);
    const t = text(ctx, args[3]);
    const e = firstError(s, a, n, t);
    if (e) return e;
    const st = Math.trunc(a) - 1;
    return s.slice(0, st) + t + s.slice(st + Math.trunc(n));
  },
  FIND: (args, ctx) => findFn(args, ctx, false),
  SEARCH: (args, ctx) => findFn(args, ctx, true),
  VALUE: (args, ctx) => {
    arity(args, 1);
    const v = sc(ctx, args[0]);
    if (typeof v === "number") return v;
    if (isError(v)) return v;
    const t = String(v ?? "");
    const n = textToNumber(t);
    if (n !== null) return n;
    const d = dateTextToSerial(t);
    if (d === null) return ERR.value;
    return date1904(ctx) && d >= 1 ? d - 1462 : d;
  },
  TEXT: (args, ctx) => {
    arity(args, 2);
    const v = sc(ctx, args[0]);
    const f = text(ctx, args[1]);
    if (isError(v)) return v;
    if (isError(f)) return f;
    if (typeof v === "boolean") return formatText(v ? "TRUE" : "FALSE", f);
    if (typeof v === "string") {
      const n2 = textToNumber(v);
      if (n2 === null) return formatText(v, f);
      return formatValue(n2, f, ctx.host.date1904);
    }
    const n = toNumber2(v);
    if (isError(n)) return n;
    return formatValue(n, f, ctx.host.date1904);
  },
  CHAR: (args, ctx) => {
    const n = num(ctx, args[0]);
    return isError(n) ? n : n < 1 || n > 255 ? ERR.value : String.fromCharCode(Math.trunc(n));
  },
  CODE: passthrough((s) => s.length ? s.charCodeAt(0) : ERR.value),
  FIXED: (args, ctx) => {
    arity(args, 1, 3);
    const x = num(ctx, args[0]);
    const d = num(ctx, args[1], 2);
    const noCommas = bool(ctx, args[2], false);
    const e = firstError(x, d, noCommas);
    if (e) return e;
    const s = roundHalfAway2(x, d).toFixed(Math.max(0, d));
    return noCommas ? s : s.replace(/^(-?\d+)/, (m) => m.replace(/\B(?=(\d{3})+(?!\d))/g, ","));
  },
  // Dates
  DATE: (args, ctx) => {
    arity(args, 3);
    const y = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const d = num(ctx, args[2]);
    const e = firstError(y, m, d);
    if (e) return e;
    let year = Math.trunc(y);
    if (year < 1900) year += 1900;
    const s = serialFromParts(year, Math.trunc(m), Math.trunc(d), date1904(ctx));
    return s < 0 ? ERR.num : s;
  },
  TIME: (args, ctx) => {
    arity(args, 3);
    const h = num(ctx, args[0]);
    const m = num(ctx, args[1]);
    const s = num(ctx, args[2]);
    const e = firstError(h, m, s);
    if (e) return e;
    const t = (h * 3600 + m * 60 + s) / 86400;
    return t - Math.floor(t);
  },
  YEAR: datePart("y"),
  MONTH: datePart("m"),
  DAY: datePart("d"),
  WEEKDAY: datePart("dow"),
  HOUR: math1((x) => Math.floor(Math.round((x - Math.floor(x)) * 86400) / 3600) % 24),
  MINUTE: math1((x) => Math.floor(Math.round((x - Math.floor(x)) * 86400) % 3600 / 60)),
  SECOND: math1((x) => Math.round((x - Math.floor(x)) * 86400) % 60),
  EDATE: edate(false),
  EOMONTH: edate(true),
  DAYS: (args, ctx) => {
    arity(args, 2);
    const a = num(ctx, args[0]);
    const b = num(ctx, args[1]);
    const e = firstError(a, b);
    return e ?? Math.trunc(a) - Math.trunc(b);
  },
  DATEVALUE: (args, ctx) => {
    arity(args, 1);
    const t = text(ctx, args[0]);
    if (isError(t)) return t;
    const n = dateTextToSerial(t);
    if (n === null || n < 1) return ERR.value;
    return Math.floor(n) - (date1904(ctx) ? 1462 : 0);
  },
  // Financial
  PMT: fin(3, 5, ([r, n, pv, fv, t]) => pmt(r, n, pv, fv, t)),
  FV: fin(3, 5, ([r, n, p, pv, t]) => fvOf(r, n, p, pv, t)),
  PV: fin(3, 5, ([r, n, p, fv, t]) => r === 0 ? -(fv + p * n) : -(fv + p * (1 + r * t) * ((1 + r) ** n - 1) / r) / (1 + r) ** n),
  NPER: fin(3, 5, ([r, p, pv, fv, t]) => {
    if (r === 0) return -(pv + fv) / p;
    const a = p * (1 + r * t);
    return Math.log((a - fv * r) / (a + pv * r)) / Math.log(1 + r);
  }),
  IPMT: fin(4, 6, ([r, per, n, pv, fv, t]) => ipmt(r, per, n, pv, fv, t)),
  PPMT: fin(4, 6, ([r, per, n, pv, fv, t]) => pmt(r, n, pv, fv, t) - ipmt(r, per, n, pv, fv, t)),
  NPV: (args, ctx) => {
    if (args.length < 2) throw new Unsupported("argument count");
    const r = num(ctx, args[0]);
    if (isError(r)) return r;
    const xs = numbers(ctx, args.slice(1), { countText: true });
    if (isError(xs)) return xs;
    let s = 0;
    xs.forEach((x, i) => s += x / (1 + r) ** (i + 1));
    return s;
  },
  IRR: (args, ctx) => {
    arity(args, 1, 2);
    const xs = numbers(ctx, [args[0]]);
    if (isError(xs)) return xs;
    const g = num(ctx, args[1], 0.1);
    if (isError(g)) return g;
    return newton((r) => xs.reduce((s, x, i) => s + x / (1 + r) ** i, 0), g);
  },
  RATE: (args, ctx) => {
    arity(args, 3, 6);
    const v = [];
    for (let i = 0; i < 6; i++) {
      const x = num(ctx, args[i], i === 5 ? 0.1 : 0);
      if (isError(x)) return x;
      v.push(x);
    }
    const [n, p, pv, fv, t, g] = v;
    return newton((r) => r === 0 ? pv + p * n + fv : pv * (1 + r) ** n + p * (1 + r * t) * ((1 + r) ** n - 1) / r + fv, g);
  }
};
function logic(args, ctx, f) {
  const xs = [];
  for (const a of args) {
    const v = ctx.ev(a);
    if (isRange(v) || isArray(v)) {
      for (const x of ctx.values(v)) {
        if (isError(x)) return x;
        if (typeof x === "boolean") xs.push(x);
        else if (typeof x === "number") xs.push(x !== 0);
      }
    } else {
      const b = toBool(v);
      if (isError(b)) return b;
      xs.push(b);
    }
  }
  return xs.length ? f(xs) : ERR.value;
}
function round(f) {
  return (args, ctx) => {
    arity(args, 1, 2);
    const x = num(ctx, args[0]);
    const d = num(ctx, args[1], 0);
    const e = firstError(x, d);
    if (e) return e;
    return f(x, Math.trunc(d));
  };
}
function ceilFloor(args, ctx, up, math) {
  arity(args, 1, math ? 3 : 2);
  const x = num(ctx, args[0]);
  const s = num(ctx, args[1], 1);
  const mode = math ? num(ctx, args[2], 0) : 0;
  const e = firstError(x, s, mode);
  if (e) return e;
  const xv = x;
  const sv = s;
  if (sv === 0) return !math && !up ? ERR.div0 : 0;
  if (!math && xv > 0 && sv < 0) return ERR.num;
  const step = Math.abs(sv);
  const q = Number((xv / step).toPrecision(15));
  let toCeil;
  if (math) toCeil = xv >= 0 ? up : up === (mode === 0);
  else toCeil = sv > 0 ? up : !up;
  const r = (toCeil ? Math.ceil(q) : Math.floor(q)) * step;
  return r === 0 ? 0 : r;
}
function nth(args, ctx, largest) {
  arity(args, 2);
  const xs = numbers(ctx, [args[0]]);
  if (isError(xs)) return xs;
  const k = num(ctx, args[1]);
  if (isError(k)) return k;
  const i = Math.ceil(k) - 1;
  if (i < 0 || i >= xs.length) return ERR.num;
  const s = [...xs].sort((a, b) => largest ? b - a : a - b);
  return s[i];
}
function rank2(args, ctx) {
  arity(args, 2, 3);
  const x = num(ctx, args[0]);
  if (isError(x)) return x;
  const xs = numbers(ctx, [args[1]]);
  if (isError(xs)) return xs;
  const asc = num(ctx, args[2], 0);
  if (isError(asc)) return asc;
  if (!xs.includes(x)) return ERR.na;
  return 1 + xs.filter((v) => asc ? v < x : v > x).length;
}
function findFn(args, ctx, insensitive) {
  arity(args, 2, 3);
  const needle = text(ctx, args[0]);
  const hay = text(ctx, args[1]);
  const start = num(ctx, args[2], 1);
  const e = firstError(needle, hay, start);
  if (e) return e;
  const from = Math.trunc(start) - 1;
  if (from < 0 || from > hay.length) return ERR.value;
  if (insensitive) {
    const src = wildcard(needle).source.slice(1, -1);
    const m = new RegExp(src, "i").exec(hay.slice(from));
    return m ? from + m.index + 1 : ERR.value;
  }
  const i = hay.indexOf(needle, from);
  return i < 0 ? ERR.value : i + 1;
}
function matrix(ctx, n) {
  const a = ctx.array(ctx.ev(n));
  const m = [];
  for (const v of a.data) {
    if (isError(v)) return v;
    if (typeof v !== "number") return ERR.value;
    m.push(v);
  }
  return { rows: a.rows, cols: a.cols, m };
}
function gauss(src, n, invert) {
  const a = [...src];
  const inv = invert ? Array.from({ length: n * n }, (_, i) => i % (n + 1) === 0 ? 1 : 0) : [];
  let det = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r * n + c]) > Math.abs(a[p * n + c])) p = r;
    if (a[p * n + c] === 0) return { det: 0 };
    if (p !== c) {
      det = -det;
      for (let k = 0; k < n; k++) {
        [a[p * n + k], a[c * n + k]] = [a[c * n + k], a[p * n + k]];
        if (invert) [inv[p * n + k], inv[c * n + k]] = [inv[c * n + k], inv[p * n + k]];
      }
    }
    const pivot = a[c * n + c];
    det *= pivot;
    for (let k = 0; k < n; k++) {
      a[c * n + k] /= pivot;
      if (invert) inv[c * n + k] /= pivot;
    }
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const f = a[r * n + c];
      if (f === 0) continue;
      for (let k = 0; k < n; k++) {
        a[r * n + k] -= f * a[c * n + k];
        if (invert) inv[r * n + k] -= f * inv[c * n + k];
      }
    }
  }
  return { det, inverse: invert ? inv : void 0 };
}
function baseToDec(args, ctx, base) {
  arity(args, 1);
  const t = text(ctx, args[0]);
  if (isError(t)) return t;
  const s = t.trim();
  if (!s || s.length > 10 || !new RegExp(`^[${"0123456789ABCDEF".slice(0, base)}]+$`, "i").test(s)) return ERR.num;
  const n = parseInt(s, base);
  return s.length === 10 && n >= base ** 10 / 2 ? n - base ** 10 : n;
}

// packages/spreadiff/src/core/engine/evaluate.ts
var BINARY_OPS = /* @__PURE__ */ new Set(["+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="]);
var Engine = class {
  constructor(host) {
    this.host = host;
  }
  parsed = /* @__PURE__ */ new Map();
  /** Range values, patched as the host's values change. */
  ranges = /* @__PURE__ */ new Map();
  /** Exact-match indexes over cached ranges (first position of each key). */
  indexes = /* @__PURE__ */ new WeakMap();
  /** Trees by relative key: a formula copied down a column is parsed once. */
  byKey = /* @__PURE__ */ new Map();
  /** Per cell and text, the compiled tree and its shift. */
  compiled = /* @__PURE__ */ new Map();
  /** Values of a range, row-major, reused across evaluations and patched when cells change. */
  rangeValues(sheet, r0, c0, r1, c1) {
    const host = this.host;
    const cells = (r1 - r0 + 1) * (c1 - c0 + 1);
    const read = () => {
      if (host.getRange) return host.getRange(sheet, r0, c0, r1, c1);
      const out = new Array(cells);
      let i = 0;
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) out[i++] = host.get(sheet, r, c);
      return out;
    };
    const version = host.version;
    if (version === void 0 || cells < 16) return read();
    const k = `${sheet}\0${r0},${c0},${r1},${c1}`;
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
    if (host.version === version) this.ranges.set(k, { sheet, r0, c0, r1, c1, values, version });
    return values;
  }
  /** Parse an absolute formula (names, INDIRECT text) with a cache; throws Unsupported. */
  parse(formula) {
    let p = this.parsed.get(formula);
    if (p === void 0) {
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
  shapeOf(formula, row, col2) {
    try {
      return relativeKey(lex(formula), row, col2);
    } catch {
      return null;
    }
  }
  /** The tree for a cell's formula, shared with copies of it; throws Unsupported. */
  compile(formula, row, col2) {
    const ck = `${row},${col2}\0${formula}`;
    let c = this.compiled.get(ck);
    if (c === void 0) {
      try {
        const toks = lex(formula);
        const key = relativeKey(toks, row, col2);
        let hit = this.byKey.get(key);
        if (hit === void 0) {
          try {
            hit = { node: parseTokens(toks), row, col: col2 };
          } catch (e) {
            hit = e instanceof Unsupported ? e : new Unsupported(String(e));
          }
          this.byKey.set(key, hit);
        }
        c = hit instanceof Unsupported ? hit : { node: hit.node, dRow: row - hit.row, dCol: col2 - hit.col, shape: key };
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
  evaluate(formula, sheet, row, col2, record = true) {
    const reads = [];
    const { node, dRow, dCol } = typeof formula === "string" ? this.compile(formula, row, col2) : formula;
    const ctx = new EvalContext(this, sheet, row, col2, record ? reads : null, 0, dRow, dCol);
    const v = ctx.ev(node);
    let out;
    if (isArray(v) || isRange(v)) {
      const block = this.host.arrayBlock?.(sheet, row, col2);
      if (block) {
        const d = ctx.dims(v);
        const r = row - block.r0;
        const c = col2 - block.c0;
        const rr = d.rows === 1 ? 0 : r;
        const cc = d.cols === 1 ? 0 : c;
        out = rr < d.rows && cc < d.cols ? ctx.at(v, rr, cc) : ERR.na;
      } else out = ctx.scalar(v);
    } else out = v;
    return { value: out === null ? 0 : out, reads };
  }
};
var EvalContext = class _EvalContext {
  constructor(engine, sheet, row, col2, reads, depth, dRow, dCol) {
    this.engine = engine;
    this.sheet = sheet;
    this.row = row;
    this.col = col2;
    this.reads = reads;
    this.depth = depth;
    this.dRow = dRow;
    this.dCol = dCol;
    this.host = engine.host;
  }
  host;
  read(r) {
    this.reads?.push(r);
  }
  rangeOf(ref) {
    const s = ref.sheet ?? this.sheet;
    const ext = this.host.extent(s);
    if (!ext) return ERR.ref;
    if (ref.sheetTo !== void 0) throw new Unsupported("3-D reference outside an aggregate");
    const rel = ref.rel ?? 0;
    let r0 = ref.r0 < 0 ? 0 : ref.r0 + (rel & 1 ? this.dRow : 0);
    let r1 = ref.r1 < 0 ? Math.max(0, ext.rows - 1) : ref.r1 + (rel & 4 ? this.dRow : 0);
    let c0 = ref.c0 < 0 ? 0 : ref.c0 + (rel & 2 ? this.dCol : 0);
    let c1 = ref.c1 < 0 ? Math.max(0, ext.cols - 1) : ref.c1 + (rel & 8 ? this.dCol : 0);
    if (r0 < 0 || c0 < 0) return ERR.ref;
    if (r0 > r1) [r0, r1] = [r1, r0];
    if (c0 > c1) [c0, c1] = [c1, c0];
    return { kind: "range", sheet: s, r0, c0, r1, c1 };
  }
  /** A 3-D reference, as a single array of every sheet's cells (aggregates only). */
  threeD(ref) {
    const sheets = this.host.sheets;
    const a = sheets.findIndex((s) => s.toUpperCase() === ref.sheet.toUpperCase());
    const b = sheets.findIndex((s) => s.toUpperCase() === ref.sheetTo.toUpperCase());
    if (a < 0 || b < 0) throw new Unsupported("3-D reference to unknown sheet");
    const data = [];
    for (let k = Math.min(a, b); k <= Math.max(a, b); k++) {
      const rg = this.rangeOf({ ...ref, sheet: sheets[k], sheetTo: void 0 });
      if (isError(rg)) throw new Unsupported("3-D reference");
      data.push(...this.values(rg));
    }
    return { kind: "array", rows: data.length, cols: 1, data };
  }
  ev(n) {
    switch (n.t) {
      case "num":
      case "str":
      case "bool":
        return n.v;
      case "err":
        return new XlError(n.v);
      case "missing":
        return null;
      case "ref":
        return n.ref.sheetTo !== void 0 ? this.threeD(n.ref) : this.rangeOf(n.ref);
      case "name": {
        const target = lookupName(this.host.namesFor ? this.host.namesFor(this.sheet) : this.host.names, n.name);
        if (target === void 0) return ERR.name;
        if (this.depth > 8) throw new Unsupported("name nesting");
        const sub = new _EvalContext(this.engine, this.sheet, this.row, this.col, this.reads, this.depth + 1, 0, 0);
        return sub.ev(this.engine.parse(target));
      }
      case "struct": {
        const r = structuredRange(n.table, n.spec, { tables: this.host.tables, row: this.row, sheets: this.host.sheets });
        if (!r) throw new Unsupported("structured reference");
        return this.rangeOf(r);
      }
      case "array": {
        const rows = n.rows.length;
        const cols = n.rows[0].length;
        const data = [];
        for (const r of n.rows) {
          if (r.length !== cols) return ERR.value;
          for (const x of r) data.push(this.scalar(this.ev(x)));
        }
        return { kind: "array", rows, cols, data };
      }
      case "unary": {
        const v = this.ev(n.arg);
        const op = n.op;
        const f = (x) => {
          const num3 = toNumber2(x);
          if (isError(num3)) return num3;
          return op === "-" ? -num3 : op === "%" ? num3 / 100 : num3;
        };
        if (isRange(v) || isArray(v)) {
          const a = this.array(v);
          return { kind: "array", rows: a.rows, cols: a.cols, data: a.data.map(f) };
        }
        return f(v);
      }
      case "binary": {
        if (n.op === ":") {
          const a2 = this.ev(n.l);
          const b3 = this.ev(n.r);
          if (!isRange(a2) || !isRange(b3) || a2.sheet !== b3.sheet) return ERR.value;
          return { kind: "range", sheet: a2.sheet, r0: Math.min(a2.r0, b3.r0), c0: Math.min(a2.c0, b3.c0), r1: Math.max(a2.r1, b3.r1), c1: Math.max(a2.c1, b3.c1) };
        }
        if (!BINARY_OPS.has(n.op)) throw new Unsupported(`operator ${n.op}`);
        const a = this.ev(n.l);
        const b = this.ev(n.r);
        if (isRange(a) || isArray(a) || isRange(b) || isArray(b)) return this.elementwise(n.op, a, b);
        return scalarOp(n.op, a, b);
      }
      case "call": {
        const fn = FUNCTIONS[n.name];
        if (!fn) throw new Unsupported(`function ${n.name}`);
        return fn(n.args, this);
      }
    }
  }
  scalar(v) {
    if (isRange(v)) {
      if (v.r0 === v.r1 && v.c0 === v.c1) {
        this.reads?.push({ sheet: v.sheet, r0: v.r0, c0: v.c0, r1: v.r0, c1: v.c0 });
        return this.host.get(v.sheet, v.r0, v.c0);
      }
      if (v.sheet === this.sheet && v.c0 === v.c1 && this.row >= v.r0 && this.row <= v.r1) return this.at(v, this.row - v.r0, 0);
      if (v.sheet === this.sheet && v.r0 === v.r1 && this.col >= v.c0 && this.col <= v.c1) return this.at(v, 0, this.col - v.c0);
      return ERR.value;
    }
    if (isArray(v)) return v.data[0] ?? null;
    return v;
  }
  values(v) {
    if (isRange(v)) {
      this.reads?.push({ sheet: v.sheet, r0: v.r0, c0: v.c0, r1: v.r1, c1: v.c1 });
      return this.engine.rangeValues(v.sheet, v.r0, v.c0, v.r1, v.c1);
    }
    if (isArray(v)) return v.data;
    return [v];
  }
  dims(v) {
    if (isRange(v)) return { rows: v.r1 - v.r0 + 1, cols: v.c1 - v.c0 + 1 };
    if (isArray(v)) return { rows: v.rows, cols: v.cols };
    return { rows: 1, cols: 1 };
  }
  at(v, r, c) {
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
  array(v) {
    if (isArray(v)) return v;
    if (isRange(v)) {
      const d = this.dims(v);
      if (d.rows * d.cols > 5e6) throw new Unsupported("range too large");
      return { kind: "array", rows: d.rows, cols: d.cols, data: this.values(v) };
    }
    return { kind: "array", rows: 1, cols: 1, data: [v] };
  }
  /** Array arithmetic: broadcast 1-row / 1-column / scalar operands, #N/A outside the overlap. */
  elementwise(op, a, b) {
    const x = this.array(a);
    const y = this.array(b);
    const rows = Math.max(x.rows, y.rows);
    const cols = Math.max(x.cols, y.cols);
    if (rows * cols > 5e6) throw new Unsupported("array too large");
    const pick = (m, r, c) => {
      const rr = m.rows === 1 ? 0 : r;
      const cc = m.cols === 1 ? 0 : c;
      return rr < m.rows && cc < m.cols ? m.data[rr * m.cols + cc] : ERR.na;
    };
    const data = new Array(rows * cols);
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) data[r * cols + c] = scalarOp(op, pick(x, r, c), pick(y, r, c));
    return { kind: "array", rows, cols, data };
  }
};
function scalarOp(op, a, b) {
  switch (op) {
    case "&": {
      const x2 = toText(a);
      if (isError(x2)) return x2;
      const y2 = toText(b);
      if (isError(y2)) return y2;
      return x2 + y2;
    }
    case "=":
    case "<>":
    case "<":
    case ">":
    case "<=":
    case ">=": {
      if (isError(a)) return a;
      if (isError(b)) return b;
      const c = compare(a, b);
      return op === "=" ? c === 0 : op === "<>" ? c !== 0 : op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0;
    }
  }
  const x = toNumber2(a);
  if (isError(x)) return x;
  const y = toNumber2(b);
  if (isError(y)) return y;
  switch (op) {
    case "+":
      return x + y;
    case "-":
      return x - y;
    case "*":
      return x * y;
    case "/":
      return y === 0 ? ERR.div0 : x / y;
    case "^": {
      if (x === 0 && y < 0) return ERR.div0;
      const r = Math.pow(x, y);
      return Number.isFinite(r) ? r : ERR.num;
    }
  }
  throw new Unsupported(`operator ${op}`);
}

// packages/spreadiff/src/core/engine/host.ts
var ROW_MUL = 2 ** 14;
var SHEET_MUL = 2 ** 34;
var sheetCellKey = (sheet, row, col2) => sheet * SHEET_MUL + row * ROW_MUL + col2;
var ERROR_TEXT = /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!)$/;
function engineValue(v, date19042) {
  if (v === void 0 || v === null || typeof v !== "string") return v ?? null;
  const c = v.charCodeAt(0);
  if (c === 35 && ERROR_TEXT.test(v)) return errorOf(v);
  if (c >= 48 && c <= 57 && v.length >= 8) {
    const s = isoToSerial(v, date19042);
    if (s !== null) return s;
  }
  return v;
}
function tableHost(tables, extra = {}) {
  const byName = /* @__PURE__ */ new Map();
  tables.forEach((t, i) => {
    byName.set(t.name, i);
    if (!byName.has(t.name.toUpperCase())) byName.set(t.name.toUpperCase(), i);
  });
  const date19042 = extra.date1904 ?? false;
  const overrides = /* @__PURE__ */ new Map();
  const hiddenSets = [];
  const scoped = /* @__PURE__ */ new Map();
  const index = (name) => byName.get(name) ?? byName.get(name.toUpperCase()) ?? -1;
  const cachedIn = (t, r, c) => {
    const o = t.origin ?? { row: 1, col: 0 };
    const tr = r - o.row;
    const tc = c - o.col;
    if (tc < 0 || tc >= t.columns.length) return null;
    if (tr === -1 && t.header) return engineValue(t.header[tc], date19042);
    if (tr < 0 || tr >= t.rows.length) return null;
    return engineValue(t.rows[tr][tc], date19042);
  };
  const LOG = 4096;
  const log = [];
  let logStart = 0;
  const host = {
    sheets: tables.map((t) => t.name),
    names: extra.names,
    namesFor: tables.some((t) => t.names) ? (sheet) => {
      let m = scoped.get(sheet);
      if (!m) scoped.set(sheet, m = scopedNames(extra.names, tables, sheet) ?? {});
      return m;
    } : void 0,
    tables: extra.tables,
    date1904: date19042,
    overrides,
    version: 0,
    override(k, v) {
      if (v === void 0) overrides.delete(k);
      else overrides.set(k, v);
      const s = Math.floor(k / 2 ** 34);
      const rest = k - s * 2 ** 34;
      if (log.length >= LOG) {
        log.length = 0;
        logStart = host.version;
      }
      log.push({ sheet: tables[s].name, row: Math.floor(rest / 2 ** 14), col: rest % 2 ** 14 });
      host.version++;
    },
    changesSince(version) {
      if (version < logStart) return null;
      return log.slice(version - logStart);
    },
    sheetIndex: index,
    rowHidden: tables.some((t) => t.hiddenRows) ? (sheet, row) => {
      const s = index(sheet);
      if (s < 0) return 0;
      let h = hiddenSets[s];
      if (h === void 0) {
        const hr = tables[s].hiddenRows;
        h = hiddenSets[s] = hr ? new Map([...hr.rows.map((r) => [r, 1]), ...hr.filtered.map((r) => [r, 2])]) : null;
      }
      return h?.get(row) ?? 0;
    } : void 0,
    formulaAt(sheet, r, c) {
      const s = index(sheet);
      return s < 0 ? null : host.formula(s, r, c);
    },
    cached: (s, r, c) => cachedIn(tables[s], r, c),
    formula(s, r, c) {
      const t = tables[s];
      const o = t.origin ?? { row: 1, col: 0 };
      return t.formulas?.[r - o.row]?.[c - o.col] ?? null;
    },
    get(sheet, r, c) {
      const s = index(sheet);
      if (s < 0) return null;
      if (overrides.size) {
        const k = sheetCellKey(s, r, c);
        const v = overrides.get(k);
        if (v !== void 0 || overrides.has(k)) return v ?? null;
      }
      return cachedIn(tables[s], r, c);
    },
    getRange(sheet, r0, c0, r1, c1) {
      const s = index(sheet);
      const out = new Array((r1 - r0 + 1) * (c1 - c0 + 1));
      if (s < 0) return out.fill(null);
      const t = tables[s];
      const o = t.origin ?? { row: 1, col: 0 };
      let i = 0;
      for (let r = r0; r <= r1; r++) {
        const tr = r - o.row;
        const row = tr >= 0 && tr < t.rows.length ? t.rows[tr] : null;
        for (let c = c0; c <= c1; c++) {
          const tc = c - o.col;
          let v;
          if (tc < 0 || tc >= t.columns.length) v = null;
          else if (row) v = engineValue(row[tc], date19042);
          else v = tr === -1 && t.header ? engineValue(t.header[tc], date19042) : null;
          if (overrides.size) {
            const k = sheetCellKey(s, r, c);
            if (overrides.has(k)) v = overrides.get(k);
          }
          out[i++] = v;
        }
      }
      return out;
    },
    extent(sheet) {
      const s = index(sheet);
      if (s < 0) return null;
      const t = tables[s];
      const o = t.origin ?? { row: 1, col: 0 };
      return { rows: o.row + t.rows.length, cols: o.col + t.columns.length };
    },
    arrayBlock(sheet, r, c) {
      const s = index(sheet);
      const blocks = s >= 0 ? tables[s].arrays : void 0;
      if (!blocks) return null;
      for (const b of blocks) if (r >= b.r0 && r <= b.r1 && c >= b.c0 && c <= b.c1) return { r0: b.r0, c0: b.c0 };
      return null;
    }
  };
  return host;
}

// packages/spreadiff/src/core/recalc.ts
function computeMissingValues(workbook) {
  const result = { computed: 0, failed: 0 };
  const missing2 = /* @__PURE__ */ new Set();
  workbook.sheets.forEach((t, s) => {
    t.formulas?.forEach((row, r) => row?.forEach((f, c) => {
      if (f && (t.rows[r]?.[c] ?? null) === null) missing2.add(`${s},${r},${c}`);
    }));
  });
  if (!missing2.size) return result;
  const host = tableHost(workbook.sheets, { names: workbook.names, tables: workbook.tables, date1904: workbook.date1904 });
  const engine = new Engine(host);
  const state = /* @__PURE__ */ new Map();
  const failed = /* @__PURE__ */ new Set();
  let depth = 0;
  const base = host.get;
  host.getRange = void 0;
  const compute = (s, r, c) => {
    const k = `${s},${r},${c}`;
    const st = state.get(k);
    if (st === 2) return;
    if (st === 1) throw new Unsupported("circular reference");
    if (depth > 2e3) throw new Unsupported("dependency chain too deep");
    state.set(k, 1);
    depth++;
    const t = workbook.sheets[s];
    const o = t.origin ?? { row: 1, col: 0 };
    try {
      const { value } = engine.evaluate(t.formulas[r][c], t.name, r + o.row, c + o.col, false);
      t.rows[r][c] = toCellValue(value);
      host.override(sheetKeyOf(s, r + o.row, c + o.col), value);
      result.computed++;
    } catch (e) {
      if (!(e instanceof Unsupported)) throw e;
      failed.add(k);
      result.failed++;
    } finally {
      depth--;
      state.set(k, 2);
    }
  };
  host.get = (sheet, row, col2) => {
    const s = host.sheetIndex(sheet);
    if (s >= 0) {
      const t = workbook.sheets[s];
      const o = t.origin ?? { row: 1, col: 0 };
      const k = `${s},${row - o.row},${col2 - o.col}`;
      if (missing2.has(k)) {
        if (state.get(k) !== 2) compute(s, row - o.row, col2 - o.col);
        if (failed.has(k)) throw new Unsupported("depends on a value that could not be computed");
      }
    }
    return base(sheet, row, col2);
  };
  for (const k of missing2) {
    if (state.get(k) === 2) continue;
    const [s, r, c] = k.split(",").map(Number);
    try {
      compute(s, r, c);
    } catch (e) {
      if (!(e instanceof Unsupported)) throw e;
    }
  }
  return result;
}
var sheetKeyOf = (s, r, c) => s * 2 ** 34 + r * 2 ** 14 + c;
function toCellValue(v) {
  return isError(v) ? v.code : v;
}

// packages/spreadiff/src/core/sheet.ts
function detectHeaders(raws, names, tables, localNames) {
  const candidate = /* @__PURE__ */ new Map();
  const result = raws.map((r, i) => {
    const ok = headerLike(r);
    if (ok) candidate.set(r.name.toUpperCase(), i);
    return ok;
  });
  if (!candidate.size) return result;
  const sheets = raws.map((r) => r.name);
  const top = Math.max(...[...candidate.values()].map((i) => raws[i].first + 1));
  const rowNumbers = Array.from({ length: top }, (_, i) => i + 1).join("|");
  const spells = top <= 64 ? new RegExp(`(?:[A-Za-z$](?:${rowNumbers})(?!\\d))|(?:(?:^|[^\\w.$])\\$?(?:${rowNumbers}):)`) : /./;
  const nameList = [...Object.keys(names), ...[...localNames.values()].flatMap((m) => Object.keys(m))].map((n) => n.toUpperCase());
  const scoped = raws.map((r) => localNames.size ? scopedNames(names, raws.map((x) => ({ name: x.name, names: localNames.get(x.name) })), r.name) : names);
  const mayHide = (f) => f.includes("[") || nameList.length > 0 && (nameList.length > 64 || nameList.some((n) => f.toUpperCase().includes(n)));
  raws.forEach((r, ri) => {
    for (let row = 0; row < r.formulaGrid.length && candidate.size; row++) {
      const fr = r.formulaGrid[row];
      if (!fr) continue;
      for (const f of fr) {
        if (!f) continue;
        if (!spells.test(f) && !mayHide(f)) continue;
        for (const rg of formulaRefs(f, scoped[ri], { tables, row, sheets }).ranges) {
          const key = (rg.sheet ?? r.name).toUpperCase();
          const i = candidate.get(key);
          if (i === void 0 || rg.r0 < 0) continue;
          const first = raws[i].first;
          if (rg.r0 <= first && first <= rg.r1) {
            result[i] = false;
            candidate.delete(key);
          }
        }
        if (!candidate.size) break;
      }
    }
  });
  return result;
}
var ISO_DATE = /^\d{4}-\d{2}-\d{2}/;
function headerLike(r) {
  if (r.maxCol < 0) return false;
  const row = r.grid[r.first];
  if (r.formulaGrid[r.first]?.some(Boolean)) return false;
  let n = 0;
  for (let c = r.minCol; c <= r.maxCol; c++) {
    const v = row?.[c];
    if (v === null || v === void 0 || v === "") continue;
    if (isYear(v)) continue;
    if (typeof v !== "string" || ISO_DATE.test(v)) return false;
    n++;
  }
  return n > 0;
}
var isYear = (v) => typeof v === "number" && Number.isInteger(v) && v >= 1900 && v <= 2100;
function tablesFromRaw(raws, header, names = {}, tables = [], localNames = /* @__PURE__ */ new Map()) {
  const headers = header === "auto" ? detectHeaders(raws, names, tables, localNames) : raws.map(() => header);
  return raws.map((r, i) => {
    const t = buildTable(r, headers[i], header === "auto");
    const own = localNames.get(r.name);
    if (own) t.names = own;
    return t;
  });
}
function rawSheet(name, grid, formulaGrid, extra = {}) {
  let minCol = Infinity;
  let maxCol = -1;
  for (const g of [grid, formulaGrid]) {
    for (const row of g) {
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v === void 0 || v === null) continue;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
      }
    }
  }
  let first = 0;
  while (first < grid.length && !grid[first]?.some((v) => v !== null && v !== void 0) && !formulaGrid[first]?.some((f) => f)) first++;
  return { name, grid, formulaGrid, minCol, maxCol, first, arrays: [], merges: [], labelStyles: [], hidden: [], filters: [], ...extra };
}
function labelRow(r) {
  const row = r.grid[r.first];
  let text3 = 0;
  let filled = 0;
  for (let c = r.minCol; c <= r.maxCol; c++) {
    const v = row?.[c];
    if (v === null || v === void 0 || v === "") continue;
    filled++;
    if (typeof v === "string" && !ISO_DATE.test(v) && !r.formulaGrid[r.first]?.[c]) text3++;
  }
  return text3 >= 2 && text3 * 2 >= filled;
}
function buildTable(raw, header, auto = false) {
  const { name, grid, formulaGrid, minCol, maxCol, first, arrays, merges, labelStyles, formats, hidden, filters } = raw;
  if (maxCol < 0) return { name, columns: [], rows: [], formulas: [], origin: { row: 0, col: 0 }, ...merges.length ? { merges } : {} };
  const width = maxCol - minCol + 1;
  const dense = (src) => {
    const out = new Array(width);
    for (let c = 0; c < width; c++) out[c] = src?.[minCol + c] ?? null;
    return out;
  };
  const columns = header ? uniqueNames(dense(grid[first]).map((v) => v === null ? "" : String(v))) : auto && labelRow(raw) ? uniqueNames(dense(grid[first]).map((v, c) => typeof v === "string" && v.trim() && !formulaGrid[first]?.[minCol + c] ? v : columnLabel(minCol + c))) : Array.from({ length: width }, (_, c) => columnLabel(minCol + c));
  const start = header ? first + 1 : first;
  const rows = [];
  const formulas = [];
  for (let r = start; r < grid.length; r++) {
    rows.push(dense(grid[r]));
    if (formulaGrid[r]) formulas[r - start] = dense(formulaGrid[r]);
  }
  while (rows.length && !formulas[rows.length - 1] && rows[rows.length - 1].every((v) => v === null)) rows.pop();
  formulas.length = Math.min(formulas.length, rows.length);
  let numberFormats;
  if (formats) {
    const fr = [];
    for (let r = start; r < start + rows.length; r++) {
      const src = formats.grid[r];
      if (!src) continue;
      const out = new Uint16Array(width);
      for (let c = 0; c < width; c++) out[c] = src[minCol + c] ?? 0;
      fr[r - start] = out;
    }
    if (fr.length) numberFormats = { codes: formats.codes, rows: fr, ...formats.date1904 ? { date1904: true } : {} };
  }
  return {
    name,
    columns,
    rows,
    formulas,
    origin: { row: start, col: minCol },
    ...header ? { header: dense(grid[first]) } : {},
    ...arrays.length ? { arrays } : {},
    ...merges.length ? { merges } : {},
    ...labelStyles.length ? { labelStyles } : {},
    ...numberFormats ? { numberFormats } : {},
    ...hidden.length ? { hiddenRows: { rows: hidden, filtered: hidden.filter((r) => filters.some((f) => r >= f.r0 && r <= f.r1)) } } : {}
  };
}

// packages/spreadiff/src/core/external.ts
var LIVE_FUNCTIONS = {
  Bloomberg: ["BDP", "BDH", "BDS", "BQL", "BQL.QUERY", "BEQS", "BSRCH", "BCURVE", "BLP", "BLPH", "BLPW"],
  LSEG: ["TR", "RDP.DATA", "RDP.HISTORY", "RDP.SEARCH", "RHISTORY", "RDATA", "RSEARCH", "DSGRID"],
  FactSet: ["FDS", "FDSB", "FDSC", "FQL"],
  "S&P Capital IQ": ["CIQ*", "SPG"],
  Morningstar: ["MSDP", "MSTS"],
  Excel: [
    "RTD",
    "WEBSERVICE",
    "STOCKHISTORY",
    "CUBEVALUE",
    "CUBEMEMBER",
    "CUBESET",
    "CUBESETCOUNT",
    "CUBERANKEDMEMBER",
    "CUBEKPIMEMBER",
    "CUBEMEMBERPROPERTY",
    "PY",
    "COPILOT",
    "IMAGE",
    "TRANSLATE",
    "DETECTLANGUAGE"
  ]
};
var compileMatcher = (table) => {
  const exact = /* @__PURE__ */ new Map();
  const prefixes = [];
  for (const [provider, names] of Object.entries(table)) {
    for (const n of names) {
      const u = n.toUpperCase();
      if (u.endsWith("*")) prefixes.push([u.slice(0, -1), provider]);
      else exact.set(u, provider);
    }
  }
  return { exact, prefixes };
};
var DEFAULT = compileMatcher(LIVE_FUNCTIONS);
var cache = /* @__PURE__ */ new WeakMap();
var STRING2 = /"(?:[^"]|"")*"/g;
var CALL = /(?<![A-Za-z0-9_.\\])((?:_xll\.|_xlfn\.(?:_xlws\.)?)?)([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
var LINK = /(?<![A-Za-z0-9_.\\\]\[])\[\d+\]/;
function externalSource(formula, extra) {
  if (!formula) return null;
  const f = formula.replace(STRING2, '""');
  let extraMatcher;
  if (extra) {
    extraMatcher = cache.get(extra);
    if (!extraMatcher) cache.set(extra, extraMatcher = compileMatcher(extra));
  }
  CALL.lastIndex = 0;
  let m;
  while (m = CALL.exec(f)) {
    const prefix = m[1].toLowerCase();
    const name = m[2].toUpperCase();
    for (const matcher of extraMatcher ? [extraMatcher, DEFAULT] : [DEFAULT]) {
      const hit = matcher.exact.get(name) ?? matcher.prefixes.find(([p]) => name.startsWith(p))?.[1];
      if (hit) return { provider: hit, function: name };
    }
    if (prefix === "_xll.") return { provider: "add-in", function: name };
  }
  const link = LINK.exec(f);
  if (link) return { provider: "link", function: link[0] };
  return null;
}

// packages/spreadiff/src/core/text.ts
function numericDelta(before, after) {
  const x = typeof before === "number" ? before : typeof before === "string" && before.trim() ? Number(before) : NaN;
  const y = typeof after === "number" ? after : typeof after === "string" && after.trim() ? Number(after) : NaN;
  if (!Number.isFinite(x) || !Number.isFinite(y)) return null;
  const scale = Math.max(Math.abs(x), Math.abs(y));
  const decimals = scale ? 14 - Math.floor(Math.log10(scale)) : 0;
  const raw = y - x;
  const delta = decimals > 0 && decimals <= 100 ? Number(raw.toFixed(decimals)) : raw;
  return { delta, ratio: x === 0 ? null : delta / Math.abs(x) };
}
function formatCell(v) {
  if (v === null || v === void 0) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") {
    if (Number.isInteger(v)) return String(v);
    return String(Number(v.toPrecision(15)));
  }
  return v;
}

// packages/spreadiff/src/core/trace.ts
var ROW_MUL2 = 2 ** 14;
var SHEET_MUL2 = 2 ** 34;
var cellId = (s, r, c) => s * SHEET_MUL2 + r * ROW_MUL2 + c;
var idSheet = (id) => Math.floor(id / SHEET_MUL2);
var idRow = (id) => Math.floor(id % SHEET_MUL2 / ROW_MUL2);
var idCol = (id) => id % ROW_MUL2;
var Graph = class {
  constructor(precedents, flags = /* @__PURE__ */ new Map()) {
    this.precedents = precedents;
    this.flags = flags;
    for (const [id, rects] of precedents) {
      for (const rc of rects) {
        if (rc.r0 === rc.r1 && rc.c0 === rc.c1) {
          const k = cellId(rc.sheet, rc.r0, rc.c0);
          const list = this.single.get(k);
          if (list) list.push(id);
          else this.single.set(k, [id]);
          continue;
        }
        const cols = this.ranges[rc.sheet] ??= /* @__PURE__ */ new Map();
        for (let c = rc.c0; c <= rc.c1; c++) {
          let list = cols.get(c);
          if (!list) cols.set(c, list = []);
          list.push(rc.r0, rc.r1, id);
        }
      }
    }
  }
  /** Single-cell precedents: cell id → formula ids that read exactly it. */
  single = /* @__PURE__ */ new Map();
  /** Range precedents: per sheet, per column, flat [r0, r1, formulaId, …]. */
  ranges = [];
  /** Formula cells that read (sheet, row, col). */
  dependents(sheet, row, col2, out) {
    const one = this.single.get(cellId(sheet, row, col2));
    if (one) for (const id of one) out.push(id);
    const list = this.ranges[sheet]?.get(col2);
    if (!list) return;
    for (let k = 0; k < list.length; k += 3) if (row >= list[k] && row <= list[k + 1]) out.push(list[k + 2]);
  }
};
var OverlayGraph = class {
  constructor(base, replaced) {
    this.base = base;
    this.replaced = replaced;
    this.extra = new Graph(replaced);
    this.flags = base.flags;
    this.precedents = { get: (id) => replaced.get(id) ?? base.precedents.get(id) };
  }
  extra;
  precedents;
  flags;
  dependents(sheet, row, col2, out) {
    const start = out.length;
    this.base.dependents(sheet, row, col2, out);
    let w = start;
    for (let i = start; i < out.length; i++) if (!this.replaced.has(out[i])) out[w++] = out[i];
    out.length = w;
    this.extra.dependents(sheet, row, col2, out);
  }
};
function topoOrder(sources, nodes, graph) {
  const preds = /* @__PURE__ */ new Map();
  const indeg = /* @__PURE__ */ new Map();
  for (const n of nodes) indeg.set(n, 0);
  const deps = [];
  const edgesFrom = (u) => {
    deps.length = 0;
    graph.dependents(idSheet(u), idRow(u), idCol(u), deps);
    for (const v of deps) {
      if (v === u || !nodes.has(v)) continue;
      let p = preds.get(v);
      if (!p) preds.set(v, p = []);
      if (p[p.length - 1] === u) continue;
      p.push(u);
      if (nodes.has(u)) indeg.set(v, indeg.get(v) + 1);
    }
  };
  for (const u of sources) if (!nodes.has(u)) edgesFrom(u);
  for (const u of nodes) edgesFrom(u);
  const order = [];
  const ready = [];
  for (const [n, d] of indeg) if (d === 0) ready.push(n);
  while (ready.length) {
    const u = ready.pop();
    order.push(u);
    deps.length = 0;
    graph.dependents(idSheet(u), idRow(u), idCol(u), deps);
    let last = -1;
    for (const v of deps) {
      if (v === u || v === last || !nodes.has(v)) continue;
      last = v;
      const d = indeg.get(v) - 1;
      indeg.set(v, d);
      if (d === 0) ready.push(v);
    }
  }
  const cyclic = order.length === nodes.size ? [] : [...nodes].filter((n) => indeg.get(n) > 0);
  return { order, preds, cyclic };
}
function clip(rg, self, byName, tables) {
  const sheet = rg.sheet === null ? self : byName.get(rg.sheet.toUpperCase());
  if (sheet === void 0) return null;
  const t = tables[sheet];
  const origin = t.origin ?? { row: 1, col: 0 };
  const nRows = t.rows.length;
  const nCols = t.columns.length;
  const r0 = rg.r0 < 0 ? 0 : Math.max(0, rg.r0 - origin.row);
  const r1 = rg.r1 < 0 ? nRows - 1 : Math.min(nRows - 1, rg.r1 - origin.row);
  const c0 = rg.c0 < 0 ? 0 : Math.max(0, rg.c0 - origin.col);
  const c1 = rg.c1 < 0 ? nCols - 1 : Math.min(nCols - 1, rg.c1 - origin.col);
  if (r0 > r1 || c0 > c1) return null;
  return { sheet, r0, r1, c0, c1 };
}
function refReader(diff, side) {
  const tables = diff.sheets.map((s) => side === "old" ? s.diff.oldTable : s.diff.newTable);
  const byName = new Map(tables.map((t, i) => [t.name.toUpperCase(), i]));
  const names = side === "old" ? diff.names?.old : diff.names?.new;
  const tableDefs = side === "old" ? diff.tables?.old : diff.tables?.new;
  const sheetNames = tables.map((t) => t.name);
  const scoped = tables.some((t) => t.names) ? tables.map((t) => scopedNames(names, tables, t.name)) : null;
  return {
    tables,
    read(s, r, c) {
      const formula = tables[s]?.formulas?.[r]?.[c];
      if (!formula) return null;
      const origin = tables[s].origin ?? { row: 1, col: 0 };
      const refs = formulaRefs(formula, scoped ? scoped[s] : names, { tables: tableDefs, row: r + origin.row, sheets: sheetNames });
      const rects = [];
      for (const rg of refs.ranges) {
        const rect = clip(rg, s, byName, tables);
        if (rect) rects.push(rect);
      }
      return { rects, dynamic: refs.dynamic, volatile: refs.volatile };
    }
  };
}
function staticGraph(diff, side) {
  const reader = refReader(diff, side);
  const precedents = /* @__PURE__ */ new Map();
  const flags = /* @__PURE__ */ new Map();
  reader.tables.forEach((t, s) => {
    const f = t.formulas;
    if (!f) return;
    for (let r = 0; r < f.length; r++) {
      const row = f[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const refs = reader.read(s, r, c);
        if (!refs) continue;
        const id = cellId(s, r, c);
        precedents.set(id, refs.rects);
        if (refs.dynamic || refs.volatile) flags.set(id, { dynamic: refs.dynamic, volatile: refs.volatile });
      }
    }
  });
  return new Graph(precedents, flags);
}
function sheetMaps(diff) {
  const inv = (arr, n) => {
    const out = new Int32Array(n).fill(-1);
    for (let i = 0; i < arr.length; i++) if (arr[i] >= 0) out[arr[i]] = i;
    return out;
  };
  const cols = (side, n) => {
    const out = new Int32Array(n).fill(-1);
    diff.columns.forEach((c, ci) => {
      const j = side === "old" ? c.oldIndex : c.newIndex;
      if (j >= 0) out[j] = ci;
    });
    return out;
  };
  return {
    diff,
    rows: { old: inv(diff.oldRows, diff.oldTable.rows.length), new: inv(diff.newRows, diff.newTable.rows.length) },
    cols: { old: cols("old", diff.oldTable.columns.length), new: cols("new", diff.newTable.columns.length) }
  };
}
function traceWorkbook(diff, options = {}) {
  const maxCauses = options.maxCauses ?? 25;
  const maxImpacts = options.maxImpacts ?? 500;
  let budget = options.budget ?? 5e6;
  let truncated = false;
  const maps = diff.sheets.map((s) => sheetMaps(s.diff));
  const traces = diff.sheets.map(() => /* @__PURE__ */ new Map());
  const key = (ref) => ref.row * maps[ref.sheet].diff.columns.length + ref.col;
  const refOfKey = (si, k) => {
    const w = maps[si].diff.columns.length;
    return { sheet: si, row: Math.floor(k / w), col: k % w };
  };
  const blank = (kind) => ({ kind, causes: [], causeCount: 0, impacts: [], impactCount: 0 });
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (const [row, mask] of d.changedCells) {
      const fm = d.formulaChanged?.get(row);
      const ni = d.newRows[row];
      for (let col2 = 0; col2 < mask.length; col2++) {
        if (!mask[col2]) continue;
        const nf = d.newTable.formulas?.[ni]?.[d.columns[col2].newIndex] ?? null;
        const ext = nf && !fm?.[col2] ? externalSource(nf, options.liveFunctions) : null;
        const t = blank(fm?.[col2] ? "formula" : nf && !ext ? "computed" : "input");
        if (ext) t.external = ext;
        traces[si].set(key({ sheet: si, row, col: col2 }), t);
      }
    }
  });
  const changedKind = (ref) => traces[ref.sheet].get(key(ref))?.kind;
  const toAligned = (side, id) => {
    const sheet = idSheet(id);
    const m = maps[sheet];
    const row = m.rows[side][idRow(id)];
    const col2 = m.cols[side][idCol(id)];
    return row >= 0 && col2 >= 0 ? { sheet, row, col: col2 } : null;
  };
  const sideId = (side, ref) => {
    const d = maps[ref.sheet].diff;
    const r = side === "new" ? d.newRows[ref.row] : d.oldRows[ref.row];
    const c = side === "new" ? d.columns[ref.col].newIndex : d.columns[ref.col].oldIndex;
    return r >= 0 && c >= 0 ? cellId(ref.sheet, r, c) : null;
  };
  const roots = [];
  const staticNew = staticGraph(diff, "new");
  let staticOld = null;
  const oldGraph = () => staticOld ??= staticGraph(diff, "old");
  let oldReader = null;
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (const [k, t] of traces[si]) {
      if (t.kind !== "input" && t.kind !== "formula") continue;
      const ref = refOfKey(si, k);
      roots.push({ ref, trace: t, side: "new", id: cellId(si, d.newRows[ref.row], d.columns[ref.col].newIndex) });
    }
    for (let row = 0; row < d.rowCount; row++) {
      const ni = d.newRows[row];
      if (ni < 0) continue;
      const rowAdded = d.status[row] === ROW_STATUS.added;
      d.columns.forEach((c, col2) => {
        if (c.newIndex < 0 || !(rowAdded || c.status === "added")) return;
        const deps = [];
        staticNew.dependents(si, ni, c.newIndex, deps);
        if (!deps.length) return;
        const ref = { sheet: si, row, col: col2 };
        roots.push({ ref, trace: traces[si].get(key(ref)) ?? blank("added"), side: "new", id: cellId(si, ni, c.newIndex) });
      });
    }
  });
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    for (let row = 0; row < d.rowCount; row++) {
      const oi = d.oldRows[row];
      if (oi < 0) continue;
      const rowRemoved = d.status[row] === ROW_STATUS.removed;
      d.columns.forEach((c, col2) => {
        if (c.oldIndex < 0 || !(rowRemoved || c.status === "removed")) return;
        const deps = [];
        oldGraph().dependents(si, oi, c.oldIndex, deps);
        if (!deps.length) return;
        roots.push({ ref: { sheet: si, row, col: col2 }, trace: blank("removed"), side: "old", id: cellId(si, oi, c.oldIndex) });
      });
    }
  });
  const analysis = options.recalculate === false ? null : recalculate(diff, maps, traces, roots, staticNew, options.evaluationBudget ?? 2e6);
  if (analysis?.truncated) truncated = true;
  const newGraph = analysis?.graph ?? staticNew;
  const causeSets = diff.sheets.map(() => /* @__PURE__ */ new Map());
  const walk = (side, startId, root) => {
    const g = side === "new" ? newGraph : oldGraph();
    const seen = /* @__PURE__ */ new Set([startId]);
    const queue = [startId];
    const deps = [];
    while (queue.length) {
      if (--budget < 0) {
        truncated = true;
        return;
      }
      const id = queue.pop();
      deps.length = 0;
      g.dependents(idSheet(id), idRow(id), idCol(id), deps);
      for (const dep of deps) {
        if (seen.has(dep)) continue;
        seen.add(dep);
        const ref = toAligned(side, dep);
        if (!ref) continue;
        const kind = changedKind(ref);
        if (kind !== "computed" && kind !== "formula") continue;
        if (kind === "computed") {
          const m = causeSets[ref.sheet];
          const k = key(ref);
          let set = m.get(k);
          if (!set) m.set(k, set = /* @__PURE__ */ new Set());
          set.add(root);
        }
        queue.push(dep);
      }
    }
  };
  roots.forEach((r, i) => walk(r.side, r.id, i));
  const selfCaused = /* @__PURE__ */ new Set();
  diff.sheets.forEach((s, si) => {
    for (const [k, t] of traces[si]) {
      if (t.kind !== "computed" || causeSets[si].get(k)?.size) continue;
      const ref = refOfKey(si, k);
      const id = sideId("new", ref);
      if (id === null || analysis?.modelled.has(id)) continue;
      const flags = staticNew.flags.get(id);
      if (!flags) continue;
      t.unexplained = flags.volatile ? "volatile" : "dynamic";
      const i = roots.push({ ref, trace: t, side: "new", id }) - 1;
      selfCaused.add(i);
      walk("new", id, i);
    }
  });
  const exact = analysis?.butFor(roots) ?? null;
  const stats = { inputs: 0, formulas: 0, computed: 0, unexplained: 0, stale: 0, recalculated: 0, unmodelled: analysis?.unmodelled ?? 0, external: 0 };
  diff.sheets.forEach((_, si) => {
    for (const [k, t] of traces[si]) {
      if (t.kind === "input") stats.inputs++;
      if (t.external) stats.external++;
      else if (t.kind === "formula") stats.formulas++;
      if (t.kind !== "computed") continue;
      stats.computed++;
      if (t.stale) stats.stale++;
      const ref = refOfKey(si, k);
      const byRefs = causeSets[si].get(k);
      const precise = exact?.get(si)?.get(k);
      let list;
      const structural2 = [...byRefs ?? []].filter((r) => roots[r].trace.kind !== "input" && roots[r].trace.kind !== "formula");
      if (precise && precise.causes.length && !precise.tainted) {
        list = [...precise.causes, ...structural2];
        t.contributions = [...precise.contributions, ...structural2.map(() => null)].slice(0, maxCauses);
        t.method = "recalculated";
        stats.recalculated++;
      } else {
        list = precise?.causes.length ? [...precise.causes, ...[...byRefs ?? []].filter((r) => !precise.causes.includes(r))] : [...byRefs ?? []];
        t.method = "references";
      }
      list = list.filter((r) => !(selfCaused.has(r) && roots[r].trace === t));
      t.causeCount = list.length;
      t.causes = list.slice(0, maxCauses).map((r) => roots[r].ref);
      for (const r of list) {
        const rt = roots[r].trace;
        rt.impactCount++;
        if (rt.impacts.length < maxImpacts) rt.impacts.push(ref);
      }
      if (!t.causeCount) {
        t.unexplained ??= "unknown";
        stats.unexplained++;
      }
    }
  });
  for (const r of roots) {
    if ((r.trace.kind === "added" || r.trace.kind === "removed") && r.trace.impactCount) traces[r.ref.sheet].set(key(r.ref), r.trace);
  }
  const structural = [];
  const structuralRows = (si) => {
    let rows = structural[si];
    if (!rows) {
      rows = structural[si] = [];
      const d = maps[si].diff;
      for (let row = 0; row < d.rowCount; row++) if (d.status[row] === ROW_STATUS.added || d.status[row] === ROW_STATUS.removed) rows.push(row);
    }
    return rows;
  };
  return {
    get: (ref) => traces[ref.sheet]?.get(key(ref)),
    sheet: (i) => traces[i] ?? /* @__PURE__ */ new Map(),
    stats,
    truncated,
    parents(ref) {
      const d = maps[ref.sheet]?.diff;
      if (!d) return [];
      const out = [];
      const seen = /* @__PURE__ */ new Set();
      const add = (r) => {
        if (!r) return;
        const s = `${r.sheet}:${r.row}:${r.col}`;
        if (seen.has(s) || r.sheet === ref.sheet && r.row === ref.row && r.col === ref.col) return;
        const t = traces[r.sheet].get(key(r));
        const st = maps[r.sheet].diff.status[r.row];
        if (t || st === ROW_STATUS.added || st === ROW_STATUS.removed) {
          seen.add(s);
          out.push(r);
        }
      };
      for (const side of ["new", "old"]) {
        const id = sideId(side, ref);
        if (id === null) continue;
        const rects = side === "new" ? newGraph.precedents.get(id) : staticOld ? staticOld.precedents.get(id) : (oldReader ??= refReader(diff, "old")).read(idSheet(id), idRow(id), idCol(id))?.rects;
        for (const rc of rects ?? []) {
          const area = (rc.r1 - rc.r0 + 1) * (rc.c1 - rc.c0 + 1);
          const target = maps[rc.sheet].diff;
          if (area <= 4096) {
            for (let r = rc.r0; r <= rc.r1; r++) for (let c = rc.c0; c <= rc.c1; c++) add(toAligned(side, cellId(rc.sheet, r, c)));
            continue;
          }
          for (const [k] of traces[rc.sheet]) {
            if (out.length >= 200) break;
            const cand = refOfKey(rc.sheet, k);
            const cid = sideId(side, cand);
            if (cid === null) continue;
            const r = idRow(cid);
            const c = idCol(cid);
            if (r >= rc.r0 && r <= rc.r1 && c >= rc.c0 && c <= rc.c1) add(cand);
          }
          for (const row of structuralRows(rc.sheet)) {
            if (out.length >= 200) break;
            const sr = side === "new" ? target.newRows[row] : target.oldRows[row];
            if (sr < rc.r0 || sr > rc.r1) continue;
            for (let col2 = 0; col2 < target.columns.length; col2++) {
              const sc2 = side === "new" ? target.columns[col2].newIndex : target.columns[col2].oldIndex;
              if (sc2 >= rc.c0 && sc2 <= rc.c1) add({ sheet: rc.sheet, row, col: col2 });
            }
          }
        }
      }
      return out.slice(0, 200);
    }
  };
}
var originOf = (t) => t.origin ?? { row: 1, col: 0 };
function recalculate(diff, maps, traces, roots, staticNew, evalBudget) {
  const oldTables = diff.sheets.map((s) => s.diff.oldTable);
  const newTables = diff.sheets.map((s) => s.diff.newTable);
  if (!newTables.some((t) => t.formulas?.some((r) => r?.some(Boolean)))) return null;
  const oldHost = tableHost(oldTables, { names: diff.names?.old, tables: diff.tables?.old, date1904: diff.date1904?.old });
  const newHost = tableHost(newTables, { names: diff.names?.new, tables: diff.tables?.new, date1904: diff.date1904?.new });
  const oldEngine = new Engine(oldHost);
  const newEngine = new Engine(newHost);
  let evaluations = 0;
  let truncated = false;
  const traceKey = (ref) => ref.row * maps[ref.sheet].diff.columns.length + ref.col;
  const tablesOf = (side) => side === "new" ? newTables : oldTables;
  const hostOf = (side) => side === "new" ? newHost : oldHost;
  const sheetKey = (side, id) => {
    const s = idSheet(id);
    const o = originOf(tablesOf(side)[s]);
    return sheetCellKey(s, idRow(id) + o.row, idCol(id) + o.col);
  };
  const formulaAt = (side, id) => tablesOf(side)[idSheet(id)].formulas?.[idRow(id)]?.[idCol(id)] ?? null;
  const cachedAt = (side, id) => {
    const s = idSheet(id);
    const o = originOf(tablesOf(side)[s]);
    return hostOf(side).cached(s, idRow(id) + o.row, idCol(id) + o.col);
  };
  const compiled = { old: /* @__PURE__ */ new Map(), new: /* @__PURE__ */ new Map() };
  const compileAt = (side, id, formula) => {
    let c = compiled[side].get(id);
    if (c === void 0) {
      const s = idSheet(id);
      const o = originOf(tablesOf(side)[s]);
      try {
        c = (side === "new" ? newEngine : oldEngine).compile(formula, idRow(id) + o.row, idCol(id) + o.col);
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e;
        c = e;
      }
      compiled[side].set(id, c);
    }
    if (c instanceof Unsupported) throw c;
    return c;
  };
  const evaluate = (engine, host, side, id, formula, record = false) => {
    if (++evaluations > evalBudget) {
      truncated = true;
      throw new Unsupported("evaluation budget");
    }
    const c = compileAt(side, id, formula);
    const s = idSheet(id);
    const o = originOf(tablesOf(side)[s]);
    return engine.evaluate(c, host.sheets[s], idRow(id) + o.row, idCol(id) + o.col, record);
  };
  const CALIBRATION_SAMPLE = 8;
  const shapes = /* @__PURE__ */ new Map();
  const toAligned = (side, id) => {
    const sheet = idSheet(id);
    const m = maps[sheet];
    const row = m.rows[side][idRow(id)];
    const col2 = m.cols[side][idCol(id)];
    return row >= 0 && col2 >= 0 ? { sheet, row, col: col2 } : null;
  };
  const oldIdOf = (ref) => {
    const d = maps[ref.sheet].diff;
    const r = d.oldRows[ref.row];
    const c = d.columns[ref.col].oldIndex;
    return r >= 0 && c >= 0 ? cellId(ref.sheet, r, c) : null;
  };
  const rectsOf = (reads2) => {
    const out = [];
    for (const rd of reads2) {
      const s = newHost.sheetIndex(rd.sheet);
      if (s < 0) continue;
      const t = newTables[s];
      const o = originOf(t);
      const r0 = Math.max(0, rd.r0 - o.row);
      const r1 = Math.min(t.rows.length - 1, rd.r1 - o.row);
      const c0 = Math.max(0, rd.c0 - o.col);
      const c1 = Math.min(t.columns.length - 1, rd.c1 - o.col);
      if (r0 <= r1 && c0 <= c1) out.push({ sheet: s, r0, r1, c0, c1 });
    }
    return out;
  };
  const MAX_CONE = 3e5;
  const cone = /* @__PURE__ */ new Set();
  const queue = [];
  const queued = /* @__PURE__ */ new Set();
  for (const r of roots) {
    if (r.side !== "new" || queued.has(r.id)) continue;
    queued.add(r.id);
    queue.push(r.id);
    if (formulaAt("new", r.id) && !r.trace.external) cone.add(r.id);
  }
  const deps = [];
  for (let qi = 0; qi < queue.length && !truncated; qi++) {
    deps.length = 0;
    const id = queue[qi];
    staticNew.dependents(idSheet(id), idRow(id), idCol(id), deps);
    for (const d of deps) {
      if (cone.has(d)) continue;
      if (cone.size >= MAX_CONE) {
        truncated = true;
        break;
      }
      cone.add(d);
      if (!queued.has(d)) {
        queued.add(d);
        queue.push(d);
      }
    }
  }
  if (truncated) return { graph: staticNew, modelled: /* @__PURE__ */ new Set(), unmodelled: 0, truncated, butFor: () => /* @__PURE__ */ new Map() };
  const modelled = /* @__PURE__ */ new Set();
  const reads = /* @__PURE__ */ new Map();
  const stale = /* @__PURE__ */ new Map();
  const state = /* @__PURE__ */ new Map();
  let unmodelled = 0;
  let depth = 0;
  const baseGet = newHost.get;
  const baseGetRange = newHost.getRange;
  newHost.getRange = void 0;
  newHost.get = (sheet, r, c) => {
    const s = newHost.sheetIndex(sheet);
    if (s >= 0) {
      const o = originOf(newTables[s]);
      const tr = r - o.row;
      const tc = c - o.col;
      if (tr >= 0 && tc >= 0) {
        const id = cellId(s, tr, tc);
        if (state.get(id) !== 2 && cone.has(id)) visit(id);
      }
    }
    return baseGet(sheet, r, c);
  };
  const visit = (id) => {
    const st = state.get(id);
    if (st === 2) return;
    if (st === 1) throw new Unsupported("circular reference");
    if (depth > 400) throw new Unsupported("dependency chain too deep");
    state.set(id, 1);
    depth++;
    try {
      const formula = formulaAt("new", id);
      if (!formula) return;
      const ref = toAligned("new", id);
      const editedFormula = ref ? traces[ref.sheet].get(traceKey(ref))?.kind === "formula" : false;
      const oid = ref ? oldIdOf(ref) : null;
      const oldFormula = oid !== null ? formulaAt("old", oid) : null;
      const calibrate = () => {
        try {
          const oldCached = cachedAt("old", oid);
          return oldCached !== null && sameResult(evaluate(oldEngine, oldHost, "old", oid, oldFormula).value, oldCached, diff.date1904?.old);
        } catch (e) {
          if (!(e instanceof Unsupported)) throw e;
          return false;
        }
      };
      let calibrated = false;
      if (oid !== null && oldFormula !== null) {
        let shapeText = null;
        try {
          const sameSpot = oldFormula === formula && sheetKey("old", oid) === sheetKey("new", id);
          shapeText = (sameSpot ? compileAt("new", id, formula) : compileAt("old", oid, oldFormula)).shape;
        } catch (e) {
          if (!(e instanceof Unsupported)) throw e;
        }
        const shape = `${idSheet(oid)}\0${shapeText}`;
        let stat = shapes.get(shape);
        if (!stat) shapes.set(shape, stat = { passed: 0, failed: false });
        if (stat.failed || stat.passed < CALIBRATION_SAMPLE) {
          calibrated = calibrate();
          if (!calibrated) {
            stat.failed = true;
            unmodelled++;
            return;
          }
          stat.passed++;
        }
      }
      let result;
      try {
        result = evaluate(newEngine, newHost, "new", id, formula, true);
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e;
        unmodelled++;
        return;
      }
      if (!sameResult(result.value, cachedAt("new", id), diff.date1904?.new)) {
        if (oldFormula === null || editedFormula || !calibrated && !calibrate()) {
          unmodelled++;
          return;
        }
        stale.set(id, result.value);
        newHost.override(sheetKey("new", id), result.value);
      }
      modelled.add(id);
      reads.set(id, rectsOf(result.reads));
    } finally {
      depth--;
      state.set(id, 2);
    }
  };
  const sorted = topoOrder(queue, cone, staticNew);
  for (const id of sorted.cyclic) state.set(id, 2);
  unmodelled += sorted.cyclic.length;
  for (const id of sorted.order) {
    if (truncated) break;
    try {
      visit(id);
    } catch (e) {
      if (!(e instanceof Unsupported)) throw e;
      state.set(id, 2);
    }
  }
  newHost.get = baseGet;
  newHost.getRange = baseGetRange;
  for (const [id, value] of stale) {
    const ref = toAligned("new", id);
    if (!ref) continue;
    const oid = oldIdOf(ref);
    const oldValue = oid !== null ? cachedAt("old", oid) : null;
    const k = traceKey(ref);
    const info = { cached: toCellValue2(cachedAt("new", id)), recalculated: toCellValue2(value) };
    const t = traces[ref.sheet].get(k);
    if (t) t.stale = info;
    else if (!sameResult(oldValue, value, diff.date1904?.new)) traces[ref.sheet].set(k, { kind: "computed", causes: [], causeCount: 0, impacts: [], impactCount: 0, stale: info });
  }
  const graph = new OverlayGraph(staticNew, reads);
  const valueOf = (id) => stale.has(id) ? stale.get(id) : cachedAt("new", id);
  let layoutSame = null;
  const sameLayout = () => layoutSame ??= diff.sheets.every(({ diff: d }) => {
    const oo = originOf(d.oldTable);
    const no = originOf(d.newTable);
    if (oo.row !== no.row || oo.col !== no.col) return false;
    if (d.columns.some((c) => c.oldIndex !== c.newIndex)) return false;
    for (let r = 0; r < d.rowCount; r++) if (d.oldRows[r] !== d.newRows[r]) return false;
    return true;
  });
  const butFor = (rootList) => {
    const out = /* @__PURE__ */ new Map();
    const record = (ref) => {
      let m = out.get(ref.sheet);
      if (!m) out.set(ref.sheet, m = /* @__PURE__ */ new Map());
      const k = traceKey(ref);
      let e = m.get(k);
      if (!e) m.set(k, e = { causes: [], contributions: [], tainted: false });
      return e;
    };
    const isComputed = (ref) => !!ref && traces[ref.sheet].get(traceKey(ref))?.kind === "computed";
    rootList.forEach((root, ri) => {
      if (truncated || root.side !== "new") return;
      if (root.trace.kind !== "input" && root.trace.kind !== "formula") return;
      const oid = oldIdOf(root.ref);
      let without = oid !== null ? cachedAt("old", oid) : null;
      const oldFormula = oid !== null && root.trace.kind === "formula" ? formulaAt("old", oid) : null;
      if (oldFormula && sameLayout()) {
        const s = idSheet(root.id);
        const o = originOf(newTables[s]);
        try {
          if (++evaluations > evalBudget) throw new Unsupported("evaluation budget");
          without = newEngine.evaluate(oldFormula, newHost.sheets[s], idRow(root.id) + o.row, idCol(root.id) + o.col, false).value;
        } catch (e) {
          if (!(e instanceof Unsupported)) throw e;
        }
      }
      const saved = /* @__PURE__ */ new Map();
      const set = (k, v) => {
        if (!saved.has(k)) saved.set(k, newHost.overrides.has(k) ? newHost.overrides.get(k) : void 0);
        newHost.override(k, v);
      };
      set(sheetKey("new", root.id), without);
      const reach = /* @__PURE__ */ new Set();
      const queue2 = [root.id];
      const deps2 = [];
      for (let qi = 0; qi < queue2.length; qi++) {
        deps2.length = 0;
        const src = queue2[qi];
        graph.dependents(idSheet(src), idRow(src), idCol(src), deps2);
        for (const d of deps2) {
          if (reach.has(d) || d === root.id || !formulaAt("new", d)) continue;
          reach.add(d);
          queue2.push(d);
        }
      }
      const { order, preds, cyclic } = topoOrder([root.id], reach, graph);
      const changed = /* @__PURE__ */ new Set([root.id]);
      const tainted = new Set(cyclic);
      try {
        for (const d of order) {
          const p = preds.get(d) ?? [];
          if (p.some((x) => tainted.has(x)) || !modelled.has(d)) {
            tainted.add(d);
            continue;
          }
          if (!p.some((x) => changed.has(x))) continue;
          const v = evaluate(newEngine, newHost, "new", d, formulaAt("new", d)).value;
          const now = valueOf(d);
          if (sameResult(v, now, diff.date1904?.new)) continue;
          set(sheetKey("new", d), v);
          changed.add(d);
          const ref = toAligned("new", d);
          if (!isComputed(ref)) continue;
          const e = record(ref);
          e.causes.push(ri);
          e.contributions.push(typeof now === "number" && typeof v === "number" ? numericDelta(v, now).delta : null);
        }
        for (const d of tainted) {
          const ref = toAligned("new", d);
          if (isComputed(ref)) record(ref).tainted = true;
        }
      } catch (e) {
        if (!(e instanceof Unsupported)) throw e;
        truncated = true;
      } finally {
        for (const [k, v] of saved) newHost.override(k, v);
      }
    });
    return out;
  };
  return { graph, modelled, unmodelled, get truncated() {
    return truncated;
  }, butFor };
}
function toCellValue2(v) {
  return isError(v) ? v.code : v;
}

// packages/spreadiff/src/core/xlsx.ts
function parseXlsx(data, options = {}) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const files = unzipSync(bytes, {
    filter: (f) => f.name === "xl/workbook.xml" || f.name === "xl/sharedStrings.xml" || f.name === "xl/styles.xml" || f.name === "xl/_rels/workbook.xml.rels" || f.name.startsWith("xl/worksheets/") || f.name.startsWith("xl/tables/") || // Charts and pivot tables (see parseObjects): their links and definitions, not their data.
    f.name.startsWith("xl/chartsheets/") || f.name.startsWith("xl/drawings/_rels/") || /^xl\/charts\/chart\d+\.xml$/.test(f.name) || f.name.startsWith("xl/pivotTables/") || /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(f.name)
  });
  let lower = null;
  const part = (path) => {
    if (files[path]) return files[path];
    lower ??= new Map(Object.keys(files).map((k2) => [k2.toLowerCase(), k2]));
    const k = lower.get(path.toLowerCase());
    return k ? files[k] : void 0;
  };
  const text3 = (path) => {
    const f = part(path);
    return f ? stripPrefixes(strFromU8(f)) : "";
  };
  const workbookXml = text3("xl/workbook.xml");
  if (!workbookXml) throw new Error("Not an XLSX file (missing xl/workbook.xml)");
  let date19042 = false;
  eachElement(workbookXml, "workbookPr", (a) => {
    const v = attr(a, "date1904");
    date19042 = v === "1" || v === "true";
  });
  const rels = /* @__PURE__ */ new Map();
  eachElement(text3("xl/_rels/workbook.xml.rels"), "Relationship", (a) => {
    const id = attr(a, "Id");
    const target = attr(a, "Target");
    if (id && target) rels.set(id, target.startsWith("/") ? target.slice(1) : `xl/${target}`);
  });
  const strings = parseSharedStrings(text3("xl/sharedStrings.xml"));
  const stylesXml = text3("xl/styles.xml");
  const dateStyles = parseDateStyles(stylesXml);
  const labelStyles = parseLabelStyles(stylesXml);
  const numberFormats = parseNumberFormats(stylesXml);
  const raws = [];
  const sheetParts = [];
  const tables = [];
  const allSheets = [];
  eachElement(workbookXml, "sheet", (a) => {
    const name = decode(attr(a, "name") ?? `Sheet${raws.length + 1}`);
    allSheets.push(name);
    if (options.sheets && !options.sheets.includes(name)) return;
    const rid = attr(a, "r:id") ?? /\s[\w.-]+:id\s*=\s*["']([^"']*)/.exec(a)?.[1];
    const path = rid ? rels.get(rid) : void 0;
    if (!path || !part(path)) return;
    raws.push(parseCells(name, text3(path), strings, dateStyles, date19042, labelStyles, numberFormats));
    sheetParts.push({ name, path });
    const relsPath = path.replace(/([^/]+)$/, "_rels/$1.rels");
    eachElement(text3(relsPath), "Relationship", (ra) => {
      const target = attr(ra, "Target");
      if (!target || !/\/table$/.test(attr(ra, "Type") ?? "")) return;
      const tablePath = resolvePath(path, target);
      const tableXml = text3(tablePath);
      const t = parseTableDefinition(tableXml, name);
      if (t) tables.push(t);
      const raw = raws[raws.length - 1];
      if (raw.hidden.length) raw.filters.push(...activeFilters(tableXml));
    });
  });
  const names = {};
  const localNames = /* @__PURE__ */ new Map();
  eachElement(workbookXml, "definedName", (a, body) => {
    const name = attr(a, "name");
    if (!name || !body || name.startsWith("_xlnm.")) return;
    const local = attr(a, "localSheetId");
    if (local === void 0) names[decode(name)] = decode(body);
    else {
      const sheet = allSheets[Number(local)];
      if (sheet === void 0) return;
      let m = localNames.get(sheet);
      if (!m) localNames.set(sheet, m = {});
      m[decode(name)] = decode(body);
    }
  });
  const sheets = tablesFromRaw(raws, options.header ?? "auto", names, tables, localNames);
  const workbook = {
    name: options.name ?? "workbook.xlsx",
    sheets,
    ...date19042 ? { date1904: date19042 } : {},
    ...Object.keys(names).length ? { names } : {},
    ...tables.length ? { tables } : {}
  };
  const { charts, pivots } = parseObjects(text3, sheetParts);
  if (charts.length) workbook.charts = charts;
  if (pivots.length) workbook.pivots = pivots;
  if (options.computeMissing !== false) {
    const m = computeMissingValues(workbook);
    if (m.computed || m.failed) workbook.computedValues = m;
  }
  return workbook;
}
function parseTableDefinition(xml, sheet) {
  let def = null;
  eachElement(xml, "table", (a) => {
    const name = attr(a, "displayName") ?? attr(a, "name");
    const ref = attr(a, "ref");
    if (!name || !ref || def) return;
    const [from, to = from] = ref.split(":");
    const [r0, c0] = decodeRef(from);
    const [r1, c1] = decodeRef(to);
    const headerRows = Number(attr(a, "headerRowCount") ?? 1);
    const totalsRows = Number(attr(a, "totalsRowCount") ?? 0);
    def = { name: decode(name), sheet, range: { r0, c0, r1, c1 }, columns: [], headerRows, totalsRows };
  });
  if (!def) return null;
  const d = def;
  eachElement(xml, "tableColumn", (a) => {
    d.columns.push(decode(attr(a, "name") ?? ""));
  });
  return d;
}
function stripPrefixes(xml) {
  if (!/<[A-Za-z_][\w.-]*:(?:workbook|worksheet|sst|styleSheet|Relationships)\b/.test(xml.slice(0, 4096))) return xml;
  return xml.replace(/<(\/?)[A-Za-z_][\w.-]*:/g, "<$1");
}
function parseSharedStrings(xml) {
  const out = [];
  eachElement(xml, "si", (_, body) => out.push(body ? textRuns(body) : ""));
  return out;
}
var BUILTIN_DATE_FORMATS = /* @__PURE__ */ new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
function parseDateStyles(xml) {
  const custom = /* @__PURE__ */ new Map();
  eachElement(xml, "numFmt", (a) => {
    custom.set(Number(attr(a, "numFmtId")), decode(attr(a, "formatCode") ?? ""));
  });
  const isDate = (id) => {
    if (BUILTIN_DATE_FORMATS.has(id)) return true;
    const code = custom.get(id);
    if (!code) return false;
    const stripped = code.replace(/"[^"]*"|\[[^\]]*\]|\\./g, "");
    return /[dmyhs]/i.test(stripped) && !/^[#0.,%\s]*$/.test(stripped);
  };
  const out = /* @__PURE__ */ new Set();
  eachElement(xml, "cellXfs", (_, body) => {
    let i = 0;
    eachElement(body ?? "", "xf", (a) => {
      if (isDate(Number(attr(a, "numFmtId") ?? 0))) out.add(i);
      i++;
    });
  });
  return out;
}
function parseNumberFormats(xml) {
  const custom = /* @__PURE__ */ new Map();
  eachElement(xml, "numFmt", (a) => {
    custom.set(Number(attr(a, "numFmtId")), decode(attr(a, "formatCode") ?? ""));
  });
  const out = /* @__PURE__ */ new Map();
  eachElement(xml, "cellXfs", (_, body) => {
    let i = 0;
    eachElement(body ?? "", "xf", (a) => {
      const id = Number(attr(a, "numFmtId") ?? 0);
      const code = custom.get(id) ?? BUILTIN_FORMATS[id];
      if (code && !/^general$/i.test(code)) out.set(i, code);
      i++;
    });
  });
  return out;
}
function parseLabelStyles(xml) {
  const bold = [];
  eachElement(xml, "fonts", (_, body) => {
    eachElement(body ?? "", "font", (_a2, fb) => {
      const b = /<b(?:\s+val\s*=\s*["']([^"']*)["'])?\s*\/?>/.exec(fb ?? "");
      bold.push(!!b && b[1] !== "0" && b[1] !== "false");
    });
  });
  const out = /* @__PURE__ */ new Map();
  eachElement(xml, "cellXfs", (_, body) => {
    let i = 0;
    eachElement(body ?? "", "xf", (a, xb) => {
      const indent = Number(/<alignment\b[^>]*\sindent\s*=\s*["'](\d+)["']/.exec(xb ?? "")?.[1] ?? 0);
      const b = bold[Number(attr(a, "fontId") ?? 0)] ?? false;
      if (indent > 0 || b) out.set(i, { ...indent > 0 ? { indent } : {}, ...b ? { bold: true } : {} });
      i++;
    });
  });
  return out;
}
function shiftFormula(formula, dRow, dCol) {
  const colName = (abs, letters) => {
    if (abs) return abs + letters;
    let c = 0;
    for (const ch of letters) c = c * 26 + (ch.charCodeAt(0) - 64);
    const nc = c - 1 + dCol;
    return nc < 0 ? null : columnLabel(nc);
  };
  const rowNum = (abs, digits) => {
    if (abs) return abs + digits;
    const nr = Number(digits) + dRow;
    return nr < 1 ? null : String(nr);
  };
  return formula.replace(
    // String literals and quoted sheet names pass through untouched. A ref
    // may follow `!` (Sheet!A1) but not be followed by `!` (it's a sheet name).
    // Whole columns (B:B) and whole rows (2:5) shift too.
    /("[^"]*")|('[^']*')|(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})(?![\d(A-Za-z_!])|(?<![A-Za-z0-9_.$:])(\$?)(\d+):(\$?)(\d+)(?![\d(A-Za-z_!.:])|(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])/g,
    (all, str2, sheet, ca1, c1, ca2, c2, ra1, r1, ra2, r2, cAbs, col2, rAbs, row) => {
      if (str2 || sheet) return all;
      if (c1 !== void 0) {
        const a = colName(ca1, c1);
        const b = colName(ca2, c2);
        return a && b ? `${a}:${b}` : all;
      }
      if (r1 !== void 0) {
        const a = rowNum(ra1, r1);
        const b = rowNum(ra2, r2);
        return a && b ? `${a}:${b}` : all;
      }
      const c = colName(cAbs, col2);
      const r = rowNum(rAbs, row);
      if (c === null || r === null) return all;
      return `${c}${r}`;
    }
  );
}
function activeFilters(xml) {
  const out = [];
  eachElement(xml, "autoFilter", (a, body) => {
    const ref = attr(a, "ref");
    if (!ref || !body || !/<(?:\w+:)?filterColumn\b/.test(body)) return;
    const [from, to = from] = ref.split(":");
    out.push({ r0: decodeRef(from)[0] + 1, r1: decodeRef(to)[0] });
  });
  return out;
}
function parseCells(name, xml, strings, dateStyles, date19042, labelStyles, numberFormats) {
  const styled = [];
  const codes = [""];
  const codeIds = /* @__PURE__ */ new Map();
  const fmtGrid = [];
  const grid = [];
  const formulaGrid = [];
  const shared = /* @__PURE__ */ new Map();
  const blocks = [];
  let minCol = Infinity;
  let maxCol = -1;
  let implicitRow = 0;
  const hidden = [];
  eachElement(xml, "row", (rowAttrs, rowBody) => {
    const rAttr = attr(rowAttrs, "r");
    const r = rAttr ? Number(rAttr) - 1 : implicitRow;
    implicitRow = r + 1;
    const h = attr(rowAttrs, "hidden");
    if ((h === "1" || h === "true") && r >= 0) hidden.push(r);
    if (!rowBody || !(r >= 0)) return;
    let implicitCol = 0;
    eachElement(rowBody, "c", (attrs, cellBody) => {
      const body = cellBody ?? "";
      const ref = attr(attrs, "r");
      const c = ref ? decodeRef(ref)[1] : implicitCol;
      implicitCol = c + 1;
      if (!(c >= 0)) return;
      const t = attr(attrs, "t") ?? "n";
      let raw = null;
      eachElement(body, "v", (_, v) => {
        raw = decode(v ?? "");
      });
      let value = null;
      switch (t) {
        case "s":
          value = raw === null ? null : strings[Number(raw)] ?? null;
          break;
        case "inlineStr":
          value = textRuns(body);
          break;
        case "str":
        case "e":
        case "d":
          value = raw;
          break;
        case "b":
          value = raw === "1";
          break;
        default:
          if (raw !== null && raw !== "") {
            const num3 = Number(raw);
            const style = Number(attr(attrs, "s") ?? -1);
            value = dateStyles.has(style) ? serialToIso(num3, date19042) ?? num3 : num3;
          }
      }
      let formula = null;
      eachElement(body, "f", (fAttrs, fBody) => {
        const text3 = fBody ? decode(fBody) : "";
        const ft = attr(fAttrs, "t");
        const si = attr(fAttrs, "si");
        if (ft === "shared" && si !== void 0) {
          if (text3) {
            shared.set(si, { formula: text3, row: r, col: c });
            formula = text3;
          } else {
            const master = shared.get(si);
            if (master) formula = shiftFormula(master.formula, r - master.row, c - master.col);
          }
        } else if (ft === "dataTable") {
          const r1 = attr(fAttrs, "r1");
          const r2 = attr(fAttrs, "r2");
          formula = `TABLE(${[r1, r2].filter(Boolean).join(",")})`;
          const blockRef = attr(fAttrs, "ref");
          if (blockRef) blocks.push({ ref: blockRef, formula, array: false });
        } else if (text3) {
          formula = text3;
          const blockRef = attr(fAttrs, "ref");
          if (ft === "array" && blockRef) blocks.push({ ref: blockRef, formula: text3, array: true });
        }
      });
      if (formula !== null && /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(formula) && value === formula) formula = null;
      if (value === null && formula === null) return;
      if (typeof value === "string" && labelStyles.size) {
        const ls = labelStyles.get(Number(attr(attrs, "s") ?? 0));
        if (ls) styled.push({ r, c, ...ls });
      }
      (grid[r] ??= [])[c] = value;
      if (formula !== null) (formulaGrid[r] ??= [])[c] = "=" + formula;
      if (numberFormats.size && value !== null && typeof value !== "boolean") {
        const code = numberFormats.get(Number(attr(attrs, "s") ?? 0));
        if (code && (typeof value !== "string" || dateStyles.has(Number(attr(attrs, "s"))) || code.includes("@"))) {
          let id = codeIds.get(code);
          if (id === void 0) {
            id = codes.length;
            if (id > 65535) id = 0;
            else {
              codes.push(code);
              codeIds.set(code, id);
            }
          }
          if (id) (fmtGrid[r] ??= [])[c] = id;
        }
      }
      if (c < minCol) minCol = c;
      if (c > maxCol) maxCol = c;
    });
  });
  const arrays = [];
  for (const b of blocks) {
    const [from, to = from] = b.ref.split(":");
    const [r0, c0] = decodeRef(from);
    const [r1, c1] = decodeRef(to);
    if (!(r1 - r0 < 1e6 && c1 - c0 < 2e4)) continue;
    if (b.array && (r1 > r0 || c1 > c0)) arrays.push({ r0, c0, r1, c1 });
    for (let r = r0; r <= r1; r++) {
      const fr = formulaGrid[r] ??= [];
      for (let c = c0; c <= c1; c++) {
        if (fr[c]) continue;
        fr[c] = "=" + b.formula;
        (grid[r] ??= [])[c] ??= null;
        if (c < minCol) minCol = c;
        if (c > maxCol) maxCol = c;
      }
    }
  }
  const merges = [];
  eachElement(xml, "mergeCell", (a) => {
    const ref = attr(a, "ref");
    if (!ref || !ref.includes(":")) return;
    const [from, to] = ref.split(":");
    const [r0, c0] = decodeRef(from);
    const [r1, c1] = decodeRef(to);
    if (r1 > r0 || c1 > c0) merges.push({ r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) });
  });
  let first = 0;
  while (first < grid.length && !grid[first]) first++;
  return {
    name,
    grid,
    formulaGrid,
    minCol,
    maxCol,
    first,
    arrays,
    merges,
    labelStyles: styled,
    formats: codes.length > 1 ? { codes, grid: fmtGrid, date1904: date19042 } : void 0,
    hidden,
    filters: hidden.length ? activeFilters(xml) : []
  };
}

// packages/spreadiff/src/core/ftab.ts
var FTAB = {
  0: "COUNT",
  1: "IF",
  2: "ISNA",
  3: "ISERROR",
  4: "SUM",
  5: "AVERAGE",
  6: "MIN",
  7: "MAX",
  8: "ROW",
  9: "COLUMN",
  10: "NA",
  11: "NPV",
  12: "STDEV",
  13: "DOLLAR",
  14: "FIXED",
  15: "SIN",
  16: "COS",
  17: "TAN",
  18: "ATAN",
  19: "PI",
  20: "SQRT",
  21: "EXP",
  22: "LN",
  23: "LOG10",
  24: "ABS",
  25: "INT",
  26: "SIGN",
  27: "ROUND",
  28: "LOOKUP",
  29: "INDEX",
  30: "REPT",
  31: "MID",
  32: "LEN",
  33: "VALUE",
  34: "TRUE",
  35: "FALSE",
  36: "AND",
  37: "OR",
  38: "NOT",
  39: "MOD",
  40: "DCOUNT",
  41: "DSUM",
  42: "DAVERAGE",
  43: "DMIN",
  44: "DMAX",
  45: "DSTDEV",
  46: "VAR",
  47: "DVAR",
  48: "TEXT",
  49: "LINEST",
  50: "TREND",
  51: "LOGEST",
  52: "GROWTH",
  56: "PV",
  57: "FV",
  58: "NPER",
  59: "PMT",
  60: "RATE",
  61: "MIRR",
  62: "IRR",
  63: "RAND",
  64: "MATCH",
  65: "DATE",
  66: "TIME",
  67: "DAY",
  68: "MONTH",
  69: "YEAR",
  70: "WEEKDAY",
  71: "HOUR",
  72: "MINUTE",
  73: "SECOND",
  74: "NOW",
  75: "AREAS",
  76: "ROWS",
  77: "COLUMNS",
  78: "OFFSET",
  82: "SEARCH",
  83: "TRANSPOSE",
  86: "TYPE",
  97: "ATAN2",
  98: "ASIN",
  99: "ACOS",
  100: "CHOOSE",
  101: "HLOOKUP",
  102: "VLOOKUP",
  105: "ISREF",
  109: "LOG",
  111: "CHAR",
  112: "LOWER",
  113: "UPPER",
  114: "PROPER",
  115: "LEFT",
  116: "RIGHT",
  117: "EXACT",
  118: "TRIM",
  119: "REPLACE",
  120: "SUBSTITUTE",
  121: "CODE",
  124: "FIND",
  125: "CELL",
  126: "ISERR",
  127: "ISTEXT",
  128: "ISNUMBER",
  129: "ISBLANK",
  130: "T",
  131: "N",
  140: "DATEVALUE",
  141: "TIMEVALUE",
  142: "SLN",
  143: "SYD",
  144: "DDB",
  148: "INDIRECT",
  162: "CLEAN",
  163: "MDETERM",
  164: "MINVERSE",
  165: "MMULT",
  167: "IPMT",
  168: "PPMT",
  169: "COUNTA",
  183: "PRODUCT",
  184: "FACT",
  189: "DPRODUCT",
  190: "ISNONTEXT",
  193: "STDEVP",
  194: "VARP",
  195: "DSTDEVP",
  196: "DVARP",
  197: "TRUNC",
  198: "ISLOGICAL",
  199: "DCOUNTA",
  204: "USDOLLAR",
  205: "FINDB",
  206: "SEARCHB",
  207: "REPLACEB",
  208: "LEFTB",
  209: "RIGHTB",
  210: "MIDB",
  211: "LENB",
  212: "ROUNDUP",
  213: "ROUNDDOWN",
  214: "ASC",
  215: "DBCS",
  216: "RANK",
  219: "ADDRESS",
  220: "DAYS360",
  221: "TODAY",
  222: "VDB",
  227: "MEDIAN",
  228: "SUMPRODUCT",
  229: "SINH",
  230: "COSH",
  231: "TANH",
  232: "ASINH",
  233: "ACOSH",
  234: "ATANH",
  235: "DGET",
  244: "INFO",
  247: "DB",
  252: "FREQUENCY",
  261: "ERROR.TYPE",
  269: "AVEDEV",
  270: "BETADIST",
  271: "GAMMALN",
  272: "BETAINV",
  273: "BINOMDIST",
  274: "CHIDIST",
  275: "CHIINV",
  276: "COMBIN",
  277: "CONFIDENCE",
  278: "CRITBINOM",
  279: "EVEN",
  280: "EXPONDIST",
  281: "FDIST",
  282: "FINV",
  283: "FISHER",
  284: "FISHERINV",
  285: "FLOOR",
  286: "GAMMADIST",
  287: "GAMMAINV",
  288: "CEILING",
  289: "HYPGEOMDIST",
  290: "LOGNORMDIST",
  291: "LOGINV",
  292: "NEGBINOMDIST",
  293: "NORMDIST",
  294: "NORMSDIST",
  295: "NORMINV",
  296: "NORMSINV",
  297: "STANDARDIZE",
  298: "ODD",
  299: "PERMUT",
  300: "POISSON",
  301: "TDIST",
  302: "WEIBULL",
  303: "SUMXMY2",
  304: "SUMX2MY2",
  305: "SUMX2PY2",
  306: "CHITEST",
  307: "CORREL",
  308: "COVAR",
  309: "FORECAST",
  310: "FTEST",
  311: "INTERCEPT",
  312: "PEARSON",
  313: "RSQ",
  314: "STEYX",
  315: "SLOPE",
  316: "TTEST",
  317: "PROB",
  318: "DEVSQ",
  319: "GEOMEAN",
  320: "HARMEAN",
  321: "SUMSQ",
  322: "KURT",
  323: "SKEW",
  324: "ZTEST",
  325: "LARGE",
  326: "SMALL",
  327: "QUARTILE",
  328: "PERCENTILE",
  329: "PERCENTRANK",
  330: "MODE",
  331: "TRIMMEAN",
  332: "TINV",
  336: "CONCATENATE",
  337: "POWER",
  342: "RADIANS",
  343: "DEGREES",
  344: "SUBTOTAL",
  345: "SUMIF",
  346: "COUNTIF",
  347: "COUNTBLANK",
  350: "ISPMT",
  351: "DATEDIF",
  352: "DATESTRING",
  353: "NUMBERSTRING",
  354: "ROMAN",
  358: "GETPIVOTDATA",
  359: "HYPERLINK",
  360: "PHONETIC",
  361: "AVERAGEA",
  362: "MAXA",
  363: "MINA",
  364: "STDEVPA",
  365: "VARPA",
  366: "STDEVA",
  367: "VARA",
  368: "BAHTTEXT"
};
var FTAB_XLSB = {
  380: "CUBEVALUE",
  381: "CUBEMEMBER",
  382: "CUBEMEMBERPROPERTY",
  383: "CUBERANKEDMEMBER",
  384: "HEX2BIN",
  385: "HEX2DEC",
  386: "HEX2OCT",
  387: "DEC2BIN",
  388: "DEC2HEX",
  389: "DEC2OCT",
  390: "OCT2BIN",
  391: "OCT2HEX",
  392: "OCT2DEC",
  393: "BIN2DEC",
  394: "BIN2OCT",
  395: "BIN2HEX",
  396: "IMSUB",
  397: "IMDIV",
  398: "IMPOWER",
  399: "IMABS",
  400: "IMSQRT",
  401: "IMLN",
  402: "IMLOG2",
  403: "IMLOG10",
  404: "IMSIN",
  405: "IMCOS",
  406: "IMEXP",
  407: "IMARGUMENT",
  408: "IMCONJUGATE",
  409: "IMAGINARY",
  410: "IMREAL",
  411: "COMPLEX",
  412: "IMSUM",
  413: "IMPRODUCT",
  414: "SERIESSUM",
  415: "FACTDOUBLE",
  416: "SQRTPI",
  417: "QUOTIENT",
  418: "DELTA",
  419: "GESTEP",
  420: "ISEVEN",
  421: "ISODD",
  422: "MROUND",
  423: "ERF",
  424: "ERFC",
  425: "BESSELJ",
  426: "BESSELK",
  427: "BESSELY",
  428: "BESSELI",
  429: "XIRR",
  430: "XNPV",
  431: "PRICEMAT",
  432: "YIELDMAT",
  433: "INTRATE",
  434: "RECEIVED",
  435: "DISC",
  436: "PRICEDISC",
  437: "YIELDDISC",
  438: "TBILLEQ",
  439: "TBILLPRICE",
  440: "TBILLYIELD",
  441: "PRICE",
  442: "YIELD",
  443: "DOLLARDE",
  444: "DOLLARFR",
  445: "NOMINAL",
  446: "EFFECT",
  447: "CUMPRINC",
  448: "CUMIPMT",
  449: "EDATE",
  450: "EOMONTH",
  451: "YEARFRAC",
  452: "COUPDAYBS",
  453: "COUPDAYS",
  454: "COUPDAYSNC",
  455: "COUPNCD",
  456: "COUPNUM",
  457: "COUPPCD",
  458: "DURATION",
  459: "MDURATION",
  460: "ODDLPRICE",
  461: "ODDLYIELD",
  462: "ODDFPRICE",
  463: "ODDFYIELD",
  464: "RANDBETWEEN",
  465: "WEEKNUM",
  466: "AMORDEGRC",
  467: "AMORLINC",
  468: "CONVERT",
  469: "ACCRINT",
  470: "ACCRINTM",
  471: "WORKDAY",
  472: "NETWORKDAYS",
  473: "GCD",
  474: "MULTINOMIAL",
  475: "LCM",
  476: "FVSCHEDULE",
  477: "CUBEKPIMEMBER",
  478: "CUBESET",
  479: "CUBESETCOUNT",
  480: "IFERROR",
  481: "COUNTIFS",
  482: "SUMIFS",
  483: "AVERAGEIF",
  484: "AVERAGEIFS"
};
var FIXED_ARITY = {
  2: 1,
  3: 1,
  10: 0,
  15: 1,
  16: 1,
  17: 1,
  18: 1,
  19: 0,
  20: 1,
  21: 1,
  22: 1,
  23: 1,
  24: 1,
  25: 1,
  26: 1,
  27: 2,
  30: 2,
  31: 3,
  32: 1,
  33: 1,
  34: 0,
  35: 0,
  38: 1,
  39: 2,
  40: 3,
  41: 3,
  42: 3,
  43: 3,
  44: 3,
  45: 3,
  47: 3,
  48: 2,
  61: 3,
  63: 0,
  65: 3,
  66: 3,
  67: 1,
  68: 1,
  69: 1,
  71: 1,
  72: 1,
  73: 1,
  74: 0,
  75: 1,
  76: 1,
  77: 1,
  83: 1,
  86: 1,
  97: 2,
  98: 1,
  99: 1,
  105: 1,
  111: 1,
  112: 1,
  113: 1,
  114: 1,
  117: 2,
  118: 1,
  119: 4,
  121: 1,
  126: 1,
  127: 1,
  128: 1,
  129: 1,
  130: 1,
  131: 1,
  140: 1,
  141: 1,
  142: 3,
  143: 4,
  162: 1,
  163: 1,
  164: 1,
  165: 2,
  184: 1,
  189: 3,
  190: 1,
  195: 3,
  196: 3,
  198: 1,
  199: 3,
  207: 4,
  210: 3,
  211: 1,
  214: 1,
  215: 1,
  221: 0,
  229: 1,
  230: 1,
  231: 1,
  232: 1,
  233: 1,
  234: 1,
  235: 3,
  244: 1,
  252: 2,
  261: 1,
  271: 1,
  273: 4,
  274: 2,
  275: 2,
  276: 2,
  277: 3,
  278: 3,
  279: 1,
  280: 3,
  281: 3,
  282: 3,
  283: 1,
  284: 1,
  285: 2,
  286: 4,
  287: 3,
  288: 2,
  289: 4,
  290: 3,
  291: 3,
  292: 3,
  293: 4,
  294: 1,
  295: 3,
  296: 1,
  297: 3,
  298: 1,
  299: 2,
  300: 3,
  301: 3,
  302: 4,
  303: 2,
  304: 2,
  305: 2,
  306: 2,
  307: 2,
  308: 2,
  309: 3,
  310: 2,
  311: 2,
  312: 2,
  313: 2,
  314: 2,
  315: 2,
  316: 4,
  325: 2,
  326: 2,
  327: 2,
  328: 2,
  331: 2,
  332: 2,
  337: 2,
  342: 1,
  343: 1,
  346: 2,
  347: 1,
  350: 4,
  351: 3,
  360: 1,
  368: 1
};

// packages/spreadiff/src/core/ptg.ts
var ERRORS = { 0: "#NULL!", 7: "#DIV/0!", 15: "#VALUE!", 23: "#REF!", 29: "#NAME?", 36: "#NUM!", 42: "#N/A", 43: "#GETTING_DATA" };
var BINARY2 = { 3: "+", 4: "-", 5: "*", 6: "/", 7: "^", 8: "&", 9: "<", 10: "<=", 11: "=", 12: ">=", 13: ">", 14: "<>", 15: " ", 16: ",", 17: ":" };
var Undecodable = class extends Error {
};
var quoteSheet = (s) => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) && !/^[A-Za-z]{1,3}\d+$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`;
var num2 = (n) => String(n).replace(/e\+/, "E+").replace(/e-/, "E-");
function decodeFormula(rgce, ctx) {
  try {
    return decode2(rgce, ctx);
  } catch (e) {
    if (e instanceof Undecodable || e instanceof RangeError) return null;
    throw e;
  }
}
function decode2(rgce, ctx) {
  const dv = new DataView(rgce.buffer, rgce.byteOffset, rgce.byteLength);
  const b12 = ctx.flavor === "biff12";
  const maxRow = b12 ? 1048575 : 65535;
  const maxCol = b12 ? 16383 : 255;
  const stack = [];
  let extraPos = 0;
  const pop = () => {
    const v = stack.pop();
    if (v === void 0) throw new Undecodable("stack");
    return v;
  };
  const cell = (row, colField, relative) => {
    const colRel = (colField & 16384) !== 0;
    const rowRel = (colField & 32768) !== 0;
    let c = colField & 16383;
    let r = row;
    if (relative) {
      if (rowRel) r = (((b12 ? r | 0 : r << 16 >> 16) + ctx.row) % (maxRow + 1) + maxRow + 1) % (maxRow + 1);
      if (colRel) c = (((b12 ? c << 18 >> 18 : (c & 255) << 24 >> 24) + ctx.col) % (maxCol + 1) + maxCol + 1) % (maxCol + 1);
    }
    if (!b12) c &= 255;
    return { text: `${colRel ? "" : "$"}${columnLabel(c)}${rowRel ? "" : "$"}${r + 1}`, r, c, colRel, rowRel };
  };
  const area = (r0, r1, c0f, c1f, relative) => {
    const a = cell(r0, c0f, relative);
    const b = cell(r1, c1f, relative);
    if (a.r === 0 && b.r === maxRow) return `${a.colRel ? "" : "$"}${columnLabel(a.c)}:${b.colRel ? "" : "$"}${columnLabel(b.c)}`;
    if (a.c === 0 && b.c === maxCol) return `${a.rowRel ? "" : "$"}${a.r + 1}:${b.rowRel ? "" : "$"}${b.r + 1}`;
    return `${a.text}:${b.text}`;
  };
  const sheet3d = (ixti) => {
    const s = ctx.sheetOf(ixti);
    if (!s) throw new Undecodable("external reference");
    const sheets = s.first === s.last ? s.first : `${s.first}:${s.last}`;
    if (s.book === void 0) return s.first === s.last ? `${quoteSheet(s.first)}!` : `${quoteSheet(s.first)}:${quoteSheet(s.last)}!`;
    return quoteSheet(sheets).startsWith("'") ? `'[${s.book}]${sheets.replace(/'/g, "''")}'!` : `[${s.book}]${sheets}!`;
  };
  const str8 = (pos) => {
    const cch = rgce[pos];
    const high = rgce[pos + 1] & 1;
    let s = "";
    let p = pos + 2;
    for (let i2 = 0; i2 < cch; i2++) {
      s += String.fromCharCode(high ? dv.getUint16(p, true) : rgce[p]);
      p += high ? 2 : 1;
    }
    return [s, p];
  };
  const arrayConstant = () => {
    const x = ctx.extra;
    if (!x) throw new Undecodable("array constant");
    const xv = new DataView(x.buffer, x.byteOffset, x.byteLength);
    let p = extraPos;
    let rows;
    let cols;
    if (b12) {
      rows = xv.getUint32(p, true);
      cols = xv.getUint32(p + 4, true);
      p += 8;
    } else {
      cols = x[p] + 1;
      rows = xv.getUint16(p + 1, true) + 1;
      p += 3;
    }
    if (rows * cols > x.length) throw new Undecodable("array constant size");
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const t = x[p++];
        const kind = b12 ? ["num", "str", "bool", "", "err"][t] : { 1: "num", 2: "str", 4: "bool", 16: "err" }[t];
        if (kind === "num") {
          row.push(num2(xv.getFloat64(p, true)));
          p += 8;
        } else if (kind === "str") {
          const cch = xv.getUint16(p, true);
          const high = b12 ? 1 : x[p + 2] & 1;
          p += b12 ? 2 : 3;
          let s = "";
          for (let i2 = 0; i2 < cch; i2++) {
            s += String.fromCharCode(high ? xv.getUint16(p, true) : x[p]);
            p += high ? 2 : 1;
          }
          row.push(`"${s.replace(/"/g, '""')}"`);
        } else if (kind === "bool") {
          row.push(x[p] ? "TRUE" : "FALSE");
          p += b12 ? 1 : 8;
        } else if (kind === "err") {
          row.push(ERRORS[x[p]] ?? "#N/A");
          p += b12 ? 4 : 8;
        } else if (!b12 && t === 0) {
          row.push("");
          p += 8;
        } else throw new Undecodable("array value");
      }
      out.push(row.join(","));
    }
    extraPos = p;
    return `{${out.join(";")}}`;
  };
  let i = 0;
  while (i < rgce.length) {
    const ptg = rgce[i++];
    const base = ptg >= 32 ? ptg & 31 | 32 : ptg;
    switch (base) {
      case 1:
        throw new Undecodable("shared or array formula reference");
      // resolved by the reader
      case 18:
        stack.push(`+${pop()}`);
        break;
      case 19:
        stack.push(`-${pop()}`);
        break;
      case 20:
        stack.push(`${pop()}%`);
        break;
      case 21:
        stack.push(`(${pop()})`);
        break;
      case 22:
        stack.push("");
        break;
      case 23: {
        let s;
        if (b12) {
          const cch = dv.getUint16(i, true);
          s = "";
          for (let k = 0; k < cch; k++) s += String.fromCharCode(dv.getUint16(i + 2 + k * 2, true));
          i += 2 + cch * 2;
        } else [s, i] = str8(i);
        stack.push(`"${s.replace(/"/g, '""')}"`);
        break;
      }
      case 24:
        throw new Undecodable("extended token");
      // structured references (.xlsb) and the like
      case 25: {
        const grbit = rgce[i];
        if (grbit & 4) {
          const cases = dv.getUint16(i + 1, true);
          i += 3 + (cases + 1) * 2;
        } else i += 3;
        if (grbit & 16) stack.push(`SUM(${pop()})`);
        break;
      }
      case 28:
        stack.push(ERRORS[rgce[i]] ?? "#N/A");
        i += 1;
        break;
      case 29:
        stack.push(rgce[i] ? "TRUE" : "FALSE");
        i += 1;
        break;
      case 30:
        stack.push(String(dv.getUint16(i, true)));
        i += 2;
        break;
      case 31:
        stack.push(num2(dv.getFloat64(i, true)));
        i += 8;
        break;
      case 32:
        i += b12 ? 14 : 7;
        stack.push(arrayConstant());
        break;
      case 33:
      case 34: {
        let argc;
        let iftab;
        if (base === 33) {
          iftab = dv.getUint16(i, true);
          i += 2;
          const n = FIXED_ARITY[iftab];
          if (n === void 0) throw new Undecodable(`function ${iftab}`);
          argc = n;
        } else {
          argc = rgce[i] & 127;
          iftab = dv.getUint16(i + 1, true) & 32767;
          i += 3;
        }
        const args = [];
        for (let k = 0; k < argc; k++) args.unshift(pop());
        let name;
        if (iftab === 255) {
          name = args.shift()?.replace(/^_xlfn\./i, "");
          if (!name) throw new Undecodable("user function");
        } else name = FTAB[iftab] ?? (b12 ? FTAB_XLSB[iftab] : void 0);
        if (!name) throw new Undecodable(`function ${iftab}`);
        stack.push(`${name}(${args.join(",")})`);
        break;
      }
      case 35: {
        const n = b12 ? dv.getUint32(i, true) : dv.getUint16(i, true);
        i += 4;
        const name = ctx.nameOf(n);
        if (name === null) throw new Undecodable("name");
        stack.push(name);
        break;
      }
      case 36:
      case 44: {
        const row = b12 ? dv.getUint32(i, true) : dv.getUint16(i, true);
        const col2 = dv.getUint16(i + (b12 ? 4 : 2), true);
        i += b12 ? 6 : 4;
        stack.push(cell(row, col2, base === 44).text);
        break;
      }
      case 37:
      case 45: {
        const [r0, r1, c0, c1] = b12 ? [dv.getUint32(i, true), dv.getUint32(i + 4, true), dv.getUint16(i + 8, true), dv.getUint16(i + 10, true)] : [dv.getUint16(i, true), dv.getUint16(i + 2, true), dv.getUint16(i + 4, true), dv.getUint16(i + 6, true)];
        i += b12 ? 12 : 8;
        stack.push(area(r0, r1, c0, c1, base === 45));
        break;
      }
      case 38:
        if (ctx.extra && extraPos < ctx.extra.length) {
          const xv = new DataView(ctx.extra.buffer, ctx.extra.byteOffset, ctx.extra.byteLength);
          extraPos += b12 ? 4 + 16 * xv.getUint32(extraPos, true) : 2 + 8 * xv.getUint16(extraPos, true);
        }
        i += 6;
        break;
      case 39:
      case 40:
        i += 6;
        break;
      case 41:
        i += 2;
        break;
      case 42:
        i += b12 ? 6 : 4;
        stack.push("#REF!");
        break;
      case 43:
        i += b12 ? 12 : 8;
        stack.push("#REF!");
        break;
      case 57: {
        const ixti = dv.getUint16(i, true);
        const n = b12 ? dv.getUint32(i + 2, true) : dv.getUint16(i + 2, true);
        i += 6;
        const name = ctx.nameXOf(ixti, n);
        if (name === null) throw new Undecodable("external name");
        stack.push(name);
        break;
      }
      case 58: {
        const ixti = dv.getUint16(i, true);
        const row = b12 ? dv.getUint32(i + 2, true) : dv.getUint16(i + 2, true);
        const col2 = dv.getUint16(i + (b12 ? 6 : 4), true);
        i += b12 ? 8 : 6;
        stack.push(sheet3d(ixti) + cell(row, col2, false).text);
        break;
      }
      case 59: {
        const ixti = dv.getUint16(i, true);
        const p = i + 2;
        const [r0, r1, c0, c1] = b12 ? [dv.getUint32(p, true), dv.getUint32(p + 4, true), dv.getUint16(p + 8, true), dv.getUint16(p + 10, true)] : [dv.getUint16(p, true), dv.getUint16(p + 2, true), dv.getUint16(p + 4, true), dv.getUint16(p + 6, true)];
        i += b12 ? 14 : 10;
        stack.push(sheet3d(ixti) + area(r0, r1, c0, c1, false));
        break;
      }
      case 60:
        i += b12 ? 8 : 6;
        stack.push("#REF!");
        break;
      case 61:
        i += b12 ? 14 : 10;
        stack.push("#REF!");
        break;
      default: {
        const op = BINARY2[ptg];
        if (op === void 0) throw new Undecodable(`token 0x${ptg.toString(16)}`);
        const b = pop();
        const a = pop();
        stack.push(`${a}${op}${b}`);
      }
    }
  }
  if (stack.length !== 1) throw new Undecodable("unbalanced");
  return `=${stack[0]}`;
}

// packages/spreadiff/src/core/xlsb.ts
var ERRORS2 = { 0: "#NULL!", 7: "#DIV/0!", 15: "#VALUE!", 23: "#REF!", 29: "#NAME?", 36: "#NUM!", 42: "#N/A", 43: "#GETTING_DATA" };
function eachRecord(part, fn) {
  let p = 0;
  while (p < part.length) {
    let type = part[p++];
    if (type & 128) type = type & 127 | (part[p++] & 127) << 7;
    let len = 0;
    for (let k = 0; k < 4; k++) {
      const b = part[p++];
      len |= (b & 127) << 7 * k;
      if (!(b & 128)) break;
    }
    const end = Math.min(p + len, part.length);
    try {
      fn(type, new Reader(part.subarray(p, end)));
    } catch (e) {
      if (!(e instanceof RangeError)) throw e;
    }
    p = end;
  }
}
var Reader = class {
  constructor(data) {
    this.data = data;
    this.dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  }
  pos = 0;
  dv;
  get left() {
    return this.data.length - this.pos;
  }
  u8() {
    const v = this.dv.getUint8(this.pos);
    this.pos += 1;
    return v;
  }
  u16() {
    const v = this.dv.getUint16(this.pos, true);
    this.pos += 2;
    return v;
  }
  u32() {
    const v = this.dv.getUint32(this.pos, true);
    this.pos += 4;
    return v;
  }
  i32() {
    const v = this.dv.getInt32(this.pos, true);
    this.pos += 4;
    return v;
  }
  f64() {
    const v = this.dv.getFloat64(this.pos, true);
    this.pos += 8;
    return v;
  }
  bytes(n) {
    if (n > this.left) throw new RangeError("past the record");
    const b = this.data.subarray(this.pos, this.pos + n);
    this.pos += n;
    return b;
  }
  /** XLWideString: a 32-bit count of UTF-16 characters (0xFFFFFFFF: none). */
  str() {
    const n = this.u32();
    if (n === 4294967295) return "";
    if (n * 2 > this.left) throw new RangeError("past the record");
    let s = "";
    for (let k = 0; k < n; k++) s += String.fromCharCode(this.dv.getUint16(this.pos + k * 2, true));
    this.pos += n * 2;
    return s;
  }
  /** A formula: its tokens, then the extra data (array constants) they use. */
  formula() {
    const rgce = this.bytes(this.u32());
    const extra = this.left >= 4 ? this.bytes(Math.min(this.u32(), this.left)) : new Uint8Array(0);
    return { rgce, extra };
  }
  /** RfX: first and last row, first and last column. */
  range() {
    const r0 = this.u32();
    const r1 = this.u32();
    const c0 = this.u32();
    const c1 = this.u32();
    return { r0, r1, c0, c1 };
  }
};
function rk(v) {
  const n = v & 2 ? v >> 2 : new DataView(new Uint32Array([0, v & 4294967292]).buffer).getFloat64(0, true);
  return v & 1 ? n / 100 : n;
}
function relsOf(files, part) {
  const out = /* @__PURE__ */ new Map();
  const rels = files[part.replace(/([^/]+)$/, "_rels/$1.rels")];
  if (!rels) return out;
  eachElement(strFromU8(rels), "Relationship", (a) => {
    const id = attr(a, "Id");
    const target = attr(a, "Target");
    if (id && target && attr(a, "TargetMode") !== "External") out.set(id, target.startsWith("/") ? target.slice(1) : resolvePath(part, target));
  });
  return out;
}
function parseXlsb(data, options = {}) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const all = unzipSync(bytes, { filter: (f) => f.name.startsWith("xl/") && !/\.(?:png|jpe?g|gif|emf|wmf|vml)$/i.test(f.name) && !f.name.includes("vbaProject") });
  const files = {};
  for (const [k, v] of Object.entries(all)) files[k.toLowerCase()] = v;
  const get = (p) => files[p.toLowerCase()];
  const wbPart = get("xl/workbook.bin");
  if (!wbPart) throw new Error("Not an Excel binary workbook (no xl/workbook.bin)");
  const wbRels = relsOf(files, "xl/workbook.bin");
  let date19042 = false;
  const sheets = [];
  const nameRecs = [];
  const supbooks = [];
  let externalBooks = 0;
  const xti = [];
  eachRecord(wbPart, (type, r) => {
    switch (type) {
      case 153:
        date19042 = (r.u32() & 1) === 1;
        break;
      case 156: {
        r.u32();
        r.u32();
        const rid = r.str();
        sheets.push({ name: r.str(), path: wbRels.get(rid) });
        break;
      }
      case 39: {
        const flags = r.u32();
        r.u8();
        const itab = r.u32();
        const name = r.str();
        const { rgce, extra } = r.formula();
        nameRecs.push({ name, sheet: itab === 4294967295 ? -1 : itab, rgce, extra, skip: (flags & 46) !== 0 || /^_xl(nm|fn|ws)\./i.test(name) });
        break;
      }
      case 357:
      // BrtSupSelf
      case 358:
      // BrtSupSame
      case 667:
        supbooks.push({ self: true, sheets: [], names: [] });
        break;
      case 355:
        supbooks.push({ self: false, book: ++externalBooks, sheets: [], names: [] });
        break;
      case 359: {
        const book = supbooks[supbooks.length - 1];
        const n = r.u32();
        for (let k = 0; k < n && r.left >= 4; k++) book?.sheets.push(r.str());
        break;
      }
      case 361:
        supbooks[supbooks.length - 1]?.names.push(r.str());
        break;
      case 362: {
        const n = r.u32();
        for (let k = 0; k < n && r.left >= 12; k++) xti.push({ book: r.u32(), first: r.i32(), last: r.i32() });
        break;
      }
    }
  });
  const allNames = sheets.map((s) => s.name);
  const sst = [];
  const sstPart = get("xl/sharedStrings.bin");
  if (sstPart)
    eachRecord(sstPart, (type, r) => {
      if (type !== 19) return;
      r.u8();
      sst.push(r.str());
    });
  const formats = /* @__PURE__ */ new Map();
  const xfFormat = [];
  const stylesPart = get("xl/styles.bin");
  if (stylesPart) {
    let inCellXfs = false;
    eachRecord(stylesPart, (type, r) => {
      if (type === 617) inCellXfs = true;
      else if (type === 618) inCellXfs = false;
      else if (type === 44) {
        const id = r.u16();
        formats.set(id, r.str());
      } else if (type === 47 && inCellXfs) {
        r.u16();
        xfFormat.push(r.u16());
      }
    });
  }
  const codeOf = (xf) => {
    const id = xfFormat[xf] ?? 0;
    return formats.get(id) ?? BUILTIN_FORMATS[id];
  };
  const baseCtx = (row, col2, extra) => ({
    flavor: "biff12",
    row,
    col: col2,
    extra,
    sheetOf: (i) => {
      const x = xti[i];
      const book = x && supbooks[x.book];
      if (!book || x.first < 0) return null;
      const list = book.self ? allNames : book.sheets;
      const first = list[x.first];
      if (first === void 0) return null;
      return { first, last: list[x.last] ?? first, ...book.self ? {} : { book: book.book } };
    },
    nameOf: (i) => nameRecs[i - 1]?.name ?? null,
    nameXOf: (i, n) => {
      const book = supbooks[xti[i]?.book ?? -1];
      const name = book?.names[n - 1];
      if (name === void 0) return null;
      return book.self ? name : `[${book.book}]!${name}`;
    }
  });
  const names = {};
  const localNames = /* @__PURE__ */ new Map();
  for (const n of nameRecs) {
    if (n.skip || !n.rgce.length) continue;
    const f = decodeFormula(n.rgce, baseCtx(0, 0, n.extra));
    if (!f) continue;
    if (n.sheet >= 0) {
      const sheetName = allNames[n.sheet];
      if (!sheetName) continue;
      let m = localNames.get(sheetName);
      if (!m) localNames.set(sheetName, m = {});
      m[n.name] = f.slice(1);
    } else names[n.name] = f.slice(1);
  }
  const raws = [];
  for (const s of sheets) {
    const part = s.path && /worksheets\//i.test(s.path) ? get(s.path) : void 0;
    if (!part) continue;
    const grid = [];
    const formulaGrid = [];
    const codes = [""];
    const codeIds = /* @__PURE__ */ new Map();
    const fmtGrid = [];
    const merges = [];
    const arrays = [];
    const hidden = [];
    const filters = [];
    const shared = [];
    const arrayFormulas = [];
    const pending = [];
    let row = 0;
    let filter = null;
    const put = (c, value, xf) => {
      const code = codeOf(xf);
      if (typeof value === "number" && code && isDateFormat(code)) value = serialToIso(value, date19042) ?? value;
      (grid[row] ??= [])[c] = value;
      if (code && !/^general$/i.test(code) && value !== null && typeof value !== "boolean") {
        let id = codeIds.get(code);
        if (id === void 0) {
          id = codes.length;
          codes.push(code);
          codeIds.set(code, id);
        }
        (fmtGrid[row] ??= [])[c] = id;
      }
    };
    eachRecord(part, (type, r) => {
      if (type === 0) {
        row = r.u32();
        if (r.left >= 8 && r.data[11] & 16) hidden.push(row);
        return;
      }
      if (type >= 2 && type <= 11 || type === 62) {
        const c = r.u32();
        const xf = r.u16() | r.u8() << 16;
        r.u8();
        let value;
        switch (type) {
          case 2:
            value = rk(r.u32());
            break;
          case 3:
          case 11:
            value = ERRORS2[r.u8()] ?? "#N/A";
            break;
          case 4:
          case 10:
            value = r.u8() === 1;
            break;
          case 5:
          case 9:
            value = r.f64();
            break;
          case 7:
            value = sst[r.u32()] ?? "";
            break;
          case 62:
            r.u8();
            value = r.str();
            break;
          default:
            value = r.str();
        }
        put(c, value, xf);
        if (type >= 8) {
          if (type === 62) return;
          r.u16();
          pending.push({ r: row, c, ...r.formula() });
        }
        return;
      }
      switch (type) {
        case 426: {
          const range = r.range();
          r.u8();
          arrayFormulas.push({ ...range, ...r.formula() });
          break;
        }
        case 427:
          shared.push({ ...r.range(), ...r.formula() });
          break;
        case 176:
          merges.push(r.range());
          break;
        case 161: {
          const { r0, r1 } = r.range();
          filter = { r0: r0 + 1, r1 };
          break;
        }
        case 163:
          if (filter && !filters.includes(filter)) filters.push(filter);
          break;
      }
    });
    for (const p of pending) {
      let f;
      if (p.rgce[0] === 1 && p.rgce.length === 5) {
        const inside = (b) => p.r >= b.r0 && p.r <= b.r1 && p.c >= b.c0 && p.c <= b.c1;
        const arr = arrayFormulas.find(inside);
        const sh = arr ? void 0 : shared.find(inside);
        if (arr) {
          f = decodeFormula(arr.rgce, baseCtx(arr.r0, arr.c0, arr.extra));
          if (!arrays.some((a) => a.r0 === arr.r0 && a.c0 === arr.c0)) arrays.push({ r0: arr.r0, c0: arr.c0, r1: arr.r1, c1: arr.c1 });
        } else f = sh ? decodeFormula(sh.rgce, baseCtx(p.r, p.c, sh.extra)) : null;
      } else f = decodeFormula(p.rgce, baseCtx(p.r, p.c, p.extra));
      if (f) (formulaGrid[p.r] ??= [])[p.c] = f;
    }
    raws.push(rawSheet(s.name, grid, formulaGrid, { merges, arrays, hidden, filters, formats: codes.length > 1 ? { codes, grid: fmtGrid, date1904: date19042 } : void 0 }));
  }
  const tables = tablesFromRaw(raws, options.header ?? "auto", names, [], localNames);
  return {
    name: options.name ?? "workbook.xlsb",
    sheets: tables,
    ...date19042 ? { date1904: date19042 } : {},
    ...Object.keys(names).length ? { names } : {}
  };
}

// packages/spreadiff/src/core/cfb.ts
var FREE = 4294967295;
var END = 4294967294;
function readCfb(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (dv.getUint32(0, true) !== 3759263696 || dv.getUint32(4, true) !== 3776623009) throw new Error("Not a compound file");
  const sectorSize = 1 << dv.getUint16(30, true);
  const miniSize = 1 << dv.getUint16(32, true);
  const dirStart = dv.getUint32(48, true);
  const miniCutoff = dv.getUint32(56, true);
  const miniFatStart = dv.getUint32(60, true);
  let difatStart = dv.getUint32(68, true);
  const difatCount = dv.getUint32(72, true);
  const sector = (i) => {
    const off = (i + 1) * sectorSize;
    return bytes.subarray(off, Math.min(off + sectorSize, bytes.length));
  };
  const fatSectors = [];
  for (let i = 0; i < 109; i++) {
    const s = dv.getUint32(76 + i * 4, true);
    if (s !== FREE && s !== END) fatSectors.push(s);
  }
  for (let n = 0; n < difatCount && difatStart !== END && difatStart !== FREE; n++) {
    const s = sector(difatStart);
    const sdv = new DataView(s.buffer, s.byteOffset, s.byteLength);
    const per = sectorSize / 4 - 1;
    for (let i = 0; i < per; i++) {
      const x = sdv.getUint32(i * 4, true);
      if (x !== FREE && x !== END) fatSectors.push(x);
    }
    difatStart = sdv.getUint32(per * 4, true);
  }
  const fat = new Uint32Array(fatSectors.length * (sectorSize / 4));
  fatSectors.forEach((s, k) => {
    const sec = sector(s);
    fat.set(new Uint32Array(sec.buffer.slice(sec.byteOffset, sec.byteOffset + sectorSize)), k * (sectorSize / 4));
  });
  const chain = (start, table, limit) => {
    const out = [];
    for (let s = start; s !== END && s !== FREE && s < table.length && out.length <= limit; s = table[s]) out.push(s);
    return out;
  };
  const maxSectors = Math.ceil(bytes.length / sectorSize) + 1;
  const read = (start, size) => {
    const out = new Uint8Array(size);
    let pos = 0;
    for (const s of chain(start, fat, maxSectors)) {
      if (pos >= size) break;
      const sec = sector(s);
      out.set(sec.subarray(0, Math.min(sec.length, size - pos)), pos);
      pos += sec.length;
    }
    return out;
  };
  const dirSectors = chain(dirStart, fat, maxSectors);
  const dir = read(dirStart, dirSectors.length * sectorSize);
  const ddv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const entries = [];
  for (let off = 0; off + 128 <= dir.length; off += 128) {
    const nameLen = ddv.getUint16(off + 64, true);
    let name = "";
    for (let i = 0; i < Math.max(0, nameLen / 2 - 1) && i < 32; i++) name += String.fromCharCode(ddv.getUint16(off + i * 2, true));
    entries.push({
      name,
      type: dir[off + 66],
      left: ddv.getUint32(off + 68, true),
      right: ddv.getUint32(off + 72, true),
      child: ddv.getUint32(off + 76, true),
      start: ddv.getUint32(off + 116, true),
      size: ddv.getUint32(off + 120, true)
    });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new Error("Compound file without a root entry");
  const miniStream = read(root.start, root.size);
  const miniFatSectors = chain(miniFatStart, fat, maxSectors);
  const miniFatBytes = read(miniFatStart, miniFatSectors.length * sectorSize);
  const miniFat = new Uint32Array(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength / 4);
  const readMini = (start, size) => {
    const out = new Uint8Array(size);
    let pos = 0;
    for (const s of chain(start, miniFat, Math.ceil(miniStream.length / miniSize) + 1)) {
      if (pos >= size) break;
      const off = s * miniSize;
      out.set(miniStream.subarray(off, Math.min(off + miniSize, off + size - pos)), pos);
      pos += miniSize;
    }
    return out;
  };
  const streams = /* @__PURE__ */ new Map();
  const seen = /* @__PURE__ */ new Set();
  const walk = (i) => {
    if (i === FREE || i >= entries.length || seen.has(i)) return;
    seen.add(i);
    const e = entries[i];
    walk(e.left);
    walk(e.right);
    if (e.type === 2) streams.set(e.name, e.size < miniCutoff ? readMini(e.start, e.size) : read(e.start, e.size));
  };
  walk(root.child);
  return streams;
}

// packages/spreadiff/src/core/xls.ts
var ERRORS3 = {
  0: "#NULL!",
  7: "#DIV/0!",
  15: "#VALUE!",
  23: "#REF!",
  29: "#NAME?",
  36: "#NUM!",
  42: "#N/A"
};
function records(stream, start, stopAtEof) {
  const dv = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const out = [];
  let p = start;
  let depth = 0;
  while (p + 4 <= stream.length) {
    const type = dv.getUint16(p, true);
    const len = dv.getUint16(p + 2, true);
    const data = stream.subarray(p + 4, p + 4 + len);
    p += 4 + len;
    if (type === 2057 && p - 4 - len > start) depth++;
    if (depth > 0) {
      if (type === 10) depth--;
      continue;
    }
    if (type === 60 && out.length) out[out.length - 1].more.push(data);
    else out.push({ type, data, more: [] });
    if (stopAtEof && type === 10) break;
  }
  return out;
}
var Reader2 = class {
  seg = 0;
  pos = 0;
  parts;
  constructor(rec) {
    this.parts = [rec.data, ...rec.more];
  }
  get cur() {
    return this.parts[this.seg];
  }
  next() {
    if (this.pos >= this.cur.length && this.seg < this.parts.length - 1) {
      this.seg++;
      this.pos = 0;
    }
  }
  /** True once every byte of the record and its continuations is read. */
  get eof() {
    return this.seg >= this.parts.length - 1 && this.pos >= this.cur.length;
  }
  u8() {
    this.next();
    return this.cur[this.pos++] ?? 0;
  }
  u16() {
    return this.u8() | this.u8() << 8;
  }
  u32() {
    return (this.u16() | this.u16() << 16) >>> 0;
  }
  /** Skip n bytes, across continuations; a length past the end (a corrupt file) stops at the end. */
  skip(n) {
    while (n > 0 && !this.eof) {
      this.next();
      const step = Math.min(n, this.cur.length - this.pos);
      this.pos += step;
      n -= step;
      if (step === 0) break;
    }
  }
  /** XLUnicodeRichExtendedString: count, flags, (runs), (extension), characters, (run data), (extension data). */
  string(cch = this.u16()) {
    let flags = this.u8();
    const runs = flags & 8 ? this.u16() : 0;
    const ext = flags & 4 ? this.u32() : 0;
    let s = "";
    for (let k = 0; k < cch; k++) {
      if (this.pos >= this.cur.length && this.seg < this.parts.length - 1) {
        this.seg++;
        this.pos = 0;
        flags = flags & ~1 | this.cur[this.pos++] & 1;
      }
      s += String.fromCharCode(flags & 1 ? this.u16() : this.u8());
    }
    this.skip(runs * 4 + ext);
    return s;
  }
};
function str(data, pos, cch) {
  const high = data[pos] & 1;
  let p = pos + 1;
  let s = "";
  for (let k = 0; k < cch; k++) {
    s += String.fromCharCode(high ? data[p] | data[p + 1] << 8 : data[p]);
    p += high ? 2 : 1;
  }
  return [s, p];
}
function rk2(v) {
  const n = v & 2 ? v >> 2 : new DataView(new Uint32Array([0, v & 4294967292]).buffer).getFloat64(0, true);
  return v & 1 ? n / 100 : n;
}
function parseXls(data, options = {}) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const streams = readCfb(bytes);
  if (streams.has("EncryptedPackage") || streams.has("EncryptionInfo")) throw new Error("This workbook is password-protected. Save a copy without a password to compare it");
  const wbs = streams.get("Workbook") ?? streams.get("Book");
  if (!wbs) throw new Error("Not an Excel workbook (no Workbook stream)");
  if (!streams.has("Workbook")) throw new Error("Excel 5 and older (.xls before Excel 97) isn't supported. Save as .xlsx");
  const dv = new DataView(wbs.buffer, wbs.byteOffset, wbs.byteLength);
  const globals = records(wbs, 0, true);
  if (globals[0]?.type !== 2057 || dv.getUint16(4, true) !== 1536) throw new Error("Only Excel 97-2003 (BIFF8) .xls files are supported. Save as .xlsx");
  let date19042 = false;
  const formats = /* @__PURE__ */ new Map();
  const xfFormat = [];
  let sst = [];
  const sheets = [];
  const supbooks = [];
  let externalBooks = 0;
  const xti = [];
  const nameRecs = [];
  for (const rec of globals) {
    const d = rec.data;
    const v = new DataView(d.buffer, d.byteOffset, d.byteLength);
    try {
      switch (rec.type) {
        case 34:
          date19042 = v.getUint16(0, true) === 1;
          break;
        case 1054: {
          const id = v.getUint16(0, true);
          formats.set(
            id,
            new Reader2({
              type: 0,
              data: d.subarray(2),
              more: rec.more
            }).string()
          );
          break;
        }
        case 224:
          xfFormat.push(v.getUint16(2, true));
          break;
        case 252: {
          const r = new Reader2(rec);
          r.u32();
          const unique = r.u32();
          sst = [];
          for (let k = 0; k < unique && !r.eof; k++) sst.push(r.string());
          break;
        }
        case 133: {
          const [name] = str(d, 7, d[6]);
          sheets.push({ name, pos: v.getUint32(0, true), type: d[5] });
          break;
        }
        case 430:
          {
            const ctab = v.getUint16(0, true);
            const cch = v.getUint16(2, true);
            if (cch === 1025 || cch === 14849) supbooks.push({ self: true, sheets: [], names: [] });
            else {
              const r = new Reader2({ type: 0, data: d.subarray(4), more: rec.more });
              r.string(cch);
              const sheetNames = [];
              for (let k = 0; k < ctab && !r.eof; k++) sheetNames.push(r.string());
              supbooks.push({ self: false, book: ++externalBooks, sheets: sheetNames, names: [] });
            }
          }
          break;
        case 35: {
          const [name] = str(d, 7, d[6]);
          supbooks[supbooks.length - 1]?.names.push(name);
          break;
        }
        case 23: {
          const n = v.getUint16(0, true);
          for (let k = 0; k < n; k++)
            xti.push({
              book: v.getUint16(2 + k * 6, true),
              first: v.getInt16(4 + k * 6, true),
              last: v.getInt16(6 + k * 6, true)
            });
          break;
        }
        case 24: {
          const grbit = v.getUint16(0, true);
          const cch = d[3];
          const cce = v.getUint16(4, true);
          const itab = v.getUint16(8, true);
          const builtin = (grbit & 32) !== 0;
          const [name, p] = builtin ? [`_xlnm.${d[15]}`, 17] : str(d, 14, cch);
          nameRecs.push({
            name,
            sheet: itab,
            rgce: d.subarray(p, p + cce),
            builtin
          });
          break;
        }
      }
    } catch (e) {
      if (!(e instanceof RangeError)) throw e;
    }
  }
  const workSheets = sheets.filter((s) => s.type === 0);
  const allNames = sheets.map((s) => s.name);
  const codeOf = (xf) => {
    const id = xfFormat[xf] ?? 0;
    return formats.get(id) ?? BUILTIN_FORMATS[id];
  };
  const baseCtx = (row, col2, extra) => ({
    flavor: "biff8",
    row,
    col: col2,
    extra,
    sheetOf: (i) => {
      const x = xti[i];
      const book = x && supbooks[x.book];
      if (!book || x.first < 0) return null;
      const list = book.self ? allNames : book.sheets;
      const first = list[x.first];
      if (first === void 0) return null;
      return { first, last: list[x.last] ?? first, ...book.self ? {} : { book: book.book } };
    },
    nameOf: (i) => nameRecs[i - 1]?.name ?? null,
    nameXOf: (i, n) => {
      const book = supbooks[xti[i]?.book ?? -1];
      const name = book?.names[n - 1];
      if (name === void 0) return null;
      return book.self ? name : `[${book.book}]!${name}`;
    }
  });
  const names = {};
  const localNames = /* @__PURE__ */ new Map();
  for (const n of nameRecs) {
    if (n.builtin || !n.rgce.length) continue;
    const f = decodeFormula(n.rgce, baseCtx(0, 0));
    if (!f) continue;
    if (n.sheet) {
      const sheetName = allNames[n.sheet - 1];
      if (!sheetName) continue;
      let m = localNames.get(sheetName);
      if (!m) localNames.set(sheetName, m = {});
      m[n.name] = f.slice(1);
    } else names[n.name] = f.slice(1);
  }
  const raws = [];
  for (const s of workSheets) {
    const recs = records(wbs, s.pos, true);
    const grid = [];
    const formulaGrid = [];
    const codes = [""];
    const codeIds = /* @__PURE__ */ new Map();
    const fmtGrid = [];
    const merges = [];
    const arrays = [];
    const hidden = [];
    const shared = [];
    const dataTables = [];
    const arrayFormulas = [];
    const pending = [];
    let stringFor = null;
    const put = (r, c, value, xf) => {
      const code = codeOf(xf);
      if (typeof value === "number" && code && isDateFormat(code)) value = serialToIso(value, date19042) ?? value;
      (grid[r] ??= [])[c] = value;
      if (code && !/^general$/i.test(code) && value !== null && typeof value !== "boolean") {
        let id = codeIds.get(code);
        if (id === void 0) {
          id = codes.length;
          codes.push(code);
          codeIds.set(code, id);
        }
        (fmtGrid[r] ??= [])[c] = id;
      }
    };
    for (const rec of recs) {
      const d = rec.data;
      const v = new DataView(d.buffer, d.byteOffset, d.byteLength);
      try {
        switch (rec.type) {
          case 520:
            if (v.getUint16(12, true) & 32) hidden.push(v.getUint16(0, true));
            break;
          case 515:
            put(v.getUint16(0, true), v.getUint16(2, true), v.getFloat64(6, true), v.getUint16(4, true));
            break;
          case 638:
            put(v.getUint16(0, true), v.getUint16(2, true), rk2(v.getUint32(6, true)), v.getUint16(4, true));
            break;
          case 189: {
            const r = v.getUint16(0, true);
            const c0 = v.getUint16(2, true);
            const n = (d.length - 6) / 6;
            for (let k = 0; k < n; k++) put(r, c0 + k, rk2(v.getUint32(6 + k * 6, true)), v.getUint16(4 + k * 6, true));
            break;
          }
          case 253:
            put(v.getUint16(0, true), v.getUint16(2, true), sst[v.getUint32(6, true)] ?? "", v.getUint16(4, true));
            break;
          case 516: {
            const r = new Reader2({
              type: 0,
              data: d.subarray(6),
              more: rec.more
            });
            put(v.getUint16(0, true), v.getUint16(2, true), r.string(), v.getUint16(4, true));
            break;
          }
          case 517: {
            const isErr = d[7] === 1;
            put(v.getUint16(0, true), v.getUint16(2, true), isErr ? ERRORS3[d[6]] ?? "#N/A" : d[6] === 1, v.getUint16(4, true));
            break;
          }
          case 6: {
            const r = v.getUint16(0, true);
            const c = v.getUint16(2, true);
            const xf = v.getUint16(4, true);
            let value;
            if (v.getUint16(12, true) === 65535) {
              const kind = d[6];
              if (kind === 0) {
                value = "";
                stringFor = { r, c };
              } else if (kind === 1) value = d[8] === 1;
              else if (kind === 2) value = ERRORS3[d[8]] ?? "#N/A";
              else value = "";
            } else value = v.getFloat64(6, true);
            put(r, c, value, xf);
            const cce = v.getUint16(20, true);
            const rgce = d.subarray(22, 22 + cce);
            pending.push({ r, c, rgce, extra: d.subarray(22 + cce) });
            break;
          }
          case 519:
            if (stringFor) {
              (grid[stringFor.r] ??= [])[stringFor.c] = new Reader2(rec).string();
              stringFor = null;
            }
            break;
          case 1212: {
            const cce = v.getUint16(8, true);
            shared.push({
              r0: v.getUint16(0, true),
              r1: v.getUint16(2, true),
              c0: d[4],
              c1: d[5],
              rgce: d.subarray(10, 10 + cce),
              extra: d.subarray(10 + cce)
            });
            break;
          }
          case 545: {
            const cce = v.getUint16(12, true);
            arrayFormulas.push({
              r0: v.getUint16(0, true),
              r1: v.getUint16(2, true),
              c0: d[4],
              c1: d[5],
              rgce: d.subarray(14, 14 + cce),
              extra: d.subarray(14 + cce)
            });
            break;
          }
          case 566: {
            const flags = v.getUint16(6, true);
            const ref = (r, c) => `${columnLabel(c)}${r + 1}`;
            const inputs = [ref(v.getUint16(8, true), v.getUint16(10, true))];
            if (flags & 8) inputs.push(ref(v.getUint16(12, true), v.getUint16(14, true)));
            dataTables.push({ r0: v.getUint16(0, true), r1: v.getUint16(2, true), c0: d[4], c1: d[5], formula: `=TABLE(${inputs.join(",")})` });
            break;
          }
          case 229: {
            const n = v.getUint16(0, true);
            for (let k = 0; k < n; k++) {
              const o = 2 + k * 8;
              merges.push({
                r0: v.getUint16(o, true),
                r1: v.getUint16(o + 2, true),
                c0: v.getUint16(o + 4, true),
                c1: v.getUint16(o + 6, true)
              });
            }
            break;
          }
        }
      } catch (e) {
        if (!(e instanceof RangeError)) throw e;
      }
    }
    for (const p of pending) {
      let f;
      if (p.rgce[0] === 2 && p.rgce.length === 5) {
        f = dataTables.find((t) => p.r >= t.r0 && p.r <= t.r1 && p.c >= t.c0 && p.c <= t.c1)?.formula ?? null;
      } else if (p.rgce[0] === 1 && p.rgce.length === 5) {
        const ar = p.rgce[1] | p.rgce[2] << 8;
        const ac = p.rgce[3] | p.rgce[4] << 8;
        const arr = arrayFormulas.find((a) => a.r0 === ar && a.c0 === ac && p.r >= a.r0 && p.r <= a.r1 && p.c >= a.c0 && p.c <= a.c1);
        const sh = arr ? void 0 : shared.find((x) => p.r >= x.r0 && p.r <= x.r1 && p.c >= x.c0 && p.c <= x.c1);
        if (arr) {
          f = decodeFormula(arr.rgce, baseCtx(arr.r0, arr.c0, arr.extra));
          if (!arrays.some((a) => a.r0 === arr.r0 && a.c0 === arr.c0)) arrays.push({ r0: arr.r0, c0: arr.c0, r1: arr.r1, c1: arr.c1 });
        } else f = sh ? decodeFormula(sh.rgce, baseCtx(p.r, p.c, sh.extra)) : null;
      } else f = decodeFormula(p.rgce, baseCtx(p.r, p.c, p.extra));
      if (f) (formulaGrid[p.r] ??= [])[p.c] = f;
    }
    raws.push(
      rawSheet(s.name, grid, formulaGrid, {
        merges,
        arrays,
        hidden,
        formats: codes.length > 1 ? { codes, grid: fmtGrid, date1904: date19042 } : void 0
      })
    );
  }
  const tables = tablesFromRaw(raws, options.header ?? "auto", names, [], localNames);
  return {
    name: options.name ?? "workbook.xls",
    sheets: tables,
    ...date19042 ? { date1904: date19042 } : {},
    ...Object.keys(names).length ? { names } : {}
  };
}

// packages/spreadiff/src/core/ods.ts
var LO_ERRORS = { 503: "#NUM!", 519: "#VALUE!", 521: "#NULL!", 524: "#REF!", 525: "#NAME?", 532: "#DIV/0!", 32767: "#N/A" };
var MAX_ROWS = 1048576;
var MAX_COLS = 16384;
var count = (v, max2) => Math.max(1, Math.min(Number(v ?? 1) || 1, max2));
function openFormulaToExcel(f) {
  let s = f.replace(/^of:/, "");
  if (!s.startsWith("=")) s = "=" + s;
  let out = "";
  let braces = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') {
      const end = s.indexOf('"', i + 1);
      const j = end < 0 ? s.length - 1 : end;
      out += s.slice(i, j + 1);
      i = j;
    } else if (ch === "[") {
      const end = s.indexOf("]", i);
      out += refToExcel(s.slice(i + 1, end < 0 ? s.length : end));
      i = end < 0 ? s.length : end;
    } else if (ch === ";") out += ",";
    else if (ch === "|" && braces > 0) out += ";";
    else {
      if (ch === "{") braces++;
      else if (ch === "}") braces--;
      out += ch;
    }
  }
  return out.replace(/\bCOM\.MICROSOFT\.(?=[A-Z][A-Z0-9.]*\()/g, "");
}
function refToExcel(ref) {
  const quote = (sh) => sh && !/^'.*'$/.test(sh) && /[^A-Za-z0-9_]/.test(sh) ? `'${sh.replace(/'/g, "''")}'` : sh;
  const sheets = [];
  const cells = ref.split(":").map((p) => {
    const dot = p.lastIndexOf(".");
    const sh = (dot > 0 ? p.slice(0, dot) : "").replace(/^\$/, "");
    sheets.push(sh);
    return dot >= 0 ? p.slice(dot + 1) : p;
  });
  const [first, last] = [sheets[0], sheets[sheets.length - 1]];
  const prefix = first ? last && last !== first ? `${quote(first)}:${quote(last)}!` : `${quote(first)}!` : "";
  return prefix + cells.join(":");
}
function cellText(body) {
  const paras = [];
  const own = body.replace(/<office:annotation\b[\s\S]*?<\/office:annotation>/g, "").replace(/<draw:([\w-]+)(?=[\s>\/])[\s\S]*?<\/draw:\1>/g, "");
  eachElement(own, "text:p", (_, b) => {
    const t = (b ?? "").replace(/<text:s(?:\s+text:c="(\d+)")?\s*\/>/g, (_m, n) => " ".repeat(Number(n ?? 1))).replace(/<text:tab\s*\/>/g, "	").replace(/<text:line-break\s*\/>/g, "\n").replace(/<[^>]+>/g, "");
    paras.push(decode(t));
  });
  return paras.join("\n");
}
function odsDate(v) {
  const iso = v.replace("T", " ").replace(/(\.\d*?)0+$/, "$1").replace(/\.$/, "").replace(/ 00:00:00$/, "");
  if (iso >= "1900-03-01") return iso;
  const ms = Date.parse(v.length <= 10 ? `${v}T00:00:00Z` : `${v}Z`);
  if (!Number.isFinite(ms)) return iso;
  const serial = (ms - Date.UTC(1899, 11, 30)) / 864e5;
  return serialToIso(serial, false) ?? serial;
}
function isoDuration(v) {
  const m = /^-?P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:([\d.]+)S)?$/.exec(v);
  if (!m) return v;
  const h = Number(m[1] ?? 0) * 24 + Number(m[2] ?? 0);
  const sec = Number(m[4] ?? 0);
  const whole = Math.floor(sec + 1e-9);
  const ms = Math.round((sec - whole) * 1e3);
  return `${String(h).padStart(2, "0")}:${String(Number(m[3] ?? 0)).padStart(2, "0")}:${String(whole).padStart(2, "0")}${ms ? `.${String(ms).padStart(3, "0").replace(/0+$/, "")}` : ""}`;
}
function parseOds(data, options = {}) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const files = unzipSync(bytes, { filter: (f) => f.name === "content.xml" || f.name === "mimetype" });
  if (!files["content.xml"]) throw new Error("Not an OpenDocument spreadsheet (missing content.xml)");
  const xml = strFromU8(files["content.xml"]);
  const raws = [];
  const names = {};
  eachElement(xml, "table:named-range", (a) => {
    const name = attr(a, "table:name");
    const range = attr(a, "table:cell-range-address");
    if (name && range) names[decode(name)] = openFormulaToExcel(`[${decode(range)}]`).slice(1);
  });
  eachElement(xml, "table:table", (tableAttrs, body) => {
    const name = decode(attr(tableAttrs, "table:name") ?? `Sheet${raws.length + 1}`);
    const grid = [];
    const formulas = [];
    const merges = [];
    const hidden = [];
    const filters = [];
    const arrays = [];
    let r = 0;
    const rowRe = /<table:table-row\b([^>]*?)(\/>|>([\s\S]*?)<\/table:table-row>)/g;
    for (const m of (body ?? "").matchAll(rowRe)) {
      const rowAttrs = m[1];
      const rowBody = m[3] ?? "";
      const repeat = count(attr(rowAttrs, "table:number-rows-repeated"), MAX_ROWS);
      const visibility = attr(rowAttrs, "table:visibility") ?? "";
      const isHidden = visibility === "collapse" || visibility === "filter";
      const row = [];
      const frow = [];
      let c = 0;
      let any = false;
      const cellRe = /<table:(table-cell|covered-table-cell)\b([^>]*?)(\/>|>([\s\S]*?)<\/table:\1>)/g;
      for (const cm of rowBody.matchAll(cellRe)) {
        const a = cm[2];
        const cb = cm[4] ?? "";
        const reps = count(attr(a, "table:number-columns-repeated"), MAX_COLS);
        let value = null;
        const type = attr(a, "office:value-type") ?? attr(a, "calcext:value-type");
        if (cm[1] === "table-cell") {
          switch (type) {
            case "float":
            case "percentage":
            case "currency":
              value = Number(attr(a, "office:value"));
              break;
            case "date":
              value = odsDate(attr(a, "office:date-value") ?? "");
              break;
            case "time":
              value = isoDuration(attr(a, "office:time-value") ?? "");
              break;
            case "boolean":
              value = attr(a, "office:boolean-value") === "true";
              break;
            case "string":
              value = attr(a, "office:string-value") !== void 0 ? decode(attr(a, "office:string-value")) : cellText(cb);
              break;
            default:
              if (cb.includes("<text:p")) value = cellText(cb) || null;
          }
          if (attr(a, "calcext:value-type") === "error") {
            const t = cellText(cb);
            const code = /^Err:(\d+)$/.exec(t)?.[1];
            value = code ? LO_ERRORS[Number(code)] ?? "#VALUE!" : t;
          }
        }
        const f = attr(a, "table:formula");
        let formula = f ? openFormulaToExcel(decode(f)) : null;
        if (formula && /^=#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(formula) && (value === null || value === formula.slice(1))) {
          value = formula.slice(1);
          formula = null;
        }
        if (formula && /^=(TRUE|FALSE)\(\)$/i.test(formula) && typeof value === "boolean") formula = null;
        const mc = Number(attr(a, "table:number-matrix-columns-spanned") ?? 0);
        const mr = Number(attr(a, "table:number-matrix-rows-spanned") ?? 0);
        if (formula && (mc > 1 || mr > 1)) arrays.push({ r0: r, c0: c, r1: r + Math.max(1, mr) - 1, c1: c + Math.max(1, mc) - 1, formula });
        const cs = Number(attr(a, "table:number-columns-spanned") ?? 1);
        const rs = Number(attr(a, "table:number-rows-spanned") ?? 1);
        if (cs > 1 || rs > 1) merges.push({ r0: r, c0: c, r1: r + rs - 1, c1: c + cs - 1 });
        if (value === null && !formula) c += reps;
        else
          for (let k = 0; k < reps && c < MAX_COLS; k++) {
            row[c] = value;
            frow[c] = formula;
            any = true;
            c++;
          }
      }
      if (!any) r += repeat;
      else
        for (let k = 0; k < repeat && r < MAX_ROWS; k++) {
          if (isHidden) hidden.push(r);
          if (visibility === "filter") filters.push({ r0: r, r1: r });
          grid[r] = row.slice();
          formulas[r] = frow.slice();
          r++;
        }
    }
    for (const b of arrays) {
      for (let rr = b.r0; rr <= b.r1; rr++) {
        const fr = formulas[rr] ??= [];
        for (let cc = b.c0; cc <= b.c1; cc++) {
          fr[cc] ??= b.formula;
          (grid[rr] ??= [])[cc] ??= null;
        }
      }
    }
    raws.push(rawSheet(name, grid, formulas, { merges, hidden, filters, arrays: arrays.map(({ r0, c0, r1, c1 }) => ({ r0, c0, r1, c1 })) }));
  });
  const sheets = tablesFromRaw(raws, options.header ?? "auto", names);
  return { name: options.name ?? "workbook.ods", sheets, ...Object.keys(names).length ? { names } : {} };
}

// packages/spreadiff/src/core/file.ts
function parseFile(name, data, options = {}) {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  const isZip = bytes[0] === 80 && bytes[1] === 75 && bytes[2] === 3 && bytes[3] === 4;
  if (isZip) {
    const head = String.fromCharCode(...bytes.subarray(30, 110));
    if (/\.ods$/i.test(name) || head.includes("opendocument.spreadsheet")) return parseOds(bytes, { name, header: options.header });
    if (/\.xlsb$/i.test(name) || hasEntry(bytes, "xl/workbook.bin")) return parseXlsb(bytes, { name, header: options.header });
    return parseXlsx(bytes, { name, header: options.header });
  }
  if (bytes[0] === 208 && bytes[1] === 207 && bytes[2] === 17 && bytes[3] === 224) {
    try {
      return parseXls(bytes, { name, header: options.header });
    } catch (e) {
      throw new Error(`${name}: ${e.message}`);
    }
  }
  if (/\.(xlsx|xlsm|xlsb|xls|ods)$/i.test(name)) throw new Error(`${name}: not a valid ${name.slice(name.lastIndexOf("."))} file`);
  const text3 = decodeText(bytes);
  const delimiter = options.delimiter ?? (/\.tsv$/i.test(name) ? "	" : void 0);
  const sheet = parseCsv(text3, { name: name.replace(/\.[^.]+$/, ""), delimiter, header: options.header !== false });
  return { name, sheets: [sheet] };
}
function hasEntry(bytes, entry) {
  const tail = bytes.subarray(Math.max(0, bytes.length - 65536));
  const needle = Array.from(entry, (ch) => ch.charCodeAt(0));
  outer: for (let i = 0; i + needle.length <= tail.length; i++) {
    for (let k = 0; k < needle.length; k++) if ((tail[i + k] | 32) !== (needle[k] | 32)) continue outer;
    return true;
  }
  return false;
}
function decodeText(bytes) {
  if (bytes[0] === 255 && bytes[1] === 254) return new TextDecoder("utf-16le").decode(bytes);
  if (bytes[0] === 254 && bytes[1] === 255) return new TextDecoder("utf-16be").decode(bytes);
  return new TextDecoder("utf-8").decode(bytes);
}

// packages/spreadiff/src/core/labels.ts
var SAMPLE = 5e3;
var TOTAL = /^\s*(?:total|sub-?total|grand total)\b|\btotal\s*$/i;
var hasText = (v) => {
  for (let i = 0; i < v.length; i++) if (v.charCodeAt(i) > 32) return true;
  return false;
};
var isBlank = (v) => v === null || v === void 0 || v === "";
var text2 = (v) => v === null || v === void 0 ? "" : typeof v === "string" ? v.replace(/\s+/g, " ").trim() : String(v);
var ERROR3 = /^#(?:VALUE!|N\/A|REF!|DIV\/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!|GETTING_DATA)$/;
var MAX_HEADING_ROWS = 3;
var yearLike = (v) => typeof v === "number" && Number.isInteger(v) && v >= 1900 && v <= 2100;
var ISO_DATE2 = /^\d{4}-\d{2}-\d{2}(?:[ T]|$)/;
function tableLabeler(table, options = {}) {
  let analysed = null;
  return (row, col2) => {
    analysed ??= sheetLabels(table, options);
    return analysed.cell(row, col2);
  };
}
function sheetLabels(table, options = {}) {
  return analyse(table, options);
}
function analyse(table, options) {
  const { rows, columns } = table;
  const width = columns.length;
  const n = rows.length;
  const o = table.origin ?? { row: 1, col: 0 };
  const lettered = columns.every((c, i) => c === columnLabel(i + o.col));
  const header = table.header?.length ? table.header : !lettered && width ? columns : null;
  const hasHeaderRow = !!header;
  const maxCells = options.maxRowCells ?? 2;
  const keyIdx = (options.keyColumns ?? []).map((k) => columns.indexOf(k)).filter((i) => i >= 0);
  const sample = Math.min(n, SAMPLE);
  const numberRows = [];
  for (let r = 0; r < sample; r++) if (rows[r].some((v) => typeof v === "number")) numberRows.push(r);
  let zone = 0;
  for (; zone < width; zone++) {
    if (keyIdx.includes(zone)) continue;
    let filled = 0;
    let numeric = 0;
    for (const r of numberRows) {
      const v = rows[r][zone];
      if (isBlank(v)) continue;
      filled++;
      if (typeof v === "number" || table.formulas?.[r]?.[zone] && typeof v !== "string") numeric++;
    }
    if (filled && numeric * 2 >= filled) break;
  }
  if (zone === width) zone = Math.min(1, width);
  const styles = /* @__PURE__ */ new Map();
  for (const s of table.labelStyles ?? []) styles.set((s.r - o.row) * width + (s.c - o.col), s);
  const merges = (table.merges ?? []).filter((m) => m.r0 === m.r1).map((m) => ({ r: m.r0 - o.row, c0: m.c0 - o.col, c1: m.c1 - o.col }));
  const isLabel = (v) => typeof v === "string" ? hasText(v) && !(v.charCodeAt(0) === 35 && ERROR3.test(v)) : typeof v === "boolean";
  const firstLabel = new Int16Array(n).fill(-1);
  const dataCount = new Uint32Array(n);
  const band = new Uint8Array(n);
  const depth = new Uint16Array(n);
  const bold = new Uint8Array(n);
  for (let r = 0; r < n; r++) {
    const cells = rows[r];
    let first = -1;
    for (let c = 0; c < zone; c++) {
      if (isLabel(cells[c])) {
        first = c;
        break;
      }
    }
    let data = 0;
    let isBand = true;
    let words = 0;
    let years = 0;
    let lastYear = NaN;
    let dates = 0;
    let lastDate = "";
    for (let c = zone; c < width; c++) {
      const v = cells[c];
      if (isBlank(v)) continue;
      data++;
      if (!isBand) continue;
      if (typeof v === "string") {
        if (v.length >= 10 && ISO_DATE2.test(v)) {
          if (v <= lastDate) isBand = false;
          lastDate = v;
          dates++;
        } else if (v.charCodeAt(0) === 35 && ERROR3.test(v)) isBand = false;
        else words++;
      } else if (typeof v === "boolean") words++;
      else if (yearLike(v) && (Number.isNaN(lastYear) || v - lastYear === 0 || v - lastYear === 1)) {
        lastYear = v;
        years++;
      } else isBand = false;
    }
    const runs = years + dates;
    if (runs === 1 && (words > 0 || first >= 0 || dates === 1)) isBand = false;
    if (hasHeaderRow) isBand = false;
    firstLabel[r] = first;
    dataCount[r] = data;
    band[r] = data > 0 && isBand ? 1 : 0;
    if (first >= 0) {
      const raw = cells[first];
      const lead = typeof raw === "string" ? raw.length - raw.trimStart().length : 0;
      const st = styles.size ? styles.get(r * width + first) : void 0;
      depth[r] = first * 64 + Math.min(63, (st?.indent ?? 0) + Math.floor(lead / 2));
      bold[r] = st?.bold ? 1 : 0;
    }
  }
  const labelOf = (r) => text2(rows[r][firstLabel[r]]);
  const hasLabel = (r) => firstLabel[r] >= 0;
  const isHeading = (r) => hasLabel(r) && dataCount[r] === 0 && !band[r];
  const nodeLabel = [];
  const nodeParent = [];
  const sectionAt = new Int32Array(n).fill(-1);
  const noSections = keyIdx.length > 0 || zone === 0;
  if (!noSections) {
    const group = new Uint8Array(n);
    let next = -1;
    for (let r = n - 1; r >= 0; r--) {
      if (isHeading(r)) group[r] = next >= 0 && isHeading(next) ? 1 : 0;
      if (hasLabel(r) || dataCount[r]) next = r;
    }
    let firstDepth = -1;
    let byDepth = false;
    for (let r = 0; r < n && !byDepth; r++) {
      if (!hasLabel(r)) continue;
      if (firstDepth < 0) firstDepth = depth[r];
      else if (depth[r] !== firstDepth) byDepth = true;
    }
    const stack = [];
    const top = () => stack[stack.length - 1];
    const innermost = () => stack.length ? top().node : -1;
    let prev = null;
    for (let r = 0; r < n; r++) {
      sectionAt[r] = innermost();
      if (!hasLabel(r) || band[r]) continue;
      const d = depth[r];
      while (stack.length && top().depth > d) stack.pop();
      if (isHeading(r)) {
        const label = labelOf(r);
        const isGroup = !!group[r] || !!bold[r];
        if (byDepth) {
          while (stack.length && top().depth >= d) stack.pop();
        } else if (prev !== "heading") {
          while (stack.length && top().depth === d && !top().group) stack.pop();
          if (isGroup && stack.length && top().depth === d) stack.pop();
        }
        sectionAt[r] = innermost();
        nodeLabel.push(label);
        nodeParent.push(innermost());
        stack.push({ node: nodeLabel.length - 1, depth: d, group: isGroup, label });
        prev = "heading";
        continue;
      }
      if (byDepth) while (stack.length && top().depth >= d) stack.pop();
      sectionAt[r] = innermost();
      const raw = rows[r][firstLabel[r]];
      if (typeof raw === "string" && TOTAL.test(raw)) {
        const rest = text2(raw).replace(TOTAL, "").trim().toLowerCase();
        let named = -1;
        if (rest) {
          for (let i = stack.length - 1; i >= 0 && named < 0; i--) if (stack[i].label.toLowerCase() === rest) named = i;
        }
        stack.length = named >= 0 ? named : Math.max(0, stack.length - 1);
      }
      prev = "data";
    }
  }
  const sectionsOf = (r) => {
    const out = [];
    for (let k = sectionAt[r] ?? -1; k >= 0; k = nodeParent[k]) out.push(nodeLabel[k]);
    return out.reverse();
  };
  const blocks = [header ? [-1] : []];
  const blockAt = new Int32Array(n);
  let current = 0;
  let open = null;
  for (let r = 0; r < n; r++) {
    blockAt[r] = current;
    if (band[r]) {
      if (!open) open = [];
      open.push(r);
    } else if (dataCount[r] > 0 || hasLabel(r)) {
      if (open) {
        blocks.push(open);
        current = blocks.length - 1;
      }
      open = null;
      blockAt[r] = current;
    }
  }
  const valueAt = (r, c) => r < 0 ? header[c] ?? null : rows[r][c];
  const heading = (r, c, spanRight) => {
    const own = valueAt(r, c);
    if (!isBlank(own)) return text2(own);
    const m = merges.find((g) => g.r === r && c >= g.c0 && c <= g.c1);
    if (m) return text2(valueAt(r, m.c0));
    if (!spanRight || c < zone) return "";
    for (let x = c - 1; x >= zone; x--) if (!isBlank(valueAt(r, x))) return text2(valueAt(r, x));
    return "";
  };
  const rowLabel = (row, col2 = -1) => {
    const own = [];
    if (row < 0 || row >= n) return own;
    const cells = rows[row];
    for (const k of keyIdx) if (typeof cells[k] === "string" && hasText(cells[k])) own.push(text2(cells[k]));
    for (let c = Math.max(0, firstLabel[row]); c < zone && own.length < maxCells && firstLabel[row] >= 0; c++) {
      if (!isLabel(cells[c])) continue;
      if (c === col2 && dataCount[row] > 0) continue;
      const t = text2(cells[c]);
      if (!own.includes(t)) own.push(t);
    }
    return own;
  };
  const columnHeading = (row, col2) => {
    const column = [];
    let b = row >= 0 && row < n ? blockAt[row] : 0;
    if (!blocks[b].length && b + 1 < blocks.length) b++;
    const block = blocks[b].slice(-MAX_HEADING_ROWS);
    block.forEach((br, i) => {
      const h = heading(br, col2, i < block.length - 1);
      if (h && column[column.length - 1] !== h) column.push(h);
    });
    if (!column.length && header && !block.includes(-1)) {
      const h = text2(header[col2]);
      if (h) column.push(h);
    }
    return column;
  };
  let absorbed = null;
  return {
    cell: (row, col2) => row >= 0 && row < n ? { sections: sectionsOf(row), row: rowLabel(row, col2), column: columnHeading(row, col2) } : { sections: [], row: [], column: [] },
    row: (row) => row >= 0 && row < n ? { sections: sectionsOf(row), row: rowLabel(row), heading: isHeading(row) } : { sections: [], row: [], heading: false },
    column: columnHeading,
    headingRow: (row) => row >= 0 && row < n && !!band[row] && !hasLabel(row),
    labelColumns() {
      if (absorbed) return absorbed;
      const textKeys = keyIdx.filter((k) => rows.every((r) => isBlank(r[k]) || typeof r[k] === "string"));
      const candidates = new Set(keyIdx.length <= maxCells ? textKeys : []);
      for (let c = 0; c < zone; c++) candidates.add(c);
      absorbed = [...candidates].filter((c) => {
        let used = false;
        const isKey = textKeys.includes(c);
        for (let r = 0; r < n; r++) {
          const v = rows[r][c];
          if (isBlank(v)) continue;
          used = true;
          if (isKey) continue;
          if (!isLabel(v)) return false;
          let before = 0;
          for (const k of keyIdx) if (typeof rows[r][k] === "string") before++;
          for (let x = Math.max(0, firstLabel[r]); x < c; x++) if (!keyIdx.includes(x) && isLabel(rows[r][x])) before++;
          if (before < maxCells) continue;
          if (!rowLabel(r).includes(text2(v))) return false;
        }
        return used;
      });
      return absorbed;
    }
  };
}
function formatLabel(label) {
  const rowPart = [...label.sections, ...label.row].join(" \u203A ");
  const colPart = label.column.join(" \u203A ");
  return rowPart && colPart ? `${rowPart} \xB7 ${colPart}` : rowPart || colPart;
}

// packages/spreadiff/src/core/report.ts
function reportOf(diff, trace, oldName = "old", newName = "new") {
  const labelers = diff.sheets.map((s) => ({
    old: tableLabeler(s.diff.oldTable, { keyColumns: s.diff.keyColumns }),
    new: tableLabeler(s.diff.newTable, { keyColumns: s.diff.keyColumns })
  }));
  const side = (d, row) => d.newRows[row] >= 0 ? "new" : "old";
  const locate = (ref) => {
    const s = diff.sheets[ref.sheet];
    const d = s.diff;
    const sd = side(d, ref.row);
    const col2 = d.columns[ref.col];
    const t = sd === "new" ? d.newTable : d.oldTable;
    const tr = sd === "new" ? d.newRows[ref.row] : d.oldRows[ref.row];
    const tc = sd === "new" ? col2.newIndex : col2.oldIndex;
    const o = t.origin ?? { row: 1, col: 0 };
    const cell = tc >= 0 ? `${columnLabel(tc + o.col)}${tr + o.row + 1}` : `?${tr + o.row + 1}`;
    const labels = tc >= 0 ? labelers[ref.sheet][sd](tr, tc) : { sections: [], row: [], column: [col2.name] };
    return { sheet: s.name, cell, id: `${s.name}!${cell}`, labels, label: formatLabel(labels) };
  };
  const valueOf = (d, row, col2, sd) => {
    const r = sd === "new" ? d.newRows[row] : d.oldRows[row];
    const c = sd === "new" ? d.columns[col2].newIndex : d.columns[col2].oldIndex;
    return r >= 0 && c >= 0 ? (sd === "new" ? d.newTable : d.oldTable).rows[r][c] : null;
  };
  const formulaOf = (d, row, col2, sd) => {
    const r = sd === "new" ? d.newRows[row] : d.oldRows[row];
    const c = sd === "new" ? d.columns[col2].newIndex : d.columns[col2].oldIndex;
    return r >= 0 && c >= 0 ? (sd === "new" ? d.newTable : d.oldTable).formulas?.[r]?.[c] ?? null : null;
  };
  const edits = [];
  const downstream = [];
  const intermediate = /* @__PURE__ */ new Set();
  const computedRefs = [];
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    const w = d.columns.length;
    const keys = [...trace.sheet(si).keys()].sort((x, y) => x - y);
    for (const k of keys) {
      const t = trace.sheet(si).get(k);
      if (t.kind === "added" || t.kind === "removed") continue;
      const ref = { sheet: si, row: Math.floor(k / w), col: k % w };
      const change = cellChange(ref, t);
      if (t.kind === "computed") {
        downstream.push(change);
        computedRefs.push(ref);
      } else edits.push(change);
    }
  });
  function cellChange(ref, t) {
    const d = diff.sheets[ref.sheet].diff;
    const before = valueOf(d, ref.row, ref.col, "old");
    const after = t.stale ? t.stale.recalculated : valueOf(d, ref.row, ref.col, "new");
    const fb = formulaOf(d, ref.row, ref.col, "old");
    const fa = formulaOf(d, ref.row, ref.col, "new");
    const kind = t.external ? "external" : t.kind === "computed" ? "computed" : t.kind === "formula" ? "formula" : "input";
    const c = { ...locate(ref), kind, before, after };
    Object.defineProperty(c, "ref", { value: ref });
    if (fb || fa) c.formula = { before: fb, after: fa };
    const nd = numericDelta(before, after);
    if (nd && typeof before === "number" && typeof after === "number") {
      c.delta = nd.delta;
      if (nd.ratio !== null) c.ratio = Number(nd.ratio.toPrecision(15));
    }
    if (t.external) c.external = t.external;
    if (t.stale) c.stale = { saved: t.stale.cached };
    if (kind !== "computed") {
      c.impact = t.impactCount;
      return c;
    }
    c.causeCount = t.causeCount;
    if (t.method) c.method = t.method;
    if (t.unexplained) c.unexplained = t.unexplained;
    const causes = t.causes.map((r, i) => Object.defineProperty({ id: locate(r).id, label: locate(r).label, contribution: t.contributions?.[i] ?? null }, "ref", { value: r }));
    causes.sort((x, y) => Math.abs(y.contribution ?? 0) - Math.abs(x.contribution ?? 0));
    c.causes = causes;
    return c;
  }
  if (computedRefs.length <= 5e4) {
    for (const ref of computedRefs) for (const p of trace.parents(ref)) intermediate.add(`${p.sheet}:${p.row}:${p.col}`);
    computedRefs.forEach((ref, i) => downstream[i].output = !intermediate.has(`${ref.sheet}:${ref.row}:${ref.col}`));
    const order = downstream.map((c, i) => ({ c, i }));
    order.sort((x, y) => Number(!!y.c.output) - Number(!!x.c.output) || x.i - y.i);
    downstream.splice(0, downstream.length, ...order.map((o) => o.c));
  }
  const structure = [];
  diff.sheets.forEach((s, si) => {
    const d = s.diff;
    if (s.status === "added" || s.status === "removed") {
      structure.push({ sheet: s.name, kind: s.status === "added" ? "sheet-added" : "sheet-removed" });
      return;
    }
    const rowImpact = /* @__PURE__ */ new Map();
    const colImpact = /* @__PURE__ */ new Map();
    const w = d.columns.length;
    for (const [k, t] of trace.sheet(si)) {
      if (t.kind !== "added" && t.kind !== "removed") continue;
      const row = Math.floor(k / w);
      const col2 = k % w;
      const st2 = d.columns[col2].status;
      if (st2 === "added" || st2 === "removed") colImpact.set(col2, (colImpact.get(col2) ?? 0) + t.impactCount);
      else rowImpact.set(row, (rowImpact.get(row) ?? 0) + t.impactCount);
    }
    for (const status of [ROW_STATUS.added, ROW_STATUS.removed]) {
      const sd = status === ROW_STATUS.added ? "new" : "old";
      let run2 = [];
      const flush = () => {
        if (!run2.length) return;
        const t = sd === "new" ? d.newTable : d.oldTable;
        const o = t.origin ?? { row: 1, col: 0 };
        const idx = run2.map((r) => sd === "new" ? d.newRows[r] : d.oldRows[r]);
        const lab = labelers[si][sd];
        const impact = run2.reduce((n, r) => n + (rowImpact.get(r) ?? 0), 0);
        structure.push({
          sheet: s.name,
          kind: status === ROW_STATUS.added ? "rows-added" : "rows-removed",
          rows: idx.map((r) => r + o.row + 1),
          labels: idx.map((r) => {
            const l = lab(r, Math.max(0, t.columns.length - 1));
            return [...l.sections, ...l.row].join(" \u203A ");
          }),
          ...impact ? { impact } : {}
        });
        run2 = [];
      };
      for (let r = 0; r < d.rowCount; r++) {
        if (d.status[r] === status) {
          if (run2.length && run2[run2.length - 1] !== r - 1) flush();
          run2.push(r);
        }
      }
      flush();
    }
    const moved = [];
    for (let r = 0; r < d.rowCount; r++) if (d.moved[r]) moved.push(r);
    if (moved.length) {
      const o = d.newTable.origin ?? { row: 1, col: 0 };
      structure.push({
        sheet: s.name,
        kind: "rows-moved",
        rows: moved.map((r) => d.newRows[r] + o.row + 1),
        labels: moved.map((r) => {
          const l = labelers[si].new(d.newRows[r], 0);
          return [...l.sections, ...l.row].join(" \u203A ");
        })
      });
    }
    d.columns.forEach((c, ci) => {
      if (c.status === "added" || c.status === "removed") {
        const impact = colImpact.get(ci);
        structure.push({ sheet: s.name, kind: c.status === "added" ? "column-added" : "column-removed", column: c.name, ...impact ? { impact } : {} });
      } else if (c.status === "renamed") structure.push({ sheet: s.name, kind: "column-renamed", column: `${c.oldName} \u2192 ${c.name}` });
    });
  });
  const st = trace.stats;
  const notes = [];
  if (st.stale) notes.push(`${st.stale} result${st.stale === 1 ? " was" : "s were"} saved in the new file without recalculating; the recalculated values are shown.`);
  if (st.external) notes.push(`${st.external} value${st.external === 1 ? "" : "s"} came from outside data (a data add-in, a web or RTD function, or another workbook). ${st.external === 1 ? "It is" : "They are"} taken as saved, not recalculated.`);
  const byRefs = downstream.filter((c) => c.method === "references").length;
  if (byRefs) notes.push(`${byRefs} downstream result${byRefs === 1 ? "'s causes come" : "s' causes come"} from formula references, not recalculation (a formula on the path couldn't be evaluated). They may list a cause that didn't matter, and have no contributions.`);
  if (st.unexplained) notes.push(`${st.unexplained} result${st.unexplained === 1 ? "" : "s"} changed with no edit found upstream (volatile functions such as NOW, or files calculated by different apps).`);
  if (trace.truncated) notes.push("The workbook was too large to trace completely; some causes may be missing.");
  const computed = [diff.computedValues?.old, diff.computedValues?.new].reduce((n, c) => n + (c?.computed ?? 0), 0);
  if (computed) notes.push(`${computed} formula result${computed === 1 ? " was" : "s were"} missing from the files (saved by a tool that doesn't calculate) and computed by spreadiff.`);
  const report = {
    old: oldName,
    new: newName,
    totals: {
      inputs: edits.filter((e) => e.kind === "input").length,
      formulas: edits.filter((e) => e.kind === "formula").length,
      external: edits.filter((e) => e.kind === "external").length,
      downstream: downstream.length,
      outputs: downstream.filter((c) => c.output).length,
      recalculated: st.recalculated,
      unexplained: st.unexplained,
      stale: st.stale
    },
    edits,
    downstream,
    structure,
    objects: diff.objects ?? [],
    notes
  };
  Object.defineProperties(report, { diff: { value: diff }, trace: { value: trace } });
  return report;
}

// packages/spreadiff-lite/src/report.ts
var isWorkbook = (f) => "sheets" in f;
var toWorkbook = (f, options) => {
  const wb = isWorkbook(f) ? f : parseFile(f.name, f.data, { header: "auto", ...options.parse });
  return options.sheets ? { ...wb, sheets: wb.sheets.filter((s) => options.sheets.includes(s.name)) } : wb;
};
function compare2(oldFile, newFile, options = {}) {
  const a = toWorkbook(oldFile, options);
  const b = toWorkbook(newFile, options);
  const diff = diffWorkbooks(a, b, options.diff);
  return reportOf(diff, traceWorkbook(diff, options.trace), a.name, b.name);
}

// packages/spreadiff-lite/src/markdown.ts
var PIPE = /\|/g;
var cellText2 = (s) => s.replace(/\r?\n/g, " ").replace(PIPE, "\\|");
function toMarkdown(report, options = {}) {
  const maxRows = options.maxRows ?? 50;
  const maxCauses = options.maxCauses ?? 3;
  const level = Math.min(5, Math.max(1, options.headingLevel ?? 2));
  const digits = options.digits ?? 15;
  const h = (n, text3) => `${"#".repeat(level + n)} ${text3}`;
  const num3 = (v) => {
    if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
    return String(Number(v.toPrecision(Math.min(15, Math.max(1, digits)))));
  };
  const val = (v) => {
    if (v === null || v === "") return "*(empty)*";
    return cellText2(typeof v === "number" ? num3(v) : formatCell(v));
  };
  const signed = (v) => v > 0 ? `+${num3(v)}` : v < 0 ? `\u2212${num3(-v)}` : "0";
  const pct = (r) => {
    const p = r * 100;
    const s = Math.abs(p) >= 10 ? p.toFixed(1) : Math.abs(p) >= 1 ? p.toFixed(2) : Number(p.toPrecision(2)).toString();
    return `${p > 0 ? "+" : p < 0 ? "\u2212" : ""}${s.replace("-", "")}%`;
  };
  const share = (v) => {
    const r = Number(v.toPrecision(Math.min(4, digits)));
    return r > 0 ? `+${r}` : r < 0 ? `\u2212${-r}` : "0";
  };
  const change = (c) => c.delta === void 0 ? "" : c.ratio === void 0 ? signed(c.delta) : `${signed(c.delta)} (${pct(c.ratio)})`;
  const code = (f) => f ? `\`${cellText2(f).replace(/`/g, "'")}\`` : "*(value)*";
  const what = (c) => c.label ? cellText2(c.label) : "";
  const { totals: t } = report;
  const out = [h(0, `${cellText2(report.old)} \u2192 ${cellText2(report.new)}`), ""];
  const parts = [];
  const plural = (n, one, many = `${one}s`) => `${n.toLocaleString("en-US")} ${n === 1 ? one : many}`;
  if (t.inputs) parts.push(plural(t.inputs, "input"));
  if (t.formulas) parts.push(plural(t.formulas, "formula"));
  if (t.external) parts.push(plural(t.external, "outside-data value"));
  const edits = report.edits.length;
  const shape = report.structure.length;
  const objects = report.objects.length;
  if (!edits && !t.downstream && !shape && !objects) {
    out.push("No changes.");
    return out.join("\n") + "\n";
  }
  let summary = edits ? `**${plural(edits, "edit")}** (${parts.join(", ")})` : "**No cell edits**";
  if (t.downstream) {
    summary += ` moved **${plural(t.downstream, "result")} downstream**`;
    if (t.outputs && t.outputs < t.downstream) summary += ` (${plural(t.outputs, "final output")})`;
  } else if (edits) summary += "; nothing downstream moved";
  if (shape) summary += `${t.downstream || edits ? "; " : ", "}${plural(shape, "change")} to rows, columns or sheets`;
  if (objects) summary += `${t.downstream || edits || shape ? "; " : ", "}${plural(objects, "change")} to names, charts or pivot tables`;
  summary += ".";
  if (t.downstream) {
    summary += t.recalculated === t.downstream ? " Every cause was found by recalculating." : t.recalculated ? ` Causes of ${plural(t.recalculated, "result")} were found by recalculating; the rest by following formula references.` : " Causes were found by following formula references (the formulas on the path couldn't be recalculated).";
  }
  out.push(summary, "");
  const table = (head, rows, total, more) => {
    out.push(`| ${head.join(" | ")} |`, `| ${head.map(() => "---").join(" | ")} |`);
    for (const r of rows) out.push(`| ${r.join(" | ")} |`);
    if (total > rows.length) out.push("", `\u2026and ${plural(total - rows.length, more)}.`);
    out.push("");
  };
  if (edits) {
    out.push(h(1, "Edits"), "");
    const rows = report.edits.slice(0, maxRows).map((c) => {
      const before = c.kind === "formula" ? code(c.formula?.before ?? null) : val(c.before);
      const after = c.kind === "formula" ? code(c.formula?.after ?? null) : val(c.after);
      const note = c.kind === "external" ? ` *(${cellText2(c.external.provider)} ${cellText2(c.external.function)})*` : c.kind === "input" && c.formula?.before ? " *(was a formula)*" : "";
      const result = c.kind === "formula" && (c.before !== null || c.after !== null) ? `${val(c.before)} \u2192 ${val(c.after)}` : change(c);
      return [`\`${cellText2(c.id)}\``, what(c) + note, before, after, result, c.impact ? String(c.impact) : ""];
    });
    table(["Cell", "What", "Before", "After", "Change", "Moves"], rows, edits, "more edit");
  }
  const downstreamTable = (title, list, more) => {
    if (!list.length) return;
    out.push(h(1, title), "");
    const rows = list.slice(0, maxRows).map((c) => {
      const causes = (c.causes ?? []).slice(0, maxCauses).map((x) => {
        const name = x.label || x.id;
        return x.contribution === null ? cellText2(name) : `${cellText2(name)} ${share(x.contribution)}`;
      });
      const extra = (c.causeCount ?? 0) - causes.length;
      let driven = causes.join("; ") + (extra > 0 ? `; ${plural(extra, "more", "more")}` : "");
      if (!c.causeCount) driven = c.unexplained === "volatile" ? "*volatile function*" : c.unexplained === "dynamic" ? "*dynamic reference*" : "*no edit found*";
      else if (c.method === "references") driven += " *(by references)*";
      const after = val(c.after) + (c.stale ? ` *(file saved ${val(c.stale.saved)})*` : "");
      return [`\`${cellText2(c.id)}\``, what(c), val(c.before), after, change(c), driven];
    });
    table(["Cell", "What", "Before", "After", "Change", "Driven by"], rows, list.length, more);
  };
  const split = t.outputs > 0 && t.outputs < t.downstream;
  if (!split) downstreamTable("Downstream", report.downstream, "more result");
  else {
    downstreamTable("Final outputs", report.downstream.filter((c) => c.output), "more output");
    if (options.intermediate !== false) downstreamTable("Intermediate results", report.downstream.filter((c) => !c.output), "more intermediate result");
  }
  if (shape) {
    out.push(h(1, "Rows, columns and sheets"), "");
    for (const s of report.structure.slice(0, maxRows)) out.push(`- ${structureLine(s)}`);
    if (shape > maxRows) out.push(`- \u2026and ${plural(shape - maxRows, "more")}.`);
    out.push("");
  }
  if (objects) {
    out.push(h(1, "Names, charts and pivot tables"), "");
    const noun = { name: "Name", chart: "Chart", pivot: "Pivot table" };
    for (const o of report.objects.slice(0, maxRows)) {
      const details = o.details.filter(Boolean).map((d) => cellText2(d)).join("; ");
      out.push(`- ${noun[o.kind]} \`${cellText2(o.name)}\` ${o.change}${details ? `: ${details}` : ""}`);
    }
    if (objects > maxRows) out.push(`- \u2026and ${plural(objects - maxRows, "more")}.`);
    out.push("");
  }
  if (report.notes.length) {
    out.push(h(1, "Notes"), "");
    for (const n of report.notes) out.push(`- ${n}`);
    out.push("");
  }
  return out.join("\n").replace(/\n+$/, "\n");
}
function structureLine(s) {
  const sheet = `**${cellText2(s.sheet)}**`;
  const moves = s.impact ? `; moves ${s.impact} result${s.impact === 1 ? "" : "s"}` : "";
  const rows = () => {
    const r = s.rows ?? [];
    const span = r.length > 1 && r[r.length - 1] - r[0] === r.length - 1 ? `rows ${r[0]}\u2013${r[r.length - 1]}` : r.length === 1 ? `row ${r[0]}` : `rows ${r.slice(0, 8).join(", ")}${r.length > 8 ? ", \u2026" : ""}`;
    const labels = (s.labels ?? []).filter(Boolean);
    const named = labels.length ? ` (${labels.slice(0, 5).map(cellText2).join(", ")}${labels.length > 5 ? ", \u2026" : ""})` : "";
    return span + named;
  };
  switch (s.kind) {
    case "sheet-added":
      return `${sheet}: sheet added`;
    case "sheet-removed":
      return `${sheet}: sheet removed`;
    case "rows-added":
      return `${sheet}: ${rows()} added${moves}`;
    case "rows-removed":
      return `${sheet}: ${rows()} removed${moves}`;
    case "rows-moved":
      return `${sheet}: ${rows()} moved`;
    case "column-added":
      return `${sheet}: column *${cellText2(s.column ?? "")}* added${moves}`;
    case "column-removed":
      return `${sheet}: column *${cellText2(s.column ?? "")}* removed${moves}`;
    case "column-renamed":
      return `${sheet}: column renamed *${cellText2(s.column ?? "")}*`;
  }
}

// action/src/main.ts
var MARKER = "<!-- spreadiff-report -->";
var LIMIT = 6e4;
var input = (name, fallback = "") => process.env[`INPUT_${name.replace(/ /g, "_").toUpperCase()}`]?.trim() || fallback;
function globs(patterns) {
  const parts = patterns.split(/[\n,]+(?![^{]*\})/).map((p) => p.trim()).filter(Boolean);
  const one = (g) => g.replace(/[.+^$()|[\]\\]/g, "\\$&").replace(/\*\*\//g, "").replace(/\*\*/g, "").replace(/\*/g, "").replace(/\?/g, "").replace(/\{([^}]*)\}/g, (_, alts) => `(?:${alts.split(",").join("|")})`).replace(/\u0001/g, "(?:.*/)?").replace(/\u0002/g, ".*").replace(/\u0003/g, "[^/]*").replace(/\u0004/g, "[^/]");
  return new RegExp(`^(?:${parts.map(one).join("|")})$`, "i");
}
var git = (args, encoding = "utf8") => execFileSync("git", args, { encoding: encoding === "utf8" ? "utf8" : void 0, maxBuffer: 512 * 1024 * 1024, stdio: ["ignore", "pipe", "pipe"] });
function blob(sha, path) {
  try {
    return new Uint8Array(git(["show", `${sha}:${path}`], "buffer"));
  } catch {
    return null;
  }
}
function reportBetween(base, head, options) {
  const names = git(["diff", "--name-only", "--no-renames", `${base}...${head}`]).split("\n").filter((p) => p && options.paths.test(p));
  const sections = [];
  let changed = 0;
  for (const path of names) {
    const empty = () => ({ name: path, sheets: [] });
    const a = blob(base, path);
    const b = blob(head, path);
    try {
      const report = compare2(a ? { name: path, data: a } : empty(), b ? { name: path, data: b } : empty(), {
        diff: options.minChange ? { relativeTolerance: options.minChange } : void 0
      });
      report.old = a ? `${path} (base)` : `${path} (new file)`;
      report.new = b ? `${path} (head)` : `${path} (deleted)`;
      const n = report.edits.length + report.downstream.length + report.structure.length + report.objects.length;
      if (n) changed++;
      sections.push(toMarkdown(report, { headingLevel: 3, maxRows: options.maxRows }));
    } catch (e) {
      sections.push(`### ${path}

Couldn't compare: ${e.message}
`);
    }
  }
  const head1 = names.length ? `## Spreadsheet changes

${changed} of ${names.length} changed spreadsheet${names.length === 1 ? "" : "s"} differ in content.

` : "## Spreadsheet changes\n\nNo spreadsheets changed.\n";
  let markdown = head1 + sections.join("\n");
  if (markdown.length > LIMIT) markdown = markdown.slice(0, LIMIT) + "\n\n\u2026cut to fit a comment. The job summary has it all.\n";
  return { markdown, files: names.length, changed };
}
async function github(path, token, init = {}) {
  const api = process.env.GITHUB_API_URL ?? "https://api.github.com";
  const res = await fetch(`${api}${path}`, {
    method: init.method ?? "GET",
    headers: { authorization: `Bearer ${token}`, accept: "application/vnd.github+json", "x-github-api-version": "2022-11-28", "content-type": "application/json" },
    body: init.body === void 0 ? void 0 : JSON.stringify(init.body)
  });
  if (!res.ok) throw Object.assign(new Error(`GitHub API ${init.method ?? "GET"} ${path}: ${res.status} ${await res.text()}`), { status: res.status });
  return res.status === 204 ? null : res.json();
}
async function upsertComment(repo, pr, token, body) {
  const text3 = `${MARKER}
${body}`;
  for (let page = 1; page < 50; page++) {
    const comments = await github(`/repos/${repo}/issues/${pr}/comments?per_page=100&page=${page}`, token);
    const mine = comments.find((c) => c.body?.startsWith(MARKER));
    if (mine) {
      await github(`/repos/${repo}/issues/comments/${mine.id}`, token, { method: "PATCH", body: { body: text3 } });
      return;
    }
    if (comments.length < 100) break;
  }
  await github(`/repos/${repo}/issues/${pr}/comments`, token, { method: "POST", body: { body: text3 } });
}
async function run() {
  const event = JSON.parse(readFileSync(process.env.GITHUB_EVENT_PATH, "utf8"));
  const pr = event.pull_request;
  const base = input("base", pr?.base?.sha ?? event.before);
  const head = input("head", pr?.head?.sha ?? process.env.GITHUB_SHA ?? "HEAD");
  if (!base) throw new Error("No base commit: run on pull_request, or set the base input");
  const pct = input("min-change");
  const minChange = pct ? pct.endsWith("%") ? Number(pct.slice(0, -1)) / 100 : Number(pct) : void 0;
  const { markdown, files, changed } = reportBetween(base, head, {
    paths: globs(input("paths", "**/*.{xlsx,xlsm,xlsb,xls,ods,csv,tsv}")),
    minChange,
    maxRows: Number(input("max-rows", "25"))
  });
  if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, markdown + "\n");
  if (process.env.GITHUB_OUTPUT) appendFileSync(process.env.GITHUB_OUTPUT, `files=${files}
changed=${changed}
`);
  const token = input("github-token");
  const comment = input("comment", "true") !== "false";
  if (process.env.SPREADIFF_DRY_RUN) {
    process.stdout.write(markdown);
    return;
  }
  if (comment && pr && files && token) {
    try {
      await upsertComment(process.env.GITHUB_REPOSITORY, pr.number, token, markdown);
    } catch (e) {
      if (e.status !== 403) throw e;
      console.log(`::warning::Couldn't comment on the pull request (the token can't write; pull requests from forks get a read-only one). The report is in the job summary.`);
      console.log(markdown);
    }
  } else console.log(markdown);
}
if (process.env.GITHUB_ACTIONS || process.env.SPREADIFF_DRY_RUN) {
  run().catch((e) => {
    console.log(`::error::${e.message}`);
    process.exitCode = 1;
  });
}
export {
  globs,
  reportBetween
};
