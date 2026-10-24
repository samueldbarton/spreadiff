/**
 * Small, allocation-light XML helpers for the spreadsheet readers: find
 * elements by tag, read attributes, decode entities. Not a general XML
 * parser: enough for the parts of OOXML and ODF the readers use.
 */

/**
 * Call `fn(attrs, body)` for each `<tag …>body</tag>` or `<tag …/>` in
 * `xml`. Stops at the first unclosed element instead of rescanning.
 */
export function eachElement(xml: string, tag: string, fn: (attrs: string, body: string | null) => void): void {
  const open = `<${tag}`;
  const close = `</${tag}>`;
  let i = xml.indexOf(open);
  while (i >= 0) {
    const next = xml.charCodeAt(i + open.length);
    // Must be the whole tag name: followed by whitespace, '>' or '/'.
    if (next !== 62 && next !== 47 && next !== 32 && next !== 9 && next !== 10 && next !== 13) {
      i = xml.indexOf(open, i + 1);
      continue;
    }
    const gt = xml.indexOf('>', i);
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

/** Concatenated text of every `<t>` in `body`, skipping phonetic `<rPh>` runs. */
export function textRuns(body: string): string {
  let s = '';
  const clean = body.indexOf('<rPh') < 0 ? body : removeElements(body, 'rPh');
  eachElement(clean, 't', (_, b) => {
    if (b) s += b;
  });
  return decode(s);
}

export function removeElements(xml: string, tag: string): string {
  let out = '';
  let pos = 0;
  const open = `<${tag}`;
  const close = `</${tag}>`;
  for (;;) {
    const i = xml.indexOf(open, pos);
    if (i < 0) break;
    const end = xml.indexOf(close, i);
    if (end < 0) break;
    out += xml.slice(pos, i);
    pos = end + close.length;
  }
  return out + xml.slice(pos);
}

/** Attribute value by name; accepts either quote style and any whitespace. */
export function attr(attrs: string, name: string): string | undefined {
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
  return undefined;
}

const ENTITY = /&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g;
export function decode(s: string): string {
  // XML parsers must read a literal CR LF (or lone CR) as LF; an escaped
  // &#13; survives. Excel writes CR LF, LibreOffice LF, for the same text.
  if (s.indexOf('\r') >= 0) s = s.replace(/\r\n?/g, '\n');
  // Excel escapes control characters in text as _xHHHH_ (OOXML ST_Xstring).
  if (s.indexOf('_x') >= 0) s = s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h: string) => String.fromCharCode(parseInt(h, 16)));
  if (s.indexOf('&') < 0) return s;
  return s.replace(ENTITY, (all, e: string) => {
    switch (e) {
      case 'amp': return '&';
      case 'lt': return '<';
      case 'gt': return '>';
      case 'quot': return '"';
      case 'apos': return "'";
      default: {
        const cp = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return cp >= 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : all;
      }
    }
  });
}

export function decodeRef(ref: string): [row: number, col: number] {
  let col = 0;
  let i = 0;
  for (; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
  }
  return [Number(ref.slice(i)) - 1, col - 1];
}

export function resolvePath(from: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = from.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}
