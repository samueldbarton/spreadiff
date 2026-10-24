/**
 * Formulas whose value comes from outside the workbook: market-data add-ins
 * (Bloomberg, LSEG, FactSet, S&P Capital IQ, …), Excel's own connected
 * functions (RTD, WEBSERVICE, STOCKHISTORY, cube functions, Python in Excel)
 * and links to other workbooks.
 *
 * Their results can't be recalculated without the connection, and they can
 * change when nothing in the file did (the data refreshed). The trace treats
 * such a cell as an input: its saved value is taken as given, and a change
 * to it is a root cause of everything downstream.
 */

/** Where an external value comes from, from {@link externalSource}. */
export interface ExternalSource {
  /** The data provider: `'Bloomberg'`, `'S&P Capital IQ'`, `'Excel'` (RTD, WEBSERVICE, …), `'add-in'` (any other XLL function), `'link'` (another workbook), or a name given in `liveFunctions`. */
  provider: string;
  /** The function called (`'BDP'`), or the linked workbook reference (`'[1]'`). */
  function: string;
}

/**
 * Live-data functions by provider. Names are matched case-insensitively,
 * with or without the `_xll.` prefix Excel writes for add-in functions. A
 * name ending in `*` matches any function starting with it.
 */
export const LIVE_FUNCTIONS: Readonly<Record<string, readonly string[]>> = {
  Bloomberg: ['BDP', 'BDH', 'BDS', 'BQL', 'BQL.QUERY', 'BEQS', 'BSRCH', 'BCURVE', 'BLP', 'BLPH', 'BLPW'],
  LSEG: ['TR', 'RDP.DATA', 'RDP.HISTORY', 'RDP.SEARCH', 'RHISTORY', 'RDATA', 'RSEARCH', 'DSGRID'],
  FactSet: ['FDS', 'FDSB', 'FDSC', 'FQL'],
  'S&P Capital IQ': ['CIQ*', 'SPG'],
  Morningstar: ['MSDP', 'MSTS'],
  Excel: [
    'RTD',
    'WEBSERVICE',
    'STOCKHISTORY',
    'CUBEVALUE',
    'CUBEMEMBER',
    'CUBESET',
    'CUBESETCOUNT',
    'CUBERANKEDMEMBER',
    'CUBEKPIMEMBER',
    'CUBEMEMBERPROPERTY',
    'PY',
    'COPILOT',
    'IMAGE',
    'TRANSLATE',
    'DETECTLANGUAGE',
  ],
};

interface Matcher {
  exact: Map<string, string>;
  prefixes: [string, string][];
}

const compileMatcher = (table: Readonly<Record<string, readonly string[]>>): Matcher => {
  const exact = new Map<string, string>();
  const prefixes: [string, string][] = [];
  for (const [provider, names] of Object.entries(table)) {
    for (const n of names) {
      const u = n.toUpperCase();
      if (u.endsWith('*')) prefixes.push([u.slice(0, -1), provider]);
      else exact.set(u, provider);
    }
  }
  return { exact, prefixes };
};

const DEFAULT = compileMatcher(LIVE_FUNCTIONS);
const cache = new WeakMap<object, Matcher>();

const STRING = /"(?:[^"]|"")*"/g;
// A name followed by "(", not preceded by a name character (so SUM.BDP( isn't BDP().
const CALL = /(?<![A-Za-z0-9_.\\])((?:_xll\.|_xlfn\.(?:_xlws\.)?)?)([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
// [1]Sheet!A1, '[1]My sheet'!A1, [1]!Name. Sheet names can't contain brackets; a
// structured reference (Sales[2024]) follows a name, so it isn't a link.
const LINK = /(?<![A-Za-z0-9_.\\\]\[])\[\d+\]/;

/**
 * The external source a formula reads, or null if it reads only the
 * workbook. The first live-data call (or workbook link) found wins.
 *
 * ```ts
 * import { externalSource } from 'spreadiff/core';
 * externalSource('=_xll.BDP("IBM US Equity","PX_LAST")'); // { provider: 'Bloomberg', function: 'BDP' }
 * externalSource('=SUM(A1:A3)'); // null
 * ```
 *
 * @param extra More live-data functions, by provider, on top of {@link LIVE_FUNCTIONS}
 *   (for an in-house add-in, or a provider not listed).
 */
export function externalSource(formula: string, extra?: Readonly<Record<string, readonly string[]>>): ExternalSource | null {
  if (!formula) return null;
  const f = formula.replace(STRING, '""');
  let extraMatcher: Matcher | undefined;
  if (extra) {
    extraMatcher = cache.get(extra);
    if (!extraMatcher) cache.set(extra, (extraMatcher = compileMatcher(extra)));
  }
  CALL.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = CALL.exec(f))) {
    const prefix = m[1].toLowerCase();
    const name = m[2].toUpperCase();
    for (const matcher of extraMatcher ? [extraMatcher, DEFAULT] : [DEFAULT]) {
      const hit = matcher.exact.get(name) ?? matcher.prefixes.find(([p]) => name.startsWith(p))?.[1];
      if (hit) return { provider: hit, function: name };
    }
    // Any other XLL add-in function needs its add-in, and usually a connection.
    if (prefix === '_xll.') return { provider: 'add-in', function: name };
  }
  const link = LINK.exec(f);
  if (link) return { provider: 'link', function: link[0] };
  return null;
}
