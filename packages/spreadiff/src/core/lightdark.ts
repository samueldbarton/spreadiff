/**
 * A fallback for browsers without CSS `light-dark()` (Safari before 17.5,
 * Chrome and Edge before 123, Firefox before 120). There, a custom property
 * whose value uses `light-dark()` is invalid at computed time, so every
 * color built on it would resolve to nothing. This derives plain light and
 * dark declarations from the same source, inside
 * `@supports not (color: light-dark(#000, #fff))`.
 */
export function lightDarkFallback(
  declarations: string,
  selectors: { light: string; darkWhenSystem: string; darkForced?: string },
): string {
  const lines = declarations
    .split(/;\s*/)
    .map((d) => d.trim())
    .filter((d) => d.includes('light-dark('));
  if (!lines.length) return '';
  const light = lines.map((d) => pick(d, 0)).join(';\n    ');
  const dark = lines.map((d) => pick(d, 1)).join(';\n    ');
  return `
@supports not (color: light-dark(#000, #fff)) {
  ${selectors.light} {
    ${light};
  }
  @media (prefers-color-scheme: dark) {
    ${selectors.darkWhenSystem} {
      ${dark};
    }
  }${selectors.darkForced ? `\n  ${selectors.darkForced} {\n    ${dark};\n  }` : ''}
}`;
}

/** `s` with every `light-dark(a, b)` replaced by its `side` argument (0 light, 1 dark). */
function pick(s: string, side: 0 | 1): string {
  let out = '';
  let i = 0;
  while (i < s.length) {
    const at = s.indexOf('light-dark(', i);
    if (at < 0) {
      out += s.slice(i);
      break;
    }
    out += s.slice(i, at);
    // Split the arguments at the top-level comma; parentheses nest.
    let depth = 0;
    let comma = -1;
    let j = at + 'light-dark('.length;
    for (; j < s.length; j++) {
      const ch = s[j];
      if (ch === '(') depth++;
      else if (ch === ')') {
        if (depth === 0) break;
        depth--;
      } else if (ch === ',' && depth === 0 && comma < 0) comma = j;
    }
    const args = [s.slice(at + 'light-dark('.length, comma), s.slice(comma + 1, j)];
    out += pick(args[side].trim(), side);
    i = j + 1;
  }
  return out;
}
