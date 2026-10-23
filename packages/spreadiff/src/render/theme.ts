/**
 * Theming. Everything visual is a CSS custom property on the host element, so
 * you theme a canvas-rendered grid exactly like a DOM one:
 *
 *   .my-diff { --sd-added: #22c55e; --sd-font-size: 12px; }
 *
 * Set tokens on any ancestor; the defaults below only apply when unset.
 *
 * Canvas can't read CSS variables, so we resolve each token through a hidden
 * probe element (which also evaluates `light-dark()` and `color-mix()`), and
 * re-resolve when the color scheme or host attributes change.
 */

export const STYLE = /* css */ `
:host, .sd-root {
  color-scheme: light dark;
  --sdi-mixer: var(--sd-mixer, light-dark(#000, #fff));
  --sdi-bg: var(--sd-bg, light-dark(#ffffff, #0b0b0c));
  --sdi-fg: var(--sd-fg, light-dark(#0b0b0c, #ededef));
  --sdi-fg-muted: var(--sd-fg-muted, color-mix(in srgb, var(--sdi-fg) 52%, var(--sdi-bg)));
  --sdi-fg-faint: var(--sd-fg-faint, color-mix(in srgb, var(--sdi-fg) 32%, var(--sdi-bg)));
  --sdi-border: var(--sd-border, color-mix(in srgb, var(--sdi-bg) 90%, var(--sdi-mixer)));
  --sdi-border-strong: var(--sd-border-strong, color-mix(in srgb, var(--sdi-bg) 80%, var(--sdi-mixer)));
  --sdi-header-bg: var(--sd-header-bg, color-mix(in srgb, var(--sdi-bg) 97%, var(--sdi-mixer)));
  --sdi-gutter-bg: var(--sd-gutter-bg, color-mix(in srgb, var(--sdi-bg) 98%, var(--sdi-mixer)));
  --sdi-separator-bg: var(--sd-separator-bg, color-mix(in srgb, var(--sdi-bg) 95.5%, var(--sdi-mixer)));
  --sdi-hover: var(--sd-hover, color-mix(in srgb, var(--sdi-fg) 4%, transparent));
  --sdi-added: var(--sd-added, light-dark(#0dbe4e, #5ecc71));
  --sdi-removed: var(--sd-removed, light-dark(#ff2e3f, #ff6762));
  --sdi-modified: var(--sd-modified, light-dark(#009fff, #69b1ff));
  --sdi-added-bg: var(--sd-added-bg, light-dark(color-mix(in srgb, var(--sdi-added) 9%, var(--sdi-bg)), color-mix(in srgb, var(--sdi-added) 12%, var(--sdi-bg))));
  --sdi-removed-bg: var(--sd-removed-bg, light-dark(color-mix(in srgb, var(--sdi-removed) 9%, var(--sdi-bg)), color-mix(in srgb, var(--sdi-removed) 12%, var(--sdi-bg))));
  --sdi-modified-bg: var(--sd-modified-bg, light-dark(color-mix(in srgb, var(--sdi-modified) 5%, var(--sdi-bg)), color-mix(in srgb, var(--sdi-modified) 8%, var(--sdi-bg))));
  --sdi-added-emphasis: var(--sd-added-emphasis, light-dark(color-mix(in srgb, var(--sdi-added) 22%, var(--sdi-bg)), color-mix(in srgb, var(--sdi-added) 28%, var(--sdi-bg))));
  --sdi-removed-emphasis: var(--sd-removed-emphasis, light-dark(color-mix(in srgb, var(--sdi-removed) 20%, var(--sdi-bg)), color-mix(in srgb, var(--sdi-removed) 26%, var(--sdi-bg))));
  --sdi-added-fg: var(--sd-added-fg, color-mix(in srgb, var(--sdi-added) 70%, var(--sdi-fg)));
  --sdi-removed-fg: var(--sd-removed-fg, color-mix(in srgb, var(--sdi-removed) 70%, var(--sdi-fg)));
  --sdi-selection: var(--sd-selection, var(--sdi-modified));
  --sdi-hatch: var(--sd-hatch, color-mix(in srgb, var(--sdi-fg) 7%, transparent));
  --sdi-font-family: var(--sd-font-family, "Inter", ui-sans-serif, system-ui, -apple-system, "Segoe UI", Roboto, sans-serif);
  --sdi-mono-font-family: var(--sd-mono-font-family, "Berkeley Mono", "JetBrains Mono", ui-monospace, "SF Mono", Menlo, Consolas, monospace);
  --sdi-font-size: var(--sd-font-size, 13px);
  --sdi-row-height: var(--sd-row-height, 30px);
  --sdi-radius: var(--sd-radius, 10px);

  display: block;
  position: relative;
  min-height: 0;
  background: var(--sdi-bg);
  color: var(--sdi-fg);
  font-family: var(--sdi-font-family);
  font-size: var(--sdi-font-size);
  overflow: hidden;
}
:host([data-theme="light"]), .sd-root[data-theme="light"] { color-scheme: light; }
:host([data-theme="dark"]), .sd-root[data-theme="dark"] { color-scheme: dark; }

.sd-scroller {
  position: absolute; inset: 0;
  overflow: auto;
  overscroll-behavior: contain;
  scrollbar-width: thin;
  scrollbar-color: var(--sdi-border-strong) transparent;
  outline: none;
}
.sd-spacer { pointer-events: none; }
.sd-canvas, .sd-minimap {
  position: absolute; top: 0; left: 0;
  pointer-events: none;
}
.sd-overlay { position: absolute; inset: 0; pointer-events: none; overflow: hidden; }
::slotted(.sd-annotation) {
  position: absolute; left: 0;
  pointer-events: auto;
  box-sizing: border-box;
  will-change: transform;
}
.sd-tooltip {
  position: absolute;
  z-index: 3;
  max-width: min(420px, 80%);
  padding: 10px 12px;
  border-radius: var(--sdi-radius);
  background: color-mix(in srgb, var(--sdi-bg) 96%, var(--sdi-mixer));
  border: 1px solid var(--sdi-border-strong);
  box-shadow: 0 8px 28px -6px rgb(0 0 0 / 0.22), 0 2px 6px -2px rgb(0 0 0 / 0.12);
  font-size: 12px;
  line-height: 1.45;
  pointer-events: none;
  opacity: 0;
  transform: translateY(2px);
  transition: opacity 90ms ease, transform 90ms ease;
}
.sd-tooltip[data-open] { opacity: 1; transform: none; }
.sd-tooltip .sd-tt-head {
  display: flex; gap: 8px; align-items: baseline;
  color: var(--sdi-fg-muted); font-size: 11px; margin-bottom: 6px;
}
.sd-tooltip .sd-tt-col { color: var(--sdi-fg); font-weight: 600; }
.sd-tooltip .sd-tt-row { display: grid; grid-template-columns: 14px 1fr; gap: 6px; align-items: start; }
.sd-tooltip .sd-tt-sign { font-family: var(--sdi-mono-font-family); color: var(--sdi-fg-muted); }
.sd-tooltip .sd-tt-val { font-family: var(--sdi-mono-font-family); word-break: break-word; white-space: pre-wrap; }
.sd-tooltip .sd-tt-old .sd-tt-sign { color: var(--sdi-removed); }
.sd-tooltip .sd-tt-new .sd-tt-sign { color: var(--sdi-added); }
.sd-tooltip del { background: var(--sdi-removed-emphasis); text-decoration: none; border-radius: 3px; }
.sd-tooltip ins { background: var(--sdi-added-emphasis); text-decoration: none; border-radius: 3px; }
.sd-tooltip .sd-tt-delta {
  margin-top: 6px; padding-top: 6px;
  border-top: 1px solid var(--sdi-border);
  font-family: var(--sdi-mono-font-family);
  color: var(--sdi-fg-muted);
}
.sd-tooltip .sd-tt-delta b[data-sign="+"] { color: var(--sdi-added); font-weight: 600; }
.sd-tooltip .sd-tt-delta b[data-sign="-"] { color: var(--sdi-removed); font-weight: 600; }
.sd-tooltip .sd-tt-formula { color: var(--sdi-fg-muted); font-family: var(--sdi-mono-font-family); font-size: 11px; }
.sd-probe { position: absolute; width: 0; height: 0; visibility: hidden; pointer-events: none; }
`;

export interface Theme {
  bg: string;
  fg: string;
  fgMuted: string;
  fgFaint: string;
  border: string;
  borderStrong: string;
  headerBg: string;
  gutterBg: string;
  separatorBg: string;
  hover: string;
  added: string;
  removed: string;
  modified: string;
  addedBg: string;
  removedBg: string;
  modifiedBg: string;
  addedEmphasis: string;
  removedEmphasis: string;
  addedFg: string;
  removedFg: string;
  selection: string;
  hatch: string;
  fontFamily: string;
  monoFontFamily: string;
  fontSize: number;
  rowHeight: number;
  dark: boolean;
}

const COLOR_TOKENS: [keyof Theme, string][] = [
  ['bg', '--sdi-bg'],
  ['fg', '--sdi-fg'],
  ['fgMuted', '--sdi-fg-muted'],
  ['fgFaint', '--sdi-fg-faint'],
  ['border', '--sdi-border'],
  ['borderStrong', '--sdi-border-strong'],
  ['headerBg', '--sdi-header-bg'],
  ['gutterBg', '--sdi-gutter-bg'],
  ['separatorBg', '--sdi-separator-bg'],
  ['hover', '--sdi-hover'],
  ['added', '--sdi-added'],
  ['removed', '--sdi-removed'],
  ['modified', '--sdi-modified'],
  ['addedBg', '--sdi-added-bg'],
  ['removedBg', '--sdi-removed-bg'],
  ['modifiedBg', '--sdi-modified-bg'],
  ['addedEmphasis', '--sdi-added-emphasis'],
  ['removedEmphasis', '--sdi-removed-emphasis'],
  ['addedFg', '--sdi-added-fg'],
  ['removedFg', '--sdi-removed-fg'],
  ['selection', '--sdi-selection'],
  ['hatch', '--sdi-hatch'],
];

/** Resolve every token to a canvas-ready color string. */
export function resolveTheme(probeHost: HTMLElement | ShadowRoot, styleSource: Element): Theme {
  const doc = styleSource.ownerDocument;
  const probe = doc.createElement('span');
  probe.className = 'sd-probe';
  probeHost.appendChild(probe);
  const ctx = doc.createElement('canvas').getContext('2d', { willReadFrequently: true })!;
  const theme = {} as Record<string, unknown>;
  for (const [key, prop] of COLOR_TOKENS) {
    probe.style.color = `var(${prop})`;
    theme[key] = toCanvasColor(ctx, getComputedStyle(probe).color);
  }
  probe.remove();
  const cs = getComputedStyle(styleSource);
  const px = (prop: string, fallback: number) => {
    const v = parseFloat(cs.getPropertyValue(prop));
    return Number.isFinite(v) ? v : fallback;
  };
  theme.fontFamily = cs.getPropertyValue('--sdi-font-family').trim() || 'system-ui, sans-serif';
  theme.monoFontFamily = cs.getPropertyValue('--sdi-mono-font-family').trim() || 'monospace';
  theme.fontSize = px('--sdi-font-size', 13);
  theme.rowHeight = px('--sdi-row-height', 30);
  theme.dark = luminance(ctx, theme.bg as string) < 0.4;
  return theme as unknown as Theme;
}

/**
 * Normalize any computed CSS color (`color(srgb …)`, `lab(…)`, …) into
 * something every canvas implementation accepts, by painting one pixel.
 */
function toCanvasColor(ctx: CanvasRenderingContext2D, css: string): string {
  ctx.canvas.width = ctx.canvas.height = 1;
  ctx.clearRect(0, 0, 1, 1);
  ctx.fillStyle = '#000';
  ctx.fillStyle = css;
  ctx.fillRect(0, 0, 1, 1);
  const [r, g, b, a] = ctx.getImageData(0, 0, 1, 1).data;
  return a === 255 ? `rgb(${r},${g},${b})` : `rgba(${r},${g},${b},${(a / 255).toFixed(3)})`;
}

function luminance(ctx: CanvasRenderingContext2D, css: string): number {
  ctx.fillStyle = css;
  ctx.fillRect(0, 0, 1, 1);
  const [r, g, b] = ctx.getImageData(0, 0, 1, 1).data;
  return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
}
