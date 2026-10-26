# Theming

Every color, font and size in the grid is a CSS custom property. Set the
properties on the grid's container, or on any ancestor, and the grid picks
them up. This works even though the grid draws on a canvas.

```css
.my-diff {
  --sd-added: #16a34a;
  --sd-removed: #dc2626;
  --sd-computed: #ca8a04;
  --sd-font-family: 'Inter', system-ui, sans-serif;
  --sd-font-size: 12px;
  --sd-row-height: 26px;
}
```

You only need to set what you want to change. Most surfaces are *derived*
from a few base colors with `color-mix()`, so changing `--sd-bg` and
`--sd-fg` re-tints borders, headers and gutters consistently. Changing
`--sd-added` re-tints added rows, new values and their highlights.

<!-- demo: theme -->

## Light and dark

The defaults use `light-dark()`, so the grid follows the page's color
scheme: `prefers-color-scheme`, or `color-scheme` set on the page. To force
one, pass `theme: 'light' | 'dark'` to the grid.

To theme each scheme separately, use `light-dark()` in your own values:

```css
.my-diff {
  --sd-bg: light-dark(#fffdf8, #12100e);
  --sd-computed: light-dark(#b45309, #fcd34d);
}
```

The grid re-reads its tokens when the color scheme changes, and when a
`class`, `style` or `data-theme` attribute changes on `<html>` or on the
container, which is how most theme toggles work. If you change variables
some other way from script, call `grid.refreshTheme()`.

## Tokens

| Token | Default (light / dark) | Used for |
| --- | --- | --- |
| `--sd-bg` | `#ffffff` / `#0b0b0c` | Background |
| `--sd-fg` | `#0b0b0c` / `#ededef` | Text |
| `--sd-mixer` | `#000` / `#fff` | Mixed into `--sd-bg` to derive surfaces below |
| `--sd-fg-muted` | 66% fg | Secondary text: removed rows, separators, old values of computed changes |
| `--sd-fg-faint` | 58% fg | Row numbers, column letters |
| `--sd-border` | bg + 10% mixer | Grid lines |
| `--sd-border-strong` | bg + 20% mixer | Header rule, frozen-column edge, split divider |
| `--sd-header-bg` | bg + 3% mixer | Column header row |
| `--sd-gutter-bg` | bg + 2% mixer | Row-number gutter, annotation rows, minimap |
| `--sd-separator-bg` | bg + 4.5% mixer | "N unchanged rows" separators; the old-value band of computed changes |
| `--sd-hover` | 4% fg | Hover highlight |
| `--sd-added` | `#0dbe4e` / `#5ecc71` | Added: row bars, tints, minimap |
| `--sd-removed` | `#ff2e3f` / `#ff6762` | Removed: row bars, tints, minimap |
| `--sd-modified` | `#009fff` / `#69b1ff` | Modified rows' bar, selection ring, badge tint, annotation markers |
| `--sd-modified-fg` | modified + 38% black / modified | Text and glyphs in the modified hue: the modified dot, `↕` for moved rows, change counts, rename arrows, cause outlines |
| `--sd-added-bg` | added at 9% / 12% | Tint of added rows and columns |
| `--sd-removed-bg` | removed at 9% / 12% | Tint of removed rows and columns |
| `--sd-added-emphasis` | added at 22% / 28% | Background of a changed cell's new value |
| `--sd-removed-emphasis` | removed at 20% / 26% | Background of a changed cell's old value |
| `--sd-added-fg` | added + 38% black / 70% added | Text in the added hue: `+` signs, positive deltas, names of added columns, NEW labels |
| `--sd-removed-fg` | removed + 38% black / 70% removed | Text in the removed hue: `−` signs, negative deltas, old values in `stacked` display, names of removed columns |
| `--sd-computed` | `#e08a00` / `#f5b83d` | Downstream changes: tint base |
| `--sd-computed-fg` | computed + 38% black / computed | Text and outlines in the computed hue: impact outlines, the tooltip's kind label |
| `--sd-computed-emphasis` | computed at 13% / 15% | Background of downstream changes |
| `--sd-formula` | `#7c3aed` / `#a78bfa` | The `ƒ` on formula edits |
| `--sd-selection` | `--sd-modified-fg` | Selected-cell ring and keyboard focus ring (3:1 or better on every background) |
| `--sd-dim` | 62% bg | Veil over unrelated cells in trace focus |
| `--sd-hatch` | 7% fg | Hatching where a row or column is missing (split layout) |
| `--sd-font-family` | Inter, system UI fonts | Cell and header text |
| `--sd-mono-font-family` | JetBrains Mono, system monospace | Row numbers, column letters, deltas, popover values |
| `--sd-font-size` | `13px` | Cell text (headers and numbers are 1–3px smaller) |
| `--sd-row-height` | `30px` | Row height (modified rows are 1.6× in `stacked` and `delta` displays) |
| `--sd-radius` | `10px` | Popover corner radius |

Percentages are `color-mix(in srgb, …)` with the background, so a token
stays legible whatever the base colors are.

**Contrast.** The defaults meet WCAG AA: text and glyphs are 4.5:1 or
better against every background they're drawn on, in light and dark. A
browser test checks the canvas colors, and axe-core checks the DOM. The
bright status colors stay on bars and tints. Wherever a status color is
*text*, the grid uses its `-fg` variant, which is darkened in light mode.
If you set `--sd-added`, `--sd-removed`, `--sd-modified` or
`--sd-computed`, the `-fg` variants follow. Check a custom palette's text
contrast yourself, though: a pale base color can still give a pale `-fg`.
See [Accessibility](accessibility.md).

**Forced colors.** Under `forced-colors: active` (Windows High Contrast),
the grid uses the user's system colors and ignores your tokens, as the
browser does for page CSS. See [Accessibility](accessibility.md#the-canvas-grid).

[`spreadiff/html`](html.md#styling) reads the base colors and fonts too:
`--sd-bg`, `--sd-fg`, `--sd-added`, `--sd-removed`, `--sd-modified`,
`--sd-computed`, `--sd-formula`, `--sd-added-fg`, `--sd-removed-fg`,
`--sd-modified-fg`, `--sd-font-family` and `--sd-mono-font-family`.

## How it works

Canvas can't read CSS variables. On mount, and on every theme change, the
grid resolves each token through a hidden probe element. The browser
evaluates `light-dark()`, `color-mix()` and inheritance. The grid then
converts the result to a color canvas accepts by painting one pixel. This
costs about a millisecond and happens only on theme changes, never per
frame.

The grid lives in a Shadow DOM, so page styles don't leak into it, except
for these custom properties (which inherit by design) and annotation
content, which stays in your DOM on purpose ([Annotations](grid.md#annotations)).
