# Accessibility

spreadiff aims for WCAG 2.1 AA in all three ways of showing a diff:
- the canvas `DiffGrid`;
- the HTML from `spreadiff/lite`;
- the viewer app.

This page says what each one provides, how that is tested, and what
isn't covered yet.

## The canvas grid

A canvas is a picture: screen readers can't read what's drawn on it. The
grid therefore keeps a **visually hidden copy of the visible window** as
an ARIA grid, inside the focusable element that has `role="grid"`:

| Part | What a screen reader gets |
| --- | --- |
| The grid | Its name, a summary such as "Spreadsheet diff: 3 rows modified, 1 row added, 7 cells changed" (`ariaLabel` to override). The full size as `aria-rowcount` / `aria-colcount`, so positions are announced against the whole diff, not the window |
| Column headers | The column name, and "renamed from …", "added column" or "removed column" |
| Row headers | "Row 12, modified", "Row 4, added", "Row 7, moved from row 2" |
| Cells | The value. For a change: the new value, the old one, the direction and size, and why. For example: "70.63, was 65.4, up 5.23 (8.0%), edited"; "129168, was 115200.96, up 13967.04 (12.1%), recalculated result" |
| Folded rows | "40 unchanged rows, collapsed" |
| The selection | The selected cell is the grid's `aria-activedescendant`, so each arrow key announces the new cell |
| Popover | While open, it's the grid's `aria-describedby`; closed, it's hidden from assistive technology |

The copy covers the rows and columns on screen, so its cost doesn't grow
with the diff. It's rebuilt when scrolling settles, not every frame.
Scrolling benchmarks with it on and off didn't differ. Pass
`accessibility: false` to skip it, for example for a decorative preview
next to an accessible view.

**Keyboard.** Everything the mouse does, the keyboard does too: `n`/`p`
for the next or previous change, the arrow keys to move, Esc to clear,
⌘C/Ctrl+C to copy. The grid is one Tab stop, and Tab leaves it again:
focus is never trapped. The full table is in [The DiffGrid → Keyboard and
mouse](grid.md#keyboard-and-mouse).

**Contrast.** Text and status glyphs meet 4.5:1 against every background
they're drawn on, in light and dark. That covers:
- values, row numbers and secondary text;
- the `+`, `−`, `↕` and modified-dot markers;
- the deltas and change counts;
- the outlines of the focus mode.

The bright status colours are kept for bars and tints, which back up
signs and text rather than carry meaning alone. If you override the
theme, the `--sd-*-fg` variables set the text versions
([Theming](theming.md)).

## spreadiff/lite

`renderHTML` produces a real `<table>`, with `<th>` headers, row numbers
in cells and the sign column labelled for screen readers. Changes are
marked with text as well as colour (`+`, `−`, `•`, `↕`, struck-through old
values), so they don't depend on colour alone. The same AA text colours
apply.

## The viewer

The viewer has a heading, `banner`, `main` and `toolbar` landmarks, and
labelled controls; segmented controls are `radiogroup`s. Its muted text
meets AA in both themes.

## How it's tested

These run with the rest of the suite (`npm run test:browser`,
`npm run test:hosted`):

- **Structure.** A browser test renders a diff and checks that:
  - the grid has its role and summary;
  - the column headers, row headers and folded-row cells are there;
  - selecting a cell makes it the active descendant;
  - that cell reads exactly "12.5, was 10, up 2.5 (25.0%), edited".
- **axe-core** with no violations, run on:
  - the grid (split layout, with a selection);
  - `spreadiff/lite` HTML, unified and split, in light and dark;
  - the whole viewer with a model loaded (serious and critical findings fail the run).
- **Canvas contrast.** axe can't see inside a canvas. So a test resolves
  the grid's theme in light and dark and composites each text colour over
  each background it's drawn on, on a real canvas. It fails if any pair
  is below 4.5:1 (3:1 for outlines).

The first run of these checks found real problems, all fixed since:
- the closed popover had `role="tooltip"` and no name;
- grey text was 3.6–4.35:1;
- row numbers were 2.0–2.6:1;
- green and red deltas were 2.5:1 and 3.7:1 on white;
- the viewer had no heading or landmarks;
- some lite header cells were empty.

## Not covered yet

- **No manual screen-reader testing.** The structure follows the ARIA grid
  pattern and passes automated checks, but it hasn't been tried with
  NVDA, JAWS or VoiceOver. Automated checks can't tell whether the
  announcements are pleasant to listen to.
- **Only the visible window is in the accessibility tree.** Moving with
  the arrow keys scrolls the grid, so the next cell is always there.
  Screen-reader commands that jump far away (to the last row, say) stop
  at the window's edge. `nextChange()` and the keyboard shortcuts are the
  intended way through a long diff.
- **Colour still separates two kinds of change.** An edited input and a
  downstream result differ visually by hue (blue against amber), which
  people with tritanopia may not tell apart. The popover and the
  screen-reader text name the kind, and formula edits carry an `ƒ`, but a
  pattern or glyph for downstream cells would be better.
- **The popover opens on hover.** Its content is in the selected cell's
  description, but sighted keyboard users don't see the popover itself.
