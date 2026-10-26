# Browser and runtime support

## What each entry point needs

| Entry | Needs | Minimum |
| --- | --- | --- |
| `spreadiff/html`, `spreadiff/core` | ES2022 JavaScript. No DOM: runs in Node, Deno, Bun, workers and edge runtimes | Node 18, or any 2022+ browser |
| `spreadiff` (`DiffGrid`) | ES2022, Shadow DOM, `ResizeObserver`, canvas 2D, CSS `color-mix()` | **Chrome / Edge 111, Firefox 113, Safari 16.2** (2023) |
| `spreadiff/worker` | Module workers (`new Worker(url, { type: 'module' })`) | Chrome / Edge 80, Firefox 114, Safari 15 |
| `spreadiff/react` | React 18 or 19 | |

The grid's colours use CSS `light-dark()` to follow the page's colour
scheme. It arrived later (Chrome / Edge 123, Firefox 120, Safari 17.5,
all 2024), so the stylesheet carries a fallback for browsers without it.
They get the same colours from plain light values plus a
`prefers-color-scheme: dark` block. A browser test simulates a missing
`light-dark()` and checks every resolved colour against the normal theme
in light, dark and forced-dark modes.

`spreadiff/html`'s `HTML_CSS` has the same fallback. Mail clients vary
far more than browsers, though. Many drop `<style>` blocks or custom
properties altogether, so for email, test in the clients you target. The
markers (`+`, `−`, `•`, `↕`, struck-through old values) keep a diff
readable without any colour.

## What's tested

- **Chromium** (headless, current): unit tests, the renderer regression
  suite, the render fuzzer and the bundling test. The hosted test runs the
  packed library and the viewer as static sites, with outside requests
  blocked.
- **Firefox and WebKit** (Playwright's builds, headless): the renderer
  regression suite, axe-core and the contrast checks included
  (`BROWSER=firefox npm run test:browser`).
- **Node 22**: the unit tests and the trace evaluation.

**Not tested: Safari itself, and real devices.** Playwright's WebKit is
close to Safari but isn't it, and nothing runs on iOS or with touch. A
canvas grid is exactly where engines differ (font metrics, subpixel text,
scroll events, the pixel-ratio handling), so report what you find.

## Server-side and CSP

- **Strict Content-Security-Policy.** spreadiff runs under
  `default-src 'self'; script-src 'self'; style-src 'self'` with no
  `'unsafe-inline'` or `'unsafe-eval'`:
  - there's no `eval`, `new Function` or WebAssembly;
  - the grid's styles are an adopted stylesheet (`adoptedStyleSheets`),
    which CSP doesn't treat as inline CSS;
  - positions are set through the CSSOM, not style attributes.

  `npm run test:hosted` serves both test sites with exactly that policy
  and fails on any violation. Browsers without constructable stylesheets
  (Safari before 16.4) get a `<style>` element instead, which needs
  `'unsafe-inline'` or a hash under such a policy.
- **Bundlers may inline small assets** as `data:` URLs (Vite does by
  default under 4 KB), which `font-src 'self'` refuses. That isn't
  spreadiff, but it bit the viewer: it sets `build.assetsInlineLimit: 0`.
- `DiffGrid` touches `document` only in its constructor and `render`, so
  importing `spreadiff` during server-side rendering is safe; create the
  grid in the browser (`useEffect`, or the React wrapper, which does).
- `spreadiff/react` starts with `'use client'`, so a Next.js App Router
  server component can render `<SpreadsheetDiff>` or `<DiffGrid>`
  directly; it renders on the client.
- The package is ES modules only. On Node 20.19+ and 22.12+, `require()`
  loads it too (the `exports` have a `default` condition), so CommonJS
  code and tools that resolve through `require` work.
- Nothing makes network requests. Files are read and processed where you
  call the library.
