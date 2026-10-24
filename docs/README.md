# spreadiff documentation

**Start here:** [Getting started](getting-started.md). It goes from two
files to an interactive diff in a few minutes. Without a UI, for an agent
or a CI job, start with [spreadiff-lite](spreadiff-lite.md). Then [Recipes](recipes.md)
has complete solutions to common jobs, and
[Troubleshooting](troubleshooting.md) answers "why does it do that?".

## Guides

| Page | Read it to… |
| --- | --- |
| [Getting started](getting-started.md) | Install, choose an entry point, diff and render your first table |
| [How diffing works](diffing.md) | Understand how columns, rows and cells are matched; tune the options; debug a surprising result |
| [Tracing downstream changes](tracing.md) | Tell edited inputs apart from formula results that moved, and find each result's causes |
| [The DiffGrid](grid.md) | Use the canvas grid: layouts, cell displays, options, methods, keyboard, annotations |
| [spreadiff/html](html.md) | Render a diff as HTML or text, on a server, in CI or in an email |
| [spreadiff-lite](spreadiff-lite.md) | Get a Markdown or JSON report of what changed and what it moved downstream, for agents, CI and scripts; the CLI |
| [Theming](theming.md) | Change colours, fonts and metrics with CSS variables; dark mode |
| [Workers](workers.md) | Parse, diff and trace off the main thread; bundler setup |
| [Accessibility](accessibility.md) | What screen-reader and keyboard users get, how it's tested, and what isn't covered |
| [Recipes](recipes.md) | Complete solutions: compare files a user picks, explain a result, PR comments, a CLI, and more |
| [Troubleshooting](troubleshooting.md) | Symptoms, causes and fixes |

## Reference

| Page | Contents |
| --- | --- |
| [API reference](api.md) | Every export, grouped by task |
| [Browser support](browser-support.md) | Minimum browsers and runtimes, what's tested, CSP and server-side rendering |
| [Performance](performance.md) | Benchmarks, method, and how the numbers could be wrong |
| [Design notes](design.md) | Why it's built this way, the trade-offs, and known gaps |

## How these docs are kept honest

- `npm run docs:check` typechecks every TypeScript block on these pages
  and in the source's JSDoc against the library. It runs the examples
  marked as such and compares what they print with the output shown. It
  checks that every export has an API entry and JSDoc, that every relative
  link and `#anchor` resolves, and that every bundle size quoted is the
  current one.
- `npm run docs:screenshots` regenerates every image in `images/` from the
  current source, so screenshots don't drift from the code.
- Performance figures are ranges across repeated runs, with the method
  and caveats on the same page.
- Correctness figures come from checks you can rerun: `npm run bench:trace`
  (against LibreOffice) and `npm run test:hosted` (in a browser, from
  committed fixtures).
