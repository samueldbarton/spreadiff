// The site's navigation: every page, in reading order, from the repository's Markdown.
// `slug` is the page's file name (slug.html); `file` is relative to the repository root.
export const NAV = [
  {
    title: 'Get started',
    pages: [
      { slug: 'overview', file: 'docs/README.md', title: 'Overview' },
      { slug: 'getting-started', file: 'docs/getting-started.md' },
      { slug: 'recipes', file: 'docs/recipes.md' },
      { slug: 'troubleshooting', file: 'docs/troubleshooting.md' },
    ],
  },
  {
    title: 'Concepts',
    pages: [
      { slug: 'diffing', file: 'docs/diffing.md' },
      { slug: 'tracing', file: 'docs/tracing.md' },
    ],
  },
  {
    title: 'spreadiff',
    pages: [
      { slug: 'grid', file: 'docs/grid.md' },
      { slug: 'html', file: 'docs/html.md' },
      { slug: 'theming', file: 'docs/theming.md' },
      { slug: 'workers', file: 'docs/workers.md' },
      { slug: 'accessibility', file: 'docs/accessibility.md' },
    ],
  },
  {
    title: 'spreadiff-lite',
    pages: [{ slug: 'spreadiff-lite', file: 'docs/spreadiff-lite.md', title: 'spreadiff-lite' }],
  },
  {
    title: 'Reference',
    pages: [
      { slug: 'api', file: 'docs/api.md' },
      { slug: 'browser-support', file: 'docs/browser-support.md' },
      { slug: 'performance', file: 'docs/performance.md' },
      { slug: 'design', file: 'docs/design.md' },
      { slug: 'changelog', file: 'CHANGELOG.md' },
      { slug: 'contributing', file: 'CONTRIBUTING.md' },
    ],
  },
];

export const REPO = 'https://github.com/samueldbarton/spreadiff';
export const BRANCH = 'main';
