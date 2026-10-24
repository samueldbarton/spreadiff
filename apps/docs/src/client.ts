// The docs site's runtime: theme, search, contents, copy buttons, mobile
// navigation, and the live demos (loaded only on pages that have one).
const root = document.documentElement;
const store = {
  get: (k: string) => {
    try {
      return localStorage.getItem(k);
    } catch {
      return null;
    }
  },
  set: (k: string, v: string | null) => {
    try {
      if (v === null) localStorage.removeItem(k);
      else localStorage.setItem(k, v);
    } catch {
      /* private mode: the choice lasts for the page */
    }
  },
};

// ---- Theme: system → light → dark --------------------------------------
const themeButton = document.querySelector<HTMLButtonElement>('[data-theme-toggle]');
const THEMES = ['system', 'light', 'dark'] as const;
const applyTheme = (t: (typeof THEMES)[number]) => {
  if (t === 'system') delete root.dataset.theme;
  else root.dataset.theme = t;
  if (themeButton) {
    themeButton.dataset.mode = t;
    themeButton.setAttribute('aria-label', `Theme: ${t}`);
    themeButton.title = `Theme: ${t}`;
  }
};
let theme = (THEMES as readonly string[]).includes(store.get('sd-docs-theme') ?? '') ? (store.get('sd-docs-theme') as (typeof THEMES)[number]) : 'system';
applyTheme(theme);
themeButton?.addEventListener('click', () => {
  theme = THEMES[(THEMES.indexOf(theme) + 1) % THEMES.length];
  store.set('sd-docs-theme', theme === 'system' ? null : theme);
  applyTheme(theme);
});

// ---- Mobile navigation ----------------------------------------------------
const menu = document.querySelector<HTMLButtonElement>('.menu');
const sidebar = document.getElementById('sidebar');
menu?.addEventListener('click', () => {
  const open = document.body.classList.toggle('nav-open');
  menu.setAttribute('aria-expanded', String(open));
});
sidebar?.addEventListener('click', (e) => {
  if ((e.target as HTMLElement).closest('a')) document.body.classList.remove('nav-open');
});
sidebar?.querySelector('[aria-current="page"]')?.scrollIntoView({ block: 'center' });

// ---- Copy buttons -------------------------------------------------------
const flash = (b: HTMLElement, text: string) => {
  const was = b.textContent;
  b.textContent = text;
  setTimeout(() => (b.textContent = was), 1400);
};
const copy = async (text: string, b: HTMLElement) => {
  try {
    await navigator.clipboard.writeText(text);
    flash(b, 'Copied');
  } catch {
    flash(b, 'Press ⌘C');
  }
};
document.addEventListener('click', (e) => {
  const b = (e.target as HTMLElement).closest<HTMLButtonElement>('button.copy');
  if (b) void copy(b.closest('figure')?.querySelector('pre')?.innerText ?? '', b);
  const page = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-copy-page]');
  if (page) {
    void fetch(page.dataset.copyPage!)
      .then((r) => r.text())
      .then((t) => copy(t, page));
  }
});

// ---- On this page: highlight the section in view ------------------------
const tocLinks = [...document.querySelectorAll<HTMLAnchorElement>('.toc a')];
if (tocLinks.length && 'IntersectionObserver' in window) {
  const targets = tocLinks.map((a) => document.getElementById(decodeURIComponent(a.hash.slice(1)))).filter((x): x is HTMLElement => !!x);
  const visible = new Set<Element>();
  const update = () => {
    const first = targets.find((t) => visible.has(t)) ?? [...targets].reverse().find((t) => t.getBoundingClientRect().top < 100);
    tocLinks.forEach((a) => a.classList.toggle('active', !!first && a.hash === `#${first.id}`));
  };
  const io = new IntersectionObserver((entries) => {
    for (const en of entries) (en.isIntersecting ? visible.add(en.target) : visible.delete(en.target));
    update();
  }, { rootMargin: '-64px 0px -60% 0px' });
  targets.forEach((t) => io.observe(t));
}

// ---- Search (⌘K, Ctrl+K or /) ------------------------------------------
type Entry = { t: string; p?: string; u: string; g: string; x: string };
let index: Entry[] | null = null;
const dialog = document.createElement('dialog');
dialog.className = 'search';
dialog.setAttribute('aria-label', 'Search the docs');
dialog.innerHTML = `<form method="dialog"><div class="search-field"><svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true"><circle cx="11" cy="11" r="7" fill="none" stroke="currentColor" stroke-width="2"/><path d="m20 20-4-4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg><input type="search" placeholder="Search the docs" aria-label="Search" autocomplete="off" spellcheck="false" role="combobox" aria-expanded="true" aria-controls="search-results"><kbd>esc</kbd></div><ul id="search-results" role="listbox"></ul><p class="search-hint">↑ ↓ to move · ↵ to open</p></form>`;
document.body.append(dialog);
const input = dialog.querySelector('input')!;
const list = dialog.querySelector('ul')!;
let active = 0;
let results: Entry[] = [];

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
const mark = (s: string, words: string[]) => {
  let out = esc(s);
  for (const w of words) if (w) out = out.replace(new RegExp(`(${w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi'), '<mark>$1</mark>');
  return out;
};
const snippet = (x: string, words: string[]) => {
  const lower = x.toLowerCase();
  const at = Math.max(0, Math.min(...words.map((w) => lower.indexOf(w)).filter((i) => i >= 0), x.length) - 40);
  return (at ? '…' : '') + x.slice(at, at + 140);
};
function runSearch() {
  const q = input.value.trim().toLowerCase();
  const words = q.split(/\s+/).filter(Boolean);
  if (!index) return;
  if (!words.length) {
    results = index.filter((e) => !e.p).slice(0, 12);
  } else {
    const scored: [number, Entry][] = [];
    for (const e of index) {
      const t = e.t.toLowerCase();
      const x = e.x.toLowerCase();
      let s = 0;
      for (const w of words) {
        if (t === w) s += 12;
        else if (t.startsWith(w)) s += 8;
        else if (t.includes(w)) s += 5;
        else if ((e.p ?? '').toLowerCase().includes(w)) s += 2;
        else if (x.includes(w)) s += 1;
        else {
          s = 0;
          break;
        }
      }
      if (s) scored.push([s + (e.p ? 0 : 1), e]);
    }
    scored.sort((a, b) => b[0] - a[0]);
    results = scored.slice(0, 20).map(([, e]) => e);
  }
  active = 0;
  list.innerHTML = results.length
    ? results
        .map(
          (e, i) =>
            `<li role="option" id="sr-${i}" aria-selected="${i === active}"><a href="${e.u}"><span class="sr-where">${esc(e.p ? `${e.g} › ${e.p}` : e.g)}</span><span class="sr-title">${mark(e.t, words)}</span>${words.length ? `<span class="sr-text">${mark(snippet(e.x, words), words)}</span>` : ''}</a></li>`,
        )
        .join('')
    : `<li class="empty">No results for “${esc(input.value)}”</li>`;
  input.setAttribute('aria-activedescendant', results.length ? 'sr-0' : '');
}
const move = (d: number) => {
  if (!results.length) return;
  active = (active + d + results.length) % results.length;
  list.querySelectorAll('li').forEach((li, i) => li.setAttribute('aria-selected', String(i === active)));
  input.setAttribute('aria-activedescendant', `sr-${active}`);
  list.querySelector(`#sr-${active}`)?.scrollIntoView({ block: 'nearest' });
};
async function openSearch() {
  if (!dialog.open) dialog.showModal();
  input.select();
  index ??= await fetch('search.json').then((r) => r.json()).catch(() => []);
  runSearch();
}
input.addEventListener('input', runSearch);
input.addEventListener('keydown', (e) => {
  if (e.key === 'ArrowDown') (e.preventDefault(), move(1));
  else if (e.key === 'ArrowUp') (e.preventDefault(), move(-1));
  else if (e.key === 'Enter' && results[active]) {
    e.preventDefault();
    location.href = results[active].u;
    dialog.close();
  }
});
dialog.addEventListener('click', (e) => {
  if (e.target === dialog) dialog.close();
});
document.querySelectorAll('[data-search]').forEach((b) => b.addEventListener('click', openSearch));
document.addEventListener('keydown', (e) => {
  const typing = /^(INPUT|TEXTAREA|SELECT)$/.test((e.target as HTMLElement).tagName) || (e.target as HTMLElement).isContentEditable;
  if ((e.key === 'k' && (e.metaKey || e.ctrlKey)) || (e.key === '/' && !typing)) {
    e.preventDefault();
    void openSearch();
  }
});
if (!/Mac|iPhone|iPad/.test(navigator.platform)) document.querySelectorAll('.search-button kbd').forEach((k) => (k.textContent = 'Ctrl K'));

// ---- Live demos -----------------------------------------------------------
const demos = document.querySelectorAll<HTMLElement>('[data-demo]');
if (demos.length) {
  const start = () => import('./demos.js').then((m) => demos.forEach((el) => m.mount(el)));
  if ('IntersectionObserver' in window) {
    const io = new IntersectionObserver((entries) => {
      if (entries.some((e) => e.isIntersecting)) {
        io.disconnect();
        void start();
      }
    }, { rootMargin: '400px' });
    demos.forEach((d) => io.observe(d));
  } else void start();
}
