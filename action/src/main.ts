/**
 * GitHub Action: on a pull request, report what changed in each changed
 * spreadsheet (and what it moved downstream) as one comment, kept up to
 * date on later pushes, and in the job summary.
 *
 * Inputs (action.yml): github-token, paths, min-change, max-rows, comment.
 */
import { execFileSync } from 'node:child_process';
import { appendFileSync, readFileSync } from 'node:fs';
import { compare, toMarkdown } from '../../packages/spreadiff-lite/src/index';
import type { Workbook } from '../../packages/spreadiff/src/core/index';

const MARKER = '<!-- spreadiff-report -->';
const LIMIT = 60_000; // GitHub's comment limit is 65,536 characters

const input = (name: string, fallback = '') => process.env[`INPUT_${name.replace(/ /g, '_').toUpperCase()}`]?.trim() || fallback;

/** `**`/`*`/`{a,b}` globs → a RegExp over repository paths. */
export function globs(patterns: string): RegExp {
  const parts = patterns.split(/[\n,]+(?![^{]*\})/).map((p) => p.trim()).filter(Boolean);
  // Wildcards first, as placeholders, so the groups added after aren't read as wildcards.
  const one = (g: string) =>
    g
      .replace(/[.+^$()|[\]\\]/g, '\\$&')
      .replace(/\*\*\//g, '\u0001')
      .replace(/\*\*/g, '\u0002')
      .replace(/\*/g, '\u0003')
      .replace(/\?/g, '\u0004')
      .replace(/\{([^}]*)\}/g, (_, alts: string) => `(?:${alts.split(',').join('|')})`)
      .replace(/\u0001/g, '(?:.*/)?')
      .replace(/\u0002/g, '.*')
      .replace(/\u0003/g, '[^/]*')
      .replace(/\u0004/g, '[^/]');
  return new RegExp(`^(?:${parts.map(one).join('|')})$`, 'i');
}

const git = (args: string[], encoding: 'utf8' | 'buffer' = 'utf8') =>
  execFileSync('git', args, { encoding: encoding === 'utf8' ? 'utf8' : undefined, maxBuffer: 512 * 1024 * 1024, stdio: ['ignore', 'pipe', 'pipe'] }) as unknown as string & Buffer;

/** A file at a commit, or null when it doesn't exist there. */
function blob(sha: string, path: string): Uint8Array | null {
  try {
    return new Uint8Array(git(['show', `${sha}:${path}`], 'buffer'));
  } catch {
    return null;
  }
}

/** The report for every changed spreadsheet between two commits, as Markdown. */
export function reportBetween(base: string, head: string, options: { paths: RegExp; minChange?: number; maxRows?: number }): { markdown: string; files: number; changed: number } {
  const names = git(['diff', '--name-only', '--no-renames', `${base}...${head}`]).split('\n').filter((p) => p && options.paths.test(p));
  const sections: string[] = [];
  let changed = 0;
  for (const path of names) {
    const empty = (): Workbook => ({ name: path, sheets: [] });
    const a = blob(base, path);
    const b = blob(head, path);
    try {
      const report = compare(a ? { name: path, data: a } : empty(), b ? { name: path, data: b } : empty(), {
        diff: options.minChange ? { relativeTolerance: options.minChange } : undefined,
      });
      report.old = a ? `${path} (base)` : `${path} (new file)`;
      report.new = b ? `${path} (head)` : `${path} (deleted)`;
      const n = report.edits.length + report.downstream.length + report.structure.length + report.objects.length;
      if (n) changed++;
      sections.push(toMarkdown(report, { headingLevel: 3, maxRows: options.maxRows }));
    } catch (e) {
      sections.push(`### ${path}\n\nCouldn't compare: ${(e as Error).message}\n`);
    }
  }
  const head1 = names.length
    ? `## Spreadsheet changes\n\n${changed} of ${names.length} changed spreadsheet${names.length === 1 ? '' : 's'} differ in content.\n\n`
    : '## Spreadsheet changes\n\nNo spreadsheets changed.\n';
  let markdown = head1 + sections.join('\n');
  if (markdown.length > LIMIT) markdown = markdown.slice(0, LIMIT) + '\n\n…cut to fit a comment. The job summary has it all.\n';
  return { markdown, files: names.length, changed };
}

async function github(path: string, token: string, init: { method?: string; body?: unknown } = {}): Promise<unknown> {
  const api = process.env.GITHUB_API_URL ?? 'https://api.github.com';
  const res = await fetch(`${api}${path}`, {
    method: init.method ?? 'GET',
    headers: { authorization: `Bearer ${token}`, accept: 'application/vnd.github+json', 'x-github-api-version': '2022-11-28', 'content-type': 'application/json' },
    body: init.body === undefined ? undefined : JSON.stringify(init.body),
  });
  if (!res.ok) throw new Error(`GitHub API ${init.method ?? 'GET'} ${path}: ${res.status} ${await res.text()}`);
  return res.status === 204 ? null : res.json();
}

/** Create the report comment, or update the one from an earlier run. */
async function upsertComment(repo: string, pr: number, token: string, body: string): Promise<void> {
  const text = `${MARKER}\n${body}`;
  for (let page = 1; page < 50; page++) {
    const comments = (await github(`/repos/${repo}/issues/${pr}/comments?per_page=100&page=${page}`, token)) as { id: number; body?: string }[];
    const mine = comments.find((c) => c.body?.startsWith(MARKER));
    if (mine) {
      await github(`/repos/${repo}/issues/comments/${mine.id}`, token, { method: 'PATCH', body: { body: text } });
      return;
    }
    if (comments.length < 100) break;
  }
  await github(`/repos/${repo}/issues/${pr}/comments`, token, { method: 'POST', body: { body: text } });
}

async function run(): Promise<void> {
  const event = JSON.parse(readFileSync(process.env.GITHUB_EVENT_PATH!, 'utf8'));
  const pr = event.pull_request;
  const base = input('base', pr?.base?.sha ?? event.before);
  const head = input('head', pr?.head?.sha ?? process.env.GITHUB_SHA ?? 'HEAD');
  if (!base) throw new Error('No base commit: run on pull_request, or set the base input');
  const pct = input('min-change');
  const minChange = pct ? (pct.endsWith('%') ? Number(pct.slice(0, -1)) / 100 : Number(pct)) : undefined;
  const { markdown, files, changed } = reportBetween(base, head, {
    paths: globs(input('paths', '**/*.{xlsx,xlsm,xlsb,xls,ods,csv,tsv}')),
    minChange,
    maxRows: Number(input('max-rows', '25')),
  });
  if (process.env.GITHUB_STEP_SUMMARY) appendFileSync(process.env.GITHUB_STEP_SUMMARY, markdown + '\n');
  if (process.env.GITHUB_OUTPUT) appendFileSync(process.env.GITHUB_OUTPUT, `files=${files}\nchanged=${changed}\n`);
  const token = input('github-token');
  const comment = input('comment', 'true') !== 'false';
  if (process.env.SPREADIFF_DRY_RUN) {
    process.stdout.write(markdown);
    return;
  }
  if (comment && pr && files && token) await upsertComment(process.env.GITHUB_REPOSITORY!, pr.number, token, markdown);
  else console.log(markdown);
}

if (process.env.GITHUB_ACTIONS || process.env.SPREADIFF_DRY_RUN) {
  run().catch((e) => {
    console.log(`::error::${(e as Error).message}`);
    process.exitCode = 1;
  });
}
