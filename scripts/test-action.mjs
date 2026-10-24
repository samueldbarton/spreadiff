// Test the GitHub Action as GitHub runs it: the committed bundle, in a real
// git repository with a pull-request event, against a fake GitHub API.
// Checks that the bundle is up to date, the report, the job summary and
// outputs, and that the comment is created once and then updated.
import { execFileSync, spawnSync } from 'node:child_process';
import { createServer } from 'node:http';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const failures = [];
const check = (name, ok, detail = '') => {
  console.log(`  ${ok ? '✓' : '✗'} ${name}${ok ? '' : `\n${detail}`}`);
  if (!ok) failures.push(name);
};

const dir = mkdtempSync(join(tmpdir(), 'spreadiff-action-'));
execFileSync('npm', ['run', 'build', '-w', 'spreadiff'], { cwd: ROOT, stdio: 'ignore' });
execFileSync(process.execPath, [join(ROOT, 'scripts/build-action.mjs'), join(dir, 'fresh.mjs')], { cwd: ROOT, stdio: 'inherit' });
const bundle = join(ROOT, 'action/dist/index.mjs');
check('the committed bundle is up to date (else: npm run build:action)', readFileSync(join(dir, 'fresh.mjs'), 'utf8') === readFileSync(bundle, 'utf8'));

// A repository with a pull request: base and head commits.
const repo = join(dir, 'repo');
execFileSync('mkdir', ['-p', repo]);
const git = (...args) => execFileSync('git', args, { cwd: repo, encoding: 'utf8' }).trim();
git('init', '-q', '-b', 'main');
git('config', 'user.email', 'test@example.com');
git('config', 'user.name', 'test');
const fixture = (f) => readFileSync(join(ROOT, 'test/hosted/fixtures/commission', f));
writeFileSync(join(repo, 'commission.xlsx'), fixture('old.xlsx'));
writeFileSync(join(repo, 'prices.csv'), 'id,price\n1,10\n2,20\n');
writeFileSync(join(repo, 'notes.txt'), 'a\n');
git('add', '.');
git('commit', '-qm', 'base');
const base = git('rev-parse', 'HEAD');
writeFileSync(join(repo, 'commission.xlsx'), fixture('new.xlsx'));
writeFileSync(join(repo, 'prices.csv'), 'id,price\n1,12\n2,20\n');
writeFileSync(join(repo, 'added.csv'), 'id\n1\n');
writeFileSync(join(repo, 'notes.txt'), 'b\n');
git('add', '.');
git('commit', '-qm', 'head');
const head = git('rev-parse', 'HEAD');
const eventPath = join(dir, 'event.json');
writeFileSync(eventPath, JSON.stringify({ pull_request: { number: 7, base: { sha: base }, head: { sha: head } } }));

const env = (extra) => ({ ...process.env, GITHUB_EVENT_PATH: eventPath, GITHUB_REPOSITORY: 'o/r', 'INPUT_GITHUB-TOKEN': 'secret-token', ...extra });
const summary = join(dir, 'summary.md');
const output = join(dir, 'output.txt');
writeFileSync(summary, '');
writeFileSync(output, '');
const dry = spawnSync(process.execPath, [bundle], { cwd: repo, encoding: 'utf8', env: env({ SPREADIFF_DRY_RUN: '1', GITHUB_STEP_SUMMARY: summary, GITHUB_OUTPUT: output }) });
check('it reports each changed spreadsheet, not other files', dry.status === 0 && dry.stdout.includes('3 of 3 changed spreadsheets differ') && dry.stdout.includes('### commission.xlsx (base) → commission.xlsx (head)') && dry.stdout.includes('### prices.csv (base)') && dry.stdout.includes('added.csv (new file)') && !dry.stdout.includes('notes.txt'), dry.stdout + dry.stderr);
check('the workbook report traces downstream results', /moved \*\*\d+ results? downstream\*\*/.test(dry.stdout), dry.stdout.slice(0, 1500));
check('job summary and outputs are written', readFileSync(summary, 'utf8').includes('## Spreadsheet changes') && readFileSync(output, 'utf8') === 'files=3\nchanged=3\n', readFileSync(output, 'utf8'));
const only = spawnSync(process.execPath, [bundle], { cwd: repo, encoding: 'utf8', env: env({ SPREADIFF_DRY_RUN: '1', INPUT_PATHS: '*.xlsx' }) });
check('paths limits the files', only.stdout.includes('1 of 1 changed spreadsheet differ'), only.stdout.slice(0, 400));

// The comment: created on the first run, updated on the next.
const requests = [];
const comments = [];
const server = createServer((req, res) => {
  let body = '';
  req.on('data', (c) => (body += c));
  req.on('end', () => {
    requests.push({ method: req.method, url: req.url, auth: req.headers.authorization, body: body && JSON.parse(body) });
    res.setHeader('content-type', 'application/json');
    if (req.method === 'GET') return res.end(JSON.stringify(comments));
    if (req.method === 'POST') {
      const c = { id: 99, body: JSON.parse(body).body };
      comments.push(c);
      return res.end(JSON.stringify(c));
    }
    if (req.method === 'PATCH') {
      comments[0].body = JSON.parse(body).body;
      return res.end(JSON.stringify(comments[0]));
    }
    res.statusCode = 404;
    res.end('{}');
  });
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const api = `http://127.0.0.1:${server.address().port}`;
// spawnSync blocks the event loop the server needs; run the action asynchronously instead.
const run = () =>
  new Promise((resolve) => {
    import('node:child_process').then(({ execFile }) => {
      execFile(process.execPath, [bundle], { cwd: repo, env: env({ GITHUB_ACTIONS: 'true', GITHUB_API_URL: api }) }, (err, stdout, stderr) => resolve({ status: err ? err.code ?? 1 : 0, stdout, stderr }));
    });
  });
const first = await run();
const second = await run();
server.close();
const methods = requests.map((r) => r.method).join(' ');
check('first run posts a comment, the next updates it', first.status === 0 && second.status === 0 && methods === 'GET POST GET PATCH', `${methods}\n${first.stdout}${first.stderr}${second.stderr}`);
check('the comment is marked and authorized', comments.length === 1 && comments[0].body.startsWith('<!-- spreadiff-report -->') && requests.every((r) => r.auth === 'Bearer secret-token') && requests[1].url === '/repos/o/r/issues/7/comments', JSON.stringify(requests.map((r) => r.url)));

rmSync(dir, { recursive: true, force: true });
if (failures.length) {
  console.error(`\n✗ ${failures.length} action check${failures.length === 1 ? '' : 's'} failed`);
  process.exit(1);
}
console.log('\n✓ action test passed');
