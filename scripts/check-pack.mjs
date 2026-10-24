// What npm would publish, checked before it does: every `exports` and `bin`
// target is in the tarball, nothing but dist/, the README, the LICENSE and
// package.json is, and spreadiff-lite depends on this spreadiff version.
// Usage: npm run build && node scripts/check-pack.mjs [--tag v1.2.3]
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const tagIndex = process.argv.indexOf('--tag');
const tag = tagIndex > 0 ? process.argv[tagIndex + 1] : undefined;
const problems = [];

const pkgs = {};
for (const name of ['spreadiff', 'spreadiff-lite']) {
  const pkg = JSON.parse(readFileSync(`${ROOT}packages/${name}/package.json`, 'utf8'));
  pkgs[name] = pkg;
  const [{ files, size }] = JSON.parse(execFileSync('npm', ['pack', '--dry-run', '--json', '-w', name], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }));
  const paths = new Set(files.map((f) => f.path));
  const targets = [];
  const walk = (v) => (typeof v === 'string' ? targets.push(v) : v && typeof v === 'object' && Object.values(v).forEach(walk));
  walk(pkg.exports);
  walk(pkg.bin);
  for (const t of targets) if (!paths.has(t.replace(/^\.\//, ''))) problems.push(`${name}: ${t} is in package.json but not in the tarball (run npm run build)`);
  for (const p of paths) if (!/^(dist\/|README\.md$|LICENSE$|package\.json$)/.test(p)) problems.push(`${name}: ${p} would be published`);
  for (const f of ['README.md', 'LICENSE']) if (!paths.has(f)) problems.push(`${name}: no ${f}`);
  for (const f of ['repository', 'license', 'description']) if (!pkg[f]) problems.push(`${name}: no "${f}" in package.json`);
  if (tag && `v${pkg.version}` !== tag) problems.push(`${name}: version ${pkg.version} doesn't match the tag ${tag}`);
  console.log(`  ${name}@${pkg.version}: ${paths.size} files, ${(size / 1024).toFixed(1)} KB packed`);
}
const dep = pkgs['spreadiff-lite'].dependencies?.spreadiff;
if (dep !== pkgs.spreadiff.version) problems.push(`spreadiff-lite depends on spreadiff ${dep}, but spreadiff is ${pkgs.spreadiff.version}`);

if (problems.length) {
  console.error(`\n✗ ${problems.length} packing problem${problems.length === 1 ? '' : 's'}:\n  ${problems.join('\n  ')}`);
  process.exit(1);
}
console.log('\n✓ pack check passed');
