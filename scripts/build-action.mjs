// Bundle the GitHub Action (action/src/main.ts) into one file, action/dist/index.mjs,
// which is committed: actions run from the repository without installing anything.
// Usage: node scripts/build-action.mjs [outfile]   (after npm run build -w spreadiff)
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const ROOT = fileURLToPath(new URL('..', import.meta.url));
const outfile = process.argv[2] ?? `${ROOT}action/dist/index.mjs`;
await esbuild.build({
  entryPoints: [`${ROOT}action/src/main.ts`],
  bundle: true,
  platform: 'node',
  format: 'esm',
  target: 'node20',
  outfile,
  legalComments: 'none',
  logLevel: 'error',
  banner: { js: '// Built by scripts/build-action.mjs from action/src/main.ts. Do not edit.' },
});
