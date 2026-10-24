#!/usr/bin/env node
// Differential test of the formula engine against LibreOffice.
// Usage: npm run bench:engine [-- count seed]   (needs LibreOffice and Python openpyxl)
// Writes probes to work/ (git-ignored), has LibreOffice recalculate them, then
// evaluates each with spreadiff's engine and prints every difference.
import { execFileSync } from 'node:child_process';
import { mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const WORK = join(HERE, 'work');
mkdirSync(WORK, { recursive: true });
const [count = '1500', seed = '7'] = process.argv.slice(2);
execFileSync('python3', [join(HERE, 'build.py'), count, seed], { cwd: WORK, stdio: 'inherit' });
execFileSync(join(HERE, '../../node_modules/.bin/vite-node'), [join(HERE, 'compare.ts')], { stdio: 'inherit' });
