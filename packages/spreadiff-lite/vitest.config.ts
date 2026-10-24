import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vitest/config';

// Test against the library's source, so a change there needs no build first.
export default defineConfig({
  resolve: { alias: [{ find: /^spreadiff\/core$/, replacement: fileURLToPath(new URL('../spreadiff/src/core/index.ts', import.meta.url)) }] },
});
