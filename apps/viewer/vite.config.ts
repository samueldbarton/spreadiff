import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';

const lib = (p: string) => fileURLToPath(new URL(`../../packages/spreadiff/src/${p}`, import.meta.url));

export default defineConfig({
  plugins: [react()],
  // Develop against the library source so edits hot-reload.
  resolve: {
    alias: [
      { find: /^spreadiff$/, replacement: lib('index.ts') },
      { find: /^spreadiff\/react$/, replacement: lib('react/index.tsx') },
      { find: /^spreadiff\/core$/, replacement: lib('core/index.ts') },
      { find: /^spreadiff\/lite$/, replacement: lib('lite/index.ts') },
      { find: /^spreadiff\/worker$/, replacement: lib('worker/index.ts') },
    ],
  },
  worker: { format: 'es' },
  build: {
    rollupOptions: { input: { main: fileURLToPath(new URL('index.html', import.meta.url)), vanilla: fileURLToPath(new URL('vanilla.html', import.meta.url)) } },
  },
  base: './',
});
