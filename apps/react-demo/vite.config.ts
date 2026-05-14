import { resolve } from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig, mergeConfig } from 'vite';
import { baseConfig } from '../vite.base.js';

// `baseConfig` aliases `@libraz/formulon-cell` to source; we layer the
// React adapter's source alias on top, plus React-specific dedupe so two
// copies of `react` from hoisting can't end up bundled.
const reactPkg = resolve(__dirname, '../../packages/formulon-cell-react');

export default defineConfig(
  mergeConfig(baseConfig(__dirname), {
    plugins: [react()],
    resolve: {
      dedupe: ['react', 'react-dom'],
      // Prepend the react-package alias so it wins over the broad
      // `@libraz/formulon-cell` rewrite in the base config.
      alias: [{ find: '@libraz/formulon-cell-react', replacement: `${reactPkg}/src/index.ts` }],
    },
    server: { port: 5174 },
    optimizeDeps: {
      exclude: ['@libraz/formulon-cell-react'],
    },
    build: {
      // React demo intentionally bundles the spreadsheet UI + WASM loader in
      // one app shell. Keep Vite's size warning meaningful for accidental
      // growth beyond that expected baseline.
      chunkSizeWarningLimit: 850,
    },
  }),
);
