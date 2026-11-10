import { resolve } from 'node:path';
import vue from '@vitejs/plugin-vue';
import { defineConfig, mergeConfig } from 'vite';
import { baseConfig } from '../vite.base.js';

// `baseConfig` aliases `@libraz/formulon-cell` to source; we layer the Vue
// adapter's source alias on top.
const vuePkg = resolve(import.meta.dirname, '../../packages/formulon-cell-vue');

export default defineConfig(
  mergeConfig(baseConfig(import.meta.dirname), {
    plugins: [vue()],
    resolve: {
      // Prepend the vue-package alias so it wins over the broad
      // `@libraz/formulon-cell` rewrite in the base config.
      alias: [{ find: '@libraz/formulon-cell-vue', replacement: `${vuePkg}/src/index.ts` }],
    },
    server: { port: 5175 },
    optimizeDeps: {
      exclude: ['@libraz/formulon-cell-vue'],
    },
    build: {
      // Vue demo intentionally bundles the spreadsheet UI + WASM loader in
      // one app shell. Keep Vite's warning reserved for unexpected future
      // growth.
      chunkSizeWarningLimit: 750,
    },
  }),
);
