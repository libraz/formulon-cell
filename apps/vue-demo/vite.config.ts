import { resolve } from 'node:path';
import vue from '@vitejs/plugin-vue';
import { defineConfig, mergeConfig } from 'vite';
import { baseConfig } from '../vite.base.js';

// `baseConfig` aliases `@libraz/formulon-cell` to source; we layer the Vue
// adapter aliases on top (including the .vue and toolbar.css subpaths that
// the Vue package publishes alongside its main entry).
const vuePkg = resolve(__dirname, '../../packages/formulon-cell-vue');

export default defineConfig(
  mergeConfig(baseConfig(__dirname), {
    plugins: [vue()],
    resolve: {
      // Prepend Vue-specific subpath aliases so they win over the broad
      // `@libraz/formulon-cell-vue` rewrite at the end of the list.
      alias: [
        {
          find: '@libraz/formulon-cell-vue/toolbar.css',
          replacement: `${vuePkg}/src/spreadsheet-toolbar.css`,
        },
        {
          find: '@libraz/formulon-cell-vue/toolbar.vue',
          replacement: `${vuePkg}/src/SpreadsheetToolbar.vue`,
        },
        { find: '@libraz/formulon-cell-vue', replacement: `${vuePkg}/src/index.ts` },
      ],
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
