import { resolve } from 'node:path';
import vue from '@vitejs/plugin-vue';
import { defineConfig } from 'vite';

const corePkg = resolve(__dirname, '../../packages/formulon-cell');
const vuePkg = resolve(__dirname, '../../packages/formulon-cell-vue');

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: {
      '@libraz/formulon-cell/styles.css': `${corePkg}/src/styles/index.css`,
      '@libraz/formulon-cell/styles/paper.css': `${corePkg}/src/styles/theme-paper.css`,
      '@libraz/formulon-cell/styles/ink.css': `${corePkg}/src/styles/theme-ink.css`,
      '@libraz/formulon-cell/styles/contrast.css': `${corePkg}/src/styles/theme-contrast.css`,
      '@libraz/formulon-cell/styles/tokens.css': `${corePkg}/src/styles/tokens.css`,
      '@libraz/formulon-cell-vue': `${vuePkg}/src/index.ts`,
      '@libraz/formulon-cell': `${corePkg}/src/index.ts`,
    },
  },
  server: {
    port: 5175,
    headers: {
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
    },
    fs: {
      allow: ['..', '../..'],
    },
  },
  optimizeDeps: {
    exclude: ['@libraz/formulon-cell', '@libraz/formulon-cell-vue'],
  },
  // formulon's pthread bundle uses top-level await and spawns its workers
  // as ES modules; both require an ES2022+ target on the main thread and
  // the worker pipeline.
  build: {
    target: 'es2022',
  },
  worker: {
    format: 'es',
  },
});
