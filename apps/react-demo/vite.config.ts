import { resolve } from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

// In monorepo dev we want both `@libraz/formulon-cell` and the React adapter
// to resolve to source so editing TS in `packages/**/src/` shows up
// immediately. The published `exports` maps point at `dist/`; here we
// override with aliases.
const corePkg = resolve(__dirname, '../../packages/formulon-cell');
const reactPkg = resolve(__dirname, '../../packages/formulon-cell-react');
const nodeShimDir = resolve(__dirname, '../vite-shims');

export default defineConfig({
  plugins: [react()],
  resolve: {
    dedupe: ['react', 'react-dom'],
    alias: {
      module: `${nodeShimDir}/module.mjs`,
      'node:module': `${nodeShimDir}/module.mjs`,
      worker_threads: `${nodeShimDir}/worker_threads.mjs`,
      'node:worker_threads': `${nodeShimDir}/worker_threads.mjs`,
      '@libraz/formulon-cell/styles.css': `${corePkg}/src/styles/index.css`,
      '@libraz/formulon-cell/styles/paper.css': `${corePkg}/src/styles/theme-paper.css`,
      '@libraz/formulon-cell/styles/ink.css': `${corePkg}/src/styles/theme-ink.css`,
      '@libraz/formulon-cell/styles/contrast.css': `${corePkg}/src/styles/theme-contrast.css`,
      '@libraz/formulon-cell/styles/tokens.css': `${corePkg}/src/styles/tokens.css`,
      '@libraz/formulon-cell/styles/toolbar.css': `${corePkg}/src/styles/toolbar.css`,
      '@libraz/formulon-cell-react': `${reactPkg}/src/index.ts`,
      '@libraz/formulon-cell': `${corePkg}/src/index.ts`,
    },
  },
  server: {
    port: 5174,
    // formulon ships a pthread-enabled WASM that uses SharedArrayBuffer.
    // Browsers require crossOriginIsolated context (COOP+COEP) to allow it.
    headers: {
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
    },
    fs: {
      allow: ['..', '../..'],
    },
  },
  optimizeDeps: {
    exclude: ['@libraz/formulon-cell', '@libraz/formulon-cell-react', '@libraz/formulon'],
  },
  // formulon's pthread bundle uses top-level await and spawns its workers
  // as ES modules; both require an ES2022+ target on the main thread and
  // the worker pipeline.
  build: {
    target: 'es2022',
    // React demo intentionally bundles the spreadsheet UI + WASM loader in
    // one app shell. Keep Vite's size warning meaningful for accidental
    // growth beyond that expected baseline.
    chunkSizeWarningLimit: 850,
  },
  worker: {
    format: 'es',
  },
});
