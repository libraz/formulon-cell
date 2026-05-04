import { resolve } from 'node:path';
import { defineConfig } from 'vite';

// In monorepo dev we want `@libraz/formulon-cell` to resolve to its source,
// not the built `dist/`. The published package.json `exports` map points at
// dist for installed consumers; here we override that with an alias so
// editing TS in `packages/formulon-cell/src/` shows up immediately.
const corePkg = resolve(__dirname, '../../packages/formulon-cell');

export default defineConfig({
  resolve: {
    alias: {
      '@libraz/formulon-cell/styles.css': `${corePkg}/src/styles/index.css`,
      '@libraz/formulon-cell/styles/paper.css': `${corePkg}/src/styles/theme-paper.css`,
      '@libraz/formulon-cell/styles/ink.css': `${corePkg}/src/styles/theme-ink.css`,
      '@libraz/formulon-cell': `${corePkg}/src/index.ts`,
    },
  },
  server: {
    port: 5173,
    // formulon ships a pthread-enabled WASM that uses SharedArrayBuffer.
    // Browsers require crossOriginIsolated context (COOP+COEP) to allow it.
    headers: {
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp',
    },
    fs: {
      // Allow serving the workspace's vendored WASM.
      allow: ['..', '../..'],
    },
  },
  optimizeDeps: {
    // Don't try to pre-bundle the formulon emscripten module — it doesn't
    // play well with esbuild's CJS interop.
    exclude: ['@libraz/formulon-cell'],
  },
});
