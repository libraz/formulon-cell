import { resolve } from 'node:path';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

// In monorepo dev we want both `@libraz/formulon-cell` and the React adapter
// to resolve to source so editing TS in `packages/**/src/` shows up
// immediately. The published `exports` maps point at `dist/`; here we
// override with aliases.
const corePkg = resolve(__dirname, '../../packages/formulon-cell');
const reactPkg = resolve(__dirname, '../../packages/formulon-cell-react');

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@libraz/formulon-cell/styles.css': `${corePkg}/src/styles/index.css`,
      '@libraz/formulon-cell/styles/paper.css': `${corePkg}/src/styles/theme-paper.css`,
      '@libraz/formulon-cell/styles/ink.css': `${corePkg}/src/styles/theme-ink.css`,
      '@libraz/formulon-cell/styles/contrast.css': `${corePkg}/src/styles/theme-contrast.css`,
      '@libraz/formulon-cell/styles/tokens.css': `${corePkg}/src/styles/tokens.css`,
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
    exclude: ['@libraz/formulon-cell', '@libraz/formulon-cell-react'],
  },
});
