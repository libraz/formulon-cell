import { resolve } from 'node:path';
import type { Alias, UserConfig } from 'vite';

/**
 * Shared Vite settings for the three demo apps (playground, react-demo,
 * vue-demo). Each app's own `vite.config.ts` calls this and then layers its
 * framework plugin, additional aliases, and port. The shared bits are:
 *
 *  - alias `@libraz/formulon-cell` to the workspace source so editing TS
 *    in `packages/formulon-cell/src/` is picked up without a build step
 *  - alias `node:module` / `node:worker_threads` to browser-safe shims so
 *    the formulon WASM bootstrap can be imported in the browser
 *  - COOP+COEP headers so `crossOriginIsolated` holds and pthread WASM works
 *  - allow Vite's dev server to read sibling workspace packages
 *  - opt formulon and its wrappers out of `optimizeDeps` (esbuild's CJS
 *    interop trips over the emscripten module)
 *  - target ES2022 so top-level await and module workers compile straight
 *    through
 *
 * Aliases use array form so the order is preserved: callers can prepend
 * framework-specific subpath aliases (e.g. `…/styles/contrast.css`) that
 * MUST match before the broad `@libraz/formulon-cell` rewrite.
 */
export function baseConfig(rootDir: string): UserConfig {
  const corePkg = resolve(rootDir, '../../packages/formulon-cell');
  const nodeShimDir = resolve(rootDir, '../vite-shims');

  const alias: Alias[] = [
    { find: 'node:module', replacement: `${nodeShimDir}/module.mjs` },
    { find: 'module', replacement: `${nodeShimDir}/module.mjs` },
    { find: 'node:worker_threads', replacement: `${nodeShimDir}/worker_threads.mjs` },
    { find: 'worker_threads', replacement: `${nodeShimDir}/worker_threads.mjs` },
    {
      find: '@libraz/formulon-cell/styles.css',
      replacement: `${corePkg}/src/styles/index.css`,
    },
    {
      find: '@libraz/formulon-cell/styles/paper.css',
      replacement: `${corePkg}/src/styles/theme-paper.css`,
    },
    {
      find: '@libraz/formulon-cell/styles/ink.css',
      replacement: `${corePkg}/src/styles/theme-ink.css`,
    },
    {
      find: '@libraz/formulon-cell/styles/contrast.css',
      replacement: `${corePkg}/src/styles/theme-contrast.css`,
    },
    {
      find: '@libraz/formulon-cell/styles/tokens.css',
      replacement: `${corePkg}/src/styles/tokens.css`,
    },
    {
      find: '@libraz/formulon-cell/styles/toolbar.css',
      replacement: `${corePkg}/src/styles/toolbar.css`,
    },
    // Broad alias goes LAST so the subpath rewrites above win.
    { find: '@libraz/formulon-cell', replacement: `${corePkg}/src/index.ts` },
  ];

  return {
    resolve: { alias },
    server: {
      // formulon ships a pthread-enabled WASM that uses SharedArrayBuffer.
      // Browsers require crossOriginIsolated (COOP+COEP) to allow it.
      headers: {
        'Cross-Origin-Opener-Policy': 'same-origin',
        'Cross-Origin-Embedder-Policy': 'require-corp',
      },
      fs: {
        // Vite's dev server is sandboxed to the app dir by default; widen
        // it so the workspace package source and the engine assets resolve.
        allow: ['..', '../..'],
      },
    },
    optimizeDeps: {
      // Don't pre-bundle the formulon emscripten module — it doesn't play
      // well with esbuild's CJS interop. Apps extend this with framework
      // wrappers via their own `optimizeDeps.exclude` entries.
      exclude: ['@libraz/formulon-cell', '@libraz/formulon'],
    },
    build: {
      target: 'es2022',
    },
    worker: {
      format: 'es',
    },
  };
}
