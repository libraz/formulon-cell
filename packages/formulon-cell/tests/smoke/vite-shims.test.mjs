import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');
const configs = [
  'apps/playground/vite.config.ts',
  'apps/react-demo/vite.config.ts',
  'apps/vue-demo/vite.config.ts',
];
const sharedBase = 'apps/vite.base.ts';
const shimPrefix = '`$' + '{nodeShimDir}/';
// Both `find` and `replacement` literals must show up so we know the alias
// was wired into the array (not just declared by name).
const requiredAliasSnippets = [
  `replacement: ${shimPrefix}module.mjs\``,
  `replacement: ${shimPrefix}worker_threads.mjs\``,
  "find: 'node:module'",
  "find: 'module'",
  "find: 'node:worker_threads'",
  "find: 'worker_threads'",
];

describe('formulon browser shims', () => {
  it('shared vite base aliases Node-only modules to browser-safe shims', () => {
    const source = readFileSync(resolve(root, sharedBase), 'utf8');
    expect(source).toContain("const nodeShimDir = resolve(rootDir, '../vite-shims');");
    for (const snippet of requiredAliasSnippets) {
      expect(source, `${sharedBase} is missing ${snippet}`).toContain(snippet);
    }
  });

  it('every demo Vite config inherits the shared base', () => {
    for (const config of configs) {
      const source = readFileSync(resolve(root, config), 'utf8');
      // Each app calls `baseConfig(__dirname)` and merges; that's what pulls
      // the shim aliases in. Loose check on both pieces keeps the smoke
      // test robust against trivial whitespace changes.
      expect(source, `${config} must import baseConfig`).toMatch(/from '\.\.\/vite\.base\.js'/);
      expect(source, `${config} must invoke baseConfig(__dirname)`).toContain(
        'baseConfig(__dirname)',
      );
    }
  });

  it('exports inert browser fallbacks for Node-only modules', async () => {
    const moduleShim = await import(resolve(root, 'apps/vite-shims/module.mjs'));
    const workerThreadsShim = await import(resolve(root, 'apps/vite-shims/worker_threads.mjs'));

    expect(() => moduleShim.createRequire()('fs')).toThrow('createRequire is unavailable');
    expect(workerThreadsShim.workerData).toBeUndefined();
    expect(workerThreadsShim.parentPort).toBeUndefined();
    expect(() => new workerThreadsShim.Worker()).toThrow('worker_threads is unavailable');
  });
});
