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
const shimPrefix = '`$' + '{nodeShimDir}/';
const requiredAliasSnippets = [
  `module: ${shimPrefix}module.mjs\``,
  `'node:module': ${shimPrefix}module.mjs\``,
  `worker_threads: ${shimPrefix}worker_threads.mjs\``,
  `'node:worker_threads': ${shimPrefix}worker_threads.mjs\``,
];

describe('formulon browser shims', () => {
  it('keeps demo Vite configs pointed at browser-safe Node builtin shims', () => {
    for (const config of configs) {
      const source = readFileSync(resolve(root, config), 'utf8');

      expect(source).toContain("const nodeShimDir = resolve(__dirname, '../vite-shims');");
      for (const snippet of requiredAliasSnippets) {
        expect(source, `${config} is missing ${snippet}`).toContain(snippet);
      }
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
