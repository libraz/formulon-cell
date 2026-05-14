import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { defineConfig, mergeConfig } from 'vite';
import { baseConfig } from '../vite.base.js';

// The published `exports` map of `@libraz/formulon-cell` points at `dist/`;
// `baseConfig` overrides that with a source alias so editing TS in
// `packages/formulon-cell/src/` shows up immediately in the playground.

const patchFormulonWorkerOptions = (): void => {
  const file = resolve(
    __dirname,
    '../../packages/formulon-cell/node_modules/@libraz/formulon/dist/formulon.js',
  );
  if (!existsSync(file)) return;
  const before = readFileSync(file, 'utf8');
  const after = before
    .replaceAll('/* -ignore */', '/* @vite-ignore */')
    .replaceAll(
      'new Worker(new URL("formulon.js",import.meta.url),workerOptions)',
      'new Worker(new URL("formulon.js",import.meta.url),/* @vite-ignore */ workerOptions)',
    )
    .replaceAll(
      'new Worker(new URL("formulon.js",import.meta.url),{type:"module",workerData:"em-pthread",name:"em-pthread"})',
      'new Worker(new URL("formulon.js",import.meta.url),/* @vite-ignore */ {type:"module",workerData:"em-pthread",name:"em-pthread"})',
    );
  if (after !== before) writeFileSync(file, after);
};

patchFormulonWorkerOptions();

export default defineConfig(
  mergeConfig(baseConfig(__dirname), {
    server: { port: 5173 },
    build: {
      // Playground bundles the full spreadsheet surface and the WASM loader;
      // warn only when a future change grows materially beyond that baseline.
      chunkSizeWarningLimit: 750,
    },
  }),
);
