#!/usr/bin/env node
// Copy `src/styles/` → `dist/styles/` after the Vite build.
//
// Vite's library mode tree-shakes JS modules but does not relocate static
// CSS files. Consumers import these via `@libraz/formulon-cell/styles.css`
// etc., so they need to live under dist/ alongside the JS bundle.
import { cpSync, existsSync, mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const SRC = resolve(here, '../src/styles');
const DST = resolve(here, '../dist/styles');

if (!existsSync(SRC)) {
  console.error(`copy-styles: source missing: ${SRC}`);
  process.exit(1);
}

mkdirSync(DST, { recursive: true });
cpSync(SRC, DST, { recursive: true });
console.log(`copy-styles: ${SRC} → ${DST}`);
