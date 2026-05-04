#!/usr/bin/env node
// Mirrors a local formulon dist into vendor/. Used during pre-publish development;
// once @libraz/formulon is on npm, this script is replaced by the peer-dep import.
import { copyFileSync, existsSync, mkdirSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const SRC = resolve(here, '../../../../formulon/packages/npm/dist');
const DST = resolve(here, '../vendor/formulon');

// formulon.web.js is a browser-only variant: same factory but with the
// node:* bridge tree-shaken out. Cell currently loads the universal
// `formulon.js` (which warns harmlessly about node:* externals in
// browser bundlers). Mirroring it as well lets consumers opt into the
// thinner build via a custom alias if they want to silence those warnings.
const files = ['formulon.js', 'formulon.web.js', 'formulon.wasm', 'formulon.d.ts'];

if (!existsSync(SRC)) {
  console.error(`[sync-formulon] not found: ${SRC}`);
  console.error('Clone https://github.com/libraz/formulon next to this repo, then re-run.');
  process.exit(1);
}

mkdirSync(DST, { recursive: true });

for (const f of files) {
  const s = resolve(SRC, f);
  const d = resolve(DST, f);
  if (!existsSync(s)) {
    console.error(`[sync-formulon] missing source: ${s}`);
    process.exit(1);
  }
  copyFileSync(s, d);
  const { size } = statSync(d);
  console.log(`[sync-formulon] ${f} (${(size / 1024).toFixed(1)} KB)`);
}
