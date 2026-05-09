#!/usr/bin/env node
// Mirrors a local formulon dist into vendor/. Used during pre-publish development;
// once @libraz/formulon is on npm, this script is replaced by the peer-dep import.
import { copyFileSync, existsSync, mkdirSync, rmSync, statSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const SRC = resolve(here, '../../../../formulon/packages/npm/dist');
const DST = resolve(here, '../vendor/formulon');

const files = ['formulon.js', 'formulon.wasm', 'formulon.d.ts'];
// `formulon.web.js` and its types were a browser-only variant that has been
// retired upstream. Sweep them out of vendor/ on every sync so stale copies
// from older builds don't hang around and risk being imported by accident.
const stale = ['formulon.web.js', 'formulon.web.d.ts'];

if (!existsSync(SRC)) {
  process.stderr.write(`[sync-formulon] not found: ${SRC}\n`);
  process.stderr.write(
    'Clone https://github.com/libraz/formulon next to this repo, then re-run.\n',
  );
  process.exit(1);
}

mkdirSync(DST, { recursive: true });

for (const f of files) {
  const s = resolve(SRC, f);
  const d = resolve(DST, f);
  if (!existsSync(s)) {
    process.stderr.write(`[sync-formulon] missing source: ${s}\n`);
    process.exit(1);
  }
  copyFileSync(s, d);
  const { size } = statSync(d);
  process.stdout.write(`[sync-formulon] ${f} (${(size / 1024).toFixed(1)} KB)\n`);
}

for (const f of stale) {
  const d = resolve(DST, f);
  if (existsSync(d)) {
    rmSync(d);
    process.stdout.write(`[sync-formulon] removed stale ${f}\n`);
  }
}
