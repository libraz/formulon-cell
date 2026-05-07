#!/usr/bin/env node
// Mirrors a local formulon dist into vendor/. Used during pre-publish development;
// once @libraz/formulon is on npm, this script is replaced by the peer-dep import.
import { copyFileSync, existsSync, mkdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const SRC = resolve(here, '../../../../formulon/packages/npm/dist');
const DST = resolve(here, '../vendor/formulon');

// formulon.web.js is the browser-only variant used by Cell's UI runtime.
// It has the same embind surface as formulon.js, so the generated
// formulon.d.ts is re-exported as formulon.web.d.ts after copying.
const files = ['formulon.js', 'formulon.web.js', 'formulon.wasm', 'formulon.d.ts'];

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

const webTypes = resolve(DST, 'formulon.web.d.ts');
writeFileSync(webTypes, "export { default } from './formulon.js';\nexport type * from './formulon.js';\n");
process.stdout.write('[sync-formulon] formulon.web.d.ts\n');
