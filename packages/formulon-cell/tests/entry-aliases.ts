import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageDir = fileURLToPath(new URL('..', import.meta.url));
const src = (path: string): string => resolve(packageDir, 'src', path);

/** The package entry points mapped to source. Each alias is anchored, so a
 *  subpath never resolves through the root entry. */
export const entryAliases = [
  { find: /^@libraz\/formulon-cell\/toolbar$/, replacement: src('toolbar/index.ts') },
  { find: /^@libraz\/formulon-cell\/adapter$/, replacement: src('adapter/index.ts') },
  { find: /^@libraz\/formulon-cell\/extensions$/, replacement: src('extensions/index.ts') },
  { find: /^@libraz\/formulon-cell\/charts$/, replacement: src('charts/index.ts') },
  { find: /^@libraz\/formulon-cell\/i18n\/(.+)$/, replacement: src('i18n/$1.ts') },
  { find: /^@libraz\/formulon-cell$/, replacement: src('index.ts') },
];
