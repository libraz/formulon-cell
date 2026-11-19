import { readdirSync, readFileSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const packageDir = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const repoDir = resolve(packageDir, '../..');
const self = fileURLToPath(import.meta.url);

interface EntryManifest {
  readonly mustNotExport?: readonly string[];
}
const manifest = JSON.parse(
  readFileSync(join(repoDir, 'scripts/api-surface/manifest.json'), 'utf8'),
) as { entries: Record<string, EntryManifest> };

/** Core source files that back a published entry point, keyed by path without extension. */
const ENTRY_OF_SOURCE: Record<string, string> = {
  'src/index': '@libraz/formulon-cell',
  'src/charts/index': '@libraz/formulon-cell/charts',
  'src/extensions/index': '@libraz/formulon-cell/extensions',
  'src/toolbar/index': '@libraz/formulon-cell/toolbar',
  'src/adapter/index': '@libraz/formulon-cell/adapter',
};

const SOURCE_FILE = /\.(ts|tsx|mts|mjs|vue)$/;
const listFiles = (dir: string, out: string[] = []): string[] => {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name === 'dist' || entry.name.startsWith('.'))
      continue;
    const path = join(dir, entry.name);
    if (entry.isDirectory()) listFiles(path, out);
    else if (SOURCE_FILE.test(entry.name)) out.push(path);
  }
  return out;
};

const TEST_ROOTS = [
  'packages/formulon-cell/tests',
  'packages/formulon-cell-react/tests',
  'packages/formulon-cell-vue/tests',
  'tests/e2e-shared',
];
/** Everything outside core src that must reach the instance through its public members. */
const CONSUMER_ROOTS = [
  ...TEST_ROOTS,
  'packages/formulon-cell-react/src',
  'packages/formulon-cell-vue/src',
  'apps',
];
const filesUnder = (roots: readonly string[]): string[] =>
  roots.flatMap((root) => listFiles(join(repoDir, root))).filter((file) => file !== self);

const entryOf = (file: string, specifier: string): string | null => {
  if (specifier.startsWith('@libraz/formulon-cell')) return specifier;
  if (!specifier.startsWith('.')) return null;
  const target = relative(packageDir, resolve(dirname(file), specifier)).replace(/\.(js|ts)$/, '');
  return ENTRY_OF_SOURCE[target] ?? null;
};

const NAMED_IMPORT = /import\s+(?:type\s+)?\{([^}]*)\}\s*from\s*['"]([^'"]+)['"]/g;

describe('public surface imports', () => {
  it('imports no name a published entry point is set to drop', () => {
    const hits: string[] = [];
    for (const file of filesUnder(TEST_ROOTS)) {
      for (const [, names, specifier] of readFileSync(file, 'utf8').matchAll(NAMED_IMPORT)) {
        const entry = specifier ? entryOf(file, specifier) : null;
        const dropped = entry ? manifest.entries[entry]?.mustNotExport : undefined;
        if (!dropped || !names) continue;
        for (const raw of names.split(',')) {
          const name = raw
            .trim()
            .replace(/^type\s+/, '')
            .split(/\s+as\s+/)[0];
          if (name && dropped.includes(name))
            hits.push(`${relative(repoDir, file)}: ${name} from ${entry}`);
        }
      }
    }
    expect(hits).toEqual([]);
  });

  it('reaches no legacy instance outside core src', () => {
    const token = ['legacy', 'InstanceOf'].join('');
    const hits = filesUnder(CONSUMER_ROOTS)
      .filter((file) => readFileSync(file, 'utf8').includes(token))
      .map((file) => relative(repoDir, file));
    expect(hits).toEqual([]);
  });
});
