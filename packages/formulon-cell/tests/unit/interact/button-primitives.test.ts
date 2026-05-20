import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const sourceRoots = ['src/interact', 'src/mount', 'src/toolbar'];
const buttonCreatePattern = /document\.createElement\((['"])button\1\)/g;

const allowedDirectButtonCreation: Record<string, number> = {
  'src/interact/chip-button.ts': 1,
  'src/interact/dialog-shell.ts': 5,
  'src/mount/chrome-buttons.ts': 1,
  'src/toolbar/dialogs/shell.ts': 2,
  'src/toolbar/ribbon/button.ts': 1,
};

function sourceFiles(dir: string): string[] {
  const out: string[] = [];
  for (const entry of readdirSync(dir)) {
    const full = join(dir, entry);
    const stat = statSync(full);
    if (stat.isDirectory()) out.push(...sourceFiles(full));
    else if (entry.endsWith('.ts')) out.push(full);
  }
  return out;
}

describe('button DOM primitives', () => {
  it('keeps direct button creation inside the approved primitive files', () => {
    const actual: Record<string, number> = {};
    for (const sourceRoot of sourceRoots) {
      for (const file of sourceFiles(join(root, sourceRoot))) {
        const source = readFileSync(file, 'utf8');
        const count = source.match(buttonCreatePattern)?.length ?? 0;
        if (count > 0) actual[relative(root, file)] = count;
      }
    }

    expect(actual).toEqual(allowedDirectButtonCreation);
  });
});
