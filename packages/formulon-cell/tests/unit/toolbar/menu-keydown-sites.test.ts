import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const src = resolve(dirname(fileURLToPath(import.meta.url)), '../../../src');

const sources = (dir: string): string[] =>
  readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return sources(path);
    return path.endsWith('.ts') ? [path] : [];
  });

// Files that build a menu and read arrow keys themselves, for a reason the shared handler lacks.
const OWN_KEYS = new Set([
  // Submenus open and close with the left and right arrows.
  'interact/context-menu.ts',
  // The arrows move between the Name Manager's list rows, which are not a menu.
  'interact/named-range-dialog.ts',
]);

describe('menu keyboard handling', () => {
  it('goes through handleMenuKeydown in every menu', () => {
    const own = sources(src)
      .map((path) => [relative(src, path), readFileSync(path, 'utf8')] as const)
      .filter(
        ([path, text]) =>
          path !== 'toolbar/menu-a11y.ts' &&
          /setAttribute\('role', 'menu'\)/.test(text) &&
          text.includes("'ArrowDown'") &&
          !OWN_KEYS.has(path),
      )
      .map(([path]) => path);
    expect(own).toEqual([]);
  });
});
