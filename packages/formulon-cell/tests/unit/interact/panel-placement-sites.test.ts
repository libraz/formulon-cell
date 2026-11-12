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

// Popovers, menus and tips sit on whole pixels so their text stays crisp; placePanel rounds the
// clamped position, so no site clamps and writes left/top on its own.
describe('panel placement', () => {
  it('goes through placePanel everywhere', () => {
    const owner = join(src, 'interact/overlay-position.ts');
    const bypassing = sources(src)
      .filter(
        (path) => path !== owner && readFileSync(path, 'utf8').includes('clampPanelToViewport('),
      )
      .map((path) => relative(src, path));
    expect(bypassing).toEqual([]);
  });
});
