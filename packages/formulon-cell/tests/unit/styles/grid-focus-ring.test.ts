import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

// The active-cell frame is the grid's focus indicator; the desktop app draws
// no focus ring around the sheet, so the browser default must not show either.
describe('grid focus ring', () => {
  it('suppresses the default outline on the focused grid', () => {
    const css = readFileSync(join(root, 'src/styles/core/surface/viewbar.css'), 'utf8');
    expect(css).toMatch(
      /\.fc-host__grid:focus,\s*\.fc-host__grid:focus-visible\s*\{\s*outline:\s*none;/,
    );
  });
});
