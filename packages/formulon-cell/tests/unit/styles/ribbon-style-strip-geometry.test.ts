import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/toolbar/ribbon/pivot-design.css'), 'utf8');

/** Merged declarations of every rule listing `selector` (whitespace-normalised). */
const decls = (selector: string): Record<string, string> => {
  const want = selector.replace(/\s+/g, ' ').trim();
  const out: Record<string, string> = {};
  for (const m of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
    const selectors = (m[1] ?? '')
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .split(',')
      .map((x) => x.replace(/\s+/g, ' ').trim());
    if (!selectors.includes(want)) continue;
    for (const d of (m[2] ?? '').split(';')) {
      const [k, v] = d.split(/:(.+)/).map((x) => x.trim());
      if (k && v) out[k] = v;
    }
  }
  return out;
};

describe('Sparkline Style strip geometry', () => {
  // Excel for Mac, ja: each chip is 80 x 56 pt and the next starts 78 pt on, overlapping by 2.
  it('sizes a chip 80 x 56 and steps chips by 78', () => {
    const chip = decls('.fc-tb__sparkline-quick-chip');
    expect(chip.width).toBe('80px');
    expect(chip.flex).toBe('0 0 80px');
    expect(chip.height).toBe('56px');
    const gap = Number.parseFloat(decls('.fc-tb__sparkline-quick-strip').gap ?? '0');
    const overlap = Number.parseFloat(
      decls('.fc-tb__sparkline-quick-chip + .fc-tb__sparkline-quick-chip')['margin-left'] ?? '0',
    );
    expect(80 + gap + overlap).toBe(78);
  });
});
