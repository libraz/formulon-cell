import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const read = (p: string): string => readFileSync(join(root, p), 'utf8');

/** Declarations of the first rule whose selector list is exactly `selector`. */
const rule = (css: string, selector: string): Record<string, string> => {
  const re = /([^{}]+)\{([^{}]*)\}/g;
  for (let m = re.exec(css); m; m = re.exec(css)) {
    const sel = (m[1] ?? '').replace(/\/\*[\s\S]*?\*\//g, '').trim();
    if (sel === selector) {
      return Object.fromEntries(
        (m[2] ?? '')
          .split(';')
          .map((d) => d.split(/:(.+)/).map((x) => x.trim()))
          .filter((d) => d[0] && d[1]),
      );
    }
  }
  throw new Error(`no rule for ${selector}`);
};

describe('status bar narrow-host overflow', () => {
  const css = read('src/styles/core/surface/statusbar.css');

  it('lets the aggregates take only the free space and the left cluster shrink after them', () => {
    expect(rule(css, '.fc-host__statusbar-aggs').flex).toBe('1 1 0');
    const left = rule(css, '.fc-host__statusbar-left');
    expect(left.flex).toBe('0 1 auto');
    expect(left['min-width']).toBe('0');
    expect(left.overflow).toBe('hidden');
    expect(left['text-overflow']).toBe('ellipsis');
    expect(left['white-space']).toBe('nowrap');
  });

  it('never shrinks the zoom cluster', () => {
    expect(rule(css, '.fc-host__statusbar-zoom').flex).toBe('0 0 auto');
  });
});

describe('name box text overflow', () => {
  const css = read('src/styles/core/surface/formulabar.css');

  it('ellipsizes inside the padding that clears the arrow', () => {
    const tag = rule(css, '.fc-host__formulabar-tag');
    expect(tag.padding).toBe('1px 17px 1px 6px');
    expect(tag['text-overflow']).toBe('ellipsis');
    expect(tag.display).not.toBe('inline-flex');
  });
});
