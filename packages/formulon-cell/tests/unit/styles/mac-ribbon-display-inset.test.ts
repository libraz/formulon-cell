import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/platform-mac.css'), 'utf8').replace(
  /\/\*[\s\S]*?\*\//g,
  '',
);

/** Declarations of the rules whose selector list holds `selector`, later rules winning. */
const decls = (selector: string): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const [, selectors = '', body = ''] of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
    if (!selectors.split(',').some((s) => s.replace(/\s+/g, ' ').trim() === selector)) continue;
    for (const d of body.split(';')) {
      const [name, value] = d.split(/:(.+)/).map((x) => x.trim());
      if (name && value) out[name] = value;
    }
  }
  return out;
};

const SHELL = '[data-fc-platform="mac"] > .fc-tb__ribbon-shell';

// The desktop draws no display-options toggle, so the panel keeps no right-edge reservation.
describe('Mac ribbon panel without a display toggle', () => {
  it('keeps the 8 px margin on both sides and no toggle reservation in the panel', () => {
    const panel = decls(`${SHELL} .fc-tb__ribbon:not([hidden])`);
    expect(panel.margin).toBe('0 8px');
    expect(panel['padding-inline-end']).toBeUndefined();
    expect(panel.padding).toBe('1px 6px');
  });

  it('keeps no toggle reservation in the single-line panel', () => {
    const panel = decls(`${SHELL}.fc-tb__ribbon-shell--singleLine .fc-tb__ribbon:not([hidden])`);
    expect(panel.padding).toBe('3px 6px 2px');
  });

  it('styles no display toggle and reserves no tab rail width for one', () => {
    expect(css).not.toContain('.fc-tb__ribbon-display');
    expect(decls(`${SHELL}.fc-tb__ribbon-shell--tabsOnly .fc-tb__ribbon-tabs`)).toEqual({});
  });
});
