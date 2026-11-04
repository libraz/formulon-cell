import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/core/app/overlays/context-menu.css'), 'utf8');

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

const MAC = '.fc-overlay-portal[data-fc-platform="mac"]';

// Excel for Mac, ja, 100%: the PivotTable menu is set in 12 pt text; a submenu
// label starts 16.5 pt in (31 pt with a check column) and ends 18 pt before the edge.
describe('Mac PivotTable context menu geometry', () => {
  it('sets the whole menu in 12 pt text', () => {
    expect(decls(`${MAC} .fc-ctxmenu[data-fc-menu="pivot"]`)['font-size']).toBe('12px');
  });

  it('insets submenu labels as measured, with no empty hint column', () => {
    const item = decls(`${MAC} .fc-ctxmenu__sub[data-fc-menu="pivot"] .fc-ctxmenu__item`);
    expect(item['padding-inline']).toBe('10.5px 12px');
    expect(item['column-gap']).toBe('0');
    const checked = decls(
      `${MAC} .fc-ctxmenu__sub[data-fc-menu="pivot"]:has([role="menuitemcheckbox"]) .fc-ctxmenu__item`,
    );
    expect(checked['padding-inline-start']).toBe('25px');
  });
});
