import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const read = (file: string): string => readFileSync(join(root, file), 'utf8');

/** The declaration block that follows the first occurrence of `selector`. */
const block = (css: string, selector: string): string => {
  const start = css.indexOf(`${selector} {`);
  expect(start).toBeGreaterThanOrEqual(0);
  return css.slice(start, css.indexOf('}', start));
};

describe('Mac in-cell editor caret', () => {
  it('takes the text colour rather than the accent', () => {
    const css = read('src/styles/platform-mac.css');
    expect(block(css, '[data-fc-platform="mac"] .fc-host__editor')).toContain(
      'caret-color: var(--fc-fg)',
    );
  });
});

describe('Mac column filter panel', () => {
  const css = read('src/styles/platform-mac.css');
  const panel = '.fc-overlay-portal[data-fc-platform="mac"] .fc-filter-dropdown';

  it('is a fixed 283 x 432 window', () => {
    const decl = block(css, panel);
    expect(decl).toContain('width: 283px');
    expect(decl).toContain('height: 432px');
  });

  it('sizes the list, condition popup and value box to the measured controls', () => {
    expect(block(css, `${panel}__list`)).toContain('height: 150px');
    expect(block(css, `${panel}__condition-op`)).toContain('width: 143px');
    expect(block(css, `${panel}__condition-value`)).toContain('width: 90px');
  });
});

describe('Mac context menu panel', () => {
  const css = read('src/styles/core/app/overlays/context-menu.css');

  it('is 287 wide for the cell menu and 224 for the header menus', () => {
    expect(css).toContain('min-width: 287px');
    expect(css).toMatch(
      /fc-ctxmenu:not\(\.fc-ctxmenu__sub\)\[data-fc-menu="col"\] \{\s*min-width: 224px/,
    );
  });

  it('sizes the paste, filter and sort submenus to 234, 249 and 308', () => {
    expect(css).toMatch(/data-fc-submenu="pasteSpecialMenu"\] \{\s*min-width: min\(234px/);
    expect(css).toMatch(/data-fc-submenu="filterMenu"\] \{\s*min-width: min\(249px/);
    expect(css).toMatch(/data-fc-submenu="sortMenu"\] \{\s*min-width: min\(308px/);
  });
});
