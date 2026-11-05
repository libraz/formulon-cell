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

  it('fits the condition popup text inside its 22px height', () => {
    expect(block(css, `${panel}__condition-op`)).toContain('padding: 0 6px');
  });

  it('seats the checklist checkboxes 45pt from the panel edge on a 20pt pitch with no rule under select-all', () => {
    const row = block(css, `${panel}__row`);
    expect(row).toContain('padding: 0 8px 0 26px');
    expect(row).toContain('height: 20px');
    const all = block(css, `${panel}__row--all`);
    expect(all).toContain('border-bottom: 0');
    expect(all).toContain('margin-bottom: 0');
  });

  it('draws a chevron on the value box and room for the search magnifier', () => {
    expect(block(css, `${panel}__condition-value`)).toContain('background-image');
    const field = block(css, `${panel}__search-field`);
    expect(field).toContain('position: relative');
    expect(block(css, `${panel}__search-icon`)).toContain('position: absolute');
    expect(block(css, `${panel}__search`)).toContain('padding-left');
  });

  it('keeps the apply and clear buttons on one row at 103 and 113', () => {
    expect(block(css, `${panel}__actions button`)).toContain('padding: 0');
    expect(block(css, `${panel}__apply`)).toContain('width: 103px');
    expect(block(css, `${panel}__clear`)).toContain('width: 113px');
  });

  it('sets the colour, condition and footer controls in the 11px small-control size', () => {
    for (const part of ['__color', '__condition-op', '__auto', '__actions button']) {
      expect(block(css, `${panel}${part}`)).toContain('font-size: 11px');
    }
  });

  it('right-aligns the 193px colour popups to the panel content edge', () => {
    expect(block(css, `${panel}__color`)).toContain('justify-content: space-between');
    expect(block(css, `${panel}__color select`)).toContain('width: 193px');
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
