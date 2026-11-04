import { describe, expect, it } from 'vitest';
import { createBordersMenu } from '../../../src/toolbar/ribbon/menus/borders.js';
import { toolbarText } from '../../../src/toolbar/ribbon-model.js';

const deps = {
  ribbonText: toolbarText('ja'),
  getBorderColor: () => '#000000',
  onPickColor: () => undefined,
};

// Excel for Mac (ja): the Borders menu opens with a greyed 罫線 heading.
describe('Mac borders menu', () => {
  it('starts with a disabled heading row', () => {
    const menu = createBordersMenu(deps, { mac: true });
    const first = menu.firstElementChild as HTMLElement | null;
    expect(first?.textContent?.trim()).toBe('罫線');
    expect(first?.classList.contains('fc-tb__menu-item')).toBe(false);
    expect(first?.getAttribute('aria-disabled')).toBe('true');
  });

  it('has no heading elsewhere', () => {
    const menu = createBordersMenu(deps);
    expect((menu.firstElementChild as HTMLElement).textContent?.trim()).not.toBe('罫線');
  });
});
