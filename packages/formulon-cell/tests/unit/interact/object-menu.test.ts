import { afterEach, describe, expect, it, vi } from 'vitest';
import { createObjectMenu } from '../../../src/interact/object-menu.js';

describe('createObjectMenu', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  const setup = () => {
    const host = document.createElement('div');
    const opener = document.createElement('button');
    const outside = document.createElement('input');
    document.body.append(host, opener, outside);
    const menu = createObjectMenu(host);
    const run = vi.fn();
    menu.open(10, 10, [{ id: 'bringToFront', label: 'Bring to Front', run }], {
      label: 'Chart 1',
      restoreFocusTo: opener,
    });
    return { menu, outside };
  };

  it('closes when focus leaves the menu', () => {
    const { menu, outside } = setup();
    expect(menu.isOpen()).toBe(true);
    outside.focus();
    expect(menu.isOpen()).toBe(false);
    expect(document.activeElement).toBe(outside);
  });

  it('stays open while focus remains inside the menu', () => {
    const { menu } = setup();
    document.querySelector<HTMLButtonElement>('.fc-ctxmenu__item')?.focus();
    expect(menu.isOpen()).toBe(true);
    menu.detach();
  });
});
