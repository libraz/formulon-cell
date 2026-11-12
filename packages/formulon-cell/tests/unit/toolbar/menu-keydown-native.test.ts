import { describe, expect, it, vi } from 'vitest';
import { handleMenuKeydown } from '../../../src/toolbar/menu-a11y.js';

const menuWith = (...labels: string[]): { menu: HTMLElement; items: HTMLButtonElement[] } => {
  const menu = document.createElement('div');
  menu.setAttribute('role', 'menu');
  const items = labels.map((label) => {
    const item = document.createElement('button');
    item.type = 'button';
    item.textContent = label;
    menu.append(item);
    return item;
  });
  document.body.append(menu);
  return { menu, items };
};

const press = (menu: HTMLElement, key: string): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true });
  handleMenuKeydown(event, menu, { close: vi.fn() });
  return event;
};

// Native menus: Return with no highlighted row does nothing, and moving through a menu never
// scrolls the page.
describe('menu keyboard handling', () => {
  it('presses nothing on Return or Space while no item has focus, but still takes the key', () => {
    const { menu, items } = menuWith('Cut', 'Copy');
    const clicked = vi.fn();
    items[0]?.addEventListener('click', clicked);
    try {
      expect(press(menu, 'Enter').defaultPrevented).toBe(true);
      press(menu, ' ');
      expect(clicked).not.toHaveBeenCalled();
      items[0]?.focus();
      press(menu, 'Enter');
      expect(clicked).toHaveBeenCalledTimes(1);
    } finally {
      menu.remove();
    }
  });

  it('moves focus without scrolling the page', () => {
    const { menu, items } = menuWith('Cut', 'Copy');
    const focus = vi.spyOn(items[1] as HTMLButtonElement, 'focus');
    try {
      items[0]?.focus();
      press(menu, 'ArrowDown');
      expect(focus).toHaveBeenCalledWith({ preventScroll: true });
      expect(document.activeElement).toBe(items[1]);
    } finally {
      menu.remove();
    }
  });
});
