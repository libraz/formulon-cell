import { describe, expect, it, vi } from 'vitest';
import { focusMenuItem, handleMenuKeydown } from '../../../src/toolbar/menu-a11y.js';

const ROW = 25;
const VISIBLE = 100;

/** A menu of `count` rows in a 100px overflow box; jsdom has no layout, so geometry is stubbed
 *  from the menu's scrollTop. */
const scrollingMenu = (count: number): { menu: HTMLElement; items: HTMLButtonElement[] } => {
  const menu = document.createElement('div');
  menu.setAttribute('role', 'menu');
  Object.defineProperty(menu, 'clientHeight', { value: VISIBLE });
  Object.defineProperty(menu, 'scrollHeight', { value: count * ROW });
  menu.getBoundingClientRect = () => new DOMRect(0, 0, 100, VISIBLE);
  const items = Array.from({ length: count }, (_, idx) => {
    const item = document.createElement('button');
    item.type = 'button';
    item.textContent = `Item ${idx}`;
    item.getBoundingClientRect = () => new DOMRect(0, idx * ROW - menu.scrollTop, 100, ROW);
    menu.append(item);
    return item;
  });
  document.body.append(menu);
  return { menu, items };
};

// The focus call skips page scrolling, so the menu scrolls its own box to the focused row.
describe('menu focus keeps the item in view', () => {
  it('scrolls the menu down to End and back up to Home', () => {
    const { menu, items } = scrollingMenu(10);
    try {
      focusMenuItem(menu, 'first');
      handleMenuKeydown(new KeyboardEvent('keydown', { key: 'End' }), menu, { close: vi.fn() });
      expect(document.activeElement).toBe(items[9]);
      expect(menu.scrollTop).toBe(10 * ROW - VISIBLE);
      handleMenuKeydown(new KeyboardEvent('keydown', { key: 'Home' }), menu, { close: vi.fn() });
      expect(document.activeElement).toBe(items[0]);
      expect(menu.scrollTop).toBe(0);
    } finally {
      menu.remove();
    }
  });

  it('leaves the scroll alone while the item is already visible', () => {
    const { menu } = scrollingMenu(10);
    try {
      focusMenuItem(menu, 2);
      expect(menu.scrollTop).toBe(0);
    } finally {
      menu.remove();
    }
  });
});
