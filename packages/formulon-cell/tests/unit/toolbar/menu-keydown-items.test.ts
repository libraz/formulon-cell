import { describe, expect, it, vi } from 'vitest';
import { focusMenuItem, handleMenuKeydown } from '../../../src/toolbar/menu-a11y.js';

const item = (text: string, className: string): HTMLButtonElement => {
  const el = document.createElement('button');
  el.type = 'button';
  el.className = className;
  el.textContent = text;
  return el;
};

describe('handleMenuKeydown items option', () => {
  it('moves only between the buttons the selector names', () => {
    const menu = document.createElement('div');
    menu.setAttribute('role', 'menu');
    const first = item('First', 'fc-test__item');
    const tool = item('Tool', 'fc-test__tool');
    const second = item('Second', 'fc-test__item');
    menu.append(first, tool, second);
    document.body.append(menu);
    try {
      focusMenuItem(menu, 'first', '.fc-test__item');
      expect(document.activeElement).toBe(first);
      const down = () =>
        handleMenuKeydown(new KeyboardEvent('keydown', { key: 'ArrowDown' }), menu, {
          close: vi.fn(),
          items: '.fc-test__item',
        });
      down();
      expect(document.activeElement).toBe(second);
      down();
      expect(document.activeElement).toBe(first);
    } finally {
      menu.remove();
    }
  });
});
