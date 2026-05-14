import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
  createSheetMenuButton,
  createSheetMenuSeparator,
  formatSheetLabel,
  positionSheetMenu,
} from '../../../src/mount/sheet-menu.js';

describe('mount/sheet-menu', () => {
  describe('formatSheetLabel', () => {
    it('substitutes {name} once', () => {
      expect(formatSheetLabel('Rename "{name}"', 'Q3')).toBe('Rename "Q3"');
    });

    it('returns the template unchanged when {name} is absent', () => {
      expect(formatSheetLabel('Add Sheet', 'Q3')).toBe('Add Sheet');
    });
  });

  describe('createSheetMenuButton', () => {
    let host: HTMLElement;
    beforeEach(() => {
      host = document.createElement('div');
      document.body.appendChild(host);
    });
    afterEach(() => host.remove());

    it('wires click → closeMenu → onClick when enabled', () => {
      const onClick = vi.fn();
      const closeMenu = vi.fn();
      const btn = createSheetMenuButton('Hide', onClick, closeMenu);
      host.appendChild(btn);

      expect(btn.tagName).toBe('BUTTON');
      expect(btn.className).toBe('fc-sheetmenu__item');
      expect(btn.getAttribute('role')).toBe('menuitem');
      expect(btn.disabled).toBe(false);
      expect(btn.textContent).toBe('Hide');

      btn.click();
      expect(closeMenu).toHaveBeenCalledTimes(1);
      expect(onClick).toHaveBeenCalledTimes(1);
      // closeMenu must run before onClick (so the click can't reopen its own
      // menu by stealing focus mid-action).
      expect(closeMenu.mock.invocationCallOrder[0]).toBeLessThan(
        onClick.mock.invocationCallOrder[0] ?? Number.POSITIVE_INFINITY,
      );
    });

    it('honours the disabled flag — no callbacks invoked', () => {
      const onClick = vi.fn();
      const closeMenu = vi.fn();
      const btn = createSheetMenuButton('Delete', onClick, closeMenu, true);
      host.appendChild(btn);
      expect(btn.disabled).toBe(true);

      btn.click();
      expect(closeMenu).not.toHaveBeenCalled();
      expect(onClick).not.toHaveBeenCalled();
    });
  });

  describe('createSheetMenuSeparator', () => {
    it('produces a non-focusable role=separator div', () => {
      const sep = createSheetMenuSeparator();
      expect(sep.tagName).toBe('DIV');
      expect(sep.className).toBe('fc-sheetmenu__sep');
      expect(sep.getAttribute('role')).toBe('separator');
    });
  });

  describe('positionSheetMenu', () => {
    let menu: HTMLDivElement;
    beforeEach(() => {
      menu = document.createElement('div');
      Object.defineProperty(menu, 'offsetWidth', { configurable: true, value: 200 });
      Object.defineProperty(menu, 'offsetHeight', { configurable: true, value: 150 });
      menu.hidden = true;
      document.body.appendChild(menu);
    });
    afterEach(() => menu.remove());

    it('clamps to the inner viewport with 8px padding', () => {
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: 1024 });
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: 768 });

      positionSheetMenu(menu, 100, 200);
      expect(menu.hidden).toBe(false);
      expect(menu.style.left).toBe('100px');
      expect(menu.style.top).toBe('200px');
    });

    it('shifts left when the requested x would overflow', () => {
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: 400 });
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: 400 });

      // 400 (innerW) - 200 (offsetW) - 8 (pad) = 192 ← max left
      positionSheetMenu(menu, 9999, 9999);
      expect(menu.style.left).toBe('192px');
      expect(menu.style.top).toBe('242px');
    });

    it('never goes below the 8px padding floor', () => {
      Object.defineProperty(window, 'innerWidth', { configurable: true, value: 800 });
      Object.defineProperty(window, 'innerHeight', { configurable: true, value: 600 });

      positionSheetMenu(menu, -100, -100);
      expect(menu.style.left).toBe('8px');
      expect(menu.style.top).toBe('8px');
    });
  });
});
