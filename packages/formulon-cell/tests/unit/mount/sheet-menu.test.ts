import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import {
  createSheetMenuButton,
  createSheetMenuColorButton,
  createSheetMenuSeparator,
  createSheetTabButton,
  formatSheetLabel,
  positionSheetMenu,
} from '../../../src/mount/sheet-menu.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

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

  describe('createSheetTabButton', () => {
    it('applies the shared sheet tab button contract', () => {
      const tab = createSheetTabButton({
        index: 2,
        label: 'Q3',
        selected: true,
        tabColor: '#70ad47',
      });

      expect(tab.type).toBe('button');
      expect(tab.className).toBe('fc-host__sheetbar-tab');
      expect(tab.getAttribute('role')).toBe('tab');
      expect(tab.dataset.fcSheetIndex).toBe('2');
      expect(tab.getAttribute('aria-selected')).toBe('true');
      expect(tab.tabIndex).toBe(0);
      expect(tab.textContent).toBe('Q3');
      expect(tab.dataset.fcSheetTabColor).toBe('true');
      expect(tab.style.getPropertyValue('--fc-sheet-tab-color')).toBe('#70ad47');
    });
  });

  describe('createSheetMenuColorButton', () => {
    it('applies the shared color swatch button contract', () => {
      const onClick = vi.fn();
      const button = createSheetMenuColorButton('Tab color', '#70ad47', true, onClick);

      expect(button.type).toBe('button');
      expect(button.className).toBe('fc-sheetmenu__swatch');
      expect(button.getAttribute('role')).toBe('menuitemradio');
      expect(button.getAttribute('aria-label')).toBe('Tab color #70ad47');
      expect(button.getAttribute('aria-checked')).toBe('true');
      expect(button.title).toBe('Tab color #70ad47');
      expect(button.style.getPropertyValue('--fc-sheet-tab-color')).toBe('#70ad47');

      button.click();
      expect(onClick).toHaveBeenCalledTimes(1);
    });
  });

  describe('DOM primitives', () => {
    it('keeps sheet menu buttons on the shared host button primitive', () => {
      const source = readFileSync(join(root, 'src/mount/sheet-menu.ts'), 'utf8');

      expect(source).toContain("import { createHostButton } from './chrome-buttons.js'");
      expect(source).toContain('const button = createHostButton({');
      expect(source).not.toContain("document.createElement('button')");
    });

    it('keeps the sheet tab menu close to Excel 365 desktop menu geometry', () => {
      const css = readFileSync(join(root, 'src/styles/core/app/overlays/sheet-menu.css'), 'utf8');

      expect(css).toMatch(
        /\.fc-sheetmenu\s*\{[\s\S]*?min-width: 184px;[\s\S]*?padding: 5px 0;[\s\S]*?border-radius: 2px;[\s\S]*?box-shadow:/,
      );
      expect(css).toMatch(
        /\.fc-sheetmenu__item\s*\{[\s\S]*?min-height: 25px;[\s\S]*?padding: 3px 12px 3px 28px;[\s\S]*?border-radius: 0;/,
      );
      expect(css).toMatch(/\.fc-sheetmenu__colors\s*\{[\s\S]*?padding: 5px 10px 7px 28px;/);
      expect(css).toMatch(
        /\.fc-sheetmenu__swatches\s*\{[\s\S]*?grid-template-columns: repeat\(4, 18px\);[\s\S]*?gap: 4px;/,
      );
      expect(css).toMatch(/\.fc-sheetmenu__swatch\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 18px;/);
      expect(css).not.toContain('background: var(--fc-accent-soft');
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
