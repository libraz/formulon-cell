import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

// The desktop keeps Format Painter on after a double-click on its button until Escape or another
// click on the button; a single click paints once.
describe('Format Painter button', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const doubleClick = (button: HTMLElement): void => {
    button.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 1 }));
    button.dispatchEvent(new MouseEvent('click', { bubbles: true, detail: 2 }));
    button.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, detail: 2 }));
  };

  it('stays on after a double-click until the button is clicked again', () => {
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    try {
      const painter = sheet.instance.formatPainter;
      const button = host.querySelector<HTMLElement>('[data-ribbon-command="formatPainter"]');
      expect(button).not.toBeNull();
      if (!painter || !button) return;
      const states: [boolean, boolean][] = [];
      painter.subscribe((active, sticky) => states.push([active, sticky]));

      doubleClick(button);
      expect(painter.isActive()).toBe(true);
      expect(states.at(-1)).toEqual([true, true]);

      button.click();
      expect(painter.isActive()).toBe(false);
    } finally {
      toolbar.dispose();
    }
  });

  it('paints once after a single click', () => {
    const toolbar = Spreadsheet.mountToolbar(host, sheet.instance, { helpers: stubHelpers() });
    try {
      const painter = sheet.instance.formatPainter;
      const states: [boolean, boolean][] = [];
      painter?.subscribe((active, sticky) => states.push([active, sticky]));
      host.querySelector<HTMLElement>('[data-ribbon-command="formatPainter"]')?.click();
      expect(painter?.isActive()).toBe(true);
      expect(states.at(-1)).toEqual([true, false]);
    } finally {
      toolbar.dispose();
    }
  });
});
