/** Escape closes the backstage and leaves focus on the File tab that opened it. */
import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { ribbon } from '@libraz/formulon-cell/toolbar';
import { afterEach, describe, expect, it } from 'vitest';
import { makeHost } from './helpers.js';

const live: SpreadsheetInstance[] = [];
afterEach(() => {
  for (const sheet of live.splice(0)) sheet.dispose();
  document.body.replaceChildren();
});

describe('the backstage', () => {
  // The Mac ribbon has no File tab or backstage.
  it.each(['default'] as const)(
    'closes on Escape and focuses the File tab on the %s profile',
    async (platform) => {
      const host = makeHost();
      const sheet = await mount(host, { locale: 'en', ui: { platform }, toolbar: ribbon() });
      live.push(sheet);
      sheet.toolbar?.update({ backstageOpen: true });
      const file = host.querySelector<HTMLElement>('[data-ribbon-tab="file"]');
      if (!file) throw new Error('No File tab.');
      file.focus();
      file.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
      );
      expect(sheet.toolbar?.backstageOpen).toBe(false);
      expect(document.activeElement).toBe(host.querySelector('[data-ribbon-tab="file"]'));
    },
  );
});
