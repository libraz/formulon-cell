import { afterEach, describe, expect, it } from 'vitest';
import { mutators } from '../../../src/store/store.js';
import { RIBBON_KEYSHORTCUTS } from '../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { sheetFocusTarget } from '../../test-utils/sheet-focus.js';

// The desktop shows Insert Function as a dialog; the Mac shows it as the Formula Builder pane.
const dialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>(
    '.fc-fxdialog:not([hidden]), .fc-mac-formula-palette:not([hidden])',
  );

// Shift+F3 opens Insert Function on both desktop apps, as the ribbon button advertises.
describe('Shift+F3', () => {
  let sheet: MountedStubSheet;
  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  it.each(['default', 'mac'] as const)(
    'opens Insert Function on the %s profile',
    async (platform) => {
      sheet = await mountStubSheet({
        locale: 'en',
        ...(platform === 'mac' ? { ui: { platform: 'mac' as const } } : {}),
      });
      expect(dialog()).toBeNull();
      sheet.instance.host.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'F3',
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
      expect(dialog()).not.toBeNull();
    },
  );

  it.each(['default', 'mac'] as const)(
    'hands focus to the sheet when Insert Function closes on the %s profile',
    async (platform) => {
      sheet = await mountStubSheet({
        locale: 'en',
        ...(platform === 'mac' ? { ui: { platform: 'mac' as const } } : {}),
      });
      sheet.instance.host.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'F3',
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
      const opened = dialog();
      if (!opened) throw new Error('Insert Function did not open.');
      // The dialog closes on Escape; the pane, which Escape leaves open, on its close button.
      const paneClose = opened.querySelector<HTMLButtonElement>('[data-action="close"]');
      if (paneClose) paneClose.click();
      else
        document.activeElement?.dispatchEvent(
          new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
        );
      expect(dialog()).toBeNull();
      expect(document.activeElement).toBe(sheetFocusTarget(sheet.instance.host));
    },
  );

  // The pane is docked inside the host: a key on one of its rows is not a key on the grid.
  it.each(['Tab', 'Shift+Tab', 'ArrowDown'])(
    'keeps the Mac pane open and the active cell put when %s is pressed on a function row',
    async (key) => {
      sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' as const } });
      // Away from the sheet's edges, so a grid Tab or Shift+Tab would move the active cell.
      mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 1 });
      const host = sheet.instance.host;
      host.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'F3',
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      );
      const row = dialog()?.querySelector<HTMLElement>('.fc-mac-formula-palette__function');
      if (!row) throw new Error('The Formula Builder lists no function.');
      const before = { ...sheet.instance.store.getState().selection.active };
      row.focus();
      const [modifier, name] = key.includes('+') ? key.split('+') : ['', key];
      row.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: name,
          shiftKey: modifier === 'Shift',
          bubbles: true,
          cancelable: true,
        }),
      );
      expect(dialog()).not.toBeNull();
      expect(sheet.instance.store.getState().selection.active).toEqual(before);
    },
  );

  it('is declared for the ribbon Insert Function buttons', () => {
    expect(RIBBON_KEYSHORTCUTS.fx).toBe('Shift+F3');
    expect(RIBBON_KEYSHORTCUTS.fxInsert).toBe('Shift+F3');
  });
});
