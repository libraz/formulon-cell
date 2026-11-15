import { afterEach, describe, expect, it } from 'vitest';
import { RIBBON_KEYSHORTCUTS } from '../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

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

  it('is declared for the ribbon Insert Function buttons', () => {
    expect(RIBBON_KEYSHORTCUTS.fx).toBe('Shift+F3');
    expect(RIBBON_KEYSHORTCUTS.fxInsert).toBe('Shift+F3');
  });
});
