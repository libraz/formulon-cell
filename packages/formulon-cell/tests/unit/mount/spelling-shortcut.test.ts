import { afterEach, describe, expect, it } from 'vitest';
import { mutators } from '../../../src/store/store.js';
import { RIBBON_KEYSHORTCUTS } from '../../../src/toolbar/ribbon-model.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { sheetFocusTarget } from '../../test-utils/sheet-focus.js';
import { seedText } from './toolbar/fixtures.js';

const dialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-spelling:not([hidden])');

// F7 opens Spelling on both desktop apps.
describe('F7', () => {
  let sheet: MountedStubSheet;
  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  it.each(['default', 'mac'] as const)('opens Spelling on the %s profile', async (platform) => {
    sheet = await mountStubSheet({
      locale: 'en',
      ...(platform === 'mac' ? { ui: { platform: 'mac' as const } } : {}),
    });
    seedText(sheet, 0, 0, 'teh');
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 0, col: 0 });
    sheetFocusTarget(sheet.instance.host).dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F7', bubbles: true, cancelable: true }),
    );
    expect(dialog()).not.toBeNull();
  });

  it('is declared for the ribbon Spelling button', () => {
    expect(RIBBON_KEYSHORTCUTS.spellingReview).toBe('F7');
  });
});
