import { afterEach, describe, expect, it, vi } from 'vitest';
import { openNewTableStyle } from '../../../src/mount/dynamic-dropdown-defaults/styles.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

vi.setConfig({ testTimeout: 20_000 });

describe('New Table Style element Format on the Mac', () => {
  let sheet: MountedStubSheet;

  afterEach(() => {
    sheet.dispose();
    document.body.replaceChildren();
  });

  // The element's Format Cells takes the desktop element form: Font, Border and Fill only.
  it('opens Format Cells in the style element form', async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    void openNewTableStyle(sheet.instance, 'table');
    await vi.waitFor(() => expect(document.querySelector('.fc-tb__dlg')).toBeTruthy());
    [...document.querySelectorAll<HTMLButtonElement>('.fc-tb__dlg button')]
      .find((button) => button.textContent === '書式')
      ?.click();
    const tabs = [...document.querySelectorAll<HTMLElement>('.fc-fmtdlg__tabs [data-fc-tab]')]
      .filter((tab) => tab.closest('[hidden]') === null)
      .map((tab) => tab.textContent);
    expect(tabs).toEqual(['フォント', '罫線', '塗りつぶし']);
  });
});
