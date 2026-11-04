import { afterEach, describe, expect, it } from 'vitest';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { projectRibbonActiveState } from '../../../src/toolbar/ribbon-active-state.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Bold and Clear on a protected sheet', () => {
  let sheet: MountedStubSheet | undefined;

  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
  });

  const ribbon = () => {
    const host = document.createElement('div');
    host.innerHTML =
      '<button data-ribbon-command="bold"></button><button data-ribbon-command="clearFormat"></button>';
    const button = (id: string) =>
      host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${id}"]`) as HTMLButtonElement;
    return { host, bold: button('bold'), clear: button('clearFormat') };
  };

  it('greys both on the Mac profile while the active cell is locked, and restores them', async () => {
    // Excel (Mac): 太字 and クリア are disabled on a protected sheet's locked cell.
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault(),
      locale: 'ja',
      ui: { platform: 'mac' },
    });
    const { store } = sheet.instance;
    const { host, bold, clear } = ribbon();
    setProtectedSheet(store, 0, true);
    projectRibbonActiveState(host, sheet.instance);
    for (const button of [bold, clear]) {
      expect(button.getAttribute('aria-disabled')).toBe('true');
      expect(button.dataset.ribbonDisabledReason).toMatch(/^変更しようとしているセルやグラフ/);
    }
    mutators.setCellFormat(store, store.getState().selection.active, { locked: false });
    projectRibbonActiveState(host, sheet.instance);
    for (const button of [bold, clear]) {
      expect(button.getAttribute('aria-disabled')).not.toBe('true');
    }
  });

  it('leaves both enabled off the Mac profile', async () => {
    sheet = await mountStubSheet({ workbook: await WorkbookHandle.createDefault(), locale: 'en' });
    const { host, bold, clear } = ribbon();
    setProtectedSheet(sheet.instance.store, 0, true);
    projectRibbonActiveState(host, sheet.instance);
    for (const button of [bold, clear]) {
      expect(button.getAttribute('aria-disabled')).not.toBe('true');
    }
  });
});
