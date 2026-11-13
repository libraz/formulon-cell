import { afterEach, describe, expect, it } from 'vitest';
import { spreadsheetFailureOf } from '../../../src/api/errors.js';
import { createDialogs } from '../../../src/api/spreadsheet/dialogs.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const dialog = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-prefs:not([hidden])');

describe('Preferences window', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.dispose();
    sheet = undefined;
    document.body.replaceChildren();
  });

  it('shows the zero values setting on the View pane, titled with the pane name', async () => {
    sheet = await mountStubSheet({
      locale: 'ja',
      ui: { platform: 'mac' },
      workbook: await WorkbookHandle.createDefault(),
    });
    const instance = sheet.instance;
    const dialogs = createDialogs(() => instance);
    expect(dialogs.canOpen('preferences')).toBe(true);
    dialogs.open('preferences');

    expect(dialog()?.getAttribute('aria-label')).toBe('表示');
    expect(dialog()?.querySelector('.fc-prefs__header')?.textContent).toBe('表示');
    expect(
      [...(dialog()?.querySelectorAll('[role="tab"]') ?? [])].map((tab) => tab.textContent),
    ).toEqual(['表示']);
    const row = dialog()?.querySelector<HTMLLabelElement>('.fc-prefs__check');
    expect(row?.textContent).toBe('ゼロ表示');
    expect(row?.title).toBe('ゼロ値が格納されているセルに 0 を表示します。');
    const box = row?.querySelector('input') as HTMLInputElement;
    expect(box.checked).toBe(true);

    box.checked = false;
    box.dispatchEvent(new Event('change'));
    expect(instance.store.getState().ui.showZeros).toBe(false);
    expect(instance.workbook.getSheetView(0)?.showZeros).toBe(false);

    dialogs.close('preferences');
    expect(dialog()).toBeNull();
    dialogs.open('preferences');
    expect(dialog()?.querySelector<HTMLInputElement>('.fc-prefs__check input')?.checked).toBe(
      false,
    );
  });

  it('is refused on the default platform, which keeps the setting on its ribbon', async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    const instance = sheet.instance;
    const dialogs = createDialogs(() => instance);
    expect(dialogs.canOpen('preferences')).toBe(false);
    let failure: unknown;
    try {
      dialogs.open('preferences');
    } catch (error) {
      failure = error;
    }
    expect(spreadsheetFailureOf(failure)).toBe('unsupported-platform');
    expect(dialog()).toBeNull();
  });

  it('closes when the platform leaves the Mac profile', async () => {
    sheet = await mountStubSheet({ locale: 'en', ui: { platform: 'mac' } });
    const instance = sheet.instance;
    createDialogs(() => instance).open('preferences');
    expect(dialog()?.getAttribute('aria-label')).toBe('View');
    instance.setUi({ platform: 'default' });
    expect(dialog()).toBeNull();
  });
});
