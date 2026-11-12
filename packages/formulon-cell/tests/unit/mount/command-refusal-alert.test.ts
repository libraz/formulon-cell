import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import { createFormulasDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/formulas.js';
import { mutators } from '../../../src/store/store.js';
import { showCommandRefusal } from '../../../src/toolbar/dialogs/command-refusal.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const alert = () => document.querySelector<HTMLElement>('[role="alertdialog"]');
const alertText = () => alert()?.querySelector('.fc-tb__dlg__message')?.textContent ?? null;
const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

afterEach(() => {
  document.body.replaceChildren();
});

describe('command refusal alert', () => {
  it('is the titled alert on the default chrome and the title-less one on the Mac', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    void showCommandRefusal({ host, strings: en, title: 'Remove Duplicates', reason: 'too-large' });
    expect(alertText()).toBe('Selection is too large.');
    expect(alert()?.classList.contains('fc-tb__dlg--alert')).toBe(false);
    document.body.replaceChildren();

    const mac = document.createElement('div');
    mac.dataset.fcPlatform = 'mac';
    document.body.appendChild(mac);
    void showCommandRefusal({
      host: mac,
      strings: ja,
      title: '重複の削除',
      reason: 'engine-failed',
    });
    expect(alertText()).toBe(ja.commandRefusal.engineFailed);
    expect(alert()?.classList.contains('fc-tb__dlg--alert')).toBe(true);
  });
});

describe('Create from Selection refusal', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  it('tells the user when the selection has too many columns to name', async () => {
    sheet = await mountStubSheet({ locale: 'en', workbook: await WorkbookHandle.createDefault() });
    const { instance } = sheet;
    mutators.setRange(instance.store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 16_383 });
    // The action waits for the alert's OK, so it is left running while the alert is checked.
    void createFormulasDropdownDefaults(instance).applyDefinedNameAction('create-top-row');
    await settle();
    expect(alertText()).toBe(en.commandRefusal.tooLarge);
  });
});
