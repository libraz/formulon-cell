import { afterEach, describe, expect, it } from 'vitest';
import { createEditingDropdownDefaults } from '../../../src/mount/dynamic-dropdown-defaults/editing.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const settle = () => new Promise((resolve) => setTimeout(resolve, 0));

const alertText = (): string | null =>
  document.querySelector('[role="alertdialog"] .fc-tb__dlg__message')?.textContent ?? null;

describe('Remove Duplicates result alert', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  const run = async (values: readonly string[], platform?: 'mac') => {
    sheet = await mountStubSheet({ locale: 'en' });
    if (platform) sheet.host.dataset.fcPlatform = platform;
    const { instance } = sheet;
    values.forEach((value, row) => {
      instance.workbook.setText({ sheet: 0, row, col: 0 }, value);
    });
    mutators.replaceCells(instance.store, instance.workbook.cells(0));
    mutators.setActive(instance.store, { sheet: 0, row: 0, col: 0 });
    const done = createEditingDropdownDefaults(instance).applySortMenuAction('dedupe');
    await settle();
    document
      .querySelector('.fc-dedupedlg__columns')
      ?.closest('[role="dialog"]')
      ?.querySelector<HTMLButtonElement>('.fc-fmtdlg__btn--primary')
      ?.click();
    await done;
    await settle();
  };

  it('reports the removed and remaining counts after OK', async () => {
    await run(['Name', 'a', 'b', 'a', 'c', 'b']);
    expect(alertText()).toBe(
      '2 duplicate values found and removed; 3 unique values remain. Note that counts may include empty cells, spaces, etc.',
    );
  });

  it('says none were found, in the title-less Mac alert', async () => {
    await run(['Name', 'a', 'b'], 'mac');
    expect(alertText()).toBe('No duplicate values found.');
    expect(
      document.querySelector('[role="alertdialog"]')?.classList.contains('fc-tb__dlg--alert'),
    ).toBe(true);
  });
});
