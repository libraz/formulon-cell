import { expect, it } from 'vitest';
import { setHyperlink } from '../../../../src/commands/hyperlinks.js';
import { createInsertDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/insert.js';
import { mountStubSheet } from '../../../test-utils/mount.js';

it('reports the localized unsupported target through the mounted link action', async () => {
  const sheet = await mountStubSheet({ locale: 'ja' });
  try {
    setHyperlink(
      sheet.instance.store,
      { sheet: 0, row: 0, col: 0 },
      'javascript:void(0)',
      sheet.workbook,
    );
    const action = createInsertDropdownDefaults(sheet.instance).applyLinksAction('open');
    const report = document.body
      .querySelector<HTMLElement>('.fc-tb__dlg__list')
      ?.closest<HTMLElement>('[role="dialog"]');
    expect(report).toBeTruthy();
    expect(report?.textContent).toContain(
      sheet.instance.i18n.strings.ribbonMenu.linkUnsupportedTarget,
    );
    expect(report?.textContent).not.toContain(
      sheet.instance.i18n.strings.ribbonMenu.linkNoHyperlink,
    );
    report?.querySelector<HTMLButtonElement>('.fc-fmtdlg__footer button')?.click();
    await action;
  } finally {
    document.body
      .querySelector<HTMLElement>('[role="dialog"]')
      ?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    sheet.dispose();
  }
});
