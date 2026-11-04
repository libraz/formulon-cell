import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: 20_000 });

// Excel for Mac leaves Hide Sheet and Delete Sheet enabled on a one-sheet workbook and answers
// the click with an alert; the desktop profile keeps them disabled.
describe('one-sheet workbook guards', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en', workbook: await WorkbookHandle.createDefault() });
    host = document.createElement('div');
    document.body.appendChild(host);
  });
  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const mount = (platform: 'mac' | 'default') => {
    sheet.instance.host.dataset.fcPlatform = platform;
    return Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      platform,
      helpers: stubHelpers(),
    });
  };
  const openDelete = (tb: ReturnType<typeof mount>) => {
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command]');
    const trigger = host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="deleteRows"], [data-ribbon-command="deleteCellsHome"], [data-ribbon-command="mac.home.delete"]',
    );
    tb.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'deleteRows', menuId: 'menu-delete-cells' },
      (trigger ?? button) as HTMLButtonElement,
    );
    return host.querySelector<HTMLButtonElement>('[data-cell-delete="sheet"]');
  };

  it('keeps Delete Sheet enabled on the Mac and answers with an alert', () => {
    const tb = mount('mac');
    const del = openDelete(tb);
    expect(del?.disabled).toBe(false);
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: del });
    tb.dropdownsApi?.dynamicRibbonDropdownClick(event);
    expect(document.body.textContent).toContain(
      sheet.instance.i18n.strings.ribbonMenu.sheetDeleteRequiresAnotherSheet,
    );
    expect(sheet.workbook.sheetCount).toBe(1);
    tb.dispose();
  });

  it('keeps Delete Sheet disabled on the desktop profile', () => {
    const tb = mount('default');
    expect(openDelete(tb)?.disabled).toBe(true);
    tb.dispose();
  });
});
