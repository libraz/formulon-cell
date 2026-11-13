import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { ja } from '../../../src/i18n/strings.js';
import { attachExternalLinksDialog } from '../../../src/interact/external-links-dialog.js';
import { Spreadsheet } from '../../../src/mount.js';
import { mutators } from '../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';
import { scaledTimeout } from '../../timeouts.js';
import { seedText, stubHelpers } from './toolbar/fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

// The Mac dictionary carries the wording, so the Mac host opens the desktop surfaces.
describe('Mac dialog overlays', () => {
  let sheet: MountedStubSheet | null = null;
  let host: HTMLElement | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
    host?.remove();
    host = null;
    document.body.replaceChildren();
  });

  const mountMac = async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    host = document.createElement('div');
    document.body.appendChild(host);
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    return { sheet, host, tb };
  };

  it('runs the three-step text to columns wizard', async () => {
    const { sheet, host, tb } = await mountMac();
    seedText(sheet, 0, 0, 'a,b');
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    tb.setActiveTab('data');
    host
      .querySelector<HTMLButtonElement>(
        '[data-ribbon-command="mac.data.textToColumns"], [data-ribbon-command="textToColumns"]',
      )
      ?.click();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(
      '区切り位置指定ウィザード - 1 / 3',
    );
    tb.dispose();
  });

  it('opens the desktop Style dialog from New Cell Style', async () => {
    const { host, tb } = await mountMac();
    host.querySelector<HTMLButtonElement>('[data-ribbon-command="cellStyles"]')?.click();
    const button = host.querySelector<HTMLButtonElement>(
      '[data-cell-style-footer="new-cell-style"]',
    );
    expect(button).toBeTruthy();
    const event = new MouseEvent('click', { bubbles: true });
    Object.defineProperty(event, 'target', { value: button });
    expect(tb.dropdownsApi?.dynamicRibbonDropdownClick(event)).toBe(true);
    await Promise.resolve();
    const dialog = document.body.querySelector<HTMLElement>('.fc-tb__dlg');
    expect(dialog?.querySelector('.fc-fmtdlg__header')?.textContent).toBe('セル スタイルの変更');
    tb.dispose();
  });

  it('docks the Workbook Links pane instead of the modal', () => {
    host = document.createElement('div');
    host.className = 'fc-host';
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
    const wb = { getExternalLinks: () => [] } as unknown as WorkbookHandle;
    const handle = attachExternalLinksDialog({ host, getWb: () => wb, strings: ja });
    handle.open();
    const pane = host.querySelector<HTMLElement>('.fc-extlinkpane');
    expect(pane?.hidden).toBe(false);
    expect(pane?.getAttribute('aria-label')).toBe('ブックのリンク');
    expect(document.body.textContent).toContain('外部ブックへのリンクはありません。');
    handle.detach();
  });
});
