import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { Spreadsheet } from '../../../../src/mount.js';
import { mutators } from '../../../../src/store/store.js';
import { applyMergeAction } from '../../../../src/toolbar/merge-action.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { seedNumber, seedText, stubHelpers, waitFor } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

describe('Mac Review > Check Accessibility menu', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault(),
      locale: 'en',
      ui: { platform: 'mac' },
    });
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const pick = (id: string): void => {
    const tb = Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    tb.setActiveTab('review');
    host
      .querySelector<HTMLButtonElement>('[data-ribbon-command="mac.review.accessibility"]')
      ?.click();
    const item = document.querySelector<HTMLButtonElement>(
      `#menu-mac-review-accessibility [data-ribbon-command="${id}"]`,
    );
    expect(item?.disabled).toBe(false);
    item?.click();
  };

  it('unmerges the selected cells from Unmerge Cells', async () => {
    const { store } = sheet.instance;
    const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 };
    mutators.setRange(store, range);
    await applyMergeAction(
      {
        store,
        workbook: sheet.instance.workbook,
        history: sheet.instance.history,
        strings: sheet.instance.i18n.strings,
        origin: 'ribbon',
        commandId: 'merge',
      },
      'mergeCells',
    );
    expect(store.getState().merges.byAnchor.size).toBe(1);
    pick('mac.review.accessibility.unmerge');
    await waitFor(() => store.getState().merges.byAnchor.size === 0);
  });

  it('opens Create Table on the current region from Format as Table', async () => {
    seedText(sheet, 0, 0, 'Item');
    seedText(sheet, 0, 1, 'Qty');
    seedText(sheet, 1, 0, 'Pen');
    seedNumber(sheet, 1, 1, 3);
    mutators.setActive(sheet.instance.store, { sheet: 0, row: 1, col: 0 });
    pick('mac.review.accessibility.formatTable');
    await waitFor(() => Boolean(document.body.querySelector('.fc-tb__dlg')));
    expect(
      document.body.querySelector<HTMLInputElement>('.fc-tb__dlg input[type="text"]')?.value,
    ).toBe('=$A$1:$B$2');
  });
});
