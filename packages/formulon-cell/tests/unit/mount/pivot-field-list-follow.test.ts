import { afterEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookObjectsPanelHandle } from '../../../src/interact/workbook-objects.js';
import { attachPivotFieldListFollow } from '../../../src/mount/pivot-field-list-follow.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { loadFixture } from '../engine/pivot-excel-parity/fixture.js';

// ja_c0_rows.xlsx: PivotTable1 at A3:C7 on the first sheet.
describe('pivot field list follow', () => {
  const disposers: (() => void)[] = [];
  afterEach(() => {
    for (const d of disposers.splice(0)) d();
  });

  const setup = async (hidden: boolean) => {
    const store = createSpreadsheetStore();
    const wb = await loadFixture('ja_c0_rows.xlsx');
    let open = false;
    const panel = {
      isPivotFieldListOpen: () => open,
      isPivotFieldListHidden: () => hidden,
      openPivotFieldList: vi.fn(() => {
        open = true;
        return true;
      }),
      close: vi.fn(() => {
        open = false;
      }),
    } as unknown as WorkbookObjectsPanelHandle;
    const follow = attachPivotFieldListFollow({
      store,
      getWb: () => wb,
      getWorkbookObjects: () => panel,
    });
    disposers.push(() => {
      follow.detach();
      wb.dispose();
    });
    return { store, panel };
  };

  it('opens the list when the selection moves onto a report', async () => {
    const { store, panel } = await setup(false);
    mutators.setActive(store, { sheet: 0, row: 4, col: 1 });
    expect(panel.openPivotFieldList).toHaveBeenCalledWith(0, 0);
  });

  it('leaves the list closed while Hide Field List is in effect', async () => {
    const { store, panel } = await setup(true);
    mutators.setActive(store, { sheet: 0, row: 4, col: 1 });
    mutators.setActive(store, { sheet: 0, row: 4, col: 0 });
    expect(panel.openPivotFieldList).not.toHaveBeenCalled();
  });
});
