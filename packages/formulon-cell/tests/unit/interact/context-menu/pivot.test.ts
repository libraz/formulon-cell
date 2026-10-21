import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { en } from '../../../../src/i18n/strings/en.js';
import { ja } from '../../../../src/i18n/strings/ja.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../../src/store/store.js';
import { fireContextMenu, item, miniItem, newWb, setRange } from './fixtures.js';

type PivotInfo = NonNullable<ReturnType<WorkbookHandle['pivotCellInfo']>>;

const info = (over: Partial<PivotInfo> = {}): PivotInfo => ({
  pivotIndex: 0,
  kind: 1,
  fieldName: 'Region',
  fieldIndex: 0,
  cacheId: 3,
  isValue: false,
  ...over,
});

describe('PivotTable context menu', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    detach = null;
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    document.body.innerHTML = '';
    vi.restoreAllMocks();
  });

  const ids = (): string[] =>
    Array.from(document.querySelectorAll<HTMLElement>('.fc-ctxmenu__item'))
      .filter((el) => !el.closest('.fc-ctxmenu__sub'))
      .map((el) => el.dataset.fcAction ?? el.dataset.fcSubmenu ?? '');

  it('replaces the cell menu on a label cell', () => {
    vi.spyOn(wb, 'pivotCellInfo').mockReturnValue(info());
    const onOpenPivotFieldList = vi.fn();
    detach = attachContextMenu({ host, store, wb, strings: en, onOpenPivotFieldList });
    setRange(store, 3, 0, 3, 0);
    fireContextMenu(host, 10, 10);
    expect(ids()).toEqual([
      'copy',
      'formatCells',
      'pivotRefresh',
      'sort',
      'pivotFieldSettings',
      'pivotShowFieldList',
    ]);
    expect(miniItem('bold')).toBeNull();
    expect(item('pivotFieldSettings')?.textContent).toContain('Field Settings…');
    item('pivotShowFieldList')?.click();
    expect(onOpenPivotFieldList).toHaveBeenCalledWith(0, 0);
  });

  it('labels value cells and omits sort for them', () => {
    vi.spyOn(wb, 'pivotCellInfo').mockReturnValue(
      info({ kind: 3, fieldName: 'Sum of Sales', fieldIndex: -1, isValue: true }),
    );
    detach = attachContextMenu({ host, store, wb, strings: ja, onOpenPivotFieldList: vi.fn() });
    fireContextMenu(host, 10, 10);
    expect(ids()).toEqual([
      'copy',
      'formatCells',
      'pivotRefresh',
      'pivotFieldSettings',
      'pivotShowFieldList',
    ]);
    expect(item('pivotFieldSettings')?.textContent).toContain('値フィールドの設定…');
    expect(item('pivotRefresh')?.textContent).toContain('更新');
  });

  it('omits entries without a backing command', () => {
    vi.spyOn(wb, 'pivotCellInfo').mockReturnValue(info({ cacheId: -1, fieldIndex: -1 }));
    detach = attachContextMenu({ host, store, wb, strings: en });
    fireContextMenu(host, 10, 10);
    expect(ids()).toEqual(['copy', 'formatCells']);
  });

  it('sorts the clicked field', () => {
    vi.spyOn(wb, 'pivotCellInfo').mockReturnValue(info());
    const sort = vi.spyOn(wb, 'setPivotFieldSort').mockReturnValue(true);
    vi.spyOn(wb, 'cells').mockReturnValue([].values() as ReturnType<WorkbookHandle['cells']>);
    detach = attachContextMenu({ host, store, wb, strings: en });
    fireContextMenu(host, 10, 10);
    document
      .querySelector<HTMLButtonElement>('.fc-ctxmenu__item[data-fc-submenu="pivotSortMenu"]')
      ?.click();
    item('pivotSortDesc')?.click();
    expect(sort).toHaveBeenCalledWith(0, 0, 0, 0, '');
  });

  it('keeps the normal menu outside a report', () => {
    vi.spyOn(wb, 'pivotCellInfo').mockReturnValue(null);
    detach = attachContextMenu({ host, store, wb, strings: en });
    fireContextMenu(host, 10, 10);
    expect(item('paste')).not.toBeNull();
    expect(item('pivotRefresh')).toBeNull();
  });
});
