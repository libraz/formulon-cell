import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { upsertDefinedName } from '../../../src/commands/named-ranges.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { followHyperlink } from '../../../src/interact/hyperlink-follow.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

describe('followHyperlink', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let onInvalidTarget: Mock<() => void>;
  let openExternal: Mock<(url: string) => void>;

  const follow = (target: string): boolean =>
    followHyperlink({ store, wb, onInvalidTarget, openExternal }, target);

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    wb.addSheet('Data Sheet');
    store = createSpreadsheetStore();
    onInvalidTarget = vi.fn<() => void>();
    openExternal = vi.fn<(url: string) => void>();
  });

  afterEach(() => wb.dispose());

  it('selects a cell on the active sheet and scrolls it into view', () => {
    expect(follow('#C200')).toBe(true);
    const s = store.getState();
    expect(s.selection.active).toEqual({ sheet: 0, row: 199, col: 2 });
    expect(s.viewport.rowStart + s.viewport.rowCount).toBeGreaterThan(199);
    expect(s.viewport.rowStart).toBeLessThanOrEqual(199);
  });

  it('activates another sheet, quoted names included', () => {
    expect(follow("#'Data Sheet'!B3")).toBe(true);
    const s = store.getState();
    expect(s.data.sheetIndex).toBe(1);
    expect(s.selection.active).toEqual({ sheet: 1, row: 2, col: 1 });
  });

  it('selects the whole target range', () => {
    expect(follow("#'Data Sheet'!B3:D5")).toBe(true);
    expect(store.getState().selection.range).toEqual({ sheet: 1, r0: 2, c0: 1, r1: 4, c1: 3 });
  });

  it('resolves a defined name to its sheet and range', () => {
    expect(upsertDefinedName(wb, 'Totals', "'Data Sheet'!$B$2:$B$4").ok).toBe(true);
    expect(follow('#Totals')).toBe(true);
    const s = store.getState();
    expect(s.data.sheetIndex).toBe(1);
    expect(s.selection.range).toEqual({ sheet: 1, r0: 1, c0: 1, r1: 3, c1: 1 });
  });

  it('reports an unknown target and leaves the selection alone', () => {
    expect(follow('#NoSuchName')).toBe(false);
    expect(follow('#Missing!A1')).toBe(false);
    expect(onInvalidTarget).toHaveBeenCalledTimes(2);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 0 });
    expect(store.getState().data.sheetIndex).toBe(0);
  });

  it('hands safe external URLs to the opener and refuses the rest', () => {
    expect(follow(' https://example.test/a ')).toBe(true);
    expect(openExternal).toHaveBeenCalledWith('https://example.test/a');
    expect(follow('javascript:alert(1)')).toBe(false);
    expect(openExternal).toHaveBeenCalledTimes(1);
    expect(onInvalidTarget).not.toHaveBeenCalled();
  });

  it('uses the supplied sheet activator', () => {
    const activateSheet = vi.fn(() => false);
    expect(followHyperlink({ store, wb, activateSheet }, "#'Data Sheet'!A1")).toBe(false);
    expect(activateSheet).toHaveBeenCalledWith(1);
    expect(store.getState().data.sheetIndex).toBe(0);
  });
});
