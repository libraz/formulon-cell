// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { findMatchingRanges } from '../../../src/commands/goto-special.js';
import { History } from '../../../src/commands/history.js';
import {
  collapseColGroup,
  collapseRowGroup,
  expandRowGroup,
  groupCols,
  groupRows,
  hideOutlineDetail,
  isRowGroupCollapsed,
  showOutlineDetail,
} from '../../../src/commands/outline.js';
import {
  applyFreezeMode,
  hiddenExtentInSelection,
  hiddenInSelection,
  hideRows,
  showRows,
  showRowsAroundSelection,
} from '../../../src/commands/row-col-layout.js';
import { toggleSplitPanes } from '../../../src/commands/split-panes.js';
import { deleteCols, insertRows } from '../../../src/commands/structure.js';
import { setOutlineSymbolsVisible, setShowFormulas } from '../../../src/commands/view.js';
import { MAX_ROW } from '../../../src/engine/address.js';
import { detectCapabilities } from '../../../src/engine/capabilities.js';
import { summarizeSpreadsheetCompatibility } from '../../../src/engine/compatibility.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import type { Workbook } from '../../../src/engine/types.js';
import { hydrateViewStateFromEngine } from '../../../src/engine/view-state-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { scrollActiveIntoView } from '../../../src/interact/navigation-policy.js';
import { rowHeight } from '../../../src/render/geometry.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const reopen = async (wb: WorkbookHandle, sheet = 0) => {
  const reloaded = await WorkbookHandle.loadBytes(wb.save());
  const store = createSpreadsheetStore();
  mutators.setSheetIndex(store, sheet);
  hydrateLayoutFromEngine(reloaded, store, sheet);
  return { reloaded, store };
};

describe('per-sheet Show Formulas', () => {
  it('saves the flag on the sheet it was set on only', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.addSheet();
    const store = createSpreadsheetStore();
    setShowFormulas(store, true, wb);

    expect((await reopen(wb, 0)).store.getState().ui.showFormulas).toBe(true);
    expect((await reopen(wb, 1)).store.getState().ui.showFormulas).toBe(false);
  });
});

describe('Show Outline Symbols', () => {
  it('round-trips the hidden outline symbols', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    expect(store.getState().ui.showOutlineSymbols).toBe(true);
    setOutlineSymbolsVisible(store, false, wb);

    const { store: opened } = await reopen(wb);
    expect(opened.getState().ui.showOutlineSymbols).toBe(false);
  });
});

describe('outline collapse', () => {
  it('flags the summary row of a collapsed row group and hydrates it back', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    groupRows(store, null, 1, 3, wb);
    collapseRowGroup(store, null, 1, 3, wb);

    const { reloaded, store: opened } = await reopen(wb);
    const rows = reloaded.getRowLayouts(0);
    expect(rows.find((r) => r.row === 4)?.collapsed).toBe(true);
    const layout = opened.getState().layout;
    expect([...layout.hiddenRows].sort()).toEqual([1, 2, 3]);
    expect([...(layout.collapsedRows ?? [])]).toEqual([4]);
    expect(layout.outlineRowGutter).toBeGreaterThan(0);
  });

  it('puts the flag above the group when summary rows sit above the detail', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    mutators.setOutlineSummary(store, 0, { below: false });
    groupRows(store, null, 2, 3, wb);
    collapseRowGroup(store, null, 2, 3, wb);
    expect([...(store.getState().layout.collapsedRows ?? [])]).toEqual([1]);
  });

  it('clears the flag on expand and restores it on undo', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const history = new History();
    groupRows(store, history, 1, 3, wb);
    collapseRowGroup(store, history, 1, 3, wb);
    expandRowGroup(store, history, 1, 3, wb);
    expect(wb.getRowLayouts(0).find((r) => r.row === 4)?.collapsed ?? false).toBe(false);

    history.undo();
    expect(store.getState().layout.collapsedRows?.has(4)).toBe(true);
    expect(wb.getRowLayouts(0).find((r) => r.row === 4)?.collapsed).toBe(true);
    const { store: opened } = await reopen(wb);
    expect(opened.getState().layout.collapsedRows?.has(4)).toBe(true);
  });

  it('moves the flag with the rows when rows are inserted above the group', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    groupRows(store, null, 1, 3, wb);
    collapseRowGroup(store, null, 1, 3, wb);
    insertRows(store, wb, null, 0, 1);
    expect([...(store.getState().layout.collapsedRows ?? [])]).toEqual([5]);
    expandRowGroup(store, null, 2, 4, wb);
    expect(wb.getRowLayouts(0).find((r) => r.row === 5)?.collapsed ?? false).toBe(false);
  });

  it('moves column flags when columns are deleted before the group', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    groupCols(store, null, 2, 4, wb);
    collapseColGroup(store, null, 2, 4, wb);
    deleteCols(store, wb, null, 0, 1);
    expect([...(store.getState().layout.collapsedCols ?? [])]).toEqual([4]);
  });

  it('flags the summary column of a collapsed column group', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    groupCols(store, null, 2, 4, wb);
    collapseColGroup(store, null, 2, 4, wb);

    const { reloaded, store: opened } = await reopen(wb);
    expect(reloaded.getColumnLayouts(0).find((c) => c.first <= 5 && c.last >= 5)?.collapsed).toBe(
      true,
    );
    const layout = opened.getState().layout;
    expect([...layout.hiddenCols].sort()).toEqual([2, 3, 4]);
    expect([...(layout.collapsedCols ?? [])]).toEqual([5]);
    expect(layout.outlineColGutter).toBeGreaterThan(0);
  });
});

describe('split panes', () => {
  const at = (store: ReturnType<typeof createSpreadsheetStore>, row: number, col: number) =>
    mutators.setActive(store, { sheet: 0, row, col });

  const reopenView = async (wb: WorkbookHandle) => {
    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    const store = createSpreadsheetStore();
    hydrateViewStateFromEngine(reloaded, store, { followActiveSheet: true });
    hydrateLayoutFromEngine(reloaded, store, 0);
    return { reloaded, store };
  };

  it('splits above and left of the active cell and saves the split in screen twips', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    wb.attachStore(store);
    at(store, 5, 2);
    toggleSplitPanes(store, wb);

    const layout = store.getState().layout;
    expect([layout.freezeRows, layout.freezeCols, layout.splitPanes]).toEqual([5, 2, true]);
    const view = wb.getSheetView(0);
    // Header plus the band, 15 twips per CSS px: (20 + 5 * 20) and (32 + 2 * 64).
    expect([view?.splitY, view?.splitX]).toEqual([1800, 2400]);
    expect([view?.freezeRows, view?.freezeCols]).toEqual([0, 0]);

    const { store: opened } = await reopenView(wb);
    const reopened = opened.getState().layout;
    expect([reopened.freezeRows, reopened.freezeCols, reopened.splitPanes]).toEqual([5, 2, true]);
  });

  it('splits the window in the middle when the active cell is the top-left cell', async () => {
    const store = createSpreadsheetStore();
    toggleSplitPanes(store);
    const { freezeRows, freezeCols } = store.getState().layout;
    expect([freezeRows, freezeCols]).toEqual([20, 8]);
  });

  it('removes the split when toggled again', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    at(store, 5, 2);
    toggleSplitPanes(store, wb);
    toggleSplitPanes(store, wb);
    expect(store.getState().layout.splitPanes).not.toBe(true);
    expect(store.getState().layout.freezeRows).toBe(0);
    expect([wb.getSheetView(0)?.splitX, wb.getSheetView(0)?.splitY]).toEqual([0, 0]);
  });

  it('freezes at the split, and splits at the freeze', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    at(store, 5, 2);
    toggleSplitPanes(store, wb);
    at(store, 9, 9);
    applyFreezeMode(store, 'panes', wb);
    expect(store.getState().layout).toMatchObject({ freezeRows: 5, freezeCols: 2 });
    expect(store.getState().layout.splitPanes).not.toBe(true);
    expect(wb.getSheetView(0)).toMatchObject({
      freezeRows: 5,
      freezeCols: 2,
      splitX: 0,
      splitY: 0,
    });

    toggleSplitPanes(store, wb);
    expect(store.getState().layout).toMatchObject({
      freezeRows: 5,
      freezeCols: 2,
      splitPanes: true,
    });
    expect(wb.getSheetView(0)).toMatchObject({ freezeRows: 0, freezeCols: 0, splitY: 1800 });
  });

  it('reads a split saved at another zoom back to the same rows', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setSheetZoom(0, 200);
    wb.setSheetSplit(0, 0, (20 + 5 * 20) * 2 * 15);
    const { store } = await reopenView(wb);
    expect(store.getState().layout).toMatchObject({
      freezeRows: 5,
      freezeCols: 0,
      splitPanes: true,
    });
  });

  it('scrolls the top and left panes back to an active cell above or left of them', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 5, 2, { row: 0, col: 0 }, true);
    mutators.scrollBy(store, 30, 10);
    mutators.scrollSplitPane(store, 6, 4);
    scrollActiveIntoView(store, { sheet: 0, row: 1, col: 3 });
    expect(store.getState().layout).toMatchObject({ freezeOriginRow: 1, freezeOriginCol: 3 });
    scrollActiveIntoView(store, { sheet: 0, row: 0, col: 0 });
    expect(store.getState().layout).toMatchObject({ freezeOriginRow: 0, freezeOriginCol: 0 });
  });

  it('leaves a frozen band where it is', () => {
    const store = createSpreadsheetStore();
    mutators.setFreezePanes(store, 5, 0, { row: 3, col: 0 });
    scrollActiveIntoView(store, { sheet: 0, row: 0, col: 0 });
    expect(store.getState().layout.freezeOriginRow).toBe(3);
  });

  it('scrolls the top pane on its own and saves its scroll origin', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    wb.attachStore(store);
    at(store, 5, 0);
    toggleSplitPanes(store, wb);
    mutators.scrollBy(store, 30, 0);
    mutators.scrollSplitPane(store, 4, 0);
    expect(store.getState().layout.freezeOriginRow).toBe(4);
    expect(store.getState().viewport.rowStart).toBe(35);
    mutators.scrollSplitPane(store, 100, 0);
    // The top pane stops where the bottom pane starts.
    expect(store.getState().layout.freezeOriginRow).toBe(30);

    const { store: opened } = await reopenView(wb);
    expect(opened.getState().layout).toMatchObject({ freezeOriginRow: 30, freezeRows: 5 });
    expect(opened.getState().viewport.rowStart).toBe(35);
  });
});

describe('rows hidden by default', () => {
  const hydrated = async (bytes: Uint8Array) => {
    const wb = await WorkbookHandle.loadBytes(bytes);
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    hydrateLayoutFromEngine(wb, store, 0);
    return { wb, store };
  };
  const shown = (store: ReturnType<typeof createSpreadsheetStore>, row: number) =>
    rowHeight(store.getState().layout, row) > 0;

  const zeroHeightBook = async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    wb.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
    wb.setRowHidden(0, 4, false);
    wb.patchSheetFormatDefaults(0, { zeroHeight: true });
    return wb;
  };

  it('shows only the rows that say otherwise, and keeps the setting through a save', async () => {
    const { wb, store } = await hydrated((await zeroHeightBook()).save());
    expect(store.getState().layout.defaultRowsHidden).toBe(true);
    expect([0, 1, 2, 3, 4, 5, 100].map((r) => shown(store, r))).toEqual([
      true,
      false,
      true,
      false,
      true,
      false,
      false,
    ]);
    const again = await hydrated(wb.save());
    expect(again.store.getState().layout.defaultRowsHidden).toBe(true);
    expect(shown(again.store, 4)).toBe(true);
  });

  it('unhides a default-hidden row as an explicitly visible row', async () => {
    const { wb, store } = await hydrated((await zeroHeightBook()).save());
    showRows(store, null, 1, 1, wb);
    expect(shown(store, 1)).toBe(true);
    const { store: opened } = await hydrated(wb.save());
    expect(shown(opened, 1)).toBe(true);
    expect(shown(opened, 3)).toBe(false);
  });

  it('drops the setting when every row past a point is unhidden', async () => {
    const { wb, store } = await hydrated((await zeroHeightBook()).save());
    showRows(store, null, 5, MAX_ROW, wb);
    expect(store.getState().layout.defaultRowsHidden).toBe(false);
    expect([1, 3, 5, 100].map((r) => shown(store, r))).toEqual([false, false, true, true]);
    const { wb: reloaded, store: opened } = await hydrated(wb.save());
    expect(reloaded.sheetRowsHiddenByDefault(0)).toBe(false);
    expect([1, 3, 5, 100].map((r) => shown(opened, r))).toEqual([false, false, true, true]);
  });

  it('undoes a Hide without turning default-hidden rows into visible ones', async () => {
    const { wb, store } = await hydrated((await zeroHeightBook()).save());
    const history = new History();
    hideRows(store, history, 3, 5, wb);
    history.undo();
    expect([3, 4, 5].map((r) => shown(store, r))).toEqual([false, true, false]);
    const { store: opened } = await hydrated(wb.save());
    expect([3, 4, 5].map((r) => shown(opened, r))).toEqual([false, true, false]);
  });

  it('undoes an Unhide to the last row back to rows hidden by default', async () => {
    const { wb, store } = await hydrated((await zeroHeightBook()).save());
    const history = new History();
    showRows(store, history, 5, MAX_ROW, wb);
    history.undo();
    expect(store.getState().layout.defaultRowsHidden).toBe(true);
    const { wb: reloaded, store: opened } = await hydrated(wb.save());
    expect(reloaded.sheetRowsHiddenByDefault(0)).toBe(true);
    expect([0, 1, 2, 3, 4, 5, 100].map((r) => shown(opened, r))).toEqual([
      true,
      false,
      true,
      false,
      true,
      false,
      false,
    ]);
  });

  it('applies an explicit unhide even across more rows than it can list', async () => {
    const { wb, store } = await hydrated((await zeroHeightBook()).save());
    hideRows(store, null, 4, 4, wb);
    showRows(store, null, 0, 200_000, wb);
    expect(store.getState().layout.hiddenRows.has(4)).toBe(false);
    expect(shown(store, 4)).toBe(true);
  });

  it('treats default-hidden rows as hidden in selection queries', async () => {
    const { wb, store } = await hydrated((await zeroHeightBook()).save());
    const layout = () => store.getState().layout;
    expect(hiddenExtentInSelection(layout(), 'row', 0, MAX_ROW)).toEqual({
      first: 1,
      last: MAX_ROW,
    });
    expect(hiddenExtentInSelection(layout(), 'row', 4, 4)).toBeNull();
    expect(hiddenInSelection(layout(), 'row', 0, 6)).toEqual([1, 3, 5, 6]);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 });
    expect(findMatchingRanges(wb, store, 'visible-cells')).toEqual({
      status: 'found',
      ranges: [
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 },
        { sheet: 0, r0: 4, c0: 0, r1: 4, c1: 0 },
      ],
    });
    // Unhide around a visible row reaches out through the default-hidden rows.
    showRowsAroundSelection(store, null, 4, 4, wb);
    expect(layout().defaultRowsHidden).toBe(false);
  });

  it('marks a group collapsed when its rows are hidden by default', async () => {
    const { store } = await hydrated((await zeroHeightBook()).save());
    expect(isRowGroupCollapsed(store.getState().layout, 5, 8)).toBe(true);
    expect(isRowGroupCollapsed(store.getState().layout, 4, 4)).toBe(false);
  });

  it('restates every sheet default when one changes', async () => {
    const wb = await WorkbookHandle.createDefault();
    wb.patchSheetFormatDefaults(0, {
      baseColWidth: 10,
      defaultColWidth: 12,
      hasDefaultColWidth: true,
    });
    wb.patchSheetFormatDefaults(0, { zeroHeight: true });
    const raw = (wb as unknown as { wb: { getSheetFormatDefaults(s: number): unknown } }).wb;
    expect(raw.getSheetFormatDefaults(0)).toMatchObject({
      baseColWidth: 10,
      defaultColWidth: 12,
      hasDefaultColWidth: true,
      zeroHeight: true,
    });
  });
});

describe('compatibility of the view features', () => {
  it('reports each one writable on the default engine', async () => {
    const wb = await WorkbookHandle.createDefault();
    const summary = summarizeSpreadsheetCompatibility(wb);
    for (const id of [
      'show-formulas',
      'outline-symbols',
      'outline-collapse',
      'split-panes',
      'hidden-default-rows',
    ] as const) {
      expect(summary.byId[id].status).toBe('writable');
    }
  });
});

describe('split panes from the Mac desktop app', () => {
  it('reads the split Mac Excel saved below row 5 and right of column B at 200%', async () => {
    // The values a Japanese-locale Mac Excel saved: 20 pt rows, base width 10.
    const wb = await WorkbookHandle.createDefault({ platform: 'mac', locale: 'ja' });
    wb.setSheetZoom(0, 200);
    wb.setSheetSplit(0, 7020, 4880);
    const reloaded = await WorkbookHandle.loadBytes(wb.save());
    reloaded.geometryPlatform = 'mac';
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, screenScale: 0.75, headerColWidth: 26 },
    }));
    hydrateLayoutFromEngine(reloaded, store, 0);
    expect(store.getState().layout).toMatchObject({
      freezeRows: 5,
      freezeCols: 2,
      splitPanes: true,
    });
  });
});

describe('view feature probes', () => {
  it('turns each flag on only when every binding it needs is present', () => {
    const fn = () => undefined;
    const caps = detectCapabilities({
      setSheetShowFormulas: fn,
      setSheetSplit: fn,
      setRowCollapsed: fn,
      getSheetFormatDefaults: fn,
    } as unknown as Workbook);
    expect(caps).toMatchObject({
      sheetShowFormulas: true,
      outlineSymbols: false,
      sheetSplit: true,
      outlineCollapse: false,
      sheetFormatDefaults: false,
    });
  });
});

describe('Hide Detail and Show Detail', () => {
  const range = (r0: number, c0: number, r1: number, c1: number) => ({ sheet: 0, r0, c0, r1, c1 });

  it('collapses the whole group around a partial selection inside it', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    groupRows(store, null, 1, 5, wb);
    hideOutlineDetail(store, null, range(2, 0, 3, 0), wb);
    expect([...store.getState().layout.hiddenRows].sort()).toEqual([1, 2, 3, 4, 5]);
    expect([...(store.getState().layout.collapsedRows ?? [])]).toEqual([6]);
  });

  it('expands from the summary row of a collapsed group', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    groupRows(store, null, 1, 5, wb);
    hideOutlineDetail(store, null, range(2, 0, 2, 0), wb);
    showOutlineDetail(store, null, range(6, 0, 6, 0), wb);
    expect(store.getState().layout.hiddenRows.size).toBe(0);
    expect(store.getState().layout.collapsedRows?.size ?? 0).toBe(0);
  });

  it('works on the innermost column group and leaves ungrouped selections alone', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    groupCols(store, null, 1, 6, wb);
    groupCols(store, null, 2, 3, wb);
    hideOutlineDetail(store, null, range(0, 3, 0, 3), wb);
    expect([...store.getState().layout.hiddenCols].sort()).toEqual([2, 3]);
    hideOutlineDetail(store, null, range(0, 9, 0, 9), wb);
    expect([...store.getState().layout.hiddenCols].sort()).toEqual([2, 3]);
  });
});
