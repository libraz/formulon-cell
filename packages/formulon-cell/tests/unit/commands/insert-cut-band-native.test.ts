// @vitest-environment node

import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { copy } from '../../../src/commands/clipboard/copy.js';
import { insertCopiedBand } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { captureSnapshotFromCopyResult } from '../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import type { Addr, Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;
type Axis = 'row' | 'col';
const band = (axis: Axis, start: number, count = 1, sheet = 0): Range =>
  axis === 'col'
    ? { sheet, r0: 0, c0: start, r1: MAX_ROW, c1: start + count - 1 }
    : { sheet, r0: start, c0: 0, r1: start + count - 1, c1: MAX_COL };
const address = (axis: Axis, index: number, perpendicular = 0, sheet = 0): Addr =>
  axis === 'col'
    ? { sheet, row: perpendicular, col: index }
    : { sheet, row: index, col: perpendicular };

describe.each(['row', 'col'] as const)('native Insert Cut Cells (%s)', (axis) => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;
  const snapshot = (start: number, count = 1) => {
    mutators.replaceCells(store, wb.cells(0));
    const source = band(axis, start, count);
    mutators.setActive(store, address(axis, start));
    mutators.setRange(store, source);
    const result = copy(store.getState());
    if (!result) throw new Error('missing copy result');
    const captured = captureSnapshotFromCopyResult(store.getState(), result, 'cut');
    if (!captured) throw new Error('missing cut snapshot');
    mutators.setCopyRange(store, source, 'cut');
    return captured;
  };
  const values = () => Array.from({ length: 8 }, (_, index) => wb.getValue(address(axis, index)));
  const seed = () => {
    for (let index = 0; index < 8; index++) wb.setNumber(address(axis, index), index + 1);
    wb.recalc();
  };
  const size = (handle: WorkbookHandle, index: number, sheet = 0) =>
    axis === 'col'
      ? handle.getColumnLayouts(sheet).find((entry) => entry.first <= index && entry.last >= index)
          ?.width
      : handle.getRowLayouts(sheet).find((entry) => entry.row === index)?.height;
  const setSize = (index: number, value: number, sheet = 0) => {
    if (axis === 'col') {
      wb.setColumnWidth(sheet, index, index, value);
      if (store.getState().data.sheetIndex === sheet) mutators.setColWidth(store, index, value);
    } else {
      wb.setRowHeight(sheet, index, value);
      if (store.getState().data.sheetIndex === sheet) mutators.setRowHeight(store, index, value);
    }
  };

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault();
    expect(wb.isStub).toBe(false);
    wb.renameSheet(0, 'Source');
    wb.attachStore(store);
    history = new History();
  });
  afterEach(() => wb.dispose());

  it('moves multiple axes after the source without losing terminal data or references', () => {
    seed();
    const sourceSize = axis === 'col' ? 82 : 38;
    setSize(1, sourceSize);
    const terminal = axis === 'col' ? MAX_COL : MAX_ROW;
    const tail = address(axis, terminal);
    wb.setNumber(tail, 99);
    const formulaAddr = { sheet: 0, row: 10, col: 10 };
    const terminalRef = axis === 'col' ? 'XFD1' : 'A1048576';
    const sourceRef = axis === 'col' ? '$B$1' : '$A$2';
    wb.setFormula(formulaAddr, `=${sourceRef}+${terminalRef}`);
    const captured = snapshot(1, 2);
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, captured, address(axis, 5))?.writtenRange).toEqual(
      band(axis, 3, 2),
    );
    expect(values()).toEqual([1, 4, 5, 2, 3, 6, 7, 8].map((value) => ({ kind: 'number', value })));
    expect(wb.getValue(tail)).toEqual({ kind: 'number', value: 99 });
    expect(wb.cellFormula(formulaAddr)).toBe(axis === 'col' ? '=$D$1+XFD1' : '=$A$4+A1048576');
    expect(size(wb, 3)).toBe(sourceSize);
    expect(store.getState().selection.active).toEqual(address(axis, 3));
    expect(store.getState().selection.range).toEqual(band(axis, 3, 2));
    expect(history.undo()).toBe(true);
    expect(history.canUndo()).toBe(false);
    expect(size(wb, 1)).toBe(sourceSize);
    expect(values()).toEqual([1, 2, 3, 4, 5, 6, 7, 8].map((value) => ({ kind: 'number', value })));
    expect(wb.cellFormula(formulaAddr)).toBe(`=${sourceRef}+${terminalRef}`);
    expect(wb.getValue(tail)).toEqual({ kind: 'number', value: 99 });
    expect(history.redo()).toBe(true);
    expect(values()).toEqual([1, 4, 5, 2, 3, 6, 7, 8].map((value) => ({ kind: 'number', value })));
    expect(wb.getValue(tail)).toEqual({ kind: 'number', value: 99 });
    expect(size(wb, 3)).toBe(sourceSize);
  });

  it('rejects source start and inside without consuming cut, and consumes adjacent no-op', () => {
    seed();
    const captured = snapshot(1, 2);
    wb.attachHistory(history);
    for (const target of [1, 2]) {
      const beforeSelection = store.getState().selection;
      expect(insertCopiedBand(store, wb, history, captured, address(axis, target))).toBeNull();
      expect(store.getState().selection).toEqual(beforeSelection);
      expect(store.getState().ui.copyMode).toBe('cut');
      expect(values()).toEqual(
        [1, 2, 3, 4, 5, 6, 7, 8].map((value) => ({ kind: 'number', value })),
      );
      expect(history.canUndo()).toBe(false);
    }
    expect(insertCopiedBand(store, wb, history, captured, address(axis, 3))?.writtenRange).toEqual(
      band(axis, 1, 2),
    );
    expect(store.getState().ui.copyMode).toBeNull();
    expect(store.getState().ui.copyRange).toBeNull();
    expect(history.canUndo()).toBe(false);
    expect(values()).toEqual([1, 2, 3, 4, 5, 6, 7, 8].map((value) => ({ kind: 'number', value })));
  });

  it('allows a multi-band move before the terminal axis', () => {
    wb.setNumber(address(axis, 1), 7);
    wb.setNumber(address(axis, 2), 8);
    const terminal = axis === 'col' ? MAX_COL : MAX_ROW;
    wb.setNumber(address(axis, terminal), 99);
    const captured = snapshot(1, 2);
    wb.attachHistory(history);
    expect(
      insertCopiedBand(store, wb, history, captured, address(axis, terminal))?.writtenRange,
    ).toEqual(band(axis, terminal - 2, 2));
    expect(wb.getValue(address(axis, terminal - 2))).toEqual({ kind: 'number', value: 7 });
    expect(wb.getValue(address(axis, terminal - 1))).toEqual({ kind: 'number', value: 8 });
    expect(wb.getValue(address(axis, terminal))).toEqual({ kind: 'number', value: 99 });
    expect(history.undo()).toBe(true);
    expect(wb.getValue(address(axis, 1))).toEqual({ kind: 'number', value: 7 });
    expect(wb.getValue(address(axis, 2))).toEqual({ kind: 'number', value: 8 });
  });

  it('keeps cross-sheet source axes, moves whole-axis references and persists notes and dimensions', async () => {
    const targetSheet = wb.addSheet('Target');
    wb.setNumber(address(axis, 1), 7);
    wb.setNumber(address(axis, 2), 8);
    const note = address(axis, 1, 40);
    wb.setCommentEntry(0, note.row, note.col, 'Alice', 'far note');
    const sourceSize = axis === 'col' ? 82 : 38;
    const targetSize = axis === 'col' ? 48 : 24;
    setSize(1, sourceSize);
    const formulaAddr = { sheet: 0, row: 10, col: 10 };
    const originalFormula = axis === 'col' ? '=SUM(B:B)' : '=SUM(2:2)';
    const movedFormula = axis === 'col' ? '=SUM(Target!D:D)' : '=SUM(Target!4:4)';
    wb.setFormula(formulaAddr, originalFormula);
    wb.setNumber(address(axis, 3, 0, targetSheet), 9);
    const captured = snapshot(1);
    mutators.setSheetIndex(store, targetSheet);
    mutators.replaceCells(store, wb.cells(targetSheet));
    hydrateLayoutFromEngine(wb, store, targetSheet);
    setSize(3, targetSize, targetSheet);
    wb.attachHistory(history);
    expect(
      insertCopiedBand(store, wb, history, captured, address(axis, 3, 0, targetSheet))
        ?.writtenRange,
    ).toEqual(band(axis, 3, 1, targetSheet));
    expect(wb.getValue(address(axis, 1))).toEqual({ kind: 'blank' });
    expect(wb.getValue(address(axis, 2))).toEqual({ kind: 'number', value: 8 });
    expect(wb.getValue(address(axis, 3, 0, targetSheet))).toEqual({ kind: 'number', value: 7 });
    expect(wb.getValue(address(axis, 4, 0, targetSheet))).toEqual({ kind: 'number', value: 9 });
    expect(wb.cellFormula(formulaAddr)).toBe(movedFormula);
    const movedNote = address(axis, 3, 40, targetSheet);
    expect(wb.getComment(targetSheet, movedNote.row, movedNote.col)).toEqual({
      author: 'Alice',
      text: 'far note',
    });
    expect(wb.getComment(0, note.row, note.col)).toBeNull();
    expect(size(wb, 1)).toBe(
      axis === 'col'
        ? store.getState().layout.defaultColWidth
        : store.getState().layout.defaultRowHeight,
    );
    expect(size(wb, 3, targetSheet)).toBe(sourceSize);
    expect(size(wb, 4, targetSheet)).toBe(targetSize);
    mutators.setSheetIndex(store, 0);
    hydrateLayoutFromEngine(wb, store, 0);
    const sourceSizes =
      axis === 'col' ? store.getState().layout.colWidths : store.getState().layout.rowHeights;
    expect(sourceSizes.get(1)).not.toBe(sourceSize);
    expect(sourceSizes.get(3)).toBeUndefined();
    mutators.setSheetIndex(store, targetSheet);
    hydrateLayoutFromEngine(wb, store, targetSheet);
    const targetSizes =
      axis === 'col' ? store.getState().layout.colWidths : store.getState().layout.rowHeights;
    expect(targetSizes.get(3)).toBe(sourceSize);
    expect(targetSizes.get(4)).toBe(targetSize);

    const restored = await WorkbookHandle.loadBytes(wb.save());
    try {
      expect(restored.getComment(targetSheet, movedNote.row, movedNote.col)).toEqual({
        author: 'Alice',
        text: 'far note',
      });
      expect(restored.cellFormula(formulaAddr)).toBe(movedFormula);
      expect(size(restored, 3, targetSheet)).toBe(sourceSize);
    } finally {
      restored.dispose();
    }
    expect(history.undo()).toBe(true);
    expect(wb.getValue(address(axis, 1))).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula(formulaAddr)).toBe(originalFormula);
    expect(wb.getComment(0, note.row, note.col)).toEqual({ author: 'Alice', text: 'far note' });
    expect(size(wb, 1)).toBe(sourceSize);
    expect(size(wb, 3, targetSheet)).toBe(targetSize);
    expect(history.redo()).toBe(true);
    expect(wb.cellFormula(formulaAddr)).toBe(movedFormula);
    expect(size(wb, 1)).toBe(
      axis === 'col'
        ? store.getState().layout.defaultColWidth
        : store.getState().layout.defaultRowHeight,
    );
  });

  it.each([1, 2])(
    'keeps native styles and validation for %i moved axes after undo and XLSX reload',
    async (count) => {
      const source = address(axis, 3, 2);
      const destination = address(axis, 1, 2);
      wb.setNumber(source, 7);
      mutators.setCellFormat(store, source, {
        bold: true,
        hyperlink: 'https://example.com/item',
        validation: { kind: 'list', source: ['7', '8'] },
      });
      const bold = (handle: WorkbookHandle, addr: Addr) => {
        const xfIndex = handle.getCellXfIndex(addr.sheet, addr.row, addr.col);
        const xf = xfIndex === null ? null : handle.getCellXf(xfIndex);
        return xf ? handle.getFontRecord(xf.fontIndex)?.bold : false;
      };
      const validate = (handle: WorkbookHandle, addr: Addr) => {
        expect(bold(handle, addr)).toBe(true);
        expect(handle.getHyperlinks(0)).toContainEqual(
          expect.objectContaining({
            row: addr.row,
            col: addr.col,
            target: 'https://example.com/item',
          }),
        );
        expect(handle.getValidationsForSheet(0).flatMap((rule) => rule.ranges)).toContainEqual({
          sheet: 0,
          r0: addr.row,
          c0: addr.col,
          r1: addr.row,
          c1: addr.col,
        });
      };
      if (count === 2) {
        mutators.setCellFormat(store, address(axis, 4, 2), { bold: true });
      }
      validate(wb, source);
      const captured = snapshot(3, count);
      wb.attachHistory(history);
      expect(insertCopiedBand(store, wb, history, captured, address(axis, 1))).not.toBeNull();
      validate(wb, destination);
      expect(history.undo()).toBe(true);
      validate(wb, source);
      expect(Array.from({ length: 8 }, (_, index) => bold(wb, address(axis, index, 2)))).toEqual(
        Array.from({ length: 8 }, (_, index) => index >= 3 && index < 3 + count),
      );
      const reloaded = await WorkbookHandle.loadBytes(wb.save());
      try {
        validate(reloaded, source);
      } finally {
        reloaded.dispose();
      }
      expect(history.redo()).toBe(true);
      validate(wb, destination);
      expect(Array.from({ length: 8 }, (_, index) => bold(wb, address(axis, index, 2)))).toEqual(
        Array.from({ length: 8 }, (_, index) => index >= 1 && index < 1 + count),
      );
    },
  );

  it.each([false, true])('moves hidden and outline attributes (cross-sheet: %s)', (crossSheet) => {
    const targetSheet = crossSheet ? wb.addSheet('Target') : 0;
    wb.setNumber(address(axis, 1), 7);
    if (axis === 'col') {
      expect(wb.setColumnHidden(0, 1, 1, true)).toBe(true);
      expect(wb.setColumnOutline(0, 1, 1, 2)).toBe(true);
    } else {
      expect(wb.setRowHidden(0, 1, true)).toBe(true);
      expect(wb.setRowOutline(0, 1, 2)).toBe(true);
    }
    hydrateLayoutFromEngine(wb, store, 0);
    const captured = snapshot(1);
    if (crossSheet) {
      mutators.setSheetIndex(store, targetSheet);
      mutators.replaceCells(store, wb.cells(targetSheet));
      hydrateLayoutFromEngine(wb, store, targetSheet);
    }
    const target = crossSheet ? 3 : 4;
    wb.attachHistory(history);
    expect(
      insertCopiedBand(store, wb, history, captured, address(axis, target, 0, targetSheet)),
    ).not.toBeNull();
    const attributes = (index: number, sheet: number) =>
      axis === 'col'
        ? wb.getColumnLayouts(sheet).find((entry) => entry.first <= index && entry.last >= index)
        : wb.getRowLayouts(sheet).find((entry) => entry.row === index);
    expect(attributes(3, targetSheet)).toMatchObject({ hidden: true, outlineLevel: 2 });
    if (crossSheet) {
      expect(attributes(1, 0)?.hidden ?? false).toBe(false);
      expect(attributes(1, 0)?.outlineLevel ?? 0).toBe(0);
    }
    const hidden =
      axis === 'col' ? store.getState().layout.hiddenCols : store.getState().layout.hiddenRows;
    const outline =
      axis === 'col' ? store.getState().layout.outlineCols : store.getState().layout.outlineRows;
    expect(hidden.has(3)).toBe(true);
    expect(outline.get(3)).toBe(2);
    expect(history.undo()).toBe(true);
    expect(attributes(1, 0)).toMatchObject({ hidden: true, outlineLevel: 2 });
    expect(attributes(3, targetSheet)?.hidden ?? false).toBe(false);
    expect(history.redo()).toBe(true);
    expect(attributes(3, targetSheet)).toMatchObject({ hidden: true, outlineLevel: 2 });
  });

  it.each(['name', 'filter', 'sparkline', 'permission', 'view', 'pivot'] as const)(
    'leaves unsupported %s metadata intact on rejection',
    (metadata) => {
      seed();
      if (metadata === 'name') {
        const formula = axis === 'col' ? '=Source!$B$1' : '=Source!$A$2';
        expect(wb.setDefinedNameEntry('CutSource', formula)).toBe(true);
      } else if (metadata === 'filter') {
        store.setState((state) => ({
          ...state,
          ui: {
            ...state.ui,
            filterRange: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 },
          },
        }));
        expect(wb.getSheetAutoFilterXml(0)).toContain('A1:C3');
      } else if (metadata === 'permission') {
        store.setState((state) => ({
          ...state,
          protection: {
            ...state.protection,
            allowedEditRanges: [{ id: 'editable', title: 'Editable', range: band(axis, 1) }],
          },
        }));
      } else if (metadata === 'view') {
        store.setState((state) => ({
          ...state,
          sheetViews: {
            ...state.sheetViews,
            views: [{ id: 'view', name: 'View', sheet: 0, hiddenRows: [1], hiddenCols: [1] }],
          },
        }));
      } else if (metadata === 'pivot') {
        const cache = wb.createPivotCache();
        wb.addPivotCacheField(cache, 'Value');
        const record = wb.addPivotCacheRecord(cache);
        wb.setPivotCacheRecordValue(cache, record, 0, { kind: 'number', value: 2 });
        expect(
          wb.createPivotTable(0, 'CutPivot', cache, { row: 20, col: 20 }),
        ).toBeGreaterThanOrEqual(0);
        expect(wb.getPivotTables()).toHaveLength(1);
      } else {
        store.setState((state) => ({
          ...state,
          sparkline: {
            ...state.sparkline,
            sparklines: new Map([
              [addrKey(address(axis, 1, 2)), { kind: 'line' as const, source: 'A1:B1' }],
            ]),
          },
        }));
      }
      const beforePermissions = store.getState().protection.allowedEditRanges;
      const beforeViews = store.getState().sheetViews.views;
      const beforePivots = wb.getPivotTables();
      const beforeSparklines = new Map(store.getState().sparkline.sparklines);
      const beforeNames = [...wb.definedNames()];
      const beforeFilter = wb.getSheetAutoFilterXml(0);
      const captured = snapshot(1);
      const beforeSelection = store.getState().selection;
      wb.attachHistory(history);
      expect(insertCopiedBand(store, wb, history, captured, address(axis, 5))).toBeNull();
      expect(values()).toEqual(
        [1, 2, 3, 4, 5, 6, 7, 8].map((value) => ({ kind: 'number', value })),
      );
      expect(store.getState().protection.allowedEditRanges).toEqual(beforePermissions);
      expect(store.getState().sheetViews.views).toEqual(beforeViews);
      expect(wb.getPivotTables()).toEqual(beforePivots);
      expect(store.getState().sparkline.sparklines).toEqual(beforeSparklines);
      expect([...wb.definedNames()]).toEqual(beforeNames);
      expect(wb.getSheetAutoFilterXml(0)).toBe(beforeFilter);
      expect(store.getState().selection).toEqual(beforeSelection);
      expect(store.getState().ui.copyMode).toBe('cut');
      expect(history.canUndo()).toBe(false);
      expect(
        insertCopiedBand(store, wb, history, captured, address(axis, 2))?.writtenRange,
      ).toEqual(band(axis, 1));
      expect(store.getState().protection.allowedEditRanges).toEqual(beforePermissions);
      expect(store.getState().sheetViews.views).toEqual(beforeViews);
      expect(wb.getPivotTables()).toEqual(beforePivots);
      expect(store.getState().sparkline.sparklines).toEqual(beforeSparklines);
      expect([...wb.definedNames()]).toEqual(beforeNames);
      expect(wb.getSheetAutoFilterXml(0)).toBe(beforeFilter);
      expect(history.canUndo()).toBe(false);
    },
  );

  it('rejects a protected source on a cross-sheet insertion before target changes', () => {
    const targetSheet = wb.addSheet('Target');
    wb.setNumber(address(axis, 1), 7);
    wb.setNumber(address(axis, 3, 0, targetSheet), 9);
    const captured = snapshot(1);
    mutators.setSheetProtected(store, 0, true);
    mutators.setSheetIndex(store, targetSheet);
    mutators.replaceCells(store, wb.cells(targetSheet));
    wb.attachHistory(history);
    expect(
      insertCopiedBand(store, wb, history, captured, address(axis, 3, 0, targetSheet)),
    ).toBeNull();
    expect(wb.getValue(address(axis, 1))).toEqual({ kind: 'number', value: 7 });
    expect(wb.getValue(address(axis, 3, 0, targetSheet))).toEqual({ kind: 'number', value: 9 });
    expect(store.getState().ui.copyMode).toBe('cut');
    expect(history.canUndo()).toBe(false);
  });

  it('rejects a partially cut merge before altering source or destination', () => {
    wb.setNumber(address(axis, 1), 7);
    const merge: Range =
      axis === 'col'
        ? { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 2 }
        : { sheet: 0, r0: 1, c0: 2, r1: 2, c1: 2 };
    wb.engineAddMerge(0, merge);
    mutators.mergeRange(store, merge);
    const captured = snapshot(1);
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, captured, address(axis, 5))).toBeNull();
    expect(wb.getMerges(0)).toEqual([merge]);
    expect(wb.getValue(address(axis, 1))).toEqual({ kind: 'number', value: 7 });
    expect(history.canUndo()).toBe(false);
  });

  it('unmerges a destination crossed by insertion and restores it with one undo', () => {
    seed();
    const merge: Range =
      axis === 'col'
        ? { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 3 }
        : { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 2 };
    wb.engineAddMerge(0, merge);
    mutators.mergeRange(store, merge);
    mutators.setCellFormat(store, address(axis, 0), { bold: true });
    const captured = snapshot(0);
    wb.attachHistory(history);
    expect(insertCopiedBand(store, wb, history, captured, address(axis, 3))?.writtenRange).toEqual(
      band(axis, 2),
    );
    expect(wb.getMerges(0)).toEqual([]);
    expect(store.getState().format.formats.get(addrKey(address(axis, 2)))?.bold).toBe(true);
    expect(values()).toEqual([2, 3, 1, 4, 5, 6, 7, 8].map((value) => ({ kind: 'number', value })));
    expect(history.undo()).toBe(true);
    expect(wb.getMerges(0)).toEqual([merge]);
    expect(values()).toEqual([1, 2, 3, 4, 5, 6, 7, 8].map((value) => ({ kind: 'number', value })));
    expect(history.redo()).toBe(true);
    expect(wb.getMerges(0)).toEqual([]);
  });
});
