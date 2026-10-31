import { describe, expect, it, vi } from 'vitest';
import { setFont } from '../../../src/commands/format.js';
import { associatePaginationEngine } from '../../../src/commands/pagination.js';
import { MAX_COL } from '../../../src/engine/address.js';
import { fontPointsToCssPx } from '../../../src/engine/font-units.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { cellRect, hitTest, layoutForView } from '../../../src/render/geometry.js';
import { MAX_ROW_HEIGHT_PX } from '../../../src/store/layout-limits.js';
import { mutators } from '../../../src/store/store.js';
import { mountStubSheet } from '../../test-utils/mount.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;

describe('automatic font row layout', () => {
  it('grows an implicit row for a directly formatted blank cell without storing a row height', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      expect(workbook.isStub).toBe(false);
      mutators.setActive(mounted.instance.store, A1);
      mutators.setViewportSize(mounted.instance.store, 10, 6, 640);
      const beforeState = mounted.instance.store.getState();
      const before = layoutForView(beforeState);
      const originalHeight = before.defaultRowHeight;
      expect(before.rowHeights.get(A1.row)).toBeUndefined();

      mutators.setCellFormat(mounted.instance.store, A1, { fontSize: fontPointsToCssPx(24) });
      const afterState = mounted.instance.store.getState();
      const after = layoutForView(afterState);

      const normalCss = fontPointsToCssPx(workbook.workbookDisplayFont?.size ?? 11);
      const largeCss = fontPointsToCssPx(24);
      expect(after.rowHeights.get(A1.row)).toBe(
        originalHeight + Math.round(largeCss * 1.28) - Math.round(normalCss * 1.28),
      );
      expect(after.rowHeights.get(1) ?? after.defaultRowHeight).toBe(originalHeight);
      const enlargedRect = cellRect(after, afterState.viewport, 0, 0);
      const nextRect = cellRect(after, afterState.viewport, 1, 0);
      expect(nextRect.y).toBe(enlargedRect.y + enlargedRect.h);
      expect(
        hitTest(
          after,
          afterState.viewport,
          enlargedRect.x + enlargedRect.w / 2,
          enlargedRect.y + enlargedRect.h / 2,
        ),
      ).toEqual({ row: 0, col: 0 });
      expect(
        hitTest(
          after,
          afterState.viewport,
          nextRect.x + nextRect.w / 2,
          nextRect.y + nextRect.h / 2,
        ),
      ).toEqual({ row: 1, col: 0 });
      expect(afterState.layout.rowHeights).toEqual(beforeState.layout.rowHeights);
      expect(afterState.format.formats.get('0:0:0')?.fontSize).toBe(fontPointsToCssPx(24));
      expect(workbook.getRowLayouts(0)).toEqual([]);

      mutators.setCellFormat(mounted.instance.store, A1, null);
      const cleared = layoutForView(mounted.instance.store.getState());
      expect(cleared.rowHeights.get(0) ?? cleared.defaultRowHeight).toBe(originalHeight);

      mutators.setCellFormat(mounted.instance.store, A1, { fontSize: fontPointsToCssPx(24) });
      const redone = layoutForView(mounted.instance.store.getState());
      expect(redone.rowHeights.get(0)).toBeGreaterThan(originalHeight);
      expect(workbook.getRowLayouts(0)).toEqual([]);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('projects authored font growth and preserves a manual height with the stub engine', async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true, locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      expect(workbook.isStub).toBe(true);
      const store = mounted.instance.store;
      mutators.setActive(store, A1);
      const original = layoutForView(store.getState()).defaultRowHeight;

      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(24) });
      const expanded = layoutForView(store.getState());
      const normalCss = fontPointsToCssPx(workbook.workbookDisplayFont?.size ?? 11);
      const expected =
        original + Math.round(fontPointsToCssPx(24) * 1.28) - Math.round(normalCss * 1.28);
      expect(expanded.rowHeights.get(0)).toBe(expected);

      mutators.setRowHeight(store, 0, 24);
      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(36) });
      expect(layoutForView(store.getState()).rowHeights.get(0)).toBe(24);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('waits for an authored format instead of measuring a pending blank-cell format', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      mutators.setActive(store, A1);
      const original = layoutForView(store.getState()).defaultRowHeight;

      expect(setFont(store.getState(), store, { fontSize: fontPointsToCssPx(24) })).toBe(true);
      const pending = layoutForView(store.getState());
      expect(store.getState().ui.pendingFormat?.format.fontSize).toBe(fontPointsToCssPx(24));
      expect(pending.rowHeights.get(0)).toBeUndefined();
      expect(pending.defaultRowHeight).toBe(original);

      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(24) });
      expect(layoutForView(store.getState()).rowHeights.get(0)).toBeGreaterThan(original);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('keeps a user-sized row fixed while font size changes', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      mutators.setActive(mounted.instance.store, A1);
      mutators.setRowHeight(mounted.instance.store, 0, 24);
      const beforeState = mounted.instance.store.getState();
      const before = layoutForView(beforeState);
      expect(before.rowHeights.get(0)).toBe(24);

      mutators.setCellFormat(mounted.instance.store, A1, { fontSize: fontPointsToCssPx(36) });
      const after = layoutForView(mounted.instance.store.getState());
      expect(after.rowHeights.get(0)).toBe(24);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('uses the tallest sparse font in the row and shrinks after the largest override is removed', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    const other = { sheet: 0, row: 0, col: 2 } as const;
    try {
      mutators.setCellFormat(mounted.instance.store, A1, { fontSize: fontPointsToCssPx(24) });
      mutators.setCellFormat(mounted.instance.store, other, { fontSize: fontPointsToCssPx(36) });
      const tallest = layoutForView(mounted.instance.store.getState()).rowHeights.get(0);
      expect(tallest).toBeDefined();
      expect(tallest).toBeGreaterThan(
        layoutForView(mounted.instance.store.getState()).defaultRowHeight,
      );

      mutators.setCellFormat(mounted.instance.store, other, null);
      const afterRemove = layoutForView(mounted.instance.store.getState()).rowHeights.get(0);
      expect(afterRemove).toBeDefined();
      expect(afterRemove).toBeLessThan(tallest ?? Number.POSITIVE_INFINITY);
      expect(afterRemove).toBeGreaterThan(
        layoutForView(mounted.instance.store.getState()).defaultRowHeight,
      );
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('projects a sparse row-axis XF without expanding other rows', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const baseXf = workbook.getCellXf(0);
      expect(baseXf).not.toBeNull();
      const baseFont = baseXf && workbook.getFontRecord(baseXf.fontIndex);
      expect(baseFont).not.toBeNull();
      if (!baseXf || !baseFont) throw new Error('workbook Normal style is unavailable');
      const fontIndex = workbook.addFontRecord({ ...baseFont, size: 36 });
      const rowXf = workbook.addXfRecord({ ...baseXf, fontIndex });
      expect(rowXf).toBeGreaterThanOrEqual(0);

      const row = 42;
      const state = mounted.instance.store.getState();
      const axisOverrides = new Map(state.format.sheetAxisOverrides ?? []);
      const snapshot = { columns: [], rows: [{ row, styleXf: rowXf }] };
      axisOverrides.set(0, { baseline: { columns: [], rows: [] }, desired: snapshot });
      mounted.instance.store.setState((current) => ({
        ...current,
        format: { ...current.format, sheetAxisOverrides: axisOverrides },
      }));

      const projected = layoutForView(mounted.instance.store.getState());
      expect(projected.rowHeights.get(row)).toBeGreaterThan(projected.defaultRowHeight);
      expect(projected.rowHeights.size).toBe(1);
      expect(mounted.instance.store.getState().layout.rowHeights.size).toBe(0);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('lets a row font replace a worksheet font base while direct cell fonts set a row floor', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const state = mounted.instance.store.getState();
      const formats = new Map(state.format.formats);
      formats.set('0:42:0', { fontSize: fontPointsToCssPx(30) });
      formats.set('0:43:0', { fontSize: fontPointsToCssPx(11) });
      const sheetFontOverrides = new Map(state.format.sheetFontOverrides ?? []);
      sheetFontOverrides.set(0, { patch: { fontSize: fontPointsToCssPx(24) } });
      const sheetAxisOverrides = new Map(state.format.sheetAxisOverrides ?? []);
      sheetAxisOverrides.set(0, {
        baseline: { columns: [], rows: [] },
        desired: {
          columns: [],
          rows: [
            { row: 42, styleXf: 0 },
            { row: 44, styleXf: 0 },
          ],
        },
      });
      const probe = {
        ...state,
        format: { ...state.format, formats, sheetFontOverrides, sheetAxisOverrides },
      };
      associatePaginationEngine(probe, workbook);

      const projected = layoutForView(probe);
      const normalCss = fontPointsToCssPx(workbook.workbookDisplayFont?.size ?? 11);
      const original = state.layout.defaultRowHeight;
      expect(projected.defaultRowHeight).toBe(
        original + Math.round(fontPointsToCssPx(24) * 1.28) - Math.round(normalCss * 1.28),
      );
      expect(projected.rowHeights.get(44) ?? projected.defaultRowHeight).toBe(original);
      expect(projected.rowHeights.get(43) ?? projected.defaultRowHeight).toBe(
        projected.defaultRowHeight,
      );
      expect(projected.rowHeights.get(42)).toBe(
        original + Math.round(fontPointsToCssPx(30) * 1.28) - Math.round(normalCss * 1.28),
      );
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('ignores authored font formats outside the worksheet column bounds', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const state = mounted.instance.store.getState();
      const formats = new Map(state.format.formats);
      formats.set('0:5:-1', { fontSize: 1000 });
      formats.set(`0:6:${MAX_COL + 1}`, { fontSize: 1000 });
      const probe = { ...state, format: { ...state.format, formats } };
      associatePaginationEngine(probe, workbook);

      const projected = layoutForView(probe);
      expect(projected.rowHeights.get(5)).toBeUndefined();
      expect(projected.rowHeights.get(6)).toBeUndefined();
      expect(projected.defaultRowHeight).toBe(state.layout.defaultRowHeight);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('changes the projected default for a compact worksheet font without expanding cell or row maps', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      mutators.selectAll(store);
      const before = layoutForView(store.getState());
      const rowCount = store.getState().layout.rowHeights.size;
      expect(setFont(store.getState(), store, { fontSize: fontPointsToCssPx(24) })).toBe(true);
      const after = layoutForView(store.getState());

      expect(after.defaultRowHeight).toBeGreaterThan(before.defaultRowHeight);
      expect(after.rowHeights.size).toBe(rowCount);
      expect(store.getState().layout.rowHeights.size).toBe(rowCount);
      expect(store.getState().format.formats.size).toBe(0);
      expect(workbook.getRowLayouts(0)).toEqual([]);

      const repeatedState = { ...store.getState(), layout: after };
      associatePaginationEngine(repeatedState, workbook);
      const repeated = layoutForView(repeatedState);
      expect(repeated.defaultRowHeight).toBe(after.defaultRowHeight);

      const state = store.getState();
      const overrides = new Map(state.format.sheetFontOverrides ?? []);
      overrides.delete(state.data.sheetIndex);
      store.setState((current) => ({
        ...current,
        format: { ...current.format, sheetFontOverrides: overrides },
      }));
      expect(layoutForView(store.getState()).defaultRowHeight).toBe(before.defaultRowHeight);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('reuses projected rows across selection and scroll updates, then invalidates after a font edit', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      const axisSnapshot = vi.spyOn(workbook, 'getSheetFontAxisSnapshot');
      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(24) });
      const initial = layoutForView(store.getState());
      const projectedRows = initial.rowHeights;
      const initialSnapshotCalls = axisSnapshot.mock.calls.length;
      expect(initialSnapshotCalls).toBe(1);

      mutators.setActive(store, { sheet: 0, row: 3, col: 1 });
      expect(layoutForView(store.getState()).rowHeights).toBe(projectedRows);
      store.setState((state) => ({
        ...state,
        viewport: { ...state.viewport, rowStart: 4 },
      }));
      expect(layoutForView(store.getState()).rowHeights).toBe(projectedRows);
      expect(axisSnapshot).toHaveBeenCalledTimes(initialSnapshotCalls);

      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(36) });
      const changed = layoutForView(store.getState());
      expect(axisSnapshot).toHaveBeenCalledTimes(initialSnapshotCalls + 1);
      expect(changed.rowHeights).not.toBe(projectedRows);
      expect(changed.rowHeights.get(0)).toBeGreaterThan(initial.rowHeights.get(0) ?? 0);
      expect(axisSnapshot.mock.calls.length).toBeGreaterThan(initialSnapshotCalls);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('ignores directly formatted cells inside merged ranges', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      mutators.setCellFormat(store, A1, { fontSize: fontPointsToCssPx(36) });
      const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
      store.setState((state) => ({
        ...state,
        merges: {
          byAnchor: new Map([['0:0:0', range]]),
          byCell: new Map([
            ['0:0:1', '0:0:0'],
            ['0:1:0', '0:0:0'],
            ['0:1:1', '0:0:0'],
          ]),
        },
      }));

      const layout = layoutForView(store.getState());
      expect(layout.rowHeights.get(0)).toBeUndefined();
      expect(layout.defaultRowHeight).toBe(store.getState().layout.defaultRowHeight);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('clamps very large fonts and skips nonfinite sizes', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      mutators.setCellFormat(store, A1, { fontSize: 1000 });
      expect(layoutForView(store.getState()).rowHeights.get(0)).toBe(MAX_ROW_HEIGHT_PX);

      mutators.setCellFormat(store, A1, { fontSize: Number.POSITIVE_INFINITY });
      const invalid = layoutForView(store.getState());
      expect(invalid.rowHeights.get(0)).toBeUndefined();
      expect(invalid.defaultRowHeight).toBe(store.getState().layout.defaultRowHeight);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });
});
