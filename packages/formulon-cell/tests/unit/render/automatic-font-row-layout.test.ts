import { describe, expect, it, vi } from 'vitest';
import { setFont } from '../../../src/commands/format.js';
import { associatePaginationEngine } from '../../../src/commands/pagination.js';
import { MAX_COL } from '../../../src/engine/address.js';
import { fontPointsToCssPx } from '../../../src/engine/font-units.js';
import { macDefaultRowPt } from '../../../src/engine/mac-row-model.js';
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

  it('grows the row as soon as a larger font is set on an empty cell', async () => {
    // Excel: 24 pt on an empty A1 heightens row 1 at once, before any input.
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US' });
    const mounted = await mountStubSheet({ workbook });
    try {
      const store = mounted.instance.store;
      mutators.setActive(store, A1);
      const original = layoutForView(store.getState()).defaultRowHeight;

      expect(setFont(store.getState(), store, { fontSize: fontPointsToCssPx(24) })).toBe(true);
      expect(store.getState().format.formats.get('0:0:0')?.fontSize).toBe(fontPointsToCssPx(24));
      const grown = layoutForView(store.getState()).rowHeights.get(0);
      expect(grown).toBeGreaterThan(original);

      // The same height an authored format gives the row.
      const authored = await WorkbookHandle.createDefault({ locale: 'en-US' });
      const reference = await mountStubSheet({ workbook: authored });
      try {
        mutators.setCellFormat(reference.instance.store, A1, { fontSize: fontPointsToCssPx(24) });
        expect(grown).toBe(layoutForView(reference.instance.store.getState()).rowHeights.get(0));
      } finally {
        reference.dispose();
        authored.dispose();
      }
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

describe('automatic font row layout on the mac profile', () => {
  it('sizes a row to the tabulated default row of its font, as Excel autofits on load', async () => {
    // Excel for Mac autofits a row to the height its default-row table gives
    // the row's font (Arial 12 / 14: 16 / 18 pt), never below the sheet default.
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US', platform: 'mac' });
    const mounted = await mountStubSheet({ workbook, ui: { platform: 'mac' } });
    try {
      const store = mounted.instance.store;
      const base = layoutForView(store.getState()).defaultRowHeight;
      const font = (row: number, fontFamily: string, points: number) =>
        mutators.setCellFormat(
          store,
          { sheet: 0, row, col: 0 },
          {
            fontFamily,
            fontSize: fontPointsToCssPx(points),
          },
        );
      font(0, 'Calibri', 18);
      font(1, 'Calibri', 16);
      font(2, 'Arial', 14);
      const layout = layoutForView(store.getState());
      const height = (row: number) => layout.rowHeights.get(row) ?? layout.defaultRowHeight;
      expect(base).toBeLessThan(fontPointsToCssPx(21));
      expect(height(0)).toBeCloseTo(fontPointsToCssPx(24), 6);
      expect(height(1)).toBeCloseTo(fontPointsToCssPx(21), 6);
      expect(height(2)).toBeCloseTo(Math.max(base, fontPointsToCssPx(18)), 6);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('sizes a row whose cell names only a family, at the Normal size', async () => {
    // Excel for Mac autofits a Courier New 11 row under MS PGothic 11 to 15 pt;
    // the store omits a size equal to the Normal font's.
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US', platform: 'mac' });
    const mounted = await mountStubSheet({ workbook, locale: 'en', ui: { platform: 'mac' } });
    try {
      const store = mounted.instance.store;
      const size = workbook.workbookDisplayFont?.size ?? 0;
      const meiryoRow = fontPointsToCssPx(macDefaultRowPt('Meiryo UI', size) ?? 0);
      mutators.setCellFormat(store, A1, { fontFamily: 'Meiryo UI' });
      mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fontFamily: 'Arial' });
      const layout = layoutForView(store.getState());
      expect(layout.defaultRowHeight).toBeLessThan(meiryoRow);
      expect(layout.rowHeights.get(0)).toBeCloseTo(meiryoRow, 6);
      expect(layout.rowHeights.get(1)).toBeUndefined();
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('autofits a content row below a pinned taller sheet default', async () => {
    // Excel for Mac: with <sheetFormatPr defaultRowHeight="20" customHeight="1">
    // empty rows stay 40 device px but a row holding Normal-font text autofits
    // to the Normal font's row (28 px under MS PGothic 11).
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US', platform: 'mac' });
    const mounted = await mountStubSheet({ workbook, locale: 'en', ui: { platform: 'mac' } });
    try {
      const store = mounted.instance.store;
      const normal = workbook.workbookDisplayFont;
      const normalRow = fontPointsToCssPx(
        macDefaultRowPt(normal?.name ?? '', normal?.size ?? 0) ?? 0,
      );
      const pinned = fontPointsToCssPx(20);
      expect(normalRow).toBeLessThan(pinned);
      store.setState((s) => ({ ...s, layout: { ...s.layout, defaultRowHeight: pinned } }));
      mutators.setCell(store, { sheet: 0, row: 2, col: 1 }, { kind: 'text', value: 'Hxg' });
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 4, col: 1 },
        {
          fontFamily: 'Arial',
          fontSize: fontPointsToCssPx(14),
        },
      );
      const layout = layoutForView(store.getState());
      expect(layout.defaultRowHeight).toBe(pinned);
      expect(layout.rowHeights.get(2)).toBeCloseTo(normalRow, 6);
      expect(layout.rowHeights.get(4)).toBeCloseTo(Math.max(normalRow, fontPointsToCssPx(18)), 6);
      expect(layout.rowHeights.get(3)).toBeUndefined();
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('autofits rows of wrapped text at load to whole lines of the font pitch', async () => {
    // Excel for Mac (no stored row heights): a wrapped cell's row is n lines of
    // the font's pitch - Arial 11 15 pt, Arial 20 26, Courier New 11 16,
    // Calibri 11 15 - and a single wrapped line already takes one pitch.
    const workbook = await WorkbookHandle.createDefault({ locale: 'en-US', platform: 'mac' });
    const mounted = await mountStubSheet({ workbook, locale: 'en', ui: { platform: 'mac' } });
    try {
      const store = mounted.instance.store;
      const wrapped = (row: number, text: string, fontFamily: string, points: number) => {
        const addr = { sheet: 0, row, col: 1 };
        mutators.setCell(store, addr, { kind: 'text', value: text });
        mutators.setCellFormat(store, addr, {
          wrap: true,
          fontFamily,
          fontSize: fontPointsToCssPx(points),
        });
      };
      wrapped(0, 'ab\nab\nab\nab', 'Arial', 11);
      wrapped(1, 'ab\nab\nab', 'Arial', 20);
      wrapped(2, 'ab\nab', 'Courier New', 11);
      wrapped(3, 'ab\nab', 'Calibri', 11);
      wrapped(4, 'ab', 'Arial', 11);
      mutators.setRowHeight(store, 5, fontPointsToCssPx(40));
      wrapped(5, 'ab\nab\nab\nab', 'Arial', 11);
      wrapped(6, 'ab', 'Arial', 8);
      wrapped(7, 'ab', 'Arial', 8);
      mutators.setCell(store, { sheet: 0, row: 7, col: 0 }, { kind: 'text', value: 'label' });
      const layout = layoutForView(store.getState());
      const pt = (row: number) =>
        (layout.rowHeights.get(row) ?? layout.defaultRowHeight) / fontPointsToCssPx(1);
      expect(pt(0)).toBeCloseTo(60, 6);
      expect(pt(1)).toBeCloseTo(78, 6);
      expect(pt(2)).toBeCloseTo(32, 6);
      expect(pt(3)).toBeCloseTo(30, 6);
      // One wrapped line is a full pitch, floored only by the row's other cells.
      expect(pt(4)).toBeCloseTo(15, 6);
      expect(pt(6)).toBeCloseTo(12, 6);
      const normal = workbook.workbookDisplayFont;
      const normalRowPt = macDefaultRowPt(normal?.name ?? '', normal?.size ?? 0) ?? 0;
      expect(pt(7)).toBeCloseTo(Math.max(12, normalRowPt), 6);
      // A row with its own height keeps it.
      expect(pt(5)).toBeCloseTo(40, 6);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });
});
