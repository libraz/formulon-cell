// @vitest-environment node

import type { GeometryMode, Workbook } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type RawStatus = { ok: boolean; message?: string };
type RawWorkbook = {
  getSheetColumns(sheet: number): {
    status: RawStatus;
    columns: ReadonlyArray<{
      first: number;
      last: number;
      width: number;
      hidden: number;
      outlineLevel: number;
      hasWidth?: number;
    }>;
  };
  getSheetRowOverrides(sheet: number): {
    status: RawStatus;
    rows: ReadonlyArray<{
      row: number;
      height: number;
      hidden: number;
      outlineLevel: number;
      hasHeight?: number;
      customHeight?: number;
    }>;
  };
  getSheetFormatDefaults(sheet: number): {
    status: RawStatus;
    defaultColWidth: number;
    defaultRowHeight: number;
    hasDefaultColWidth: boolean;
    hasDefaultRowHeight: boolean;
    effectiveDefaultColWidth: number;
  };
  setSheetFormatDefaults(
    sheet: number,
    defaults: {
      defaultRowHeight?: number;
      hasDefaultRowHeight?: boolean;
    },
  ): RawStatus;
  columnCharsToPt(sheet: number, mode: number, chars: number): { status: RawStatus; value: number };
  columnPtToChars(
    sheet: number,
    mode: number,
    points: number,
  ): { status: RawStatus; value: number };
  getRowHeightPt(sheet: number, row: number): { status: RawStatus; value: number };
};

/** The engine's sheet geometry, read in the shapes these tests compare. */
const rawWorkbook = (handle: WorkbookHandle): RawWorkbook => {
  const wb = (handle as unknown as { wb: Workbook }).wb;
  const ok = { ok: true };
  return {
    getSheetColumns: (sheet) => ({
      status: ok,
      columns: wb
        .sheet(sheet)
        .listColumnOverrides()
        .map((column) => ({
          first: column.first,
          last: column.last,
          width: column.widthChars ?? 0,
          hidden: column.hidden ? 1 : 0,
          outlineLevel: column.outlineLevel,
          hasWidth: column.widthChars === null ? 0 : 1,
        })),
    }),
    getSheetRowOverrides: (sheet) => ({
      status: ok,
      rows: wb
        .sheet(sheet)
        .listRowOverrides()
        .map((row) => ({
          row: row.row,
          height: row.heightPt ?? 0,
          hidden: row.hidden ? 1 : 0,
          outlineLevel: row.outlineLevel,
          hasHeight: row.heightPt === null ? 0 : 1,
          customHeight: row.customHeight ? 1 : 0,
        })),
    }),
    getSheetFormatDefaults: (sheet) => {
      const defaults = wb.sheet(sheet).getFormatDefaults();
      return {
        status: ok,
        defaultColWidth: defaults.defaultColumnWidthChars ?? 0,
        defaultRowHeight: defaults.defaultRowHeightPt ?? 0,
        hasDefaultColWidth: defaults.defaultColumnWidthChars !== null,
        hasDefaultRowHeight: defaults.defaultRowHeightPt !== null,
        effectiveDefaultColWidth: defaults.effectiveDefaultColumnWidthChars,
      };
    },
    setSheetFormatDefaults: (sheet, defaults) => {
      wb.sheet(sheet).updateFormatDefaults({
        ...(defaults.hasDefaultRowHeight === false
          ? { defaultRowHeightPt: null }
          : defaults.defaultRowHeight !== undefined
            ? { defaultRowHeightPt: defaults.defaultRowHeight }
            : {}),
      });
      return ok;
    },
    columnCharsToPt: (sheet, mode, chars) => ({
      status: ok,
      value: wb.sheet(sheet).convertColumnCharsToPt(chars, mode as GeometryMode),
    }),
    columnPtToChars: (sheet, mode, points) => ({
      status: ok,
      value: wb.sheet(sheet).convertColumnPtToChars(points, mode as GeometryMode),
    }),
    getRowHeightPt: (sheet, row) => ({ status: ok, value: wb.sheet(sheet).getRowHeightPt(row) }),
  };
};

describe('WorkbookHandle geometry round-trip', () => {
  it('keeps public CSS pixels while native columns use characters and rows use points', async () => {
    const authored = await WorkbookHandle.createDefault();
    let bytes: Uint8Array;
    try {
      expect(authored.capabilities.geometry).toBe(true);
      expect(authored.capabilities.colRowSize).toBe(true);

      const widthPx = 120;
      const heightPx = 40;
      expect(authored.setColumnWidth(0, 0, 0, widthPx)).toBe(true);
      expect(authored.setRowHeight(0, 3, heightPx)).toBe(true);

      const raw = rawWorkbook(authored);
      const nativeColumn = raw.getSheetColumns(0).columns.find((entry) => entry.first === 0);
      const nativeRow = raw.getSheetRowOverrides(0).rows.find((entry) => entry.row === 3);
      const expectedChars = raw.columnPtToChars(0, 0, widthPx * (72 / 96));
      expect(nativeColumn).toBeDefined();
      expect(nativeRow).toBeDefined();
      expect(nativeColumn?.width).toBeCloseTo(expectedChars.value, 8);
      expect(nativeRow?.height).toBeCloseTo(heightPx * (72 / 96), 8);

      const hydrated = createSpreadsheetStore();
      hydrateLayoutFromEngine(authored, hydrated, 0);
      expect(hydrated.getState().layout.colWidths.get(0)).toBeCloseTo(widthPx, 8);
      expect(hydrated.getState().layout.rowHeights.get(3)).toBeCloseTo(heightPx, 8);

      bytes = authored.save();
    } finally {
      authored.dispose();
    }

    const reloaded = await WorkbookHandle.loadBytes(bytes);
    try {
      expect(reloaded.capabilities.geometry).toBe(true);
      const hydrated = createSpreadsheetStore();
      hydrateLayoutFromEngine(reloaded, hydrated, 0);
      expect(hydrated.getState().layout.colWidths.get(0)).toBeCloseTo(120, 8);
      expect(hydrated.getState().layout.rowHeights.get(3)).toBeCloseTo(40, 8);

      const raw = rawWorkbook(reloaded);
      const nativeColumn = raw.getSheetColumns(0).columns.find((entry) => entry.first === 0);
      const nativeRow = raw.getSheetRowOverrides(0).rows.find((entry) => entry.row === 3);
      expect(nativeColumn?.width).toBeCloseTo(raw.columnPtToChars(0, 0, 120 * (72 / 96)).value, 8);
      expect(nativeRow?.height).toBeCloseTo(40 * (72 / 96), 8);
    } finally {
      reloaded.dispose();
    }
  });

  it('hydrates default metrics and does not invent size overrides for hidden-only rows or columns', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.setColumnHidden(0, 4, 4, true)).toBe(true);
      expect(wb.setRowHidden(0, 7, true)).toBe(true);

      const raw = rawWorkbook(wb);
      const defaults = raw.getSheetFormatDefaults(0);
      expect(defaults.status.ok).toBe(true);
      const expectedColDefault = raw.columnCharsToPt(0, 0, defaults.effectiveDefaultColWidth);
      expect(expectedColDefault.status.ok).toBe(true);
      const defaultLayout = wb.getSheetDefaultLayoutPx(0);
      expect(defaultLayout?.defaultColWidth).toBeCloseTo(expectedColDefault.value * (96 / 72), 8);
      // New default-profile books state the row height Windows writes for Calibri 11.
      expect(defaults.hasDefaultRowHeight).toBe(true);
      const freshRowHeight = raw.getRowHeightPt(0, 0);
      expect(freshRowHeight.status.ok).toBe(true);
      expect(freshRowHeight.value).toBe(15);
      expect(defaultLayout?.defaultRowHeight).toBeCloseTo(freshRowHeight.value * (96 / 72), 8);

      const rawColumn = raw
        .getSheetColumns(0)
        .columns.find((entry) => entry.first <= 4 && entry.last >= 4);
      const rawRow = raw.getSheetRowOverrides(0).rows.find((entry) => entry.row === 7);
      expect(rawColumn?.hidden).toBe(1);
      expect(rawColumn?.hasWidth).toBe(0);
      expect(rawRow?.hidden).toBe(1);
      expect(rawRow?.hasHeight).toBe(0);

      const store = createSpreadsheetStore();
      hydrateLayoutFromEngine(wb, store, 0);
      const { layout } = store.getState();
      expect(layout.defaultColWidth).toBeCloseTo(defaultLayout?.defaultColWidth ?? 0, 8);
      expect(layout.defaultRowHeight).toBeCloseTo(defaultLayout?.defaultRowHeight ?? 0, 8);
      expect(layout.hiddenCols.has(4)).toBe(true);
      expect(layout.hiddenRows.has(7)).toBe(true);
      expect(layout.colWidths.has(4)).toBe(false);
      expect(layout.rowHeights.has(7)).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('clears an explicit row height so the sheet default is used after save and reload', async () => {
    const wb = await WorkbookHandle.createDefault();
    let bytes: Uint8Array;
    try {
      const raw = rawWorkbook(wb);
      expect(
        raw.setSheetFormatDefaults(0, { defaultRowHeight: 30, hasDefaultRowHeight: true }).ok,
      ).toBe(true);
      expect(wb.setRowHeight(0, 8, 45)).toBe(true);
      expect(wb.getRowLayouts(0).find((entry) => entry.row === 8)?.height).toBeCloseTo(45, 8);
      expect(wb.clearRowHeight(0, 8)).toBe(true);
      expect(wb.getRowLayouts(0).some((entry) => entry.row === 8)).toBe(false);
      expect(raw.getRowHeightPt(0, 8).value).toBeCloseTo(30, 8);
      bytes = wb.save();
    } finally {
      wb.dispose();
    }

    const reloaded = await WorkbookHandle.loadBytes(bytes);
    try {
      expect(reloaded.getRowLayouts(0).some((entry) => entry.row === 8)).toBe(false);
      expect(rawWorkbook(reloaded).getRowHeightPt(0, 8).value).toBeCloseTo(30, 8);
    } finally {
      reloaded.dispose();
    }
  });
});
