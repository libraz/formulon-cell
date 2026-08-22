import { describe, expect, it } from 'vitest';

import { syncCellFormatsToEngine } from '../../../src/engine/cell-format-sync.js';
import { SheetVisibility } from '../../../src/engine/types.js';
import { defaultFontForLocale, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

/**
 * Capability-gated feature methods live on the WorkbookHandle. Under the stub
 * engine those that need first-class engine support no-op (return false / []).
 * We pin that contract here so future refactors can't silently flip a stub
 * call to a thrown error.
 */
describe('engine/workbook-handle-features (stub) — capability gates', () => {
  it('setColumnWidth no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.capabilities.colRowSize).toBe(false);
    expect(wb.setColumnWidth(0, 0, 0, 120)).toBe(false);
  });

  it('setRowHeight no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setRowHeight(0, 0, 24)).toBe(false);
  });

  it('getColumnLayouts returns [] on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.getColumnLayouts(0)).toEqual([]);
  });

  it('setSheetFreeze no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setSheetFreeze(0, 1, 1)).toBe(false);
  });

  it('setSheetZoom no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setSheetZoom(0, 150)).toBe(false);
  });

  it('setSheetTabHidden no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.setSheetTabHidden(0, true)).toBe(false);
  });

  it('capabilities snapshot reports the stub flavor', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.isStub).toBe(true);
    // The stub publishes a capabilities object so callers can branch.
    expect(wb.capabilities).toBeDefined();
    expect(typeof wb.capabilities).toBe('object');
  });

  it('selects a locale-appropriate default font for a new workbook', () => {
    expect(defaultFontForLocale('ja-JP')).toEqual({ name: '游ゴシック', size: 11 });
    expect(defaultFontForLocale('en-US')).toEqual({ name: 'Calibri', size: 11 });
  });

  it('setSheetVisibility no-ops on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.capabilities.sheetVisibility).toBe(false);
    expect(wb.setSheetVisibility(0, SheetVisibility.VeryHidden)).toBe(false);
  });

  it('getIterative returns null without the engine readback', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.capabilities.iterativeSettings).toBe(false);
    expect(wb.getIterative()).toBeNull();
  });

  it('getComments returns [] on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.capabilities.commentsEnumerable).toBe(false);
    expect(wb.getComments(0)).toEqual([]);
  });
});

/** The engine seeds a new workbook's style table with the record a spreadsheet
 *  writes and offers no way to replace font 0, so the locale default has to be
 *  carried by the handle rather than read back out of the table. */
describe('engine/workbook-handle — workbook default font', () => {
  it('reports the locale default for a workbook this handle created', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja-JP' });
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.workbookDefaultFont).toEqual({ name: '游ゴシック', size: 11 });
      // The engine's own seeded font 0 stays what it always was.
      expect(wb.getFontRecord(0)?.name).toBe('Calibri');
    } finally {
      wb.dispose();
    }
  });

  it('names the locale font on a cell that states no family of its own', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja-JP' });
    try {
      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      syncCellFormatsToEngine(wb, store, 0);

      const xfIndex = wb.getCellXfIndex(0, 0, 0);
      expect(xfIndex).not.toBeNull();
      const xf = wb.getCellXf(xfIndex ?? 0);
      expect(wb.getFontRecord(xf?.fontIndex ?? 0)?.name).toBe('游ゴシック');
    } finally {
      wb.dispose();
    }
  });

  it('falls back to a loaded workbook’s own font 0', async () => {
    const authored = await WorkbookHandle.createDefault({ locale: 'ja-JP' });
    let bytes: Uint8Array;
    try {
      bytes = authored.save();
    } finally {
      authored.dispose();
    }
    const reloaded = await WorkbookHandle.loadBytes(bytes);
    try {
      expect(reloaded.workbookDefaultFont?.name).toBe(reloaded.getFontRecord(0)?.name);
    } finally {
      reloaded.dispose();
    }
  });
});
