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

/** Font 0 is the record an unstyled cell resolves to, so a new workbook has to
 *  state the locale default there — naming it on formatted cells alone would
 *  still save every untouched cell as the engine's seeded font. */
describe('engine/workbook-handle — workbook default font', () => {
  it('states the locale default as font 0 on a workbook this handle created', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja-JP' });
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.defaultFont).toBe(true);
      expect(wb.getFontRecord(0)?.name).toBe('游ゴシック');
      expect(wb.workbookDefaultFont?.name).toBe('游ゴシック');
      expect(wb.workbookDefaultFont?.size).toBe(11);
    } finally {
      wb.dispose();
    }
  });

  it('keeps a cell with no format of its own on the locale font through a save', async () => {
    const authored = await WorkbookHandle.createDefault({ locale: 'ja-JP' });
    let bytes: Uint8Array;
    try {
      authored.setText({ sheet: 0, row: 0, col: 0 }, '無書式');
      bytes = authored.save();
    } finally {
      authored.dispose();
    }
    const reloaded = await WorkbookHandle.loadBytes(bytes);
    try {
      expect(reloaded.getFontRecord(0)?.name).toBe('游ゴシック');
    } finally {
      reloaded.dispose();
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

  it('keeps the fields the caller left out of a restated default', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'en-US' });
    try {
      const before = wb.getFontRecord(0);
      expect(before).not.toBeNull();
      expect(wb.setWorkbookDefaultFont({ name: 'Meiryo' })).toBe(true);
      const after = wb.getFontRecord(0);
      expect(after?.name).toBe('Meiryo');
      expect(after?.size).toBe(before?.size);
      expect(after?.scheme).toBe(before?.scheme);
      expect(after?.family).toBe(before?.family);
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

describe('engine/workbook-handle-features — phonetic guides', () => {
  it('reads back the spans a guide was written with', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.phoneticRuns).toBe(true);
      wb.setText({ sheet: 0, row: 0, col: 0 }, '東京都');
      const runs = [
        { start: 0, end: 2, text: 'とうきょう' },
        { start: 2, end: 3, text: 'と' },
      ];
      expect(wb.setCellPhoneticRuns(0, 0, 0, runs)).toBe(true);
      expect(wb.getCellPhoneticRuns(0, 0, 0)).toEqual(runs);
      // The whole-cell reader still reports the readings run together.
      expect(wb.getCellPhonetic(0, 0, 0)).toBe('とうきょうと');
    } finally {
      wb.dispose();
    }
  });

  it('reports an unannotated cell as no runs rather than as no surface', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setText({ sheet: 0, row: 0, col: 0 }, '東京都');
      expect(wb.getCellPhoneticRuns(0, 0, 0)).toEqual([]);
    } finally {
      wb.dispose();
    }
  });

  it('no-ops the per-run surface on stub', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    expect(wb.capabilities.phoneticRuns).toBe(false);
    expect(wb.getCellPhoneticRuns(0, 0, 0)).toBeNull();
    expect(wb.setCellPhoneticRuns(0, 0, 0, [{ start: 0, end: 1, text: 'あ' }])).toBe(false);
  });
});
