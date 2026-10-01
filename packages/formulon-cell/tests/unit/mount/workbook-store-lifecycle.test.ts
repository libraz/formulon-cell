import { describe, expect, it } from 'vitest';

import { syncCellFormatsToEngine } from '../../../src/engine/cell-format-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { mountStubSheet } from '../../test-utils/mount.js';

describe('external workbook store binding lifecycle', () => {
  it('keeps the current store subscribed when an older instance is disposed', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const older = await mountStubSheet({ workbook });
    const current = await mountStubSheet({ workbook });
    try {
      older.instance.dispose();
      mutators.setCellFormat(current.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      expect(workbook.getCellXfIndex(0, 0, 0)).toBeGreaterThan(0);
      current.instance.dispose();
      mutators.setCellFormat(
        current.instance.store,
        { sheet: 0, row: 1, col: 0 },
        { italic: true },
      );
      expect(workbook.getCellXfIndex(0, 1, 0)).toBe(0);
    } finally {
      older.dispose();
      current.dispose();
      workbook.dispose();
    }
  });

  it('does not carry old cell formats into a plain replacement workbook', async () => {
    const original = await WorkbookHandle.createDefault();
    const next = await WorkbookHandle.createDefault();
    original.setText({ sheet: 0, row: 0, col: 0 }, 'Old');
    next.setText({ sheet: 0, row: 0, col: 0 }, 'New');
    const source = createSpreadsheetStore();
    mutators.setCellFormat(
      source,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        fontFamily: 'Arial',
      },
    );
    syncCellFormatsToEngine(original, source, 0);
    const sheet = await mountStubSheet({ workbook: original });
    try {
      expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
      await sheet.instance.setWorkbook(next);
      expect(sheet.instance.store.getState().format.formats.get('0:0:0')).toBeUndefined();
      mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 1, col: 0 }, { italic: true });
      expect(next.getCellXfIndex(0, 0, 0)).toBe(0);
      const loaded = await WorkbookHandle.loadBytes(next.save());
      try {
        expect(loaded.getCellXfIndex(0, 0, 0)).toBe(0);
      } finally {
        loaded.dispose();
      }
    } finally {
      sheet.dispose();
      original.dispose();
      next.dispose();
    }
  });

  it('detaches the old external workbook before hydrating its replacement', async () => {
    const original = await WorkbookHandle.createDefault();
    const next = await WorkbookHandle.createDefault();
    original.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
    next.setNumber({ sheet: 0, row: 0, col: 0 }, 20);
    const source = createSpreadsheetStore();
    mutators.setCellFormat(
      source,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      },
    );
    syncCellFormatsToEngine(next, source, 0);
    const sheet = await mountStubSheet({ workbook: original });
    try {
      expect(original.getCellXfIndex(0, 0, 0)).toBe(0);
      await sheet.instance.setWorkbook(next);
      expect(sheet.instance.workbook).toBe(next);
      expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.numFmt).toMatchObject({
        kind: 'currency',
        symbol: '$',
      });
      expect(original.getCellXfIndex(0, 0, 0)).toBe(0);
      mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      expect(original.getCellXfIndex(0, 0, 0)).toBe(0);
      expect(original.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 10,
      });
    } finally {
      sheet.dispose();
      original.dispose();
      next.dispose();
    }
  });

  it('detaches an externally owned workbook when the spreadsheet is disposed', async () => {
    const original = await WorkbookHandle.createDefault();
    original.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
    const sheet = await mountStubSheet({ workbook: original });
    try {
      sheet.instance.dispose();
      mutators.setCellFormat(sheet.instance.store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      expect(original.getCellXfIndex(0, 0, 0)).toBe(0);
      expect(original.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 10,
      });
    } finally {
      sheet.dispose();
      original.dispose();
    }
  });
});
