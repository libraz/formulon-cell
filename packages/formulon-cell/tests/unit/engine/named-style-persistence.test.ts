import { describe, expect, it } from 'vitest';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const addr = { sheet: 0, row: 0, col: 0 } as const;

/** Author a custom named style plus a cell that uses it, and return the saved bytes. */
const savedWorkbookWithCustomStyle = async (): Promise<{ bytes: Uint8Array; xfId: number }> => {
  const wb = await WorkbookHandle.createDefault();
  try {
    wb.setNumber(addr, 1);
    const normal = wb.getCellStyleXf(0);
    if (!normal) throw new Error('no Normal style xf');
    const font = wb.getFontRecord(normal.fontIndex);
    if (!font) throw new Error('no default font');
    const boldFont = wb.addFontRecord({ ...font, bold: true });
    const xfId = wb.addCellStyleXfRecord({ ...normal, fontIndex: boldFont, applyFont: true });
    expect(wb.setNamedCellStyle('Probe Custom', xfId, null)).toBe(true);
    const cellXf = wb.addXfRecord({ ...normal, fontIndex: boldFont, applyFont: true, xfId });
    expect(wb.setCellXfIndex(0, 0, 0, cellXf)).toBe(true);
    return { bytes: wb.save(), xfId };
  } finally {
    wb.dispose();
  }
};

describe('named style persistence', () => {
  it('keeps an imported custom style reference when the sheet is flushed again', async () => {
    const { bytes } = await savedWorkbookWithCustomStyle();
    const wb = await WorkbookHandle.loadBytes(bytes);
    try {
      const style = wb.getNamedCellStyles().find((s) => s.name === 'Probe Custom');
      expect(style).toBeDefined();
      const store = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(wb, store, 0);
      expect(store.getState().format.formats.get(addrKey(addr))?.cellStyle).toBe('Probe Custom');

      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const xf = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(xf?.xfId).toBe(style?.xfId);

      const again = await WorkbookHandle.loadBytes(wb.save());
      try {
        const reStyle = again.getNamedCellStyles().find((s) => s.name === 'Probe Custom');
        expect(reStyle).toBeDefined();
        expect(again.getCellXf(again.getCellXfIndex(0, 0, 0) ?? -1)?.xfId).toBe(reStyle?.xfId);
      } finally {
        again.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('writes a user-created custom style that no cell uses yet', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      const store = createSpreadsheetStore();
      mutators.upsertCustomCellStyle(store, {
        id: 'custom:Unused Probe',
        label: 'Unused Probe',
        format: { bold: true, fill: '#ffeecc' },
      });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      expect(wb.getNamedCellStyles().some((s) => s.name === 'Unused Probe')).toBe(true);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      expect(reloaded.getNamedCellStyles().some((s) => s.name === 'Unused Probe')).toBe(true);
      const fresh = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(reloaded, fresh, 0);
      expect(
        fresh.getState().format.customCellStyles?.find((s) => s.label === 'Unused Probe')?.format,
      ).toMatchObject({ bold: true });
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('creating a custom style does not flush a sheet the store never hydrated', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.addSheet('Second')).toBe(1);
      expect(wb.addHyperlink(0, 0, 0, 'https://a.example', '', '')).toBe(true);
      const store = createSpreadsheetStore();
      mutators.setSheetIndex(store, 1);
      wb.attachStore(store);

      mutators.upsertCustomCellStyle(store, {
        id: 'custom:Flush Probe',
        label: 'Flush Probe',
        format: { bold: true },
      });

      expect(wb.getHyperlinks(0)).toHaveLength(1);
      expect(wb.getNamedCellStyles().some((s) => s.name === 'Flush Probe')).toBe(true);
    } finally {
      wb.dispose();
    }
  });
});
