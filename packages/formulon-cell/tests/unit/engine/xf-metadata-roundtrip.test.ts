import { describe, expect, it } from 'vitest';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const addr = { sheet: 0, row: 0, col: 0 } as const;
const key = addrKey(addr);

describe('CellXf metadata round-trip', () => {
  it('keeps raw apply flags, protection and quotePrefix through a cell edit', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      wb.setNumber(addr, 42);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeTruthy();
      if (!normal) return;

      const sourceXf = wb.addXfRecord({
        ...normal,
        applyNumberFormat: true,
        applyFont: false,
        applyFill: true,
        applyBorder: false,
        applyAlignment: false,
        applyProtection: true,
        quotePrefix: true,
        hasProtection: true,
        locked: false,
        hidden: true,
      });
      expect(wb.setCellXfIndex(0, 0, 0, sourceXf)).toBe(true);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      const loadedSource = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(loadedSource).toMatchObject({
        applyNumberFormat: true,
        applyFont: false,
        applyFill: true,
        applyBorder: false,
        applyAlignment: false,
        applyProtection: true,
        quotePrefix: true,
        hasProtection: true,
        locked: false,
        hidden: true,
      });

      const store = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(reloaded, store, 0);
      expect(store.getState().format.formats.get(key)).toMatchObject({
        locked: false,
        formulaHidden: true,
      });

      // Filling the cell authors only the fill group. All unrelated XF
      // metadata must come from the loaded source record.
      mutators.setCellFormat(store, addr, { fill: '#d9ead3' });
      syncCellFormatsToEngine(reloaded, store, 0, { strict: true });
      const edited = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(edited).toMatchObject({
        applyNumberFormat: true,
        applyFont: false,
        applyFill: true,
        applyBorder: false,
        applyAlignment: false,
        applyProtection: true,
        quotePrefix: true,
        hasProtection: true,
        locked: false,
        hidden: true,
      });

      // Protection is a represented CellFormat group, so changing it must
      // author the new values while retaining the opaque quotePrefix bit.
      mutators.setCellFormat(store, addr, { locked: true, formulaHidden: false });
      syncCellFormatsToEngine(reloaded, store, 0, { strict: true });
      const protectionEdited = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(protectionEdited).toMatchObject({
        applyProtection: true,
        quotePrefix: true,
        hasProtection: true,
        locked: true,
        hidden: false,
      });

      const savedAgain = await WorkbookHandle.loadBytes(reloaded.save());
      try {
        const savedXf = savedAgain.getCellXf(savedAgain.getCellXfIndex(0, 0, 0) ?? -1);
        expect(savedXf).toMatchObject({
          applyProtection: true,
          quotePrefix: true,
          hasProtection: true,
          locked: true,
          hidden: false,
        });
      } finally {
        savedAgain.dispose();
      }
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('retains raw cell metadata when a named-style cell receives an unrelated edit', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      wb.setNumber(addr, 7);
      const normal = wb.getCellStyleXf(0);
      const defaultFont = wb.getFontRecord(0);
      expect(normal).toBeTruthy();
      expect(defaultFont).toBeTruthy();
      if (!normal || !defaultFont) return;

      const boldFont = wb.addFontRecord({ ...defaultFont, bold: true });
      const namedStyleXf = wb.addCellStyleXfRecord({
        ...normal,
        fontIndex: boldFont,
        applyFont: false,
      });
      expect(namedStyleXf).toBeGreaterThanOrEqual(0);
      expect(wb.setNamedCellStyle('Raw Style', namedStyleXf, null)).toBe(true);

      const sourceXf = wb.addXfRecord({
        ...normal,
        fontIndex: boldFont,
        xfId: namedStyleXf,
        applyNumberFormat: true,
        applyFont: false,
        applyFill: false,
        applyBorder: true,
        applyAlignment: false,
        applyProtection: true,
        quotePrefix: true,
        hasProtection: true,
        locked: false,
        hidden: true,
      });
      expect(wb.setCellXfIndex(0, 0, 0, sourceXf)).toBe(true);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      const store = createSpreadsheetStore();
      store.setState((state) => ({
        ...state,
        format: {
          ...state.format,
          customCellStyles: [{ id: 'raw-style', label: 'Raw Style', format: { bold: true } }],
        },
      }));
      hydrateCellFormatsFromEngine(reloaded, store, 0);
      expect(store.getState().format.formats.get(key)).toMatchObject({
        cellStyle: 'Raw Style',
        bold: true,
        locked: false,
        formulaHidden: true,
      });

      mutators.setCellFormat(store, addr, { fill: '#fff2cc' });
      syncCellFormatsToEngine(reloaded, store, 0, { strict: true });
      const edited = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(edited).toMatchObject({
        xfId: namedStyleXf,
        applyNumberFormat: true,
        applyFont: false,
        applyBorder: true,
        applyProtection: true,
        quotePrefix: true,
        hasProtection: true,
        locked: false,
        hidden: true,
      });
      expect(edited?.fillIndex).not.toBe(normal.fillIndex);
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('authors an explicit protection child when the loaded XF had only defaults', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setNumber(addr, 1);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeTruthy();
      if (!normal) return;
      const sourceXf = wb.addXfRecord({
        ...normal,
        applyProtection: false,
        hasProtection: false,
        quotePrefix: true,
      });
      expect(wb.setCellXfIndex(0, 0, 0, sourceXf)).toBe(true);

      const store = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(wb, store, 0);
      expect(store.getState().format.formats.get(key)?.locked).toBeUndefined();

      mutators.setCellFormat(store, addr, { locked: true });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const authored = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(authored).toMatchObject({
        applyProtection: true,
        hasProtection: true,
        locked: true,
        hidden: false,
        quotePrefix: true,
      });
    } finally {
      wb.dispose();
    }
  });
});
