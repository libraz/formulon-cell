import { describe, expect, it, vi } from 'vitest';
import { History, recordFormatChange } from '../../../src/commands/history.js';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { hydrateActiveSheetFromEngine } from '../../../src/mount/hydration.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const setFormats = (store: SpreadsheetStore, formats: Map<string, CellFormat>): void => {
  store.setState((state) => ({
    ...state,
    format: { ...state.format, formats },
  }));
};

const formatAt = (sheet: number, row = 0, col = 0): string => addrKey({ sheet, row, col });

describe('WorkbookHandle store format writeback', () => {
  it('writes bound store formats on both sheets before save and reload', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.isStub).toBe(false);
      expect(wb.capabilities.cellFormatting).toBe(true);
      expect(wb.addSheet('Second')).toBe(1);
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 125);
      wb.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1*2');
      wb.setNumber({ sheet: 1, row: 0, col: 0 }, 250);
      wb.setFormula({ sheet: 1, row: 0, col: 1 }, '=A1*2');

      const store = createSpreadsheetStore();
      wb.attachStore(store);
      setFormats(
        store,
        new Map([
          [
            formatAt(0),
            {
              numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
              fontFamily: 'Arial',
              fontSize: 14,
              bold: true,
            },
          ],
          [
            formatAt(1),
            {
              numFmt: { kind: 'currency', decimals: 0, symbol: '¥' },
              fontFamily: 'Times New Roman',
              fontSize: 16,
              italic: true,
            },
          ],
        ]),
      );

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      expect(reloaded.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 125,
      });
      expect(reloaded.cellFormula({ sheet: 0, row: 0, col: 1 })).toBe('=A1*2');
      expect(reloaded.getValue({ sheet: 1, row: 0, col: 0 })).toEqual({
        kind: 'number',
        value: 250,
      });
      expect(reloaded.cellFormula({ sheet: 1, row: 0, col: 1 })).toBe('=A1*2');

      const hydrated = createSpreadsheetStore();
      hydrateActiveSheetFromEngine(reloaded, hydrated);
      expect(hydrated.getState().format.formats.get(formatAt(0))).toMatchObject({
        numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
        fontFamily: 'Arial',
        fontSize: 14,
        bold: true,
      });
      mutators.setSheetIndex(hydrated, 1);
      hydrateActiveSheetFromEngine(reloaded, hydrated);
      expect(hydrated.getState().format.formats.get(formatAt(1))).toMatchObject({
        numFmt: { kind: 'currency', decimals: 0, symbol: '¥' },
        fontFamily: 'Times New Roman',
        fontSize: 16,
        italic: true,
      });
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('resets a removed store format to XF 0 before saving', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      const store = createSpreadsheetStore();
      wb.attachStore(store);
      setFormats(
        store,
        new Map([[formatAt(0), { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } }]]),
      );
      expect(wb.getCellXfIndex(0, 0, 0)).toBeGreaterThan(0);

      setFormats(store, new Map());
      expect(wb.getCellXfIndex(0, 0, 0)).toBe(0);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      expect(reloaded.getCellXfIndex(0, 0, 0)).toBe(0);
      const hydrated = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(reloaded, hydrated, 0);
      expect(hydrated.getState().format.formats.get(formatAt(0))).toBeUndefined();
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('writeback follows format undo and redo snapshots', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const history = new History();
      wb.attachStore(store);
      recordFormatChange(history, store, () => {
        mutators.setCellFormat(
          store,
          { sheet: 0, row: 0, col: 0 },
          {
            numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
            fontFamily: 'Arial',
          },
        );
      });
      expect(wb.getCellXfIndex(0, 0, 0)).toBeGreaterThan(0);

      expect(history.undo()).toBe(true);
      expect(wb.getCellXfIndex(0, 0, 0)).toBe(0);
      expect(store.getState().format.formats.get(formatAt(0))).toBeUndefined();

      expect(history.redo()).toBe(true);
      expect(wb.getCellXfIndex(0, 0, 0)).toBeGreaterThan(0);
      expect(store.getState().format.formats.get(formatAt(0))).toMatchObject({
        numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
        fontFamily: 'Arial',
      });
    } finally {
      wb.dispose();
    }
  });

  it('does not write hydrated formats back while rebasing the store baseline', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const source = createSpreadsheetStore();
      setFormats(
        source,
        new Map([[formatAt(0), { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } }]]),
      );
      syncCellFormatsToEngine(wb, source, 0);
      const store = createSpreadsheetStore();
      wb.attachStore(store);
      const setCellXfIndex = vi.spyOn(wb, 'setCellXfIndex');

      hydrateActiveSheetFromEngine(wb, store);

      expect(setCellXfIndex).not.toHaveBeenCalled();
      expect(store.getState().format.formats.get(formatAt(0))).toMatchObject({
        numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      });
      store.setState((state) => ({ ...state, data: { ...state.data, sheetIndex: 0 } }));
      expect(setCellXfIndex).not.toHaveBeenCalled();
    } finally {
      wb.dispose();
    }
  });

  it('does not flush unchanged formats when optional custom styles are absent', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      store.setState((state) => {
        const format = { ...state.format };
        delete format.customCellStyles;
        format.formats = new Map([[formatAt(0), { bold: true }]]);
        return { ...state, format };
      });
      syncCellFormatsToEngine(wb, store, 0);
      wb.attachStore(store);
      const setCellXfIndex = vi.spyOn(wb, 'setCellXfIndex');
      mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
      expect(setCellXfIndex).not.toHaveBeenCalled();
    } finally {
      wb.dispose();
    }
  });
});
