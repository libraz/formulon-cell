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

  it('writes authored alignment defaults over stale raw XF fields and keeps selectors', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.isStub).toBe(false);
      const normal = wb.getCellStyleXf(0);
      const defaultFont = wb.getFontRecord(0);
      const defaultFill = wb.getFillRecord(0);
      expect(normal).toBeDefined();
      expect(defaultFont).toBeDefined();
      expect(defaultFill).toBeDefined();
      if (!normal || !defaultFont || !defaultFill) return;

      const fontIndex = wb.addFontRecord({ ...defaultFont, name: 'Alignment Font', bold: true });
      const fillIndex = wb.addFillRecord({
        ...defaultFill,
        pattern: 1,
        fgArgb: 0xffd9ead3,
        bgArgb: 0,
      });
      expect(fontIndex).toBeGreaterThanOrEqual(0);
      expect(fillIndex).toBeGreaterThanOrEqual(0);
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 12);
      const sourceXf = wb.addXfRecord({
        ...normal,
        fontIndex,
        fillIndex,
        horizontalAlign: 1,
        verticalAlign: 1,
        wrapText: true,
        justifyLastLine: false,
        hasAlignment: true,
        hasHorizontalAlign: true,
        hasVerticalAlign: true,
        hasWrapText: true,
        hasJustifyLastLine: true,
        textRotation: 37,
        indent: 2,
        relativeIndent: -1,
        shrinkToFit: true,
        readingOrder: 2,
      });
      expect(wb.setCellXfIndex(0, 0, 0, sourceXf)).toBe(true);

      const store = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(wb, store, 0);
      const key = formatAt(0);
      expect(store.getState().format.formats.get(key)).toMatchObject({
        bold: true,
        fill: '#d9ead3',
      });
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 0 },
        {
          rotation: 0,
          indent: 0,
          shrinkToFit: false,
          textDirection: 'context',
        },
      );
      syncCellFormatsToEngine(wb, store, 0, { strict: true });

      const reset = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(reset).toMatchObject({
        fontIndex,
        fillIndex,
        textRotation: 0,
        indent: 0,
        shrinkToFit: false,
        readingOrder: 0,
      });
      expect(reset?.hasAlignment).toBe(true);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      const loaded = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(loaded).toMatchObject({
        textRotation: 0,
        indent: 0,
        shrinkToFit: false,
        readingOrder: 0,
      });
      expect(reloaded.getFontRecord(loaded?.fontIndex ?? -1)?.name).toBe('Alignment Font');
      expect(reloaded.getFillRecord(loaded?.fillIndex ?? -1)?.fgArgb).toBe(0xffd9ead3);
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('preserves raw vertical-text rotation through unrelated edits but not an explicit degree reset', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeDefined();
      if (!normal) return;
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 7);
      const sourceXf = wb.addXfRecord({
        ...normal,
        hasAlignment: true,
        hasHorizontalAlign: true,
        hasVerticalAlign: true,
        hasWrapText: true,
        hasJustifyLastLine: true,
        textRotation: 255,
        indent: 0,
        shrinkToFit: false,
        readingOrder: 0,
      });
      expect(wb.setCellXfIndex(0, 0, 0, sourceXf)).toBe(true);

      const store = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(wb, store, 0);
      const key = formatAt(0);
      expect(store.getState().format.formats.get(key)?.rotation).toBeUndefined();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const preserved = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(preserved?.textRotation).toBe(255);

      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { rotation: 0 });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const reset = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(reset?.textRotation).toBe(0);
      expect(reset?.textRotation).not.toBe(255);
    } finally {
      wb.dispose();
    }
  });

  it('round-trips authored rotations, indent, shrink and reading direction through the real engine', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.isStub).toBe(false);
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 45);
      wb.setNumber({ sheet: 0, row: 1, col: 0 }, -45);
      wb.setNumber({ sheet: 0, row: 2, col: 0 }, -90);
      const store = createSpreadsheetStore();
      setFormats(
        store,
        new Map([
          [formatAt(0, 0), { rotation: 45, indent: 8, shrinkToFit: true, textDirection: 'rtl' }],
          [formatAt(0, 1), { rotation: -45, indent: 0, shrinkToFit: false, textDirection: 'ltr' }],
          [formatAt(0, 2), { rotation: -90, textDirection: 'context' }],
        ]),
      );
      syncCellFormatsToEngine(wb, store, 0, { strict: true });

      expect(wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1)).toMatchObject({
        textRotation: 45,
        indent: 8,
        shrinkToFit: true,
        readingOrder: 2,
      });
      expect(wb.getCellXf(wb.getCellXfIndex(0, 1, 0) ?? -1)).toMatchObject({
        textRotation: 135,
        indent: 0,
        shrinkToFit: false,
        readingOrder: 1,
      });
      expect(wb.getCellXf(wb.getCellXfIndex(0, 2, 0) ?? -1)).toMatchObject({
        textRotation: 180,
        readingOrder: 0,
      });

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      const loaded = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(reloaded, loaded, 0);
      expect(loaded.getState().format.formats.get(formatAt(0, 0))).toMatchObject({
        rotation: 45,
        indent: 8,
        shrinkToFit: true,
        textDirection: 'rtl',
      });
      expect(loaded.getState().format.formats.get(formatAt(0, 1))).toMatchObject({
        rotation: -45,
        indent: 0,
        shrinkToFit: false,
        textDirection: 'ltr',
      });
      expect(loaded.getState().format.formats.get(formatAt(0, 2))).toMatchObject({
        rotation: -90,
        textDirection: 'context',
      });
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('retains explicit zero and false alignment attributes while compact records stay absent', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.isStub).toBe(false);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeDefined();
      if (!normal) return;
      wb.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
      wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
      const explicit = wb.addXfRecord({
        ...normal,
        hasAlignment: true,
        hasHorizontalAlign: false,
        hasVerticalAlign: false,
        hasWrapText: false,
        hasJustifyLastLine: false,
        textRotation: 0,
        indent: 0,
        shrinkToFit: false,
        readingOrder: 0,
      });
      const compact = wb.addXfRecord({
        ...normal,
        hasAlignment: false,
        hasHorizontalAlign: false,
        hasVerticalAlign: false,
        hasWrapText: false,
        hasJustifyLastLine: false,
      });
      expect(wb.setCellXfIndex(0, 0, 0, explicit)).toBe(true);
      expect(wb.setCellXfIndex(0, 0, 1, compact)).toBe(true);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      const explicitLoaded = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(explicitLoaded).toMatchObject({
        hasAlignment: true,
        hasHorizontalAlign: false,
        hasVerticalAlign: false,
        hasWrapText: false,
        hasJustifyLastLine: false,
        textRotation: 0,
        indent: 0,
        shrinkToFit: false,
        readingOrder: 0,
      });
      const compactLoaded = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 1) ?? -1);
      expect(compactLoaded).toMatchObject({
        hasAlignment: false,
        hasHorizontalAlign: false,
        hasVerticalAlign: false,
        hasWrapText: false,
        hasJustifyLastLine: false,
      });
      expect(compactLoaded).not.toHaveProperty('textRotation');
      expect(compactLoaded).not.toHaveProperty('indent');
      expect(compactLoaded).not.toHaveProperty('shrinkToFit');
      expect(compactLoaded).not.toHaveProperty('readingOrder');
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });
});
