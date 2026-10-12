import { describe, expect, it } from 'vitest';
import { computeAutofitColWidth } from '../../src/commands/autofit-measurement.js';
import { MAX_COL } from '../../src/engine/address.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../src/store/store.js';

const measureContext = (): CanvasRenderingContext2D => {
  const ctx = {
    font: '',
    measureText(text: string) {
      return { width: text.length * 7 } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
  return ctx;
};

describe('real WASM inherited typography in AutoFit', () => {
  it('measures inherited 24pt text, then layers a direct family without flattening it', async () => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      const baseFont = workbook.getFontRecord(0);
      const baseXf = workbook.getCellXf(0);
      expect(baseFont).not.toBeNull();
      expect(baseXf).not.toBeNull();
      if (!baseFont || !baseXf) throw new Error('missing workbook defaults');

      const inheritedXf = workbook.addXfRecord({
        ...baseXf,
        fontIndex: workbook.addFontRecord({
          ...baseFont,
          name: 'Aptos Display',
          size: 24,
        }),
      });
      expect(workbook.setColumnStyle(0, 0, MAX_COL, inheritedXf)).toBe(true);

      const store = createSpreadsheetStore();
      store.setState((state) => ({
        ...state,
        data: {
          ...state.data,
          cells: new Map([
            ['0:0:0', { value: { kind: 'text', value: 'inherited text' }, formula: null }],
          ]),
        },
      }));

      const inheritedCtx = measureContext();
      computeAutofitColWidth(store.getState(), 0, inheritedCtx, {
        workbook,
        workbookDisplayFont: workbook.workbookDisplayFont,
      });
      expect(inheritedCtx.font).toBe('400 32px "Aptos Display", system-ui, sans-serif');
      expect(store.getState().format.formats.size).toBe(0);

      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { fontFamily: 'Georgia' });
      const directCtx = measureContext();
      computeAutofitColWidth(store.getState(), 0, directCtx, {
        workbook,
        workbookDisplayFont: workbook.workbookDisplayFont,
      });
      expect(directCtx.font).toBe('400 32px Georgia, system-ui, sans-serif');
      expect(store.getState().format.formats.get('0:0:0')).toEqual({ fontFamily: 'Georgia' });
    } finally {
      workbook.dispose();
    }
  });
});
