import { describe, expect, it } from 'vitest';
import {
  applyCellStyle,
  applyCellStyleByName,
  CELL_STYLES,
  createCellStyleFromActiveFormat,
  customCellStyleId,
  getCellStyle,
  listCustomCellStyles,
  mergeCellStylesFromWorkbook,
} from '../../../src/commands/cell-styles.js';
import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/index.js';

const namedStyleWorkbook = (): WorkbookHandle =>
  ({
    getNamedCellStyles: () => [
      {
        index: 0,
        name: 'Normal',
        xfId: 0,
        builtinId: 0,
        iLevel: 0,
        customBuiltin: false,
      },
      {
        index: 1,
        name: 'Imported Review',
        xfId: 2,
        builtinId: -1,
        iLevel: 0,
        customBuiltin: false,
      },
    ],
    getCellStyleXf: (xfId: number) =>
      xfId === 2
        ? {
            fontIndex: 1,
            fillIndex: 1,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 2,
            verticalAlign: 1,
            wrapText: true,
          }
        : {
            fontIndex: 0,
            fillIndex: 0,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 0,
            verticalAlign: 2,
            wrapText: false,
          },
    getFontRecord: (fontIndex: number) => ({
      name: 'Calibri',
      size: 11,
      bold: fontIndex === 1,
      italic: false,
      strike: false,
      underline: 0,
      colorArgb: fontIndex === 1 ? 0xff006100 : 0xff000000,
    }),
    getFillRecord: (fillIndex: number) => ({
      pattern: fillIndex === 1 ? 1 : 0,
      fgArgb: fillIndex === 1 ? 0xffc6efce : 0,
      bgArgb: 0,
    }),
    getBorderRecord: () => ({
      left: { style: 0, colorArgb: 0xff000000 },
      right: { style: 0, colorArgb: 0xff000000 },
      top: { style: 0, colorArgb: 0xff000000 },
      bottom: { style: 0, colorArgb: 0xff000000 },
      diagonal: { style: 0, colorArgb: 0xff000000 },
      diagonalUp: false,
      diagonalDown: false,
    }),
    getNumFmtCode: () => null,
  }) as unknown as WorkbookHandle;

describe('CELL_STYLES', () => {
  it('contains the spreadsheet-flavored presets', () => {
    const ids = CELL_STYLES.map((s) => s.id);
    expect(ids).toContain('normal');
    expect(ids).toContain('heading1');
    expect(ids).toContain('good');
    expect(ids).toContain('checkCell');
    expect(ids).toContain('explanatoryText');
    expect(ids).toContain('accent1');
    expect(ids).toContain('accent6_20');
    expect(ids).toContain('currency');
  });
});

describe('getCellStyle', () => {
  it('returns the matching def', () => {
    expect(getCellStyle('good')?.format.fill).toBe('#c6efce');
    expect(getCellStyle('accent1')?.format.fill).toBe('#4472c4');
    expect(getCellStyle('accent4_20')?.format.fill).toBe('#fff2cc');
  });

  it('returns undefined for unknown ids', () => {
    // @ts-expect-error — testing invalid id rejection
    expect(getCellStyle('not-a-style')).toBeUndefined();
  });
});

describe('applyCellStyle', () => {
  it('writes the named-style fields onto every cell in range', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, 'good');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.fill).toBe('#c6efce');
    expect(fmt?.color).toBe('#006100');
    expect(fmt?.cellStyle).toBe('good');
    const fmtCorner = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }));
    expect(fmtCorner?.fill).toBe('#c6efce');
    expect(fmtCorner?.cellStyle).toBe('good');
  });

  it('clears every format field for the "normal" preset', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'good');
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'normal');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    // Either the entry is gone, or all visible style fields are undefined.
    if (fmt) {
      expect(fmt.fill).toBeUndefined();
      expect(fmt.color).toBeUndefined();
      expect(fmt.bold).toBeUndefined();
      expect(fmt.cellStyle).toBeUndefined();
    }
  });

  it('applies Excel-style accent and explanatory presets', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'accent5_20');
    let fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#1f4e79', fill: '#ddebf7' });

    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'explanatoryText');
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#7f7f7f', italic: true });
  });

  it('no-ops on unknown style id', () => {
    const store = createSpreadsheetStore();
    const before = store.getState();
    // @ts-expect-error — testing invalid id rejection
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'not-a-style');
    expect(store.getState()).toBe(before);
  });

  it('creates a named style from the active cell format and records history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        fill: '#c6efce',
        color: '#006100',
      },
    );

    expect(
      createCellStyleFromActiveFormat(
        store,
        history,
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        'Review OK',
        { include: { fill: false } },
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
    ).toMatchObject({
      cellStyle: 'Review OK',
      bold: true,
      color: '#006100',
    });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))?.fill,
    ).toBeUndefined();
    expect(listCustomCellStyles(store.getState())).toMatchObject([
      {
        id: customCellStyleId('Review OK'),
        label: 'Review OK',
        format: {
          bold: true,
          color: '#006100',
        },
      },
    ]);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        customCellStyleId('Review OK'),
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 })),
    ).toMatchObject({
      cellStyle: 'Review OK',
      bold: true,
      color: '#006100',
    });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }))?.fill,
    ).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 })),
    ).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(listCustomCellStyles(store.getState())).toEqual([]);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
    ).toBeUndefined();
  });

  it('merges visible workbook named styles into the session custom style registry', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    expect(mergeCellStylesFromWorkbook(store, history, namedStyleWorkbook())).toEqual({
      imported: 1,
      skipped: 1,
    });
    expect(listCustomCellStyles(store.getState())).toMatchObject([
      {
        id: customCellStyleId('Imported Review'),
        label: 'Imported Review',
        format: {
          bold: true,
          color: '#006100',
          fill: '#c6efce',
          align: 'center',
          vAlign: 'middle',
          wrap: true,
        },
      },
    ]);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 },
        customCellStyleId('Imported Review'),
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 })),
    ).toMatchObject({
      cellStyle: 'Imported Review',
      bold: true,
      fill: '#c6efce',
    });
    expect(history.undo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(listCustomCellStyles(store.getState())).toEqual([]);
  });
});
