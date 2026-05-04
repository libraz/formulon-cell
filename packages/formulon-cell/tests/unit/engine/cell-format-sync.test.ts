import { describe, expect, it } from 'vitest';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import type {
  Addr,
  BorderRecord,
  CellXf,
  FillRecord,
  FontRecord,
} from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { type CellFormat, createSpreadsheetStore } from '../../../src/store/store.js';

interface FakeOpts {
  cellFormatting?: boolean;
  /** Pre-populated cells the engine reports via cells(). */
  cells?: Addr[];
  /** XF index per cell, keyed by addrKey. */
  xfIndices?: Map<string, number>;
  /** XF table — index → record. */
  xfTable?: Map<number, CellXf>;
  fonts?: Map<number, FontRecord>;
  fills?: Map<number, FillRecord>;
  borders?: Map<number, BorderRecord>;
  numFmts?: Map<number, string>;
}

interface FakeLog {
  setCellXf: { sheet: number; row: number; col: number; xfIndex: number }[];
  addFont: FontRecord[];
  addFill: FillRecord[];
  addBorder: BorderRecord[];
  addNumFmt: string[];
  addXf: CellXf[];
}

const makeFake = (opts: FakeOpts = {}): { wb: WorkbookHandle; log: FakeLog } => {
  const log: FakeLog = {
    setCellXf: [],
    addFont: [],
    addFill: [],
    addBorder: [],
    addNumFmt: [],
    addXf: [],
  };
  const caps = { cellFormatting: opts.cellFormatting ?? true };
  const xfIndices = opts.xfIndices ?? new Map<string, number>();
  const xfTable = opts.xfTable ?? new Map<number, CellXf>();
  const fonts = opts.fonts ?? new Map<number, FontRecord>();
  const fills = opts.fills ?? new Map<number, FillRecord>();
  const borders = opts.borders ?? new Map<number, BorderRecord>();
  const numFmts = opts.numFmts ?? new Map<number, string>();

  // Counters drive freshly-allocated indices for added records.
  let nextFontIdx = 100;
  let nextFillIdx = 100;
  let nextBorderIdx = 100;
  let nextXfIdx = 100;
  let nextNumFmtId = 200;

  const fake = {
    capabilities: caps,
    *cells(_sheet: number) {
      for (const a of opts.cells ?? []) {
        yield { addr: a, value: { kind: 'blank' as const }, formula: null };
      }
    },
    getCellXfIndex(_sheet: number, row: number, col: number): number | null {
      if (!caps.cellFormatting) return null;
      return xfIndices.get(`${_sheet}:${row}:${col}`) ?? null;
    },
    setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): boolean {
      if (!caps.cellFormatting) return false;
      log.setCellXf.push({ sheet, row, col, xfIndex });
      return true;
    },
    getCellXf(xfIndex: number) {
      if (!caps.cellFormatting) return null;
      return xfTable.get(xfIndex) ?? null;
    },
    getFontRecord(idx: number) {
      if (!caps.cellFormatting) return null;
      return fonts.get(idx) ?? null;
    },
    getFillRecord(idx: number) {
      if (!caps.cellFormatting) return null;
      return fills.get(idx) ?? null;
    },
    getBorderRecord(idx: number) {
      if (!caps.cellFormatting) return null;
      return borders.get(idx) ?? null;
    },
    getNumFmtCode(numFmtId: number) {
      if (!caps.cellFormatting) return null;
      return numFmts.get(numFmtId) ?? null;
    },
    addFontRecord(record: FontRecord): number {
      if (!caps.cellFormatting) return -1;
      log.addFont.push(record);
      const idx = nextFontIdx++;
      fonts.set(idx, record);
      return idx;
    },
    addFillRecord(record: FillRecord): number {
      if (!caps.cellFormatting) return -1;
      log.addFill.push(record);
      const idx = nextFillIdx++;
      fills.set(idx, record);
      return idx;
    },
    addBorderRecord(record: BorderRecord): number {
      if (!caps.cellFormatting) return -1;
      log.addBorder.push(record);
      const idx = nextBorderIdx++;
      borders.set(idx, record);
      return idx;
    },
    addNumFmtCode(code: string): number {
      if (!caps.cellFormatting) return -1;
      log.addNumFmt.push(code);
      const id = nextNumFmtId++;
      numFmts.set(id, code);
      return id;
    },
    addXfRecord(record: CellXf): number {
      if (!caps.cellFormatting) return -1;
      log.addXf.push(record);
      const idx = nextXfIdx++;
      xfTable.set(idx, record);
      return idx;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, log };
};

describe('syncCellFormatsToEngine', () => {
  it('writes cell XF for each formatted cell', () => {
    const { wb, log } = makeFake();
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [addrKey({ sheet: 0, row: 0, col: 0 }), { bold: true, fill: '#ffff00' } as CellFormat],
        ]),
      },
    }));
    syncCellFormatsToEngine(wb, store, 0);
    expect(log.addFont).toHaveLength(1);
    expect(log.addFont[0]?.bold).toBe(true);
    expect(log.addFill).toHaveLength(1);
    expect(log.addFill[0]?.pattern).toBe(1);
    expect(log.addXf).toHaveLength(1);
    expect(log.setCellXf).toHaveLength(1);
    expect(log.setCellXf[0]).toMatchObject({ sheet: 0, row: 0, col: 0 });
  });

  it('skips cells on other sheets', () => {
    const { wb, log } = makeFake();
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[addrKey({ sheet: 1, row: 0, col: 0 }), { bold: true } as CellFormat]]),
      },
    }));
    syncCellFormatsToEngine(wb, store, 0);
    expect(log.setCellXf).toHaveLength(0);
  });

  it('no-op when capability is off', () => {
    const { wb, log } = makeFake({ cellFormatting: false });
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[addrKey({ sheet: 0, row: 0, col: 0 }), { bold: true } as CellFormat]]),
      },
    }));
    syncCellFormatsToEngine(wb, store, 0);
    expect(log.setCellXf).toHaveLength(0);
    expect(log.addFont).toHaveLength(0);
  });

  it('writes a numFmt entry for non-general formats', () => {
    const { wb, log } = makeFake();
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [
            addrKey({ sheet: 0, row: 0, col: 0 }),
            { numFmt: { kind: 'percent' as const, decimals: 2 } } as CellFormat,
          ],
        ]),
      },
    }));
    syncCellFormatsToEngine(wb, store, 0);
    expect(log.addNumFmt).toEqual(['0.00%']);
  });

  it('skips numFmt registration for general format', () => {
    const { wb, log } = makeFake();
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [
            addrKey({ sheet: 0, row: 0, col: 0 }),
            { numFmt: { kind: 'general' as const } } as CellFormat,
          ],
        ]),
      },
    }));
    syncCellFormatsToEngine(wb, store, 0);
    expect(log.addNumFmt).toEqual([]);
  });
});

describe('hydrateCellFormatsFromEngine', () => {
  it('reads font + fill back into FormatSlice', () => {
    const xfTable = new Map<number, CellXf>([
      [
        7,
        {
          fontIndex: 1,
          fillIndex: 2,
          borderIndex: 3,
          numFmtId: 0,
          horizontalAlign: 2,
          verticalAlign: 0,
          wrapText: true,
        },
      ],
    ]);
    const fonts = new Map<number, FontRecord>([
      [
        1,
        {
          name: 'Calibri',
          size: 11,
          bold: true,
          italic: false,
          strike: false,
          underline: 0,
          colorArgb: 0xff000000,
        },
      ],
    ]);
    const fills = new Map<number, FillRecord>([[2, { pattern: 1, fgArgb: 0xffff00ff, bgArgb: 0 }]]);
    const borders = new Map<number, BorderRecord>([
      [
        3,
        {
          left: { style: 0, colorArgb: 0xff000000 },
          right: { style: 0, colorArgb: 0xff000000 },
          top: { style: 0, colorArgb: 0xff000000 },
          bottom: { style: 0, colorArgb: 0xff000000 },
          diagonal: { style: 0, colorArgb: 0xff000000 },
          diagonalDown: false,
          diagonalUp: false,
        },
      ],
    ]);
    const xfIndices = new Map<string, number>([['0:0:0', 7]]);
    const { wb } = makeFake({
      cells: [{ sheet: 0, row: 0, col: 0 }],
      xfIndices,
      xfTable,
      fonts,
      fills,
      borders,
    });
    const store = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(wb, store, 0);
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.bold).toBe(true);
    expect(fmt?.fill).toBe('#ff00ff');
    expect(fmt?.align).toBe('center');
    expect(fmt?.vAlign).toBe('top');
    expect(fmt?.wrap).toBe(true);
  });

  it('skips cells whose xfIndex is 0', () => {
    const xfIndices = new Map<string, number>([['0:0:0', 0]]);
    const { wb } = makeFake({ cells: [{ sheet: 0, row: 0, col: 0 }], xfIndices });
    const store = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });

  it('no-op when capability is off', () => {
    const { wb } = makeFake({
      cellFormatting: false,
      cells: [{ sheet: 0, row: 0, col: 0 }],
    });
    const store = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(wb, store, 0);
    expect(store.getState().format.formats.size).toBe(0);
  });
});
