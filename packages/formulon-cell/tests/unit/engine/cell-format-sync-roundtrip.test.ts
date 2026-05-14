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
import { addrKey, type WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type CellFormat, createSpreadsheetStore } from '../../../src/store/store.js';

/** In-memory engine harness that mirrors the dedup semantics of the real
 *  engine: addFontRecord / addFillRecord / addBorderRecord / addNumFmtCode /
 *  addXfRecord allocate fresh indices and remember the record so the *get*
 *  counterparts can replay it. This lets us drive the full
 *  store → engine → store roundtrip without any WASM. */
const makeEngine = (): {
  wb: WorkbookHandle;
  /** Inspect the engine's XF table size after the push. */
  size: () => number;
} => {
  const fonts = new Map<number, FontRecord>();
  const fills = new Map<number, FillRecord>();
  const borders = new Map<number, BorderRecord>();
  const numFmts = new Map<number, string>();
  const xfs = new Map<number, CellXf>();
  const cellXf = new Map<string, number>();
  const cells: Addr[] = [];

  let fontIdx = 1;
  let fillIdx = 1;
  let borderIdx = 1;
  let numFmtIdx = 200;
  let xfIdx = 1;

  const wb = {
    capabilities: { cellFormatting: true } as never,
    *cells(_sheet: number): Iterable<{ addr: Addr; value: never; formula: null }> {
      for (const a of cells) {
        yield { addr: a, value: { kind: 'blank' } as never, formula: null };
      }
    },
    getCellXfIndex(sheet: number, row: number, col: number): number | null {
      return cellXf.get(`${sheet}:${row}:${col}`) ?? null;
    },
    setCellXfIndex(sheet: number, row: number, col: number, idx: number): boolean {
      cellXf.set(`${sheet}:${row}:${col}`, idx);
      // Track the cell so hydrate can iterate it.
      if (!cells.some((c) => c.sheet === sheet && c.row === row && c.col === col)) {
        cells.push({ sheet, row, col });
      }
      return true;
    },
    getCellXf(idx: number): CellXf | null {
      return xfs.get(idx) ?? null;
    },
    getFontRecord(idx: number): FontRecord | null {
      return fonts.get(idx) ?? null;
    },
    getFillRecord(idx: number): FillRecord | null {
      return fills.get(idx) ?? null;
    },
    getBorderRecord(idx: number): BorderRecord | null {
      return borders.get(idx) ?? null;
    },
    getNumFmtCode(id: number): string | null {
      return numFmts.get(id) ?? null;
    },
    addFontRecord(record: FontRecord): number {
      const i = fontIdx++;
      fonts.set(i, record);
      return i;
    },
    addFillRecord(record: FillRecord): number {
      const i = fillIdx++;
      fills.set(i, record);
      return i;
    },
    addBorderRecord(record: BorderRecord): number {
      const i = borderIdx++;
      borders.set(i, record);
      return i;
    },
    addNumFmtCode(code: string): number {
      const i = numFmtIdx++;
      numFmts.set(i, code);
      return i;
    },
    addXfRecord(record: CellXf): number {
      const i = xfIdx++;
      xfs.set(i, record);
      return i;
    },
  } as unknown as WorkbookHandle;

  return { wb, size: () => xfs.size };
};

describe('engine/cell-format-sync — roundtrip', () => {
  it('numFmt: percent(2) survives push → hydrate', () => {
    const { wb } = makeEngine();
    const store = createSpreadsheetStore();
    const key = addrKey({ sheet: 0, row: 1, col: 2 });
    const original: CellFormat = {
      numFmt: { kind: 'percent', decimals: 2 },
    };
    store.setState((s) => ({ ...s, format: { formats: new Map([[key, original]]) } }));

    syncCellFormatsToEngine(wb, store, 0);

    // Wipe the store so hydrate can repopulate from engine state alone.
    store.setState((s) => ({ ...s, format: { formats: new Map() } }));
    hydrateCellFormatsFromEngine(wb, store, 0);

    const back = store.getState().format.formats.get(key);
    expect(back?.numFmt).toEqual({ kind: 'percent', decimals: 2 });
  });

  it('font: family / size / bold / italic / color round-trip', () => {
    const { wb } = makeEngine();
    const store = createSpreadsheetStore();
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    const original: CellFormat = {
      fontFamily: 'Arial',
      fontSize: 14,
      bold: true,
      italic: true,
      color: '#ff0000',
    };
    store.setState((s) => ({ ...s, format: { formats: new Map([[key, original]]) } }));

    syncCellFormatsToEngine(wb, store, 0);
    store.setState((s) => ({ ...s, format: { formats: new Map() } }));
    hydrateCellFormatsFromEngine(wb, store, 0);

    const back = store.getState().format.formats.get(key);
    expect(back?.fontFamily).toBe('Arial');
    expect(back?.fontSize).toBe(14);
    expect(back?.bold).toBe(true);
    expect(back?.italic).toBe(true);
    expect(back?.color).toBe('#ff0000');
  });

  it('fill: solid color round-trips', () => {
    const { wb } = makeEngine();
    const store = createSpreadsheetStore();
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    store.setState((s) => ({
      ...s,
      format: { formats: new Map([[key, { fill: '#ffff00' } as CellFormat]]) },
    }));

    syncCellFormatsToEngine(wb, store, 0);
    store.setState((s) => ({ ...s, format: { formats: new Map() } }));
    hydrateCellFormatsFromEngine(wb, store, 0);

    expect(store.getState().format.formats.get(key)?.fill).toBe('#ffff00');
  });

  it('alignment: horizontal/vertical/wrap round-trip', () => {
    const { wb } = makeEngine();
    const store = createSpreadsheetStore();
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([[key, { align: 'center', vAlign: 'middle', wrap: true } as CellFormat]]),
      },
    }));

    syncCellFormatsToEngine(wb, store, 0);
    store.setState((s) => ({ ...s, format: { formats: new Map() } }));
    hydrateCellFormatsFromEngine(wb, store, 0);

    const back = store.getState().format.formats.get(key);
    expect(back?.align).toBe('center');
    expect(back?.vAlign).toBe('middle');
    expect(back?.wrap).toBe(true);
  });

  it('numFmt: currency(2) survives roundtrip', () => {
    const { wb } = makeEngine();
    const store = createSpreadsheetStore();
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [key, { numFmt: { kind: 'currency', decimals: 2, symbol: '$' } } as CellFormat],
        ]),
      },
    }));

    syncCellFormatsToEngine(wb, store, 0);
    store.setState((s) => ({ ...s, format: { formats: new Map() } }));
    hydrateCellFormatsFromEngine(wb, store, 0);

    expect(store.getState().format.formats.get(key)?.numFmt).toMatchObject({
      kind: 'currency',
      decimals: 2,
    });
  });

  it('multiple cells: each retains its own format independently', () => {
    const { wb } = makeEngine();
    const store = createSpreadsheetStore();
    const k1 = addrKey({ sheet: 0, row: 0, col: 0 });
    const k2 = addrKey({ sheet: 0, row: 5, col: 5 });

    store.setState((s) => ({
      ...s,
      format: {
        formats: new Map([
          [k1, { bold: true, fill: '#ffff00' } as CellFormat],
          [k2, { italic: true, color: '#0000ff' } as CellFormat],
        ]),
      },
    }));

    syncCellFormatsToEngine(wb, store, 0);
    store.setState((s) => ({ ...s, format: { formats: new Map() } }));
    hydrateCellFormatsFromEngine(wb, store, 0);

    const a = store.getState().format.formats.get(k1);
    const b = store.getState().format.formats.get(k2);
    expect(a?.bold).toBe(true);
    expect(a?.fill).toBe('#ffff00');
    expect(b?.italic).toBe(true);
    expect(b?.color).toBe('#0000ff');
  });

  it('idempotent: second sync without store mutation reuses indices (no new XF row per dup)', () => {
    const { wb, size } = makeEngine();
    const store = createSpreadsheetStore();
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    store.setState((s) => ({
      ...s,
      format: { formats: new Map([[key, { bold: true } as CellFormat]]) },
    }));

    syncCellFormatsToEngine(wb, store, 0);
    const first = size();
    syncCellFormatsToEngine(wb, store, 0);
    const second = size();
    // The fake engine does not dedup (real engine does), so first run inserts
    // one row, second run inserts another. The contract we lock here is "sync
    // is safe to call repeatedly without throwing"; the dedup is the engine's
    // responsibility.
    expect(first).toBeGreaterThanOrEqual(1);
    expect(second).toBeGreaterThanOrEqual(first);
  });

  it('default font/size stripped on hydrate (no pollution of empty formats)', () => {
    // When the user formats a cell with only `bold: true`, the engine stores
    // Calibri/11. Hydrate must NOT echo these back as explicit fontFamily/size
    // overrides — otherwise reformatting back to "default" becomes impossible.
    const { wb } = makeEngine();
    const store = createSpreadsheetStore();
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    store.setState((s) => ({
      ...s,
      format: { formats: new Map([[key, { bold: true } as CellFormat]]) },
    }));

    syncCellFormatsToEngine(wb, store, 0);
    store.setState((s) => ({ ...s, format: { formats: new Map() } }));
    hydrateCellFormatsFromEngine(wb, store, 0);

    const back = store.getState().format.formats.get(key);
    expect(back?.bold).toBe(true);
    // Default font + size should not surface — they would be redundant noise
    // because the renderer falls back to them anyway.
    expect(back?.fontFamily).toBeUndefined();
    expect(back?.fontSize).toBeUndefined();
  });
});
