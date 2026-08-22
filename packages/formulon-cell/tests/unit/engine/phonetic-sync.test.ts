import { describe, expect, it } from 'vitest';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import type { Addr, CellValue, PhoneticRun } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

interface PhoneticLog {
  setRuns: Array<{ addr: string; runs: readonly PhoneticRun[] }>;
  setWhole: Array<{ addr: string; reading: string }>;
}

/** Minimal engine stand-in for the phonetic leg of the format sync: the XF
 *  table is stubbed out to a single index so only the guide is under test. */
const makeFake = (opts: {
  runs?: boolean;
  cells?: Array<{ addr: Addr; value: CellValue }>;
  guides?: Map<string, PhoneticRun[]>;
}): { wb: WorkbookHandle; log: PhoneticLog } => {
  const log: PhoneticLog = { setRuns: [], setWhole: [] };
  const guides = opts.guides ?? new Map<string, PhoneticRun[]>();
  const values = new Map((opts.cells ?? []).map((c) => [addrKey(c.addr), c.value]));
  const fake = {
    capabilities: { cellFormatting: true, phonetic: true, phoneticRuns: opts.runs === true },
    *cells(_sheet: number) {
      for (const c of opts.cells ?? []) yield { addr: c.addr, value: c.value, formula: null };
    },
    getValue(addr: Addr): CellValue {
      return values.get(addrKey(addr)) ?? { kind: 'blank' };
    },
    getCellXfIndex: () => 0,
    setCellXfIndex: () => true,
    getCellXf: () => null,
    getFontRecord: () => null,
    getFillRecord: () => null,
    getBorderRecord: () => null,
    getNumFmtCode: () => null,
    addFontRecord: () => 1,
    addFillRecord: () => 1,
    addBorderRecord: () => 1,
    addNumFmtCode: () => 1,
    addXfRecord: () => 1,
    addCellStyleXfRecord: () => 1,
    get workbookDefaultFont() {
      return null;
    },
    getCellPhonetic(sheet: number, row: number, col: number): string | null {
      const runs = guides.get(`${sheet}:${row}:${col}`);
      return runs === undefined ? null : runs.map((run) => run.text).join('');
    },
    setCellPhonetic(sheet: number, row: number, col: number, reading: string): boolean {
      log.setWhole.push({ addr: `${sheet}:${row}:${col}`, reading });
      return true;
    },
    getCellPhoneticRuns(sheet: number, row: number, col: number): PhoneticRun[] | null {
      return guides.get(`${sheet}:${row}:${col}`) ?? [];
    },
    setCellPhoneticRuns(
      sheet: number,
      row: number,
      col: number,
      runs: readonly PhoneticRun[],
    ): boolean {
      log.setRuns.push({ addr: `${sheet}:${row}:${col}`, runs });
      return true;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, log };
};

const TOKYO: PhoneticRun[] = [
  { start: 0, end: 2, text: 'とうきょう' },
  { start: 2, end: 3, text: 'と' },
];

describe('engine/cell-format-sync — phonetic guides', () => {
  it('writes a partially annotated guide span by span', () => {
    const { wb, log } = makeFake({
      runs: true,
      cells: [{ addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'text', value: '東京都' } }],
    });
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { phonetic: TOKYO });
    syncCellFormatsToEngine(wb, store, 0);

    expect(log.setRuns).toEqual([{ addr: '0:0:0', runs: TOKYO }]);
    // The whole-cell entry point would have collapsed the two spans into one.
    expect(log.setWhole).toEqual([]);
  });

  it('falls back to the concatenated reading when the engine has no run surface', () => {
    const { wb, log } = makeFake({ runs: false });
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { phonetic: TOKYO });
    syncCellFormatsToEngine(wb, store, 0);

    expect(log.setRuns).toEqual([]);
    expect(log.setWhole).toEqual([{ addr: '0:0:0', reading: 'とうきょうと' }]);
  });

  it('clears the guide of a cell whose format no longer carries one', () => {
    const { wb, log } = makeFake({ runs: true });
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
    syncCellFormatsToEngine(wb, store, 0);

    expect(log.setRuns).toEqual([{ addr: '0:0:0', runs: [] }]);
  });

  it('trims a guide pasted onto a shorter cell to the text it covers', () => {
    const { wb, log } = makeFake({
      runs: true,
      cells: [{ addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'text', value: '東京' } }],
    });
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { phonetic: TOKYO });
    syncCellFormatsToEngine(wb, store, 0);

    // The engine takes an out-of-range span verbatim and writes it into the
    // file, so the run past the end of '東京' has to be dropped here.
    expect(log.setRuns).toEqual([
      { addr: '0:0:0', runs: [{ start: 0, end: 2, text: 'とうきょう' }] },
    ]);
  });

  it('drops a guide carried onto a cell that holds no text', () => {
    const { wb, log } = makeFake({
      runs: true,
      cells: [{ addr: { sheet: 0, row: 0, col: 0 }, value: { kind: 'number', value: 42 } }],
    });
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { phonetic: TOKYO });
    syncCellFormatsToEngine(wb, store, 0);

    expect(log.setRuns).toEqual([{ addr: '0:0:0', runs: [] }]);
  });

  it('hydrates the spans an engine reports rather than one flattened reading', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    const { wb } = makeFake({
      runs: true,
      cells: [{ addr, value: { kind: 'text', value: '東京都' } }],
      guides: new Map([['0:0:0', TOKYO]]),
    });
    const store = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(wb, store, 0);

    expect(store.getState().format.formats.get(addrKey(addr))?.phonetic).toEqual(TOKYO);
  });

  it('spans the whole cell text when the engine reports only a reading', () => {
    const addr = { sheet: 0, row: 0, col: 0 };
    const { wb } = makeFake({
      runs: false,
      cells: [{ addr, value: { kind: 'text', value: '東京都' } }],
      guides: new Map([['0:0:0', TOKYO]]),
    });
    const store = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(wb, store, 0);

    expect(store.getState().format.formats.get(addrKey(addr))?.phonetic).toEqual([
      { start: 0, end: 3, text: 'とうきょうと' },
    ]);
  });
});
