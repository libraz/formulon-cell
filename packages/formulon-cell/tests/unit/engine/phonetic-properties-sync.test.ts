import { describe, expect, it } from 'vitest';
import { phoneticSettingsAt, setPhoneticSettings } from '../../../src/commands/phonetic.js';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import {
  phoneticAlignment,
  phoneticKanaForm,
  phoneticPropertiesFor,
} from '../../../src/engine/phonetic-sync.js';
import type {
  Addr,
  CellValue,
  PhoneticProperties,
  PhoneticRun,
} from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const A1: Addr = { sheet: 0, row: 0, col: 0 };
const B1: Addr = { sheet: 0, row: 0, col: 1 };
const RUNS: PhoneticRun[] = [{ start: 0, end: 3, text: 'とうきょうと' }];
const FULL_LEFT: PhoneticProperties = { fontId: 1, type: 1, alignment: 1 };

/** Engine stand-in with a phonetic property surface and a stubbed XF table. */
const makeFake = (
  opts: { properties?: boolean; stored?: Map<string, PhoneticProperties> } = {},
) => {
  const stored = opts.stored ?? new Map<string, PhoneticProperties>();
  const writes: Array<{ addr: string; properties: PhoneticProperties }> = [];
  const cells = [A1].map((addr) => ({
    addr,
    value: { kind: 'text', value: '東京都' } as CellValue,
    formula: null,
  }));
  const fake = {
    capabilities: {
      cellFormatting: true,
      phonetic: true,
      phoneticRuns: true,
      phoneticProperties: opts.properties !== false,
    },
    *cells() {
      yield* cells;
    },
    getValue: () => ({ kind: 'text', value: '東京都' }) as CellValue,
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
    getCellPhoneticRuns: () => RUNS,
    setCellPhoneticRuns: () => true,
    getCellPhoneticProperties: (s: number, r: number, c: number) =>
      stored.get(`${s}:${r}:${c}`) ?? { fontId: 0, type: 0, alignment: 0 },
    setCellPhoneticProperties(s: number, r: number, c: number, properties: PhoneticProperties) {
      writes.push({ addr: `${s}:${r}:${c}`, properties });
      stored.set(`${s}:${r}:${c}`, properties);
      return true;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, writes };
};

describe('engine/phonetic-sync — property codes', () => {
  it('maps the engine codes to kana forms and alignments', () => {
    expect(phoneticKanaForm({ fontId: 0, type: 0, alignment: 0 })).toBe('halfKatakana');
    expect(phoneticKanaForm({ fontId: 0, type: 1, alignment: 0 })).toBe('katakana');
    expect(phoneticKanaForm({ fontId: 0, type: 2, alignment: 0 })).toBe('hiragana');
    expect(phoneticKanaForm(undefined)).toBe('katakana');
    expect(phoneticAlignment({ fontId: 0, type: 0, alignment: 3 })).toBe('distributed');
    expect(phoneticAlignment(undefined)).toBe('left');
  });

  it('builds engine properties and keeps the font', () => {
    expect(phoneticPropertiesFor('hiragana', 'center', 1)).toEqual({
      fontId: 1,
      type: 2,
      alignment: 2,
    });
  });
});

describe('engine/cell-format-sync — phonetic properties', () => {
  it('writes the properties after the runs', () => {
    const { wb, writes } = makeFake();
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { phonetic: RUNS, phoneticProperties: FULL_LEFT });
    syncCellFormatsToEngine(wb, store, 0);
    expect(writes).toEqual([{ addr: '0:0:0', properties: FULL_LEFT }]);
  });

  it('writes nothing for a cell with properties but no guide', () => {
    const { wb, writes } = makeFake();
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { phoneticProperties: FULL_LEFT });
    syncCellFormatsToEngine(wb, store, 0);
    expect(writes).toEqual([]);
  });

  it('skips the write when the engine has no property surface', () => {
    const { wb, writes } = makeFake({ properties: false });
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { phonetic: RUNS, phoneticProperties: FULL_LEFT });
    syncCellFormatsToEngine(wb, store, 0);
    expect(writes).toEqual([]);
  });

  it('hydrates stored properties and treats the all-zero triple as unset', () => {
    const stored = new Map([['0:0:0', FULL_LEFT]]);
    const store = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(makeFake({ stored }).wb, store, 0);
    expect(store.getState().format.formats.get(addrKey(A1))?.phoneticProperties).toEqual(FULL_LEFT);

    const blank = createSpreadsheetStore();
    hydrateCellFormatsFromEngine(makeFake().wb, blank, 0);
    const format = blank.getState().format.formats.get(addrKey(A1));
    expect(format?.phonetic).toEqual(RUNS);
    expect(format?.phoneticProperties).toBeUndefined();
  });
});

describe('commands/phonetic — settings', () => {
  const range = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };

  it('opens on full-width katakana, left aligned, when nothing is set', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { phonetic: RUNS });
    expect(phoneticSettingsAt(store, A1)).toEqual({ form: 'katakana', alignment: 'left' });
  });

  it('applies to guided cells in the range and keeps their font', () => {
    const { wb, writes } = makeFake();
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { phonetic: RUNS, phoneticProperties: FULL_LEFT });
    mutators.setCellFormat(store, B1, { bold: true });
    const changed = setPhoneticSettings(
      store,
      wb,
      { ...range, c1: 1 },
      {
        form: 'hiragana',
        alignment: 'distributed',
      },
    );
    expect(changed).toBe(1);
    expect(writes).toEqual([{ addr: '0:0:0', properties: { fontId: 1, type: 2, alignment: 3 } }]);
    expect(store.getState().format.formats.get(addrKey(A1))?.phoneticProperties).toEqual({
      fontId: 1,
      type: 2,
      alignment: 3,
    });
    expect(phoneticSettingsAt(store, A1)).toEqual({ form: 'hiragana', alignment: 'distributed' });
  });

  it('leaves cells outside the range alone', () => {
    const { wb, writes } = makeFake();
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, B1, { phonetic: RUNS });
    expect(setPhoneticSettings(store, wb, range, { form: 'katakana', alignment: 'left' })).toBe(0);
    expect(writes).toEqual([]);
  });

  it('does nothing without an engine property surface', () => {
    const { wb } = makeFake({ properties: false });
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, A1, { phonetic: RUNS });
    expect(setPhoneticSettings(store, wb, range, { form: 'katakana', alignment: 'left' })).toBe(0);
  });
});

describe('new phonetic guides', () => {
  it('states the desktop default instead of the engine unset state', async () => {
    const { WorkbookHandle } = await import('../../../src/engine/workbook-handle.js');
    const { createSpreadsheetStore, mutators } = await import('../../../src/store/store.js');
    const { setPhoneticReading } = await import('../../../src/commands/phonetic.js');
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 0, col: 0 };
    wb.setText(addr, '漢字');
    mutators.replaceCells(store, wb.cells(0));
    expect(setPhoneticReading(store, wb, addr, 'カンジ', '')).toBe(true);
    expect(wb.getCellPhoneticProperties(0, 0, 0)).toEqual({ fontId: 0, type: 1, alignment: 1 });
    wb.dispose();
  });
});
