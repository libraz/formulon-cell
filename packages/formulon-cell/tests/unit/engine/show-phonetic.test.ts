import { describe, expect, it } from 'vitest';
import { computeAutofitRowHeight } from '../../../src/commands/autofit-measurement.js';
import { showPhoneticAt, toggleShowPhonetic } from '../../../src/commands/phonetic.js';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import type { Addr } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { canLoadWasm } from './xlsx-roundtrip/fixtures.js';

const A1: Addr = { sheet: 0, row: 0, col: 0 };
const A2: Addr = { sheet: 0, row: 1, col: 0 };
const A3: Addr = { sheet: 0, row: 2, col: 0 };
const A4: Addr = { sheet: 0, row: 3, col: 0 };
const RUNS = [{ start: 0, end: 2, text: 'かんじ' }];

const select = (store: ReturnType<typeof createSpreadsheetStore>, active: Addr, last: Addr) => {
  mutators.setActive(store, active);
  mutators.extendRangeTo(store, last);
};

describe.skipIf(!canLoadWasm())('Show Phonetic', () => {
  it('survives save and reload through the format sync', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.showPhonetic).toBe(true);
      wb.setText(A1, '漢字');
      wb.setText(A2, '東京');
      const store = createSpreadsheetStore();
      mutators.replaceCells(store, wb.cells(0));
      mutators.setCellFormat(store, A1, { phonetic: RUNS, showPhonetic: true });
      mutators.setCellFormat(store, A2, { phonetic: RUNS });
      syncCellFormatsToEngine(wb, store, 0);
      expect(wb.getCellShowPhonetic(0, 0, 0)).toBe(true);
      expect(wb.getCellShowPhonetic(0, 1, 0)).toBe(false);

      const reloaded = await WorkbookHandle.loadBytes(wb.save());
      try {
        const next = createSpreadsheetStore();
        hydrateCellFormatsFromEngine(reloaded, next, 0);
        const formats = next.getState().format.formats;
        expect(formats.get(addrKey(A1))?.showPhonetic).toBe(true);
        expect(formats.get(addrKey(A1))?.phonetic).toEqual(RUNS);
        expect(formats.get(addrKey(A2))?.showPhonetic).toBeUndefined();
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps the flag on a blank cell through save and reload', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, A3, { showPhonetic: true });
      syncCellFormatsToEngine(wb, store, 0);
      const reloaded = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(reloaded.getCellShowPhonetic(0, 2, 0)).toBe(true);
        const next = createSpreadsheetStore();
        hydrateCellFormatsFromEngine(reloaded, next, 0);
        expect(next.getState().format.formats.get(addrKey(A3))?.showPhonetic).toBe(true);
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('toggles over a selection by the active cell and skips formula cells', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      wb.setText(A1, '漢字');
      wb.setText(A2, '東京');
      wb.setFormula(A3, '=A1');
      const store = createSpreadsheetStore();
      mutators.replaceCells(store, wb.cells(0));
      select(store, A1, A4);

      expect(toggleShowPhonetic(store, wb)).toBe(true);
      const formats = () => store.getState().format.formats;
      expect(formats().get(addrKey(A1))?.showPhonetic).toBe(true);
      expect(formats().get(addrKey(A2))?.showPhonetic).toBe(true);
      expect(formats().get(addrKey(A3))?.showPhonetic).toBeUndefined();
      expect(showPhoneticAt(store, A1)).toBe(true);

      syncCellFormatsToEngine(wb, store, 0);
      expect(wb.getCellShowPhonetic(0, 1, 0)).toBe(true);
      expect(wb.getCellShowPhonetic(0, 2, 0)).toBe(false);

      // The active cell is now on, so the same command turns the selection off.
      expect(toggleShowPhonetic(store, wb)).toBe(true);
      expect(showPhoneticAt(store, A1)).toBe(false);
      expect(showPhoneticAt(store, A2)).toBe(false);
    } finally {
      wb.dispose();
    }
  });
});

describe('Show Phonetic without an engine surface', () => {
  it('does not toggle', () => {
    const store = createSpreadsheetStore();
    const wb = { capabilities: {} } as unknown as WorkbookHandle;
    expect(toggleShowPhonetic(store, wb)).toBe(false);
  });
});

describe('Show Phonetic row height', () => {
  const rowHeight = (showPhonetic: boolean) => {
    const store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      data: {
        ...s.data,
        cells: new Map([
          [addrKey(A1), { value: { kind: 'text' as const, value: '漢字' }, formula: null }],
        ]),
      },
    }));
    mutators.setCellFormat(store, A1, { phonetic: RUNS, showPhonetic });
    return computeAutofitRowHeight(store.getState(), 0, null);
  };

  it('reserves room for a shown guide only', () => {
    expect(rowHeight(true)).toBeGreaterThan(rowHeight(false));
  });
});
