import { describe, expect, it } from 'vitest';
import {
  phoneticReading,
  phoneticReadingAt,
  setPhoneticReading,
} from '../../../src/commands/phonetic.js';
import type { Addr, PhoneticRun } from '../../../src/engine/types.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { addrKey } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const ADDR: Addr = { sheet: 0, row: 0, col: 0 };

const TOKYO: PhoneticRun[] = [
  { start: 0, end: 2, text: 'とうきょう' },
  { start: 2, end: 3, text: 'と' },
];

/** Stands in for the whole-cell entry point: a written reading replaces the
 *  guide with one run over the cell text, which is what the engine does. */
const makeFake = (opts: { runs?: boolean; text?: string } = {}) => {
  const log: string[] = [];
  let guide: PhoneticRun[] = TOKYO;
  const fake = {
    setCellPhonetic(_s: number, _r: number, _c: number, reading: string): boolean {
      log.push(reading);
      guide = reading ? [{ start: 0, end: (opts.text ?? '東京都').length, text: reading }] : [];
      return true;
    },
    getCellPhoneticRuns(): PhoneticRun[] | null {
      return opts.runs === false ? null : guide;
    },
  };
  return { wb: fake as unknown as WorkbookHandle, log };
};

const storeWithGuide = () => {
  const store = createSpreadsheetStore();
  store.setState((s) => ({
    ...s,
    data: {
      ...s.data,
      cells: new Map([
        [addrKey(ADDR), { value: { kind: 'text' as const, value: '東京都' }, formula: null }],
      ]),
    },
  }));
  mutators.setCellFormat(store, ADDR, { phonetic: TOKYO });
  return store;
};

describe('commands/phonetic', () => {
  it('runs the readings together for a single-field editor', () => {
    expect(phoneticReading(TOKYO)).toBe('とうきょうと');
    expect(phoneticReading(undefined)).toBe('');
    expect(phoneticReadingAt(storeWithGuide(), ADDR)).toBe('とうきょうと');
  });

  it('leaves a partially annotated guide alone when the reading was not edited', () => {
    const store = storeWithGuide();
    const { wb, log } = makeFake();

    expect(setPhoneticReading(store, wb, ADDR, 'とうきょうと', 'とうきょうと')).toBe(false);
    expect(log).toEqual([]);
    expect(store.getState().format.formats.get(addrKey(ADDR))?.phonetic).toEqual(TOKYO);
  });

  it('replaces the guide with the engine’s own view of an edited reading', () => {
    const store = storeWithGuide();
    const { wb, log } = makeFake();

    expect(setPhoneticReading(store, wb, ADDR, 'とうきょうふ', 'とうきょうと')).toBe(true);
    expect(log).toEqual(['とうきょうふ']);
    expect(store.getState().format.formats.get(addrKey(ADDR))?.phonetic).toEqual([
      { start: 0, end: 3, text: 'とうきょうふ' },
    ]);
  });

  it('clears the guide when the reading is emptied', () => {
    const store = storeWithGuide();
    const { wb } = makeFake();

    expect(setPhoneticReading(store, wb, ADDR, '', 'とうきょうと')).toBe(true);
    expect(store.getState().format.formats.get(addrKey(ADDR))?.phonetic).toBeUndefined();
  });

  it('spans the cell text itself when the engine has no run surface', () => {
    const store = storeWithGuide();
    const { wb } = makeFake({ runs: false });

    expect(setPhoneticReading(store, wb, ADDR, 'とうきょうふ', 'とうきょうと')).toBe(true);
    expect(store.getState().format.formats.get(addrKey(ADDR))?.phonetic).toEqual([
      { start: 0, end: 3, text: 'とうきょうふ' },
    ]);
  });
});
