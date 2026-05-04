import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { pasteTSV } from '../../../src/commands/clipboard/paste.js';
import { writeInput, writeInputValidated } from '../../../src/commands/coerce-input.js';
import { setFillColor, setNumFmt, toggleBold } from '../../../src/commands/format.js';
import {
  gateProtection,
  isCellLocked,
  isCellWritable,
  isSheetProtected,
  setCellLocked,
} from '../../../src/commands/protection.js';
import { deleteRows, insertRows } from '../../../src/commands/structure.js';
import type { Addr, Range } from '../../../src/engine/types.js';
import { WorkbookHandle, addrKey } from '../../../src/engine/workbook-handle.js';
import {
  type SpreadsheetStore,
  createSpreadsheetStore,
  mutators,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const stubHandle = (): WorkbookHandle & {
  setBlank: ReturnType<typeof vi.fn>;
  setFormula: ReturnType<typeof vi.fn>;
  setNumber: ReturnType<typeof vi.fn>;
  setBool: ReturnType<typeof vi.fn>;
  setText: ReturnType<typeof vi.fn>;
} => {
  return {
    setBlank: vi.fn(),
    setFormula: vi.fn(),
    setNumber: vi.fn(),
    setBool: vi.fn(),
    setText: vi.fn(),
    capabilities: { insertDeleteRowsCols: false },
    cells: () => [] as Iterable<{ addr: Addr; value: { kind: 'blank' }; formula: null }>,
    recalc: vi.fn(),
  } as unknown as WorkbookHandle & {
    setBlank: ReturnType<typeof vi.fn>;
    setFormula: ReturnType<typeof vi.fn>;
    setNumber: ReturnType<typeof vi.fn>;
    setBool: ReturnType<typeof vi.fn>;
    setText: ReturnType<typeof vi.fn>;
  };
};

const setRange = (store: SpreadsheetStore, range: Range): void => {
  store.setState((s) => ({ ...s, selection: { ...s.selection, range } }));
};

const fmtAt = (store: SpreadsheetStore, row: number, col: number) =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

describe('protection helpers', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('isSheetProtected reflects the slice', () => {
    expect(isSheetProtected(store.getState(), 0)).toBe(false);
    mutators.setSheetProtected(store, 0, true);
    expect(isSheetProtected(store.getState(), 0)).toBe(true);
    mutators.setSheetProtected(store, 0, false);
    expect(isSheetProtected(store.getState(), 0)).toBe(false);
  });

  it('isCellLocked defaults to true (Excel default)', () => {
    const a: Addr = { sheet: 0, row: 0, col: 0 };
    expect(isCellLocked(store.getState(), a)).toBe(true);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    expect(isCellLocked(store.getState(), a)).toBe(false);
  });

  it('isCellWritable bypasses the gate when sheet is unprotected', () => {
    const a: Addr = { sheet: 0, row: 0, col: 0 };
    // Sheet not protected → writable regardless of locked flag.
    expect(isCellWritable(store.getState(), a)).toBe(true);
    mutators.setSheetProtected(store, 0, true);
    // Now protected, default-locked → not writable.
    expect(isCellWritable(store.getState(), a)).toBe(false);
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    // Explicitly unlocked → writable.
    expect(isCellWritable(store.getState(), a)).toBe(true);
  });

  it('gateProtection returns null when entire range is locked on protected sheet', () => {
    mutators.setSheetProtected(store, 0, true);
    const range: Range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    expect(gateProtection(store.getState(), range)).toBeNull();
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    // Now A1 is unlocked → range survives the gate.
    expect(gateProtection(store.getState(), range)).toEqual(range);
  });
});

describe('writeInput protection gate', () => {
  let store: SpreadsheetStore;
  let warnSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    store = createSpreadsheetStore();
    warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  });

  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('blocks writes to locked cells when sheet is protected', () => {
    const wb = stubHandle();
    mutators.setSheetProtected(store, 0, true);
    writeInput(wb, { sheet: 0, row: 0, col: 0 }, '42', store);
    expect(wb.setNumber).not.toHaveBeenCalled();
    expect(warnSpy).toHaveBeenCalledTimes(1);
  });

  it('writes through when the cell is explicitly unlocked', () => {
    const wb = stubHandle();
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    mutators.setSheetProtected(store, 0, true);
    writeInput(wb, { sheet: 0, row: 0, col: 0 }, '42', store);
    expect(wb.setNumber).toHaveBeenCalledWith({ sheet: 0, row: 0, col: 0 }, 42);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  it('writes through when the sheet is unprotected (default state)', () => {
    const wb = stubHandle();
    writeInput(wb, { sheet: 0, row: 0, col: 0 }, '7', store);
    expect(wb.setNumber).toHaveBeenCalledWith({ sheet: 0, row: 0, col: 0 }, 7);
  });

  it('writeInputValidated honors the gate', () => {
    const wb = stubHandle();
    mutators.setSheetProtected(store, 0, true);
    const outcome = writeInputValidated(wb, { sheet: 0, row: 0, col: 0 }, '42', undefined, store);
    expect(outcome.ok).toBe(true);
    expect(wb.setNumber).not.toHaveBeenCalled();
    expect(warnSpy).toHaveBeenCalledTimes(1);
  });
});

describe('format mutators on protected sheet', () => {
  let store: SpreadsheetStore;

  beforeEach(() => {
    store = createSpreadsheetStore();
  });

  it('skips bold toggle wholesale when entire range is locked on protected sheet', () => {
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    mutators.setSheetProtected(store, 0, true);
    toggleBold(store.getState(), store);
    // No format entries should have been written.
    expect(fmtAt(store, 0, 0)?.bold).toBeUndefined();
    expect(fmtAt(store, 1, 1)?.bold).toBeUndefined();
  });

  it('applies fill only to explicitly-unlocked cells inside the range', () => {
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    setCellLocked(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, false);
    mutators.setSheetProtected(store, 0, true);
    setFillColor(store.getState(), store, '#abcdef');
    // A1 is unlocked → fill written.
    expect(fmtAt(store, 0, 0)?.fill).toBe('#abcdef');
    // B2 stays locked → no fill.
    expect(fmtAt(store, 1, 1)?.fill).toBeUndefined();
  });

  it('passes formats through normally when sheet is unprotected', () => {
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    setNumFmt(store.getState(), store, { kind: 'fixed', decimals: 2 });
    expect(fmtAt(store, 0, 0)?.numFmt).toEqual({ kind: 'fixed', decimals: 2 });
  });
});

describe('structure mutators on protected sheet', () => {
  let store: SpreadsheetStore;
  let warnSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    store = createSpreadsheetStore();
    warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
  });

  afterEach(() => {
    warnSpy.mockRestore();
  });

  it('rejects insertRows / deleteRows on protected sheet', () => {
    const wb = stubHandle();
    mutators.setSheetProtected(store, 0, true);
    insertRows(store, wb, null, 1, 1);
    deleteRows(store, wb, null, 1, 1);
    // recalc should never be reached when blocked.
    expect(wb.recalc).not.toHaveBeenCalled();
    expect(warnSpy).toHaveBeenCalledTimes(2);
  });
});

describe('pasteTSV protection gate', () => {
  it('skips locked destinations silently while writing through unlocked ones', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    // Activate paste at A1; row spans A1..C1.
    setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    store.setState((s) => ({
      ...s,
      selection: { ...s.selection, active: { sheet: 0, row: 0, col: 0 } },
    }));
    // Unlock B1 only.
    setCellLocked(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }, false);
    mutators.setSheetProtected(store, 0, true);
    const result = pasteTSV(store.getState(), wb, 'foo\tbar\tbaz');
    wb.recalc();
    // A1 (locked) — engine cell should remain blank.
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 }).kind).toBe('blank');
    // B1 (unlocked) — paste landed.
    const b1 = wb.getValue({ sheet: 0, row: 0, col: 1 });
    expect(b1.kind === 'text' && b1.value).toBe('bar');
    // C1 (locked) — blank.
    expect(wb.getValue({ sheet: 0, row: 0, col: 2 }).kind).toBe('blank');
    expect(result?.writtenRange).toBeDefined();
  });
});
