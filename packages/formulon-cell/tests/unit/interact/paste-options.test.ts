import { afterEach, beforeEach, describe, expect, it, type Mock, vi } from 'vitest';
import { pasteSpecial } from '../../../src/commands/clipboard/paste-special.js';
import { captureSnapshot } from '../../../src/commands/clipboard/snapshot.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { attachPasteOptions } from '../../../src/interact/paste-options.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

const seedNumber = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  row: number,
  col: number,
  value: number,
  format?: CellFormat,
): void => {
  const addr = { sheet: 0, row, col };
  wb.setNumber(addr, value);
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    const formats = new Map(s.format.formats);
    cells.set(addrKey(addr), { value: { kind: 'number', value }, formula: null });
    if (format) formats.set(addrKey(addr), format);
    return { ...s, data: { ...s.data, cells }, format: { formats } };
  });
};

const formatAt = (store: SpreadsheetStore, row: number, col: number): CellFormat | undefined =>
  store.getState().format.formats.get(addrKey({ sheet: 0, row, col }));

describe('attachPasteOptions', () => {
  let host: HTMLElement;
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: Mock<() => void>;

  beforeEach(async () => {
    host = document.createElement('div');
    grid = document.createElement('div');
    host.appendChild(grid);
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('opens a Paste Options smart button and can switch to Values', () => {
    seedNumber(store, wb, 0, 0, 7, { bold: true, fill: '#fff2cc' });
    seedNumber(store, wb, 2, 2, 2, { italic: true, fill: '#ddebf7' });
    const source = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const before = captureSnapshot(store.getState(), { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    expect(source).not.toBeNull();
    expect(before).not.toBeNull();
    if (!source || !before) throw new Error('missing clipboard snapshots');
    setRange(store, 2, 2, 2, 2);
    pasteSpecial(store.getState(), store, wb, source, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    mutators.replaceCells(store, wb.cells(0));

    const handle = attachPasteOptions({
      host,
      grid,
      store,
      wb,
      strings: defaultStrings,
      onAfterCommit,
    });
    handle.show({
      source,
      before,
      range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
    });

    const button = document.querySelector<HTMLButtonElement>('.fc-paste-options__button');
    expect(button?.style.display).toBe('block');
    button?.click();
    document.querySelector<HTMLButtonElement>('[data-fc-mode="values"]')?.click();

    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 7 });
    expect(formatAt(store, 2, 2)).toEqual({ italic: true, fill: '#ddebf7' });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });

  it('can switch the pasted result to Formatting Only', () => {
    seedNumber(store, wb, 0, 0, 7, { bold: true, fill: '#fff2cc' });
    seedNumber(store, wb, 2, 2, 2, { italic: true, fill: '#ddebf7' });
    const source = captureSnapshot(store.getState(), { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const before = captureSnapshot(store.getState(), { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 });
    expect(source).not.toBeNull();
    expect(before).not.toBeNull();
    if (!source || !before) throw new Error('missing clipboard snapshots');
    setRange(store, 2, 2, 2, 2);
    pasteSpecial(store.getState(), store, wb, source, {
      what: 'all',
      operation: 'none',
      skipBlanks: false,
      transpose: false,
    });
    mutators.replaceCells(store, wb.cells(0));

    const handle = attachPasteOptions({
      host,
      grid,
      store,
      wb,
      strings: defaultStrings,
      onAfterCommit,
    });
    handle.show({
      source,
      before,
      range: { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
    });

    document.querySelector<HTMLButtonElement>('.fc-paste-options__button')?.click();
    document.querySelector<HTMLButtonElement>('[data-fc-mode="formatting"]')?.click();

    wb.recalc();
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'number', value: 2 });
    expect(formatAt(store, 2, 2)).toEqual({ bold: true, fill: '#fff2cc' });
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    handle.detach();
  });
});
