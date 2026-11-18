import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { applyRangeDrop, type RangeDropAction } from '../../../src/commands/range-drop.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// Source B2:C3 = 22,23 / 32,33 (B2 yellow); destination E5:F6 holds E5 'target', F6 'tgt2'.
const SRC = { sheet: 0, r0: 1, c0: 1, r1: 2, c1: 2 };
const DEST = { sheet: 0, r0: 4, c0: 4, r1: 5, c1: 5 };

describe('secondary-button drop of a selection', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;

  const put = (row: number, col: number, value: number | string): void => {
    const addr = { sheet: 0, row, col };
    if (typeof value === 'number') wb.setNumber(addr, value);
    else wb.setText(addr, value);
  };
  const value = (row: number, col: number) => {
    const v = wb.getValue({ sheet: 0, row, col });
    return v.kind === 'blank' ? '' : (v as { value: number | string }).value;
  };

  beforeEach(async () => {
    store = createSpreadsheetStore();
    // Shift drops shift cells natively; the stub engine cannot.
    wb = await WorkbookHandle.createDefault();
    history = new History();
    put(1, 1, 22);
    put(1, 2, 23);
    put(2, 1, 32);
    put(2, 2, 33);
    put(3, 1, 'keepme');
    put(4, 4, 'target');
    put(5, 5, 'tgt2');
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      formats.set(addrKey({ sheet: 0, row: 1, col: 1 }), { fill: '#ffff00' });
      return { ...s, format: { ...s.format, formats } };
    });
  });

  const drop = (action: RangeDropAction) =>
    applyRangeDrop({ store, wb, history }, SRC, DEST, action);
  const fillAt = (row: number, col: number) =>
    store.getState().format.formats.get(addrKey({ sheet: 0, row, col }))?.fill;

  it('move carries values and formats and empties the source', async () => {
    expect(await drop('move')).toBe(true);
    expect([value(4, 4), value(4, 5), value(5, 4), value(5, 5)]).toEqual([22, 23, 32, 33]);
    expect(fillAt(4, 4)).toBe('#ffff00');
    expect(value(1, 1)).toBe('');
    expect(fillAt(1, 1)).toBeUndefined();
    expect(store.getState().selection.range).toEqual(DEST);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 4, col: 4 });
  });

  it('copy keeps the source', async () => {
    expect(await drop('copy')).toBe(true);
    expect([value(4, 4), value(5, 5)]).toEqual([22, 33]);
    expect(fillAt(4, 4)).toBe('#ffff00');
    expect(value(1, 1)).toBe(22);
    expect(store.getState().selection.range).toEqual(DEST);
  });

  it('copy values only leaves the destination formats', async () => {
    expect(await drop('copyValues')).toBe(true);
    expect([value(4, 4), value(5, 5)]).toEqual([22, 33]);
    expect(fillAt(4, 4)).toBeUndefined();
  });

  it('copy formats only leaves the destination values', async () => {
    expect(await drop('copyFormats')).toBe(true);
    expect([value(4, 4), value(5, 5)]).toEqual(['target', 'tgt2']);
    expect(fillAt(4, 4)).toBe('#ffff00');
  });

  it('link writes references to the source cells', async () => {
    expect(await drop('link')).toBe(true);
    const formula = (row: number, col: number) => wb.cellFormula({ sheet: 0, row, col });
    expect([formula(4, 4), formula(4, 5), formula(5, 4), formula(5, 5)]).toEqual([
      '=B2',
      '=C2',
      '=B3',
      '=C3',
    ]);
  });

  it('hyperlink points the first destination cell at the source and keeps its text', async () => {
    expect(await drop('hyperlink')).toBe(true);
    expect(value(4, 4)).toBe('target');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 4 }));
    expect(fmt?.hyperlinkLocation).toBe('Sheet1!B2');
    expect(store.getState().selection.range).toEqual(DEST);
  });

  it('shift and copy pushes the destination cells down by the range height', async () => {
    expect(await drop('copyShiftDown')).toBe(true);
    expect([value(4, 4), value(5, 5)]).toEqual([22, 33]);
    expect(value(6, 4)).toBe('target');
    expect(value(7, 5)).toBe('tgt2');
    expect(value(1, 1)).toBe(22);
  });

  it('shift and move to the right empties the source without closing the gap', async () => {
    expect(await drop('moveShiftRight')).toBe(true);
    expect([value(4, 4), value(5, 5)]).toEqual([22, 33]);
    expect(value(4, 6)).toBe('target');
    expect(value(5, 7)).toBe('tgt2');
    expect(value(1, 1)).toBe('');
    expect(value(3, 1)).toBe('keepme');
  });
});
