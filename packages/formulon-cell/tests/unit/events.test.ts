import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ChangeEvent } from '../../src/engine/workbook-handle.js';
import {
  type SelectionChangeEvent,
  SpreadsheetEmitter,
  selectionEquals,
  toCellChangeEvent,
} from '../../src/events.js';
import type { State } from '../../src/store/store.js';

const selection = (overrides: Partial<SelectionChangeEvent> = {}): State['selection'] => ({
  active: overrides.active ?? { sheet: 0, row: 1, col: 2 },
  anchor: overrides.anchor ?? { sheet: 0, row: 1, col: 2 },
  range: overrides.range ?? { sheet: 0, r0: 1, c0: 2, r1: 1, c1: 2 },
});

describe('SpreadsheetEmitter', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('emits to registered handlers and supports unsubscribe/off', () => {
    const emitter = new SpreadsheetEmitter();
    const first = vi.fn();
    const second = vi.fn();
    const unsubscribeFirst = emitter.on('themeChange', first);
    emitter.on('themeChange', second);

    emitter.emit('themeChange', { theme: 'paper' });
    unsubscribeFirst();
    emitter.off('themeChange', second);
    emitter.emit('themeChange', { theme: 'ink' });

    expect(first).toHaveBeenCalledTimes(1);
    expect(first).toHaveBeenCalledWith({ theme: 'paper' });
    expect(second).toHaveBeenCalledTimes(1);
  });

  it('snapshots handlers so unsubscription during emit does not skip later handlers', () => {
    const emitter = new SpreadsheetEmitter();
    const second = vi.fn();
    const unsubscribeFirst = emitter.on('themeChange', () => {
      unsubscribeFirst();
    });
    emitter.on('themeChange', second);

    emitter.emit('themeChange', { theme: 'paper' });
    emitter.emit('themeChange', { theme: 'ink' });

    expect(second).toHaveBeenCalledTimes(2);
  });

  it('continues after throwing handlers and clears listeners on dispose', () => {
    const emitter = new SpreadsheetEmitter();
    const error = new Error('handler failed');
    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const good = vi.fn();
    emitter.on('themeChange', () => {
      throw error;
    });
    emitter.on('themeChange', good);

    emitter.emit('themeChange', { theme: 'paper' });
    emitter.dispose();
    emitter.emit('themeChange', { theme: 'ink' });

    expect(good).toHaveBeenCalledTimes(1);
    expect(consoleError).toHaveBeenCalledWith(
      'formulon-cell: event "themeChange" handler threw',
      error,
    );
  });
});

describe('event helpers', () => {
  it('maps workbook value changes into public cellChange events', () => {
    const change: ChangeEvent = {
      kind: 'value',
      addr: { sheet: 0, row: 1, col: 2 },
      next: { kind: 'number', value: 9 },
    };

    expect(toCellChangeEvent(change, () => '=A1+1')).toEqual({
      addr: { sheet: 0, row: 1, col: 2 },
      value: { kind: 'number', value: 9 },
      formula: '=A1+1',
    });
  });

  it('ignores non-value workbook changes for cellChange', () => {
    expect(toCellChangeEvent({ kind: 'sheet-add', index: 1, name: 'Data' }, () => null)).toBeNull();
  });

  it('compares active cell, anchor, and selected range', () => {
    const base = selection();

    expect(selectionEquals(base, selection())).toBe(true);
    expect(selectionEquals(base, selection({ active: { sheet: 0, row: 9, col: 2 } }))).toBe(false);
    expect(selectionEquals(base, selection({ anchor: { sheet: 0, row: 1, col: 9 } }))).toBe(false);
    expect(
      selectionEquals(base, selection({ range: { sheet: 0, r0: 1, c0: 2, r1: 3, c1: 2 } })),
    ).toBe(false);
  });
});
