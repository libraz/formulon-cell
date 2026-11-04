import { describe, expect, it } from 'vitest';
import { fillLastValue } from '../../../src/commands/fill.js';
import { addrKey } from '../../../src/engine/address.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const withCells = (entries: [number, number, number | string][]) => {
  const store = createSpreadsheetStore();
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    for (const [row, col, v] of entries) {
      cells.set(addrKey({ sheet: 0, row, col }), {
        value: typeof v === 'number' ? { kind: 'number', value: v } : { kind: 'text', value: v },
        formula: null,
      });
    }
    return { ...s, data: { ...s.data, cells } };
  });
  return store.getState();
};

describe('fillLastValue', () => {
  it('projects the furthest cell of a series drag: 1, 2 dragged to row 8 gives 8', () => {
    // Excel (Mac): dragging L1:L2 = 1, 2 down to L8 shows "8" by the cursor.
    const state = withCells([
      [0, 11, 1],
      [1, 11, 2],
    ]);
    const src = { sheet: 0, r0: 0, c0: 11, r1: 1, c1: 11 };
    expect(fillLastValue(state, src, { ...src, r1: 7 }, 'en-US')).toBe('8');
    expect(fillLastValue(state, src, { ...src, r0: 0, r1: 1 }, 'en-US')).toBeNull();
  });

  it('follows the drag direction and repeats a text value', () => {
    const state = withCells([
      [4, 2, 'abc'],
      [4, 3, 'x'],
    ]);
    const src = { sheet: 0, r0: 4, c0: 3, r1: 4, c1: 3 };
    expect(fillLastValue(state, src, { ...src, c1: 6 }, 'en-US')).toBe('x');
    const up = { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 };
    expect(fillLastValue(state, up, { ...up, r0: 1 }, 'en-US')).toBe('abc');
  });
});
