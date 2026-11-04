import { describe, expect, it } from 'vitest';
import { applyMerge } from '../../../src/commands/merge.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type NumberCell = { value: { kind: 'number'; value: number }; formula: null };

describe('applyMerge', () => {
  it('blanks the covered content without probing every cell of the rectangle', async () => {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    let probes = 0;
    class CountingMap extends Map<string, NumberCell> {
      override get(key: string): NumberCell | undefined {
        probes += 1;
        return super.get(key);
      }
    }
    const cells = new CountingMap();
    for (const [row, col] of [
      [0, 0],
      [0, 5],
      [30, 30],
    ] as const) {
      wb.setNumber({ sheet: 0, row, col }, row + col + 1);
      cells.set(addrKey({ sheet: 0, row, col }), {
        value: { kind: 'number', value: row + col + 1 },
        formula: null,
      });
    }
    store.setState((s) => ({ ...s, data: { ...s.data, cells } }));

    probes = 0;
    const ok = applyMerge(store, wb, null, { sheet: 0, r0: 0, c0: 0, r1: 40, c1: 40 });
    expect(ok).toBe(true);
    expect(wb.getValue({ sheet: 0, row: 0, col: 5 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 30, col: 30 })).toEqual({ kind: 'blank' });
    expect(probes).toBeLessThan(500);
  });
});
