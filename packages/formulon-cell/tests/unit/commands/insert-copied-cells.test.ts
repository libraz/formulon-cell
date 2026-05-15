import { describe, expect, it, vi } from 'vitest';
import { insertCopiedCellsFromTSV } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const setActive = (
  store: ReturnType<typeof createSpreadsheetStore>,
  row: number,
  col: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row, col },
      anchor: { sheet: 0, row, col },
      range: { sheet: 0, r0: row, c0: col, r1: row, c1: col },
    },
  }));
};

describe('insertCopiedCellsFromTSV', () => {
  it('shifts only the target columns down before writing the copied cells', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    wb.setText({ sheet: 0, row: 1, col: 1 }, 'B2');
    wb.setText({ sheet: 0, row: 1, col: 2 }, 'C2');
    wb.setText({ sheet: 0, row: 1, col: 3 }, 'D2');
    setActive(store, 1, 1);

    const result = insertCopiedCellsFromTSV(store, wb, null, 'x\ty', 'down');

    expect(result?.writtenRange).toEqual({ sheet: 0, r0: 1, c0: 1, r1: 1, c1: 2 });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'text', value: 'x' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 2 })).toEqual({ kind: 'text', value: 'y' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'text', value: 'B2' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 2 })).toEqual({ kind: 'text', value: 'C2' });
    expect(wb.getValue({ sheet: 0, row: 1, col: 3 })).toEqual({ kind: 'text', value: 'D2' });
  });

  it('shifts formats and whole merges with the moved cells', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    setActive(store, 1, 1);
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 1 }, { bold: true });
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 1, r1: 2, c1: 2 });

    insertCopiedCellsFromTSV(store, wb, null, 'x\ty', 'down');

    expect(store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 1 }))?.bold).toBe(
      true,
    );
    expect(Array.from(store.getState().merges.byAnchor.values())).toContainEqual({
      sheet: 0,
      r0: 3,
      c0: 1,
      r1: 3,
      c1: 2,
    });
  });

  it('blocks a partial merge split instead of corrupting merge indexes', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    setActive(store, 1, 1);
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 1 });
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});

    const result = insertCopiedCellsFromTSV(store, wb, null, 'x', 'down');

    expect(result).toBeNull();
    expect(warn).toHaveBeenCalledWith(
      'formulon-cell: insert copied cells blocked — merge would be split',
    );
    expect(Array.from(store.getState().merges.byAnchor.values())).toContainEqual({
      sheet: 0,
      r0: 2,
      c0: 0,
      r1: 2,
      c1: 1,
    });
  });
});
