// @vitest-environment node

import { beforeEach, describe, expect, it } from 'vitest';
import { deleteCells, insertCells } from '../../../src/commands/cell-shift.js';
import { copy } from '../../../src/commands/clipboard/copy.js';
import { insertCopiedBand } from '../../../src/commands/clipboard/insert-copied-cells.js';
import { captureSnapshotFromCopyResult } from '../../../src/commands/clipboard/snapshot.js';
import { History } from '../../../src/commands/history.js';
import { type PivotGuardAlert, setPivotGuardSink } from '../../../src/commands/pivot-guard.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { deleteRows, insertCols, insertRows } from '../../../src/commands/structure.js';
import { PivotAggregation } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const MAX_COL = 16_383;

describe('row / column / cell edits around a PivotTable', () => {
  let wb: WorkbookHandle;
  let store: SpreadsheetStore;
  let history: History;
  let alerts: PivotGuardAlert[];
  /** The report A6:B9 (rows 5..8, columns 0..1). */
  const bounds = () => wb.pivotBounds(0)[0]?.range;

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    store = createSpreadsheetStore();
    history = new History();
    alerts = [];
    setPivotGuardSink(wb, (alert) => alerts.push(alert));
    const rows: (string | number)[][] = [
      ['Region', 'Amt'],
      ['East', 5],
      ['West', 7],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        const addr = { sheet: 0, row: r, col: 10 + c };
        if (typeof v === 'number') wb.setNumber(addr, v);
        else wb.setText(addr, v);
      }
    }
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 10, r1: 2, c1: 11 },
      destination: { sheet: 0, row: 5, col: 0 },
      rowField: 'Region',
      valueField: 'Amt',
      aggregation: PivotAggregation.Sum,
    });
    expect(created.ok).toBe(true);
    mutators.replaceCells(store, wb.cells(0));
    expect(bounds()).toEqual({ sheet: 0, r0: 5, c0: 0, r1: 8, c1: 1 });
  });

  it('refuses inserting rows through a report, with the alert', () => {
    expect(insertRows(store, wb, history, 6, 1)).toBe(false);
    expect(alerts).toHaveLength(1);
    expect(bounds()?.r0).toBe(5);
  });

  it('inserts rows above a report and moves it down', () => {
    expect(insertRows(store, wb, history, 5, 2)).toBe(true);
    expect(alerts).toHaveLength(0);
    expect(bounds()?.r0).toBe(7);
  });

  it('refuses deleting rows that hold part of a report', () => {
    expect(deleteRows(store, wb, history, 7, 1)).toBe(false);
    expect(alerts).toHaveLength(1);
  });

  it('refuses inserting a column through a report but not beside it', () => {
    expect(insertCols(store, wb, history, 1, 1)).toBe(false);
    expect(alerts).toHaveLength(1);
    expect(insertCols(store, wb, history, 3, 1)).toBe(true);
  });

  it('refuses shifting cells that would move part of a report', () => {
    expect(insertCells(store, wb, history, { sheet: 0, r0: 3, c0: 0, r1: 3, c1: 0 }, 'down')).toBe(
      false,
    );
    expect(deleteCells(store, wb, history, { sheet: 0, r0: 6, c0: 3, r1: 6, c1: 3 }, 'left')).toBe(
      true,
    );
    expect(alerts).toHaveLength(1);
    expect(insertCells(store, wb, history, { sheet: 0, r0: 3, c0: 4, r1: 3, c1: 4 }, 'down')).toBe(
      true,
    );
  });

  const cutRows = (r0: number, r1: number) => {
    mutators.replaceCells(store, wb.cells(0));
    const source = { sheet: 0, r0, c0: 0, r1, c1: MAX_COL };
    mutators.setActive(store, { sheet: 0, row: r0, col: 0 });
    mutators.setRange(store, source);
    const result = copy(store.getState());
    if (!result) throw new Error('missing copy result');
    const snap = captureSnapshotFromCopyResult(store.getState(), result, 'cut');
    if (!snap) throw new Error('missing cut snapshot');
    mutators.setCopyRange(store, source, 'cut');
    return snap;
  };

  it('moves a cut row band that does not touch the report', () => {
    wb.setNumber({ sheet: 0, row: 20, col: 0 }, 42);
    const snap = cutRows(20, 20);
    const result = insertCopiedBand(store, wb, history, snap, { sheet: 0, row: 25, col: 0 });
    expect(result).not.toBeNull();
    expect(wb.getValue({ sheet: 0, row: 24, col: 0 })).toEqual({ kind: 'number', value: 42 });
    expect(bounds()?.r0).toBe(5);
  });

  it('refuses a cut row band taken from inside the report', () => {
    const snap = cutRows(6, 6);
    expect(insertCopiedBand(store, wb, history, snap, { sheet: 0, row: 25, col: 0 }).status).toBe(
      'refused',
    );
  });

  it('refuses inserting a cut row band inside the report', () => {
    wb.setNumber({ sheet: 0, row: 20, col: 0 }, 42);
    const snap = cutRows(20, 20);
    expect(insertCopiedBand(store, wb, history, snap, { sheet: 0, row: 7, col: 0 }).status).toBe(
      'refused',
    );
  });
});
