import { SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  groupSparklines,
  setSparklines,
  ungroupSparklines,
} from '../../../src/commands/sparkline.js';
import {
  groupSelectedSparklinesAction,
  ungroupSelectedSparklinesAction,
} from '../../../src/commands/sparkline-design.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const A = { sheet: 0, row: 0, col: 0 };
const B = { sheet: 0, row: 1, col: 0 };
const LINE = defaultSparklineGroupSettings(SparklineType.Line);

describe('group and ungroup that change nothing', () => {
  it('Group on cells that already share one group keeps its id and writes no history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setSparklines(
      store,
      [
        { addr: A, source: 'B1:B4' },
        { addr: B, source: 'C1:C4' },
      ],
      LINE,
    );
    const before = groupOfCell(store.getState(), A)?.id;
    expect(groupSparklines(store, [A, B])).toBe(before);
    expect(groupOfCell(store.getState(), B)?.id).toBe(before);

    mutators.setActive(store, A);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });
    const ctx = { store, history, origin: 'ribbon', commandId: 'sparklineGroup' } as const;
    expect(groupSelectedSparklinesAction(ctx)).toBe(true);
    expect(history.canUndo()).toBe(false);
  });

  it('Ungroup on single-sparkline groups regroups nothing and writes no history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    setSparklines(store, [{ addr: A, source: 'B1:B4' }], LINE);
    const before = groupOfCell(store.getState(), A)?.id;
    expect(ungroupSparklines(store, [A])).toBe(0);
    expect(groupOfCell(store.getState(), A)?.id).toBe(before);

    mutators.setActive(store, A);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    const ctx = { store, history, origin: 'ribbon', commandId: 'sparklineUngroup' } as const;
    expect(ungroupSelectedSparklinesAction(ctx)).toBe(true);
    expect(history.canUndo()).toBe(false);
  });
});
