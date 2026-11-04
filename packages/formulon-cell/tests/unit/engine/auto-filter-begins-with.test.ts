import { describe, expect, it } from 'vitest';
import { applyConditionFilter } from '../../../src/commands/filter.js';
import {
  autoFilterXmlFromState,
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
type Op = 'beginsWith' | 'notBeginsWith' | 'endsWith' | 'notEndsWith';

const seeded = async (preferStub: boolean) => {
  const store = createSpreadsheetStore();
  const wb = await WorkbookHandle.createDefault({ preferStub });
  const cells = new Map(store.getState().data.cells);
  ['Item', 'Apple', 'Banana', 'Avocado', 'Grape'].forEach((value, row) => {
    wb.setText({ sheet: 0, row, col: 0 }, value);
    cells.set(addrKey({ sheet: 0, row, col: 0 }), {
      value: { kind: 'text', value },
      formula: null,
    });
  });
  store.setState((s) => ({ ...s, data: { ...s.data, cells } }));
  return { store, wb };
};

describe('begins-with / ends-with AutoFilter conditions', () => {
  it.each([
    ['beginsWith', 'equal', 'A*'],
    ['notBeginsWith', 'notEqual', 'A*'],
    ['endsWith', 'equal', '*A'],
    ['notEndsWith', 'notEqual', '*A'],
  ] as const)('%s serialises to a wildcard customFilter', async (op, operator, val) => {
    const { store } = await seeded(true);
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        filterRanges: [range],
        filterCriteria: [{ range, byCol: 0, hiddenValues: [], condition: { op, value: 'A' } }],
      },
    }));
    expect(autoFilterXmlFromState(store.getState(), range)).toContain(
      `operator="${operator}" val="${val}"`,
    );
  });

  it.each([
    ['beginsWith', ['Apple', 'Avocado']],
    ['notBeginsWith', ['Banana', 'Grape']],
    ['endsWith', ['Banana']],
    ['notEndsWith', ['Apple', 'Avocado', 'Grape']],
  ] as [Op, string[]][])('%s keeps the matching rows', async (op, kept) => {
    const { store } = await seeded(true);
    applyConditionFilter(store.getState(), store, range, 0, { op, value: 'A' });
    const hidden = [...store.getState().layout.hiddenRows];
    const shown = ['Apple', 'Banana', 'Avocado', 'Grape'].filter((_, i) => !hidden.includes(i + 1));
    expect(shown).toEqual(kept);
  });

  it('loads a begins-with criterion back from the engine', async () => {
    const { store, wb } = await seeded(false);
    store.setState((s) => ({
      ...s,
      ui: {
        ...s.ui,
        filterRanges: [range],
        filterCriteria: [
          { range, byCol: 0, hiddenValues: [], condition: { op: 'beginsWith', value: 'A' } },
        ],
      },
    }));
    syncAutoFilterToEngine(wb, store.getState(), 0, store, 'en-US', range);
    const fresh = createSpreadsheetStore();
    hydrateAutoFilterFromEngine(wb, fresh, 0);
    expect(fresh.getState().ui.filterCriteria).toEqual([
      { range, byCol: 0, hiddenValues: [], condition: { op: 'beginsWith', value: 'A' } },
    ]);
  });
});
