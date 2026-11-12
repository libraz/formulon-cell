import { IconSetName, SortBy } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { applyValueFilter, setFilterSort } from '../../../src/commands/filter.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { FilterSortKey } from '../../../src/store/store.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

const seeded = async () => {
  const wb = await WorkbookHandle.createDefault();
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    (['N', 1, 2, 3] as const).forEach((value, row) => {
      const addr = { sheet: 0, row, col: 0 };
      if (typeof value === 'number') wb.setNumber(addr, value);
      else wb.setText(addr, value);
      cells.set(addrKey(addr), {
        value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
        formula: null,
      });
    });
    return { ...state, data: { ...state.data, cells } };
  });
  applyValueFilter(store.getState(), store, range, 0, [], wb);
  return { wb, store };
};

describe('AutoFilter sort levels by custom list, colour and icon', () => {
  const cases: { key: FilterSortKey; native: object }[] = [
    {
      key: { byCol: 0, direction: 'asc', sortOn: 'cellColor', color: '#ff0000' },
      native: { sortBy: SortBy.CellColor, descending: false },
    },
    {
      key: { byCol: 0, direction: 'desc', sortOn: 'fontColor', color: '#0000ff' },
      native: { sortBy: SortBy.FontColor, descending: true },
    },
    {
      key: {
        byCol: 0,
        direction: 'desc',
        sortOn: 'customList',
        customList: ['Low', 'Mid', 'High'],
      },
      native: { sortBy: SortBy.Value, descending: true, customList: 'Low,Mid,High' },
    },
    {
      key: { byCol: 0, direction: 'asc', sortOn: 'cellIcon', icon: { set: 'arrows3', slot: 1 } },
      native: { sortBy: SortBy.Icon, iconSet: IconSetName.ThreeArrows, iconIndex: 1 },
    },
  ];

  it('writes the sort condition and reads the same level back after a save', async () => {
    for (const { key, native } of cases) {
      const { wb, store } = await seeded();
      try {
        setFilterSort(store, null, range, [key], wb);
        expect(syncAutoFilterToEngine(wb, store.getState(), 0, store), key.sortOn).toBe(true);
        const condition = wb.getAutoFilter(0)?.sort?.conditions[0];
        expect(condition, key.sortOn).toMatchObject(native);
        if (key.sortOn === 'cellColor' || key.sortOn === 'fontColor') {
          expect(condition?.dxfIndex).not.toBeNull();
        }
        const reopened = await WorkbookHandle.loadBytes(wb.save());
        try {
          const reloaded = createSpreadsheetStore();
          reloaded.setState((state) => ({ ...state, data: store.getState().data }));
          hydrateAutoFilterFromEngine(reopened, reloaded, 0);
          expect(reloaded.getState().ui.filterSorts?.[0]?.keys, key.sortOn).toEqual([key]);
        } finally {
          reopened.dispose();
        }
      } finally {
        wb.dispose();
      }
    }
  });
});
