import { type AutoFilterColumnInput, FilterKind } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { applyValueFilter } from '../../../src/commands/filter.js';
import { filterButtonHidden } from '../../../src/commands/filter-ranges.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  autoFilterXmlFromState,
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { cellRect, hitZone, layoutForView } from '../../../src/render/geometry.js';
import { filterButtonRect } from '../../../src/render/painters/filter-button.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
const ROWS = [
  ['Name', 'Zone'],
  ['East', 'N'],
  ['West', 'S'],
  ['East', 'S'],
];

const seeded = async () => {
  const wb = await WorkbookHandle.createDefault();
  const store = createSpreadsheetStore();
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    ROWS.forEach((cols, row) => {
      cols.forEach((value, col) => {
        wb.setText({ sheet: 0, row, col }, value);
        cells.set(addrKey({ sheet: 0, row, col }), {
          value: { kind: 'text', value },
          formula: null,
        });
      });
    });
    return { ...state, data: { ...state.data, cells } };
  });
  // Column A keeps only East with its button hidden; column B has no criterion and no button.
  const columns: AutoFilterColumnInput[] = [
    {
      columnOffset: 0,
      kind: FilterKind.Values,
      values: ['East'],
      hiddenButton: true,
      showButton: false,
    },
    { columnOffset: 1, kind: FilterKind.None, hiddenButton: true, showButton: false },
  ];
  expect(
    wb.setAutoFilter(0, {
      range: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 },
      columns,
      sort: null,
    }),
  ).toBe(true);
  hydrateAutoFilterFromEngine(wb, store, 0);
  return { wb, store };
};

describe('AutoFilter columns with a hidden button', () => {
  it('keeps the filter applied but shows no button on the header', async () => {
    const { wb, store } = await seeded();
    try {
      const state = store.getState();
      expect(state.ui.filterCriteria).toEqual([{ range, byCol: 0, hiddenValues: ['West'] }]);
      expect([filterButtonHidden(state, 0, 0, 0), filterButtonHidden(state, 0, 0, 1)]).toEqual([
        true,
        true,
      ]);
      expect(filterButtonHidden(state, 0, 1, 0)).toBe(false);
      const layout = layoutForView(state);
      const btn = filterButtonRect(cellRect(layout, state.viewport, 0, 0), layout.rtl);
      const zone = hitZone(layout, state.viewport, btn.x + 2, btn.y + 2, state.ui.filterRanges, {
        filterButtonHidden: (row, col) => filterButtonHidden(state, 0, row, col),
      });
      expect(zone?.kind).toBe('cell');
    } finally {
      wb.dispose();
    }
  });

  it('writes the hidden buttons back after an edit and keeps them through a save', async () => {
    const { wb, store } = await seeded();
    try {
      applyValueFilter(store.getState(), store, range, 0, ['East'], wb);
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      const columns = wb.getAutoFilter(0)?.columns ?? [];
      expect(columns.map((c) => [c.columnOffset, c.hiddenButton, c.showButton])).toEqual([
        [0, true, false],
        [1, true, false],
      ]);
      expect(autoFilterXmlFromState(store.getState(), range)).toMatch(
        /<filterColumn colId="0" hiddenButton="1">.*<filterColumn colId="1" hiddenButton="1"\/>/,
      );
      const reopened = await WorkbookHandle.loadBytes(wb.save());
      try {
        const reloaded = createSpreadsheetStore();
        reloaded.setState((state) => ({ ...state, data: store.getState().data }));
        hydrateAutoFilterFromEngine(reopened, reloaded, 0);
        expect(reloaded.getState().ui.filterHiddenButtons).toEqual([{ range, cols: [0, 1] }]);
        expect(reloaded.getState().ui.filterCriteria[0]?.hiddenValues).toEqual(['East']);
      } finally {
        reopened.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
