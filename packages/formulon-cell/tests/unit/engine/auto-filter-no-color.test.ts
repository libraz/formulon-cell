import { describe, expect, it } from 'vitest';
import { applyColorFilter, NO_COLOR, setAutoFilter } from '../../../src/commands/filter.js';
import { sortRange } from '../../../src/commands/sort.js';
import { hydrateAutoFilterFromEngine } from '../../../src/engine/auto-filter-sync.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import type { FilterColorSpec } from '../../../src/store/types.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };

/** Whether the engine writes AutoFilter to `.xlsb`. It does not yet: the save reports the filter as
 *  deferred and the reopened book has none. Flip this once it does; the xlsb case then expects the
 *  same round trip as `.xlsx`. */
const XLSB_WRITES_AUTOFILTER = false;

/** A header and three rows: row 1 filled red with a blue font, rows 2 and 3 with neither. */
const seeded = async (): Promise<{ wb: WorkbookHandle; store: SpreadsheetStore }> => {
  const wb = await WorkbookHandle.createDefault();
  for (const [row, text] of ['H', 'a', 'b', 'c'].entries()) {
    wb.setText({ sheet: 0, row, col: 0 }, text);
  }
  const store = createSpreadsheetStore();
  mutators.replaceCells(store, wb.cells(0));
  store.setState((s) => ({
    ...s,
    format: { ...s.format, formats: new Map([['0:1:0', { fill: '#ff0000', color: '#0000ff' }]]) },
  }));
  setAutoFilter(store, range, wb);
  return { wb, store };
};

const reopened = async (bytes: Uint8Array): Promise<SpreadsheetStore> => {
  const wb = await WorkbookHandle.loadBytes(bytes);
  try {
    const store = createSpreadsheetStore();
    mutators.replaceCells(store, wb.cells(0));
    hydrateLayoutFromEngine(wb, store, 0);
    hydrateAutoFilterFromEngine(wb, store, 0);
    return store;
  } finally {
    wb.dispose();
  }
};

const NO_COLOR_FILTERS: FilterColorSpec[] = [
  { kind: 'cellColor', color: NO_COLOR },
  { kind: 'fontColor', color: NO_COLOR },
];

describe('No Fill and Automatic colour filters', () => {
  for (const color of NO_COLOR_FILTERS) {
    it(`keeps a ${color.kind} No Fill / Automatic filter through an .xlsx save`, async () => {
      const { wb, store } = await seeded();
      try {
        applyColorFilter(store.getState(), store, range, 0, color, wb);
        expect([...store.getState().layout.hiddenRows]).toEqual([1]);
        const next = await reopened(wb.save());
        expect(next.getState().ui.filterCriteria).toEqual([
          { range, byCol: 0, hiddenValues: [], color },
        ]);
      } finally {
        wb.dispose();
      }
    });
  }

  it('keeps a No Fill sort level through an .xlsx save', async () => {
    const { wb, store } = await seeded();
    try {
      sortRange(store.getState(), store, wb, range, {
        byCol: 0,
        direction: 'asc',
        hasHeader: true,
        keys: [{ byCol: 0, direction: 'asc', sortOn: 'cellColor', color: NO_COLOR }],
      });
      const next = await reopened(wb.save());
      expect(next.getState().ui.filterSorts?.[0]?.keys).toEqual([
        { byCol: 0, direction: 'asc', sortOn: 'cellColor', color: NO_COLOR },
      ]);
    } finally {
      wb.dispose();
    }
  });

  it('follows the engine AutoFilter contract on an .xlsb save', async () => {
    const { wb, store } = await seeded();
    try {
      applyColorFilter(
        store.getState(),
        store,
        range,
        0,
        NO_COLOR_FILTERS[0] as FilterColorSpec,
        wb,
      );
      const { bytes, diagnostics } = wb.saveWithDiagnostics('xlsb');
      const criteria = (await reopened(bytes)).getState().ui.filterCriteria;
      if (XLSB_WRITES_AUTOFILTER) {
        expect(criteria).toEqual([
          { range, byCol: 0, hiddenValues: [], color: NO_COLOR_FILTERS[0] },
        ]);
      } else {
        expect([diagnostics.deferredFeatureCount, criteria]).toEqual([1, []]);
      }
    } finally {
      wb.dispose();
    }
  });
});
