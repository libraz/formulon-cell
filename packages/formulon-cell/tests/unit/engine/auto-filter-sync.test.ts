import { describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import {
  autoFilterRangeFromXml,
  autoFilterXmlFromState,
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };

const seedText = (
  store: ReturnType<typeof createSpreadsheetStore>,
  row: number,
  col: number,
  value: string,
): void => {
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), { value: { kind: 'text', value }, formula: null });
    return { ...state, data: { ...state.data, cells } };
  });
};

describe('AutoFilter engine sync', () => {
  it('serializes the filter range, visible value checklist, and custom condition', () => {
    const store = createSpreadsheetStore();
    seedText(store, 1, 0, 'East');
    seedText(store, 2, 0, 'West');
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterRange: range,
        filterCriteria: [
          { range, byCol: 0, hiddenValues: ['West'] },
          { range, byCol: 1, hiddenValues: [], condition: { op: 'contains', value: 'x&y' } },
        ],
      },
    }));

    expect(autoFilterXmlFromState(store.getState(), 0)).toBe(
      '<autoFilter ref="A1:B4"><filterColumn colId="0"><filters blank="1"><filter val="East"/></filters></filterColumn><filterColumn colId="1"><customFilters><customFilter operator="equal" val="*x&amp;y*"/></customFilters></filterColumn></autoFilter>',
    );
  });

  it('parses the imported AutoFilter range without reinterpreting its criteria', () => {
    expect(
      autoFilterRangeFromXml(
        '<autoFilter ref="C3:E9"><filterColumn colId="0"><filters><filter val="x"/></filters></filterColumn></autoFilter>',
        2,
      ),
    ).toEqual({ sheet: 2, r0: 2, c0: 2, r1: 8, c1: 4 });
    expect(autoFilterRangeFromXml('', 0)).toBeNull();
  });

  it('hydrates imported filter affordances and only writes when the capability exists', () => {
    const store = createSpreadsheetStore();
    let muted = 0;
    const fake = {
      capabilities: { autoFilter: true },
      getSheetAutoFilterXml: () => '<autoFilter ref="A1:B4"/>',
      setSheetAutoFilterXml: (_sheet: number, xml: string) => xml === '<autoFilter ref="A1:B4"/>',
      withAutoFilterSyncMuted: (fn: () => void) => {
        muted += 1;
        fn();
      },
    } as unknown as WorkbookHandle;

    hydrateAutoFilterFromEngine(fake, store, 0);
    expect(muted).toBe(1);
    expect(store.getState().ui.filterRange).toEqual(range);
    expect(syncAutoFilterToEngine(fake, store.getState(), 0)).toBe(true);

    const unavailable = { capabilities: { autoFilter: false } } as unknown as WorkbookHandle;
    expect(syncAutoFilterToEngine(unavailable, store.getState(), 0)).toBe(false);
  });
});
