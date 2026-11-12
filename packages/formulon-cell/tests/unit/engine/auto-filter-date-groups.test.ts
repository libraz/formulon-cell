import { DateTimeGrouping, FilterKind } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { applyValueFilter } from '../../../src/commands/filter.js';
import {
  dateGroupsForVisible,
  serialInDateGroups,
} from '../../../src/commands/filter-date-groups.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  autoFilterXmlFromState,
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// 2025-12-31, 2026-05-08, 2026-05-18, 2026-06-07 in the 1900 date system.
const DEC_31_2025 = 46022;
const MAY_8 = 46150;
const MAY_18 = 46160;
const JUN_7 = 46180;
const ALL = [DEC_31_2025, MAY_8, MAY_18, JUN_7];

describe('date-group items for a checklist', () => {
  it('keeps a whole year, a whole month or single days, whichever is fewest', () => {
    expect(dateGroupsForVisible(ALL, ALL, false)).toEqual([
      { year: 2025, grouping: DateTimeGrouping.Year },
      { year: 2026, grouping: DateTimeGrouping.Year },
    ]);
    expect(dateGroupsForVisible([MAY_8, MAY_18], ALL, false)).toEqual([
      { year: 2026, month: 5, grouping: DateTimeGrouping.Month },
    ]);
    expect(dateGroupsForVisible([MAY_8, JUN_7], ALL, false)).toEqual([
      { year: 2026, month: 5, day: 8, grouping: DateTimeGrouping.Day },
      { year: 2026, month: 6, grouping: DateTimeGrouping.Month },
    ]);
    expect(dateGroupsForVisible([], ALL, false)).toEqual([]);
  });

  it('matches a date against year, month and day items', () => {
    const groups = [
      { year: 2025, grouping: DateTimeGrouping.Year },
      { year: 2026, month: 5, day: 18, grouping: DateTimeGrouping.Day },
    ];
    expect(ALL.map((serial) => serialInDateGroups(serial, groups, false))).toEqual([
      true,
      false,
      true,
      false,
    ]);
    expect(serialInDateGroups(MAY_18 + 0.5, groups, false)).toBe(true);
  });
});

describe('AutoFilter checklist on a date column', () => {
  const range = { sheet: 0, r0: 0, c0: 0, r1: 5, c1: 0 };
  const seeded = async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      const formats = new Map(state.format.formats);
      const put = (row: number, value: string | number, dated = false) => {
        const addr = { sheet: 0, row, col: 0 };
        if (typeof value === 'number') wb.setNumber(addr, value);
        else wb.setText(addr, value);
        cells.set(addrKey(addr), {
          value: typeof value === 'number' ? { kind: 'number', value } : { kind: 'text', value },
          formula: null,
        });
        if (dated) formats.set(addrKey(addr), { numFmt: { kind: 'date', pattern: 'yyyy-mm-dd' } });
      };
      put(0, 'Date');
      for (const [i, serial] of ALL.entries()) put(i + 1, serial, true);
      put(5, 'n/a');
      return { ...state, data: { ...state.data, cells }, format: { ...state.format, formats } };
    });
    return { wb, store };
  };

  it('saves kept dates as date-group items and reads the same checklist back', async () => {
    const { wb, store } = await seeded();
    try {
      // Hide 2025-12-31 and 2026-06-07; keep May 2026 and the text item.
      applyValueFilter(store.getState(), store, range, 0, [String(DEC_31_2025), String(JUN_7)]);
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({
        kind: FilterKind.Values,
        values: ['n/a'],
        dateGroups: [
          expect.objectContaining({ year: 2026, month: 5, grouping: DateTimeGrouping.Month }),
        ],
      });
      expect(wb.evaluateAutoFilter(0)?.match).toEqual([false, true, true, false, true]);
      expect(autoFilterXmlFromState(store.getState(), range)).toContain(
        '<filters><filter val="n/a"/><dateGroupItem year="2026" month="5" dateTimeGrouping="month"/></filters>',
      );
      const reopened = await WorkbookHandle.loadBytes(wb.save());
      try {
        hydrateAutoFilterFromEngine(reopened, store, 0);
        expect(store.getState().ui.filterCriteria[0]?.hiddenValues.sort()).toEqual(
          [String(DEC_31_2025), String(JUN_7)].sort(),
        );
      } finally {
        reopened.dispose();
      }
    } finally {
      wb.dispose();
    }
  });
});
