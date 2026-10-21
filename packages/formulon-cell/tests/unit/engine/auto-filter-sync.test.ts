import type { AutoFilterColumn } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import {
  applyAdvancedFilter,
  applyColorFilter,
  applyConditionFilter,
  applyFilterColumns,
  applyValueFilter,
  filterValueKey,
  reapplyFilters,
  recordFilterChange,
} from '../../../src/commands/filter.js';
import { History } from '../../../src/commands/history.js';
import { addrKey } from '../../../src/engine/address.js';
import {
  autoFilterRangeFromXml,
  autoFilterXmlFromState,
  hydrateAutoFilterFromEngine,
  syncAutoFilterToEngine,
} from '../../../src/engine/auto-filter-sync.js';
import { flushFormatToEngine } from '../../../src/engine/cell-format-sync.js';
import { hydrateLayoutFromEngine } from '../../../src/engine/layout-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

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

const seedNumber = (
  store: ReturnType<typeof createSpreadsheetStore>,
  row: number,
  col: number,
  value: number,
): void => {
  store.setState((state) => {
    const cells = new Map(state.data.cells);
    cells.set(addrKey({ sheet: 0, row, col }), { value: { kind: 'number', value }, formula: null });
    return { ...state, data: { ...state.data, cells } };
  });
};

describe('AutoFilter engine sync', () => {
  it('keeps checklist and custom comparisons case-insensitive like native filters', async () => {
    const valuesWb = await WorkbookHandle.createDefault();
    try {
      expect(valuesWb.capabilities.typedAutoFilter).toBe(true);
      const js = createSpreadsheetStore();
      const native = createSpreadsheetStore();
      const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
      for (const [row, value] of ['Header', 'East', 'east', 'West'].entries()) {
        seedText(js, row, 0, value);
        seedText(native, row, 0, value);
        valuesWb.setText({ sheet: 0, row, col: 0 }, value);
      }
      applyValueFilter(js.getState(), js, range, 0, ['East']);
      applyValueFilter(native.getState(), native, range, 0, ['East'], valuesWb);
      expect(native.getState().layout.hiddenRows).toEqual(js.getState().layout.hiddenRows);
      expect(native.getState().layout.hiddenRows).toEqual(new Set([1, 2]));
    } finally {
      valuesWb.dispose();
    }

    const cases = [
      ['equals', 'east'],
      ['notEquals', 'east'],
      ['contains', 'AS'],
      ['notContains', 'AS'],
      ['greaterThan', '10'],
      ['greaterThanOrEqual', '10'],
      ['lessThan', '10'],
      ['lessThanOrEqual', '10'],
    ] as const;
    for (const [op, value] of cases) {
      const wb = await WorkbookHandle.createDefault();
      try {
        expect(wb.capabilities.typedAutoFilter).toBe(true);
        const js = createSpreadsheetStore();
        const native = createSpreadsheetStore();
        const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
        if (
          op === 'greaterThan' ||
          op === 'greaterThanOrEqual' ||
          op === 'lessThan' ||
          op === 'lessThanOrEqual'
        ) {
          for (const [row, number] of [0, 5, 10, 15].entries()) {
            seedNumber(js, row, 0, number);
            seedNumber(native, row, 0, number);
            wb.setNumber({ sheet: 0, row, col: 0 }, number);
          }
        } else {
          for (const [row, text] of ['Header', 'East', 'east', 'West'].entries()) {
            seedText(js, row, 0, text);
            seedText(native, row, 0, text);
            wb.setText({ sheet: 0, row, col: 0 }, text);
          }
        }
        const condition = { op, value } as const;
        applyConditionFilter(js.getState(), js, range, 0, condition);
        applyConditionFilter(native.getState(), native, range, 0, condition, wb);
        expect(native.getState().layout.hiddenRows, op).toEqual(js.getState().layout.hiddenRows);
      } finally {
        wb.dispose();
      }
    }
  });

  it('serializes the filter range, visible value checklist, and custom condition', () => {
    const store = createSpreadsheetStore();
    seedText(store, 1, 0, 'East');
    seedText(store, 2, 0, 'West');
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterRanges: [range],
        filterCriteria: [
          { range, byCol: 0, hiddenValues: ['West'] },
          { range, byCol: 1, hiddenValues: [], condition: { op: 'contains', value: 'x&y' } },
        ],
      },
    }));

    expect(autoFilterXmlFromState(store.getState(), range)).toBe(
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
      withEngineSyncMuted: (fn: () => void) => {
        muted += 1;
        fn();
      },
    } as unknown as WorkbookHandle;

    hydrateAutoFilterFromEngine(fake, store, 0);
    expect(muted).toBe(1);
    expect(store.getState().ui.filterRanges).toEqual([range]);
    expect(syncAutoFilterToEngine(fake, store.getState(), 0)).toBe(true);

    const unavailable = { capabilities: { autoFilter: false } } as unknown as WorkbookHandle;
    expect(syncAutoFilterToEngine(unavailable, store.getState(), 0)).toBe(false);
  });

  it('falls back to the raw range when a typed imported filter is opaque', () => {
    const store = createSpreadsheetStore();
    const fake = {
      capabilities: { typedAutoFilter: true, autoFilter: true },
      getAutoFilter: () => null,
      getSheetAutoFilterXml: () => '<autoFilter ref="C3:E9"><filterColumn colId="0"/></autoFilter>',
      withEngineSyncMuted: (fn: () => void) => fn(),
    } as unknown as WorkbookHandle;

    hydrateAutoFilterFromEngine(fake, store, 0);
    expect(store.getState().ui.filterRanges).toEqual([
      {
        sheet: 0,
        r0: 2,
        c0: 2,
        r1: 8,
        c1: 4,
      },
    ]);
    expect(store.getState().ui.filterCriteria).toEqual([]);
  });

  it('uses formulon 0.13 typed values filters and rehydrates their criteria', async () => {
    const first = await WorkbookHandle.createDefault();
    try {
      expect(first.capabilities.typedAutoFilter).toBe(true);
      const typedRange = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
      const store = createSpreadsheetStore();
      seedText(store, 0, 0, 'Region');
      seedText(store, 1, 0, 'East');
      seedText(store, 2, 0, 'West');
      seedText(store, 3, 0, 'East');
      first.setText({ sheet: 0, row: 0, col: 0 }, 'Region');
      first.setText({ sheet: 0, row: 1, col: 0 }, 'East');
      first.setText({ sheet: 0, row: 2, col: 0 }, 'West');
      first.setText({ sheet: 0, row: 3, col: 0 }, 'East');
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterRanges: [typedRange],
          filterCriteria: [{ range: typedRange, byCol: 0, hiddenValues: ['West'] }],
        },
      }));

      expect(syncAutoFilterToEngine(first, store.getState(), 0, store)).toBe(true);
      expect(first.getAutoFilter(0)?.columns[0]).toMatchObject({
        colId: 0,
        kind: 1,
        values: ['East'],
        filterBlank: true,
      });
      expect(first.evaluateAutoFilter(0)?.match).toEqual([true, false, true, true]);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2]));

      const restored = await WorkbookHandle.loadBytes(first.save());
      try {
        const restoredStore = createSpreadsheetStore();
        seedText(restoredStore, 0, 0, 'Region');
        seedText(restoredStore, 1, 0, 'East');
        seedText(restoredStore, 2, 0, 'West');
        seedText(restoredStore, 3, 0, 'East');
        hydrateLayoutFromEngine(restored, restoredStore, 0);
        hydrateAutoFilterFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().ui.filterRanges).toEqual([typedRange]);
        expect(restoredStore.getState().ui.filterCriteria).toEqual([
          { range: typedRange, byCol: 0, hiddenValues: ['West'] },
        ]);
        expect(restored.evaluateAutoFilter(0)?.match).toEqual([true, false, true, true]);
        expect(restoredStore.getState().layout.hiddenRows).toEqual(new Set([2]));
      } finally {
        restored.dispose();
      }
    } finally {
      first.dispose();
    }
  });

  it('maps one typed custom criterion to the existing condition filter', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.typedAutoFilter).toBe(true);
      const typedRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
      const store = createSpreadsheetStore();
      seedText(store, 0, 0, 'Name');
      seedText(store, 1, 0, 'East');
      seedText(store, 2, 0, 'West');
      wb.setText({ sheet: 0, row: 0, col: 0 }, 'Name');
      wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
      wb.setText({ sheet: 0, row: 2, col: 0 }, 'West');
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterRanges: [typedRange],
          filterCriteria: [
            {
              range: typedRange,
              byCol: 0,
              hiddenValues: [],
              condition: { op: 'contains', value: 'a' },
            },
          ],
        },
      }));

      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({
        kind: 2,
        customCount: 1,
        op1: 0,
        val1: '*a*',
      });
      expect(wb.evaluateAutoFilter(0)?.match).toEqual([true, false, false]);

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(restored.capabilities.typedAutoFilter).toBe(true);
        expect(restored.getAutoFilter(0)?.columns[0]).toMatchObject({
          kind: 2,
          customCount: 1,
          val1: '*a*',
        });
        expect(restored.evaluateAutoFilter(0)?.match).toEqual([true, false, false]);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('escapes AutoFilter wildcard literals and decodes them after reload', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.typedAutoFilter).toBe(true);
      const typedRange = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 };
      const value = 'a*b?~';
      const rows = ['Name', value, 'axb?~', 'a*bX~', 'a*b?x'];
      const store = createSpreadsheetStore();
      for (let row = 0; row < rows.length; row += 1) {
        const current = rows[row] ?? '';
        wb.setText({ sheet: 0, row, col: 0 }, current);
        seedText(store, row, 0, current);
      }
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterRanges: [typedRange],
          filterCriteria: [
            {
              range: typedRange,
              byCol: 0,
              hiddenValues: [],
              condition: { op: 'contains', value },
            },
          ],
        },
      }));

      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({
        kind: 2,
        op1: 0,
        val1: '*a~*b~?~~*',
      });
      // The native matcher must see the escaped pattern: only the literal
      // `*`, `?`, and `~` row matches.
      expect(wb.evaluateAutoFilter(0)?.match).toEqual([true, false, false, false]);
      expect(autoFilterXmlFromState(store.getState(), typedRange)).toContain('val="*a~*b~?~~*"');

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = createSpreadsheetStore();
        for (let row = 0; row < rows.length; row += 1) {
          seedText(restoredStore, row, 0, rows[row] ?? '');
        }
        hydrateAutoFilterFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().ui.filterCriteria).toEqual([
          {
            range: typedRange,
            byCol: 0,
            hiddenValues: [],
            condition: { op: 'contains', value },
          },
        ]);
      } finally {
        restored.dispose();
      }

      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterCriteria: [
            {
              range: typedRange,
              byCol: 0,
              hiddenValues: [],
              condition: { op: 'equals', value },
            },
          ],
        },
      }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({
        kind: 2,
        op1: 0,
        val1: 'a~*b~?~~',
      });
      expect(wb.evaluateAutoFilter(0)?.match).toEqual([true, false, false, false]);
      expect(autoFilterXmlFromState(store.getState(), typedRange)).toContain('val="a~*b~?~~"');
      hydrateAutoFilterFromEngine(wb, store, 0);
      expect(store.getState().ui.filterCriteria).toEqual([
        {
          range: typedRange,
          byCol: 0,
          hiddenValues: [],
          condition: { op: 'equals', value },
        },
      ]);
    } finally {
      wb.dispose();
    }
  });

  it('replays filter history through the native definition and visibility paths', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.typedAutoFilter).toBe(true);
      const typedRange = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
      const store = createSpreadsheetStore();
      seedText(store, 0, 0, 'Name');
      seedText(store, 1, 0, 'East');
      seedText(store, 2, 0, 'West');
      wb.setText({ sheet: 0, row: 0, col: 0 }, 'Name');
      wb.setText({ sheet: 0, row: 1, col: 0 }, 'East');
      wb.setText({ sheet: 0, row: 2, col: 0 }, 'West');
      wb.attachStore(store);
      const history = new History();

      recordFilterChange(
        history,
        store,
        () => {
          applyValueFilter(store.getState(), store, typedRange, 0, ['West'], wb);
        },
        wb,
      );
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({ kind: 1, values: ['East'] });
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2]));
      expect(
        wb
          .getRowLayouts(0)
          .filter((row) => row.hidden)
          .map((row) => row.row),
      ).toEqual([2]);

      expect(history.undo()).toBe(true);
      expect(wb.getAutoFilter(0)).toBeNull();
      expect(store.getState().layout.hiddenRows).toEqual(new Set());
      expect(wb.getRowLayouts(0).filter((row) => row.hidden)).toHaveLength(0);

      expect(history.redo()).toBe(true);
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({ kind: 1, values: ['East'] });
      expect(wb.evaluateAutoFilter(0)?.match).toEqual([true, false]);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2]));
      expect(
        wb
          .getRowLayouts(0)
          .filter((row) => row.hidden)
          .map((row) => row.row),
      ).toEqual([2]);
    } finally {
      wb.dispose();
    }
  });

  it('restores exact native row flags through reapply undo and redo', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      const history = new History();
      const typedRange = { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 };
      for (const [row, value] of ['Region', 'East', 'West'].entries()) {
        seedText(store, row, 0, value);
        wb.setText({ sheet: 0, row, col: 0 }, value);
      }
      wb.attachStore(store);
      applyValueFilter(store.getState(), store, typedRange, 0, ['West'], wb);
      seedText(store, 1, 0, 'West');
      wb.setText({ sheet: 0, row: 1, col: 0 }, 'West');

      recordFilterChange(history, store, () => reapplyFilters(store.getState(), store, wb), wb);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([1, 2]));
      expect(
        new Set(
          wb
            .getRowLayouts(0)
            .filter((row) => row.hidden)
            .map((row) => row.row),
        ),
      ).toEqual(new Set([1, 2]));

      expect(history.undo()).toBe(true);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2]));
      expect(
        wb
          .getRowLayouts(0)
          .filter((row) => row.hidden)
          .map((row) => row.row),
      ).toEqual([2]);
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = createSpreadsheetStore();
        hydrateLayoutFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().layout.hiddenRows).toEqual(new Set([2]));
      } finally {
        restored.dispose();
      }

      expect(history.redo()).toBe(true);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([1, 2]));
      expect(
        new Set(
          wb
            .getRowLayouts(0)
            .filter((row) => row.hidden)
            .map((row) => row.row),
        ),
      ).toEqual(new Set([1, 2]));
    } finally {
      wb.dispose();
    }
  });

  it('persists advanced-filter hidden rows alongside an empty typed definition', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.typedAutoFilter).toBe(true);
      expect(wb.capabilities.hiddenRowsCols).toBe(true);
      const store = createSpreadsheetStore();
      const listRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
      const rows = ['Name', 'Smith', 'Smart', 'Jones'];
      for (let row = 0; row < rows.length; row += 1) {
        const value = rows[row] ?? '';
        seedText(store, row, 0, value);
        wb.setText({ sheet: 0, row, col: 0 }, value);
      }
      seedText(store, 5, 0, 'Name');
      seedText(store, 6, 0, '=smith');
      wb.setText({ sheet: 0, row: 5, col: 0 }, 'Name');
      wb.setText({ sheet: 0, row: 6, col: 0 }, '=smith');
      wb.attachStore(store);

      expect(
        applyAdvancedFilter(store.getState(), store, listRange, {
          sheet: 0,
          r0: 5,
          c0: 0,
          r1: 6,
          c1: 0,
        }),
      ).toBe(2);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2, 3]));
      expect(wb.getAutoFilter(0)?.columns).toHaveLength(0);
      expect(wb.evaluateAutoFilter(0)?.match).toEqual([true, true, true]);
      expect(
        wb
          .getRowLayouts(0)
          .filter((row) => row.hidden)
          .map((row) => row.row),
      ).toEqual([2, 3]);

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = createSpreadsheetStore();
        hydrateLayoutFromEngine(restored, restoredStore, 0);
        hydrateAutoFilterFromEngine(restored, restoredStore, 0);
        expect(restored.getAutoFilter(0)?.columns).toHaveLength(0);
        expect(restored.evaluateAutoFilter(0)?.match).toEqual([true, true, true]);
        expect(restoredStore.getState().layout.hiddenRows).toEqual(new Set([2, 3]));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('mirrors repeated advanced-filter results when the definition is unchanged', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.hiddenRowsCols).toBe(true);
      const store = createSpreadsheetStore();
      const listRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
      const rows = ['Name', 'Smith', 'Smart', 'Jones'];
      for (let row = 0; row < rows.length; row += 1) {
        const value = rows[row] ?? '';
        seedText(store, row, 0, value);
        wb.setText({ sheet: 0, row, col: 0 }, value);
      }
      seedText(store, 5, 0, 'Name');
      seedText(store, 6, 0, '=smith');
      wb.setText({ sheet: 0, row: 5, col: 0 }, 'Name');
      wb.setText({ sheet: 0, row: 6, col: 0 }, '=smith');
      wb.attachStore(store);
      const criteriaRange = { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 };

      expect(applyAdvancedFilter(store.getState(), store, listRange, criteriaRange)).toBe(2);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2, 3]));

      seedText(store, 6, 0, '=smart');
      wb.setText({ sheet: 0, row: 6, col: 0 }, '=smart');
      expect(applyAdvancedFilter(store.getState(), store, listRange, criteriaRange)).toBe(2);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([1, 3]));
      expect(
        wb
          .getRowLayouts(0)
          .filter((row) => row.hidden)
          .map((row) => row.row),
      ).toEqual([1, 3]);

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = createSpreadsheetStore();
        hydrateLayoutFromEngine(restored, restoredStore, 0);
        hydrateAutoFilterFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().layout.hiddenRows).toEqual(new Set([1, 3]));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('mirrors repeated slicer-style results when only hidden rows change', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.hiddenRowsCols).toBe(true);
      const store = createSpreadsheetStore();
      const filterRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
      const rows = ['Region', 'East', 'West', 'North'];
      for (let row = 0; row < rows.length; row += 1) {
        const value = rows[row] ?? '';
        seedText(store, row, 0, value);
        wb.setText({ sheet: 0, row, col: 0 }, value);
      }
      wb.attachStore(store);
      expect(
        applyFilterColumns(store.getState(), store, filterRange, [
          {
            byCol: 0,
            predicate: (cell) => filterValueKey(cell?.value) === 'East',
          },
        ]),
      ).toBe(2);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2, 3]));

      expect(
        applyFilterColumns(store.getState(), store, filterRange, [
          {
            byCol: 0,
            predicate: (cell) => filterValueKey(cell?.value) === 'West',
          },
        ]),
      ).toBe(2);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([1, 3]));
      expect(
        new Set(
          wb
            .getRowLayouts(0)
            .filter((row) => row.hidden)
            .map((row) => row.row),
        ),
      ).toEqual(new Set([1, 3]));

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = createSpreadsheetStore();
        hydrateLayoutFromEngine(restored, restoredStore, 0);
        hydrateAutoFilterFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().layout.hiddenRows).toEqual(new Set([1, 3]));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('round-trips a JS color mask alongside the native color column', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.typedAutoFilter).toBe(true);
      expect(wb.capabilities.hiddenRowsCols).toBe(true);
      const store = createSpreadsheetStore();
      const filterRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
      const rows = ['Name', 'East', 'West', 'South'];
      for (let row = 0; row < rows.length; row += 1) {
        const value = rows[row] ?? '';
        seedText(store, row, 0, value);
        wb.setText({ sheet: 0, row, col: 0 }, value);
      }
      mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#ff0000' });
      mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { fill: '#00ff00' });
      mutators.setCellFormat(store, { sheet: 0, row: 3, col: 0 }, { fill: '#00ff00' });
      flushFormatToEngine(wb, store, 0);

      expect(
        applyColorFilter(
          store.getState(),
          store,
          filterRange,
          0,
          { kind: 'cellColor', color: '#ff0000' },
          wb,
        ),
      ).toBe(2);
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2, 3]));
      expect(wb.getAutoFilter(0)?.columns).toHaveLength(1);
      expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({ colId: 0, cellColor: true });
      expect(
        wb
          .getRowLayouts(0)
          .filter((row) => row.hidden)
          .map((row) => row.row),
      ).toEqual([2, 3]);

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const restoredStore = createSpreadsheetStore();
        hydrateLayoutFromEngine(restored, restoredStore, 0);
        hydrateAutoFilterFromEngine(restored, restoredStore, 0);
        expect(restoredStore.getState().layout.hiddenRows).toEqual(new Set([2, 3]));
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps JS color visibility while preserving imported native columns', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.typedAutoFilter).toBe(true);
      const filterRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
      const store = createSpreadsheetStore();
      const rows = [
        ['Name', 'Score'],
        ['East', '1'],
        ['West', '2'],
        ['East', '3'],
      ];
      for (let row = 0; row < rows.length; row += 1) {
        for (let col = 0; col < 2; col += 1) {
          const value = rows[row]?.[col] ?? '';
          wb.setText({ sheet: 0, row, col }, value);
          seedText(store, row, col, value);
        }
      }
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterRanges: [filterRange],
          filterCriteria: [{ range: filterRange, byCol: 0, hiddenValues: ['West'] }],
        },
      }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      const valuesColumn = wb.getAutoFilter(0)?.columns[0];
      if (!valuesColumn) throw new Error('native values filter column missing');
      const importedUnsupported = {
        ...valuesColumn,
        colId: 1,
        kind: 3,
        top: true,
        topVal: 1,
        percent: false,
      } as AutoFilterColumn;
      expect(
        wb.setAutoFilter(0, {
          range: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 },
          columns: [valuesColumn, importedUnsupported],
          sort: null,
        }),
      ).toBe(true);
      hydrateAutoFilterFromEngine(wb, store, 0);
      mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#ff0000' });
      mutators.setCellFormat(store, { sheet: 0, row: 2, col: 0 }, { fill: '#00ff00' });
      mutators.setCellFormat(store, { sheet: 0, row: 3, col: 0 }, { fill: '#ff0000' });

      applyColorFilter(
        store.getState(),
        store,
        filterRange,
        0,
        { kind: 'cellColor', color: '#ff0000' },
        wb,
      );
      expect(store.getState().layout.hiddenRows).toEqual(new Set([2]));
      expect(wb.getAutoFilter(0)?.columns.find((column) => column.colId === 1)).toMatchObject({
        colId: 1,
        kind: 3,
        top: true,
        topVal: 1,
      });
    } finally {
      wb.dispose();
    }
  });

  it('keeps an imported unsupported criterion while a supported column is edited', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.capabilities.typedAutoFilter).toBe(true);
      const filterRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 };
      const store = createSpreadsheetStore();
      const values = [
        ['Name', 'Score'],
        ['East', '1'],
        ['West', '2'],
        ['East', '3'],
      ];
      for (let row = 0; row < values.length; row += 1) {
        const rowValues = values[row] ?? [];
        for (let col = 0; col < rowValues.length; col += 1) {
          const value = rowValues[col] ?? '';
          wb.setText({ sheet: 0, row, col }, value);
          seedText(store, row, col, value);
        }
      }
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterRanges: [filterRange],
          filterCriteria: [{ range: filterRange, byCol: 0, hiddenValues: ['West'] }],
        },
      }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      const valuesColumn = wb.getAutoFilter(0)?.columns[0];
      expect(valuesColumn?.kind).toBe(1);
      if (!valuesColumn) throw new Error('native values filter column missing');
      const importedUnsupported = {
        ...valuesColumn,
        colId: 1,
        kind: 3,
        top: true,
        topVal: 1,
        percent: false,
      };
      expect(
        wb.setAutoFilter(0, {
          range: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 },
          columns: [valuesColumn, importedUnsupported as AutoFilterColumn],
          sort: null,
        }),
      ).toBe(true);
      expect(wb.applyAutoFilter(0)).toBe(true);
      hydrateAutoFilterFromEngine(wb, store, 0);
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterCriteria: [{ range: filterRange, byCol: 0, hiddenValues: [] }],
        },
      }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      const after = wb.getAutoFilter(0);
      expect(after?.columns.find((column) => column.colId === 1)).toMatchObject({
        colId: 1,
        kind: 3,
        top: true,
        topVal: 1,
      });

      const importedDateGroups = {
        ...valuesColumn,
        colId: 1,
        kind: 1,
        values: [],
        dateGroups: [{ year: 2024, month: 0, day: 0, hour: 0, minute: 0, second: 0, grouping: 0 }],
      } as AutoFilterColumn;
      expect(
        wb.setAutoFilter(0, {
          range: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 },
          columns: [valuesColumn, importedDateGroups],
          sort: null,
        }),
      ).toBe(true);
      expect(wb.applyAutoFilter(0)).toBe(true);
      hydrateAutoFilterFromEngine(wb, store, 0);
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterCriteria: [{ range: filterRange, byCol: 0, hiddenValues: [] }],
        },
      }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getAutoFilter(0)?.columns.find((column) => column.colId === 1)).toMatchObject({
        colId: 1,
        kind: 1,
        dateGroups: [{ year: 2024, grouping: 0 }],
      });

      const importedButtonFlags = {
        ...valuesColumn,
        colId: 1,
        hiddenButton: true,
        showButton: false,
      } as AutoFilterColumn;
      expect(
        wb.setAutoFilter(0, {
          range: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 },
          columns: [valuesColumn, importedButtonFlags],
          sort: null,
        }),
      ).toBe(true);
      hydrateAutoFilterFromEngine(wb, store, 0);
      store.setState((state) => ({
        ...state,
        ui: {
          ...state.ui,
          filterCriteria: [{ range: filterRange, byCol: 0, hiddenValues: [] }],
        },
      }));
      expect(syncAutoFilterToEngine(wb, store.getState(), 0, store)).toBe(true);
      expect(wb.getAutoFilter(0)?.columns.find((column) => column.colId === 1)).toMatchObject({
        colId: 1,
        hiddenButton: true,
        showButton: false,
      });
    } finally {
      wb.dispose();
    }
  });
});
