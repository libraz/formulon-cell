// @vitest-environment node

import { afterEach, describe, expect, it } from 'vitest';
import {
  applyAdvancedFilter,
  applyConditionFilter,
  distinctFilterItems,
  distinctFilterItemsResult,
} from '../../src/commands/filter.js';
import { removeDuplicates, sortRange } from '../../src/commands/sort.js';
import { hydrateAutoFilterFromEngine } from '../../src/engine/auto-filter-sync.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators, type SpreadsheetStore } from '../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 };

const seedText = (wb: WorkbookHandle, row: number, col: number, value: string): void => {
  wb.setText({ sheet: 0, row, col }, value);
};

const seedNumber = (wb: WorkbookHandle, row: number, col: number, value: number): void => {
  wb.setNumber({ sheet: 0, row, col }, value);
};

const sync = (store: SpreadsheetStore, wb: WorkbookHandle): void => {
  mutators.replaceCells(store, wb.cells(0));
};

const rawNumber = (value: number) => ({
  kind: 1,
  number: value,
  boolean: 0,
  text: '',
  errorCode: 0,
});

const PROFILE_IDS = [
  'windows-en_US',
  'windows-ja_JP',
  'windows-de_DE',
  'windows-fr_FR',
  'windows-zh_CN',
  'windows-ko_KR',
  'windows-th_TH',
  'mac-en_US',
  'mac-ja_JP',
  'mac-de_DE',
  'mac-fr_FR',
  'mac-zh_CN',
  'mac-ko_KR',
  'mac-th_TH',
] as const;

describe('locale sort and filter boundaries', () => {
  const workbooks: WorkbookHandle[] = [];

  afterEach(() => {
    for (const wb of workbooks.splice(0)) wb.dispose();
  });

  it('uses one read-only native SORTBY call for mixed multi-key sorting', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Region');
    seedText(wb, 0, 1, 'Item');
    seedText(wb, 1, 0, 'West');
    seedText(wb, 1, 1, 'a');
    seedText(wb, 2, 0, 'East');
    seedText(wb, 2, 1, 'b');
    seedText(wb, 3, 0, 'East');
    seedText(wb, 3, 1, 'a');
    seedText(wb, 4, 0, 'West');
    seedText(wb, 4, 1, 'b');
    sync(store, wb);

    const original = wb.evaluateFormulaArray.bind(wb);
    let calls = 0;
    wb.evaluateFormulaArray = ((addr, formula) => {
      calls += 1;
      expect(formula).toContain('SORTBY(SEQUENCE(4)');
      return original(addr, formula);
    }) as WorkbookHandle['evaluateFormulaArray'];

    expect(
      sortRange(store.getState(), store, wb, range, {
        byCol: 0,
        direction: 'asc',
        hasHeader: true,
        keys: [
          { byCol: 0, direction: 'asc' },
          { byCol: 1, direction: 'desc' },
        ],
      }),
    ).toBe(true);
    expect(calls).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({
      kind: 'text',
      value: 'East',
    });
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({
      kind: 'text',
      value: 'b',
    });
  });

  it('keeps SORTBY permutation parity across all fourteen profiles', async () => {
    for (const profile of PROFILE_IDS) {
      const wb = await WorkbookHandle.createDefault();
      workbooks.push(wb);
      expect(wb.setSpreadsheetProfileId(profile), profile).toBe(true);
      const store = createSpreadsheetStore();
      const values = ['ä', 'a', 'b'];
      seedText(wb, 0, 0, 'Value');
      seedText(wb, 0, 1, 'Rank');
      seedText(wb, 1, 0, values[0] ?? '');
      seedNumber(wb, 1, 1, 3);
      seedText(wb, 2, 0, values[1] ?? '');
      seedNumber(wb, 2, 1, 2);
      seedText(wb, 3, 0, values[2] ?? '');
      seedNumber(wb, 3, 1, 1);
      sync(store, wb);

      const direct = wb.evaluateFormulaArray(
        { sheet: 0, row: 0, col: 2 },
        '=SORTBY(SEQUENCE(3),$A$2:$A$4,1)',
      );
      expect(direct.status.ok, profile).toBe(true);
      const expectedIndexes = direct.cells.map((row) => row[0]?.number ?? 0);
      expect(
        sortRange(
          store.getState(),
          store,
          wb,
          { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
          { byCol: 0, direction: 'asc', hasHeader: true },
        ),
        profile,
      ).toBe(true);
      expect(
        [1, 2, 3].map((row) => wb.getValue({ sheet: 0, row, col: 0 })),
        profile,
      ).toEqual(expectedIndexes.map((index) => ({ kind: 'text', value: values[index - 1] })));
    }
  });

  it('precomputes color and custom ranks while keeping native secondary text keys', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Sort');
    seedText(wb, 0, 1, 'Secondary');
    seedText(wb, 1, 0, 'Mar');
    seedText(wb, 1, 1, 'a');
    seedText(wb, 2, 0, 'Jan');
    seedText(wb, 2, 1, 'z');
    seedText(wb, 3, 0, 'Feb');
    seedText(wb, 3, 1, 'b');
    mutators.setCellFormat(store, { sheet: 0, row: 1, col: 0 }, { fill: '#ff0000' });
    sync(store, wb);
    let formula = '';
    const original = wb.evaluateFormulaArray.bind(wb);
    wb.evaluateFormulaArray = ((addr, text) => {
      formula = text;
      return original(addr, text);
    }) as WorkbookHandle['evaluateFormulaArray'];

    expect(
      sortRange(
        store.getState(),
        store,
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 },
        {
          byCol: 0,
          direction: 'asc',
          hasHeader: true,
          keys: [
            { byCol: 0, direction: 'asc', sortOn: 'cellColor', color: '#ff0000' },
            { byCol: 0, direction: 'asc', sortOn: 'customList', customList: ['Jan', 'Feb', 'Mar'] },
            { byCol: 1, direction: 'desc' },
          ],
        },
      ),
    ).toBe(true);
    expect(formula).toContain('SORTBY(SEQUENCE(3)');
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'text', value: 'Mar' });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'text', value: 'Jan' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'text', value: 'Feb' });
  });

  it('rejects an invalid native permutation without writing cells', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedNumber(wb, 0, 0, 3);
    seedNumber(wb, 1, 0, 1);
    seedNumber(wb, 2, 0, 2);
    sync(store, wb);
    const before = [0, 1, 2].map((row) => wb.getValue({ sheet: 0, row, col: 0 }));
    wb.evaluateFormulaArray = (() => ({
      status: { ok: true, status: 0, message: '', context: '' },
      rows: 3,
      cols: 1,
      cells: [[rawNumber(1)], [rawNumber(1)], [rawNumber(3)]],
    })) as WorkbookHandle['evaluateFormulaArray'];

    expect(
      sortRange(
        store.getState(),
        store,
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
        { byCol: 0, direction: 'asc' },
      ),
    ).toBe(false);
    expect([0, 1, 2].map((row) => wb.getValue({ sheet: 0, row, col: 0 }))).toEqual(before);
  });

  it('uses native UNIQUE for Remove Duplicates and keeps first representatives', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Name');
    seedNumber(wb, 0, 1, 1);
    seedText(wb, 1, 0, 'A');
    seedNumber(wb, 1, 1, 10);
    seedText(wb, 2, 0, 'A');
    seedNumber(wb, 2, 1, 20);
    seedText(wb, 3, 0, 'B');
    seedNumber(wb, 3, 1, 30);
    seedText(wb, 4, 0, 'B');
    seedNumber(wb, 4, 1, 40);
    sync(store, wb);
    const original = wb.evaluateFormulaArray.bind(wb);
    let formula = '';
    wb.evaluateFormulaArray = ((addr, text) => {
      formula = text;
      return original(addr, text);
    }) as WorkbookHandle['evaluateFormulaArray'];

    expect(
      removeDuplicates(store.getState(), store, wb, range, {
        columns: [0],
        hasHeader: true,
      }),
    ).toBe(2);
    expect(formula).toContain('UNIQUE');
    expect(wb.getValue({ sheet: 0, row: 1, col: 1 })).toEqual({ kind: 'number', value: 10 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 1 })).toEqual({ kind: 'number', value: 30 });
  });

  it('recovers native UNIQUE blank representatives without confusing numeric zero', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Value');
    seedNumber(wb, 1, 0, 0);
    // Row 2 stays blank in both the engine and the hydrated store.
    seedNumber(wb, 3, 0, 0);
    seedText(wb, 4, 0, 'x');
    sync(store, wb);

    expect(
      removeDuplicates(
        store.getState(),
        store,
        wb,
        { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
        {
          columns: [0],
          hasHeader: true,
        },
      ),
    ).toBe(1);
    expect(wb.getValue({ sheet: 0, row: 1, col: 0 })).toEqual({ kind: 'number', value: 0 });
    expect(wb.getValue({ sheet: 0, row: 2, col: 0 })).toEqual({ kind: 'blank' });
    expect(wb.getValue({ sheet: 0, row: 3, col: 0 })).toEqual({ kind: 'text', value: 'x' });
  });

  it('canonicalizes German numeric criteria and Advanced Filter thresholds', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    expect(wb.setSpreadsheetProfileId('windows-de_DE')).toBe(true);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Amount');
    seedNumber(wb, 1, 0, 1);
    seedNumber(wb, 2, 0, 2);
    seedNumber(wb, 3, 0, 3);
    sync(store, wb);
    const dataRange = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
    expect(
      applyConditionFilter(
        store.getState(),
        store,
        dataRange,
        0,
        {
          op: 'greaterThan',
          value: '1,5',
        },
        wb,
      ),
    ).toBe(1);
    expect(store.getState().ui.filterCriteria).toMatchObject([
      { condition: { op: 'greaterThan', value: '1.5' } },
    ]);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(false);
    expect(wb.getAutoFilter(0)?.columns[0]?.val1).toBe('1.5');
    expect(wb.getAutoFilter(0)?.columns[0]).toMatchObject({ kind: 2, op1: 5 });
    hydrateAutoFilterFromEngine(wb, store, 0);
    expect(store.getState().ui.filterCriteria).toMatchObject([
      { condition: { op: 'greaterThan', value: '1.5' } },
    ]);

    const criteriaRange = { sheet: 0, r0: 5, c0: 0, r1: 6, c1: 0 };
    seedText(wb, 5, 0, 'Amount');
    seedText(wb, 6, 0, '>1,5');
    sync(store, wb);
    expect(applyAdvancedFilter(store.getState(), store, dataRange, criteriaRange, wb)).toBe(1);
    expect(store.getState().layout.hiddenRows.has(1)).toBe(true);
    expect(store.getState().layout.hiddenRows.has(2)).toBe(false);
  });

  it('uses native UNIQUE/SORT for a checklist while preserving value types', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Value');
    seedText(wb, 1, 0, 'ä');
    seedText(wb, 2, 0, 'a');
    seedText(wb, 3, 0, 'a');
    seedNumber(wb, 4, 0, 2);
    sync(store, wb);
    const original = wb.evaluateFormulaArray.bind(wb);
    let calls = 0;
    wb.evaluateFormulaArray = ((addr, formula) => {
      calls += 1;
      expect(formula).toContain('SORT(UNIQUE');
      return original(addr, formula);
    }) as WorkbookHandle['evaluateFormulaArray'];
    const items = distinctFilterItems(
      store.getState(),
      { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 0 },
      0,
      'de-DE',
      wb,
    );
    expect(calls).toBe(1);
    expect(items.map((item) => item.key)).toEqual(['2', 'a', 'ä']);
  });

  it('keeps native UNIQUE type identities before the public string-key boundary', async () => {
    const wb = await WorkbookHandle.createDefault();
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Value');
    seedNumber(wb, 1, 0, 1);
    seedText(wb, 2, 0, '1');
    wb.setBool({ sheet: 0, row: 3, col: 0 }, true);
    seedText(wb, 4, 0, 'TRUE');
    // Row 5 is blank; the engine's array ABI reports it as numeric zero.
    seedNumber(wb, 6, 0, 0);
    sync(store, wb);

    const items = distinctFilterItems(
      store.getState(),
      { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 0 },
      0,
      'en-US',
      wb,
    );
    expect(items).toHaveLength(6);
    expect(items.filter((item) => item.key === '1')).toHaveLength(2);
    expect(items.filter((item) => item.key === 'TRUE')).toHaveLength(2);
    expect(items.map((item) => item.key).sort()).toEqual(['', '0', '1', '1', 'TRUE', 'TRUE']);
  });

  it('refuses oversized native evaluation before calling the engine', async () => {
    const sortWb = await WorkbookHandle.createDefault();
    workbooks.push(sortWb);
    const sortStore = createSpreadsheetStore();
    let sortCalls = 0;
    sortWb.evaluateFormulaArray = (() => {
      sortCalls += 1;
      throw new Error('oversized SORTBY must not be evaluated');
    }) as WorkbookHandle['evaluateFormulaArray'];
    expect(
      sortRange(
        sortStore.getState(),
        sortStore,
        sortWb,
        { sheet: 0, r0: 0, c0: 0, r1: 50_000, c1: 0 },
        { byCol: 0, direction: 'asc' },
      ),
    ).toBe(false);
    expect(sortCalls).toBe(0);

    const uniqueWb = await WorkbookHandle.createDefault();
    workbooks.push(uniqueWb);
    const uniqueStore = createSpreadsheetStore();
    let uniqueCalls = 0;
    uniqueWb.evaluateFormulaArray = (() => {
      uniqueCalls += 1;
      throw new Error('oversized UNIQUE must not be evaluated');
    }) as WorkbookHandle['evaluateFormulaArray'];
    expect(
      removeDuplicates(
        uniqueStore.getState(),
        uniqueStore,
        uniqueWb,
        { sheet: 0, r0: 0, c0: 0, r1: 20_001, c1: 0 },
        { hasHeader: true, columns: [0] },
      ),
    ).toBe(0);
    expect(uniqueCalls).toBe(0);

    const filterWb = await WorkbookHandle.createDefault();
    workbooks.push(filterWb);
    const filterStore = createSpreadsheetStore();
    seedText(filterWb, 0, 0, 'Value');
    seedText(filterWb, 1, 0, 'a');
    sync(filterStore, filterWb);
    let filterCalls = 0;
    filterWb.evaluateFormulaArray = (() => {
      filterCalls += 1;
      throw new Error('oversized checklist must not be evaluated');
    }) as WorkbookHandle['evaluateFormulaArray'];
    const items = distinctFilterItems(
      filterStore.getState(),
      { sheet: 0, r0: 0, c0: 0, r1: 20_001, c1: 0 },
      0,
      'en-US',
      filterWb,
    );
    expect(filterCalls).toBe(0);
    expect(items.map((item) => item.key)).toContain('a');
    expect(
      distinctFilterItemsResult(
        filterStore.getState(),
        { sheet: 0, r0: 0, c0: 0, r1: 20_001, c1: 0 },
        0,
        'en-US',
        filterWb,
      ).failure,
    ).toBe('native-limit');
  });

  it('uses explicit locale semantics in the stub fallback', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    workbooks.push(wb);
    const store = createSpreadsheetStore();
    seedText(wb, 0, 0, 'Value');
    seedText(wb, 1, 0, 'ä');
    seedText(wb, 2, 0, 'a');
    sync(store, wb);
    expect(
      distinctFilterItems(
        store.getState(),
        { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 0 },
        0,
        'de-DE',
      ).map((item) => item.key),
    ).toEqual(['ä', 'a']);
  });
});
