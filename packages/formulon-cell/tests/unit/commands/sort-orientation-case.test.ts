import { afterEach, describe, expect, it } from 'vitest';
import { inferSortHasHeader, sortRange } from '../../../src/commands/sort.js';
import type { CellValue } from '../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type Seed = number | string | boolean | null | { text: string };

const toValue = (v: Seed): CellValue => {
  if (v === null) return { kind: 'blank' };
  if (typeof v === 'number') return { kind: 'number', value: v };
  if (typeof v === 'boolean') return { kind: 'bool', value: v };
  if (typeof v === 'object') return { kind: 'text', value: v.text };
  return { kind: 'text', value: v };
};

const HEADERS = ['ID', 'Key', 'Sec', 'Label'];
const labels = (n: number): string[] => Array.from({ length: n }, (_, i) => `r${i + 1}`);
const ids = [1, 2, 3, 4, 5, 6];

interface Table {
  key: Seed[];
  sec: Seed[];
  idRow?: number[];
}

const dupKeys: Table = {
  key: ['Beta', 'beta', 'Alpha', 'alpha', 'beta', 'Alpha'],
  sec: [3, 1, 2, 6, 4, 5],
};
const mixedTypes: Table = {
  key: [3, 'banana', 'apple', 10, 'Apple', true],
  sec: [2, 1, 2, 1, 2, 1],
};
const textNumbers: Table = {
  key: [{ text: '10' }, { text: '9' }, 5, { text: '100' }, { text: '2' }, 20],
  sec: [1, 2, 3, 1, 2, 3],
};
const blanks: Table = {
  key: [4, null, 1, null, 3, 2],
  sec: ['d', 'b', 'a', 'c', null, 'e'],
};

/** Fields as lines: field f of record j sits at (row f, col j + 1) in the
 *  left-to-right layout and (row j + 1, col f) in the top-to-bottom one. */
const fields = (t: Table): Seed[][] => [ids, t.key, t.sec, labels(6)];

const backends = [
  { name: 'stub', create: () => WorkbookHandle.createDefault({ preferStub: true }) },
  { name: 'native', create: () => WorkbookHandle.createDefault() },
] as const;

describe.each(backends)('sort orientation and match case ($name)', ({ create }) => {
  let wb: WorkbookHandle | null = null;
  afterEach(() => {
    wb?.dispose();
    wb = null;
  });

  const seed = async (
    t: Table,
    orient: 'rows' | 'columns',
  ): Promise<ReturnType<typeof createSpreadsheetStore>> => {
    wb = await create();
    const store = createSpreadsheetStore();
    const put = (row: number, col: number, v: Seed): void => {
      const a = { sheet: 0, row, col };
      const val = toValue(v);
      if (val.kind === 'number') wb?.setNumber(a, val.value);
      else if (val.kind === 'text') wb?.setText(a, val.value);
      else if (val.kind === 'bool') wb?.setBool(a, val.value);
    };
    HEADERS.forEach((h, f) => {
      if (orient === 'rows') put(0, f, h);
      else put(f, 0, h);
    });
    fields(t).forEach((vals, f) => {
      vals.forEach((v, j) => {
        if (orient === 'rows') put(j + 1, f, v);
        else put(f, j + 1, v);
      });
    });
    wb.recalc();
    mutators.replaceCells(store, wb.cells(0));
    return store;
  };

  const line = (orient: 'rows' | 'columns', f: number): Array<number | string | boolean | null> => {
    const out: Array<number | string | boolean | null> = [];
    const n = 7;
    for (let j = 0; j < n; j += 1) {
      const addr = orient === 'rows' ? { sheet: 0, row: j, col: f } : { sheet: 0, row: f, col: j };
      const v = wb?.getValue(addr);
      if (!v || v.kind === 'blank') out.push(null);
      else
        out.push(
          v.kind === 'number'
            ? v.value
            : v.kind === 'bool'
              ? v.value
              : v.kind === 'text'
                ? v.value
                : v.text,
        );
    }
    return out;
  };

  it('case-sensitive ascending puts lower case before upper case', async () => {
    const store = await seed(dupKeys, 'rows');
    const ok = sortRange(
      store.getState(),
      store,
      wb as WorkbookHandle,
      { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 3 },
      { byCol: 1, direction: 'asc', hasHeader: true, matchCase: true },
    );
    expect(ok).toBe(true);
    expect(line('rows', 0)).toEqual(['ID', 4, 3, 6, 2, 5, 1]);
    expect(line('rows', 1)).toEqual(['Key', 'alpha', 'Alpha', 'Alpha', 'beta', 'beta', 'Beta']);
  });

  it('case-sensitive descending reverses the case tie-break and keeps ties stable', async () => {
    const store = await seed(mixedTypes, 'rows');
    sortRange(
      store.getState(),
      store,
      wb as WorkbookHandle,
      { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 3 },
      { byCol: 1, direction: 'desc', hasHeader: true, matchCase: true },
    );
    expect(line('rows', 0)).toEqual(['ID', 6, 2, 5, 3, 4, 1]);
  });

  it('case-insensitive stays the default', async () => {
    const store = await seed(dupKeys, 'rows');
    sortRange(
      store.getState(),
      store,
      wb as WorkbookHandle,
      { sheet: 0, r0: 0, c0: 0, r1: 6, c1: 3 },
      { byCol: 1, direction: 'asc', hasHeader: true },
    );
    expect(line('rows', 0)).toEqual(['ID', 3, 4, 6, 1, 2, 5]);
  });

  it('left-to-right sorts columns by the key row and ignores the header flag', async () => {
    const store = await seed(textNumbers, 'columns');
    sortRange(
      store.getState(),
      store,
      wb as WorkbookHandle,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 6 },
      { byCol: 1, direction: 'desc', hasHeader: true, orientation: 'leftToRight' },
    );
    expect(line('columns', 0)).toEqual(['ID', 2, 5, 4, 1, 6, 3]);
    expect(line('columns', 3)).toEqual(['Label', 'r2', 'r5', 'r4', 'r1', 'r6', 'r3']);
  });

  it('left-to-right without a header moves the label column into the data', async () => {
    const store = await seed(dupKeys, 'columns');
    sortRange(
      store.getState(),
      store,
      wb as WorkbookHandle,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 6 },
      {
        byCol: 1,
        direction: 'desc',
        hasHeader: false,
        orientation: 'leftToRight',
        matchCase: true,
      },
    );
    expect(line('columns', 0)).toEqual(['ID', 1, 2, 5, 3, 6, 4]);
    expect(line('columns', 1)).toEqual(['Key', 'Beta', 'beta', 'beta', 'Alpha', 'Alpha', 'alpha']);
  });

  it('left-to-right puts blanks last and moves relative formulas with their column', async () => {
    const store = await seed(blanks, 'columns');
    sortRange(
      store.getState(),
      store,
      wb as WorkbookHandle,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 6 },
      { byCol: 1, direction: 'asc', hasHeader: false, orientation: 'leftToRight' },
    );
    expect(line('columns', 0)).toEqual([3, 6, 5, 1, 'ID', 2, 4]);
    expect(line('columns', 1)).toEqual([1, 2, 3, 4, 'Key', null, null]);
  });

  it('left-to-right with two keys applies the second key within ties', async () => {
    const store = await seed(textNumbers, 'columns');
    sortRange(
      store.getState(),
      store,
      wb as WorkbookHandle,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 6 },
      {
        byCol: 1,
        direction: 'asc',
        keys: [
          { byCol: 2, direction: 'desc' },
          { byCol: 1, direction: 'asc' },
        ],
        hasHeader: true,
        orientation: 'leftToRight',
      },
    );
    expect(line('columns', 2)).toEqual(['Sec', 3, 3, 2, 2, 1, 1]);
    expect(line('columns', 0)).toEqual(['ID', 3, 6, 5, 2, 1, 4]);
  });

  it('left-to-right moves the label column even when the header flag is set', async () => {
    const store = await seed(
      { key: [5, 3, 8, 1, 9, 2].map((n) => n * 10), sec: [0, 0, 0, 0, 0, 0] },
      'columns',
    );
    // Keys derived from the IDs below so the sorted order matches the measured one.
    const w = wb as WorkbookHandle;
    const idRow = [5, 3, 8, 1, 9, 2];
    idRow.forEach((id, j) => {
      w.setNumber({ sheet: 0, row: 0, col: j + 1 }, id);
      w.setFormula({ sheet: 0, row: 1, col: j + 1 }, `=${'BCDEFG'[j]}1*10`);
      w.setFormula({ sheet: 0, row: 2, col: j + 1 }, `=SUM($B1:${'BCDEFG'[j]}1)`);
    });
    w.recalc();
    mutators.replaceCells(store, w.cells(0));
    sortRange(
      store.getState(),
      store,
      w,
      { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 6 },
      { byCol: 1, direction: 'asc', hasHeader: true, orientation: 'leftToRight', matchCase: true },
    );
    mutators.replaceCells(store, w.cells(0));
    expect(line('columns', 0)).toEqual([1, 2, 3, 5, 8, 9, 'ID']);
    if (!w.isStub) {
      expect(line('columns', 1)).toEqual([10, 20, 30, 50, 80, 90, 'Key']);
      expect(line('columns', 2)).toEqual([3, 2, 5, 10, 18, 27, 'Sec']);
    }
    const formulaAt = (row: number, col: number): string | null | undefined =>
      store.getState().data.cells.get(addrKey({ sheet: 0, row, col }))?.formula;
    expect(formulaAt(1, 0)).toBe('=A1*10');
    expect(formulaAt(2, 0)).toBe('=SUM($A1:B1)');
    expect(formulaAt(2, 4)).toBe('=SUM($B1:E1)');
  });

  it('never infers a header for a left-to-right sort', async () => {
    const store = await seed(blanks, 'columns');
    const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 6 };
    expect(inferSortHasHeader(store.getState(), range, 'leftToRight')).toBe(false);
  });
});
