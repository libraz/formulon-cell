import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { moveSheet, removeSheet, renameSheet } from '../../../src/commands/sheet-mutate.js';
import { setSparkline } from '../../../src/commands/sparkline.js';
import { deleteRows, insertCols, insertRows } from '../../../src/commands/structure.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';
import { newWb } from './structure/fixtures.js';

interface TableDef {
  name: string;
  displayName: string;
  ref: string;
  sheetIndex: number;
  columns: string[];
}

/** A three-sheet workbook double whose tables disappear with their sheet. */
const makeWb = (tables: TableDef[] = []): { wb: WorkbookHandle; names: string[] } => {
  const names = ['Alpha', 'Beta', 'Gamma'];
  let live = tables;
  const wb = {
    capabilities: { sheetMutate: true },
    get sheetCount() {
      return names.length;
    },
    sheetName: (i: number) => names[i] ?? '',
    renameSheet: (i: number, name: string) => {
      names[i] = name;
      return true;
    },
    removeSheet: (i: number) => {
      names.splice(i, 1);
      live = live
        .filter((t) => t.sheetIndex !== i)
        .map((t) => (t.sheetIndex > i ? { ...t, sheetIndex: t.sheetIndex - 1 } : t));
      return true;
    },
    moveSheet: (from: number, to: number) => {
      const [name] = names.splice(from, 1);
      if (name) names.splice(to, 0, name);
      return true;
    },
    getTables: () => live,
  } as unknown as WorkbookHandle;
  return { wb, names };
};

const shape = (sheet: number, id: string) => ({
  id,
  kind: 'shape' as const,
  shape: 'rectangle' as const,
  sheet,
  x: 10,
  y: 10,
  w: 40,
  h: 20,
});

const sheetsOf = (store: SpreadsheetStore): Record<string, number> =>
  Object.fromEntries(store.getState().illustrations.illustrations.map((i) => [i.id, i.sheet]));

describe('session objects across sheet moves and removals', () => {
  it('carries session shapes along with a moved sheet', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    mutators.upsertIllustration(store, shape(0, 'a'));
    mutators.upsertIllustration(store, shape(1, 'b'));
    mutators.upsertIllustration(store, shape(2, 'c'));
    moveSheet(store, wb, 0, 2);
    expect(sheetsOf(store)).toEqual({ a: 2, b: 0, c: 1 });
  });

  it('drops the shapes of a removed sheet and renumbers the later ones', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    mutators.upsertIllustration(store, shape(0, 'a'));
    mutators.upsertIllustration(store, shape(1, 'b'));
    mutators.upsertIllustration(store, shape(2, 'c'));
    expect(removeSheet(store, wb, 1)).toBe(true);
    expect(sheetsOf(store)).toEqual({ a: 0, c: 1 });
  });

  it('moves sparkline hosts, watches, traces and ignored-error marks with their sheet', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb();
    setSparkline(store, { sheet: 0, row: 1, col: 1 }, { kind: 'line', source: 'A1:A5' });
    setSparkline(store, { sheet: 1, row: 2, col: 2 }, { kind: 'line', source: 'B1:B5' });
    mutators.addWatch(store, { sheet: 0, row: 0, col: 0 });
    store.setState((s) => ({
      ...s,
      traces: {
        items: [
          {
            kind: 'precedent',
            from: { sheet: 1, row: 0, col: 0 },
            to: { sheet: 1, row: 1, col: 1 },
          },
        ],
      },
      errorIndicators: {
        ignoredErrors: new Set(['0:3:3']),
        validationCircles: new Set(['1:4:4']),
      },
    }));
    moveSheet(store, wb, 0, 1);
    const s = store.getState();
    expect([...s.sparkline.sparklines.keys()].sort()).toEqual(['0:2:2', '1:1:1']);
    expect(s.watch.watches).toEqual([{ sheet: 1, row: 0, col: 0 }]);
    expect(s.traces.items[0]?.from.sheet).toBe(0);
    expect([...s.errorIndicators.ignoredErrors]).toEqual(['1:3:3']);
    expect([...s.errorIndicators.validationCircles]).toEqual(['0:4:4']);

    removeSheet(store, wb, 0);
    const after = store.getState();
    expect([...after.sparkline.sparklines.keys()]).toEqual(['0:1:1']);
    expect(after.traces.items).toEqual([]);
    expect([...after.errorIndicators.validationCircles]).toEqual([]);
  });

  it('removes slicers whose table went with the removed sheet', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb([
      { name: 'T0', displayName: 'T0', ref: 'A1:B3', sheetIndex: 0, columns: ['X', 'Y'] },
      { name: 'T1', displayName: 'T1', ref: 'A1:B3', sheetIndex: 1, columns: ['X', 'Y'] },
    ]);
    mutators.addSlicer(store, { id: 's0', tableName: 'T0', column: 'X', selected: [] });
    mutators.addSlicer(store, { id: 's1', tableName: 'T1', column: 'Y', selected: ['a'] });
    removeSheet(store, wb, 1);
    expect(store.getState().slicers.slicers.map((s) => s.id)).toEqual(['s0']);
  });

  it('keeps slicers across a move and a rename', () => {
    const store = createSpreadsheetStore();
    const { wb } = makeWb([
      { name: 'T1', displayName: 'T1', ref: 'A1:B3', sheetIndex: 1, columns: ['X', 'Y'] },
    ]);
    mutators.addSlicer(store, { id: 's1', tableName: 'T1', column: 'Y', selected: [] });
    moveSheet(store, wb, 1, 2);
    renameSheet(wb, 2, 'Renamed', store);
    expect(store.getState().slicers.slicers).toHaveLength(1);
  });
});

describe('sparklines across renames and row/column edits', () => {
  it('rewrites qualified sources when their sheet is renamed, and undoes it', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const { wb } = makeWb();
    setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'line', source: 'Beta!B2:B9' });
    setSparkline(store, { sheet: 0, row: 1, col: 0 }, { kind: 'line', source: 'B2:B9' });
    setSparkline(store, { sheet: 0, row: 2, col: 0 }, { kind: 'line', source: "'Beta'!C2:C9" });

    expect(renameSheet(wb, 1, 'Q3 Sales', store, history)).toBe(true);
    const sources = () =>
      [...store.getState().sparkline.sparklines.values()].map((spark) => spark.source);
    expect(sources()).toEqual(["'Q3 Sales'!B2:B9", 'B2:B9', "'Q3 Sales'!C2:C9"]);

    history.undo();
    expect(sources()[0]).toBe('Beta!B2:B9');
    history.redo();
    expect(sources()[0]).toBe("'Q3 Sales'!B2:B9");
  });

  it('moves hosts and re-points sources for a row insert and delete', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    setSparkline(store, { sheet: 0, row: 4, col: 5 }, { kind: 'column', source: 'A2:C2' });
    setSparkline(store, { sheet: 0, row: 8, col: 5 }, { kind: 'line', source: 'A6:A9' });

    expect(insertRows(store, wb, null, 3, 2)).toBe(true);
    let s = store.getState().sparkline.sparklines;
    expect(s.get('0:6:5')?.source).toBe('A2:C2');
    expect(s.get('0:10:5')?.source).toBe('A8:A11');

    expect(deleteRows(store, wb, null, 1, 1)).toBe(true);
    s = store.getState().sparkline.sparklines;
    expect(s.get('0:5:5')?.source).toBe('#REF!');
    expect(s.get('0:9:5')?.source).toBe('A7:A10');
  });

  it('follows column inserts, drops a deleted host, and undoes', async () => {
    const store = createSpreadsheetStore();
    const wb = await newWb();
    const history = new History();
    setSparkline(store, { sheet: 0, row: 0, col: 4 }, { kind: 'line', source: 'B2:D2' });
    setSparkline(store, { sheet: 0, row: 1, col: 2 }, { kind: 'line', source: 'Sheet1!B2:B9' });

    expect(insertCols(store, wb, history, 2, 1)).toBe(true);
    const s = store.getState().sparkline.sparklines;
    expect(s.get('0:0:5')?.source).toBe('B2:E2');
    expect(s.get('0:1:3')?.source).toBe('Sheet1!B2:B9');

    history.undo();
    expect([...store.getState().sparkline.sparklines.keys()].sort()).toEqual(['0:0:4', '0:1:2']);
    expect(store.getState().sparkline.sparklines.get('0:0:4')?.source).toBe('B2:D2');
  });
});
