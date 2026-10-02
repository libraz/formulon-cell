import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import type { Addr, FormulonModule, Value, Workbook } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const good = { ok: true, status: 0, message: '', context: '' } as const;
const failed = (message: string) => ({ ok: false, status: 1, message, context: '' }) as const;

interface PhysicalCell {
  row: number;
  col: number;
  formula: string | null;
  value: Value;
}

const blank = (): Value => ({ kind: 0, number: 0, boolean: 0, text: '', errorCode: 0 });
const numberValue = (value: number): Value => ({
  kind: 1,
  number: value,
  boolean: 0,
  text: '',
  errorCode: 0,
});

const makeHandle = (raw: object): WorkbookHandle => {
  const Ctor = WorkbookHandle as unknown as new (
    module: FormulonModule,
    wb: Workbook,
  ) => WorkbookHandle;
  return new Ctor({ versionString: () => 'test' } as unknown as FormulonModule, raw as Workbook);
};

const addr = (sheet: number, row: number, col: number): Addr => ({ sheet, row, col });

const makeRaw = (
  sheets: PhysicalCell[][],
  options: {
    failCellAtIndex?: number;
    failCellAtCall?: number;
    failCellCountCall?: number;
    failGetValueCall?: number;
    canonicalizeFormula?: (formula: string) => string;
    formulaValues?: Record<string, Value>;
    manual?: boolean;
  } = {},
): {
  raw: object;
  cellCountCalls: () => number;
  cellAtCalls: () => number;
  getValueCalls: () => number;
  recalcCalls: () => number;
  setCalls: () => number;
} => {
  let cellCountCalls = 0;
  let cellAtCalls = 0;
  let getValueCalls = 0;
  let recalcCalls = 0;
  let setCalls = 0;
  const raw = {
    getValue: (sheet: number, row: number, col: number) => {
      getValueCalls += 1;
      if (options.failGetValueCall === getValueCalls) {
        return { status: failed('cell value read failed'), value: blank() };
      }
      const entry = sheets[sheet]?.find((cell) => cell.row === row && cell.col === col);
      return { status: good, value: entry?.value ?? blank() };
    },
    cellCount: (sheet: number) => {
      cellCountCalls += 1;
      if (options.failCellCountCall === cellCountCalls) {
        return { status: failed('post-write cell count failed'), value: 0 };
      }
      return { status: good, value: sheets[sheet]?.length ?? 0 };
    },
    cellAt: (sheet: number, index: number) => {
      cellAtCalls += 1;
      if (options.failCellAtCall === cellAtCalls || options.failCellAtIndex === index) {
        return {
          status: failed('cell entry failed'),
          row: 0,
          col: 0,
          formula: null,
          value: blank(),
        };
      }
      const entry = sheets[sheet]?.[index];
      return entry
        ? {
            status: good,
            row: entry.row,
            col: entry.col,
            formula: entry.formula,
            value: entry.value,
          }
        : { status: failed('cell index'), row: 0, col: 0, formula: null, value: blank() };
    },
    setNumber: (sheet: number, row: number, col: number, value: number) => {
      setCalls += 1;
      const entry = sheets[sheet]?.find((cell) => cell.row === row && cell.col === col);
      if (!entry) return failed('missing cell');
      entry.value = numberValue(value);
      entry.formula = null;
      return good;
    },
    setFormula: (sheet: number, row: number, col: number, formula: string) => {
      setCalls += 1;
      const entry = sheets[sheet]?.find((cell) => cell.row === row && cell.col === col);
      if (!entry) return failed('missing cell');
      const canonical = options.canonicalizeFormula?.(formula) ?? formula;
      entry.formula = canonical;
      entry.value = options.formulaValues?.[canonical] ?? numberValue(8);
      return good;
    },
    recalc: () => {
      recalcCalls += 1;
      return good;
    },
    calcMode: () => ({ status: good, value: options.manual ? 1 : 0 }),
    setCalcMode: () => good,
    cells: () => {
      throw new Error('pivot overlay enumeration is not allowed');
    },
  };
  return {
    raw,
    cellCountCalls: () => cellCountCalls,
    cellAtCalls: () => cellAtCalls,
    getValueCalls: () => getValueCalls,
    recalcCalls: () => recalcCalls,
    setCalls: () => setCalls,
  };
};

describe('WorkbookHandle atomic cell patches', () => {
  it('fails before writing when the value snapshot read fails', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: '=OLD', value: numberValue(7) }]], {
      failGetValueCall: 1,
      formulaValues: { '=OLD': numberValue(7) },
    });
    const wb = makeHandle(fixture.raw);
    const history = new History();
    const historyNotifications: number[] = [];
    history.subscribe(() => historyNotifications.push(historyNotifications.length + 1));
    wb.attachHistory(history);
    const events: unknown[] = [];
    wb.subscribe((event) => events.push(event));
    const internals = wb as unknown as {
      pendingRecalc: boolean;
      dirtySinceRecalc: Set<string>;
    };
    internals.pendingRecalc = true;
    internals.dirtySinceRecalc = new Set(['0:9:9']);

    expect(() =>
      wb.applyCellPatchAtomic([
        { addr: addr(0, 0, 0), value: { kind: 'number', value: 9 }, formula: null },
      ]),
    ).toThrow('cell value read failed');
    expect(fixture.setCalls()).toBe(0);
    expect(fixture.recalcCalls()).toBe(0);
    expect(wb.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula(addr(0, 0, 0))).toBe('=OLD');
    expect(events).toEqual([]);
    expect(history.canUndo()).toBe(false);
    expect(history.canRedo()).toBe(false);
    expect(historyNotifications).toEqual([]);
    expect(internals.pendingRecalc).toBe(true);
    expect([...internals.dirtySinceRecalc]).toEqual(['0:9:9']);
  });

  it('fails before writing when the formula snapshot entry fails', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: '=OLD', value: numberValue(7) }]], {
      failCellAtCall: 1,
      formulaValues: { '=OLD': numberValue(7) },
    });
    const wb = makeHandle(fixture.raw);
    const history = new History();
    wb.attachHistory(history);
    const events: unknown[] = [];
    wb.subscribe((event) => events.push(event));

    expect(() =>
      wb.applyCellPatchAtomic([
        { addr: addr(0, 0, 0), value: { kind: 'number', value: 9 }, formula: null },
      ]),
    ).toThrow('cell entry failed');
    expect(fixture.setCalls()).toBe(0);
    expect(fixture.recalcCalls()).toBe(0);
    expect(wb.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula(addr(0, 0, 0))).toBe('=OLD');
    expect(events).toEqual([]);
    expect(history.canUndo()).toBe(false);
    expect(history.canRedo()).toBe(false);
  });

  it('keeps public failed reads best-effort outside atomic snapshots', () => {
    const valueFixture = makeRaw([[{ row: 0, col: 0, formula: null, value: numberValue(7) }]], {
      failGetValueCall: 1,
    });
    const valueWb = makeHandle(valueFixture.raw);
    expect(valueWb.getValue(addr(0, 0, 0))).toEqual({ kind: 'blank' });

    const formulaFixture = makeRaw([[{ row: 0, col: 0, formula: '=A1', value: numberValue(1) }]], {
      failCellAtCall: 1,
    });
    const formulaWb = makeHandle(formulaFixture.raw);
    expect([...formulaWb.cellFormulas([addr(0, 0, 0)])]).toEqual([['0:0:0', null]]);
  });

  it('rolls back the original formula and value after a post-write value read failure', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: '=OLD', value: numberValue(7) }]], {
      failGetValueCall: 2,
      formulaValues: { '=OLD': numberValue(7), '=NEW': numberValue(8) },
    });
    const wb = makeHandle(fixture.raw);
    const history = new History();
    wb.attachHistory(history);
    const events: unknown[] = [];
    wb.subscribe((event) => events.push(event));
    const internals = wb as unknown as {
      pendingRecalc: boolean;
      dirtySinceRecalc: Set<string>;
    };
    internals.pendingRecalc = true;
    internals.dirtySinceRecalc = new Set(['0:9:9']);

    expect(() =>
      wb.applyCellPatchAtomic([
        { addr: addr(0, 0, 0), value: { kind: 'number', value: 8 }, formula: '=NEW' },
      ]),
    ).toThrow('cell value read failed');
    expect(fixture.setCalls()).toBe(2);
    expect(fixture.recalcCalls()).toBe(2);
    expect(wb.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula(addr(0, 0, 0))).toBe('=OLD');
    expect(events).toEqual([]);
    expect(history.canUndo()).toBe(false);
    expect(internals.pendingRecalc).toBe(true);
    expect([...internals.dirtySinceRecalc]).toEqual(['0:9:9']);

    const next = wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 9 }, formula: null },
    ]);
    const valueEvent = events.find(
      (event): event is { kind: 'value'; atomicBatch?: { id?: number } } =>
        typeof event === 'object' &&
        event !== null &&
        (event as { kind?: string }).kind === 'value',
    );
    expect(next.after[0]?.value).toEqual({ kind: 'number', value: 9 });
    expect(valueEvent?.atomicBatch?.id).toBe(1);
  });

  it('rolls back the original formula and value after a post-write formula read failure', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: '=OLD', value: numberValue(7) }]], {
      failCellAtCall: 2,
      formulaValues: { '=OLD': numberValue(7), '=NEW': numberValue(8) },
    });
    const wb = makeHandle(fixture.raw);
    const history = new History();
    wb.attachHistory(history);
    const events: unknown[] = [];
    wb.subscribe((event) => events.push(event));

    expect(() =>
      wb.applyCellPatchAtomic([
        { addr: addr(0, 0, 0), value: { kind: 'number', value: 8 }, formula: '=NEW' },
      ]),
    ).toThrow('cell entry failed');
    expect(fixture.setCalls()).toBe(2);
    expect(fixture.recalcCalls()).toBe(2);
    expect(wb.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 7 });
    expect(wb.cellFormula(addr(0, 0, 0))).toBe('=OLD');
    expect(events).toEqual([]);
    expect(history.canUndo()).toBe(false);
  });

  it('rolls back when authoritative post-write capture fails', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: null, value: numberValue(7) }]], {
      failCellCountCall: 2,
    });
    const wb = makeHandle(fixture.raw);
    const events: unknown[] = [];
    wb.subscribe((event) => events.push(event));

    expect(() =>
      wb.applyCellPatchAtomic([
        { addr: addr(0, 0, 0), value: { kind: 'number', value: 9 }, formula: null },
      ]),
    ).toThrow('post-write cell count failed');
    expect(wb.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 7 });
    expect(events).toEqual([]);
    expect(fixture.cellCountCalls()).toBe(2);
  });

  it('bulk-reads unique requested formulas with one physical pass per sheet', () => {
    const fixture = makeRaw([
      [
        { row: 0, col: 0, formula: '=SUM(B1:B2)', value: numberValue(3) },
        { row: 1, col: 0, formula: null, value: numberValue(1) },
      ],
      [{ row: 2, col: 2, formula: '=A1', value: numberValue(3) }],
    ]);
    const wb = makeHandle(fixture.raw);

    const formulas = wb.cellFormulas([
      addr(0, 0, 0),
      addr(0, 0, 0),
      addr(0, 9, 9),
      addr(1, 2, 2),
      addr(1, 9, 9),
    ]);

    expect([...formulas]).toEqual([
      ['0:0:0', '=SUM(B1:B2)'],
      ['0:9:9', null],
      ['1:2:2', '=A1'],
      ['1:9:9', null],
    ]);
    expect(fixture.cellCountCalls()).toBe(2);
    expect(fixture.cellAtCalls()).toBe(3);
  });

  it('does not enumerate an empty request and skips failed physical entries', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: '=A1', value: numberValue(1) }]], {
      failCellAtIndex: 0,
    });
    const wb = makeHandle(fixture.raw);
    expect([...wb.cellFormulas([])]).toEqual([]);
    expect(fixture.cellCountCalls()).toBe(0);
    expect([...wb.cellFormulas([addr(0, 0, 0)])]).toEqual([['0:0:0', null]]);
    expect(fixture.cellCountCalls()).toBe(1);
    expect(fixture.cellAtCalls()).toBe(1);
  });

  it('stops a physical scan once the requested formula is found', () => {
    const cells = Array.from({ length: 1000 }, (_, row) => ({
      row,
      col: 0,
      formula: row === 0 ? '=A1' : null,
      value: numberValue(row + 1),
    }));
    const fixture = makeRaw([cells]);
    const wb = makeHandle(fixture.raw);

    expect([...wb.cellFormulas([addr(0, 0, 0)])]).toEqual([['0:0:0', '=A1']]);
    expect(fixture.cellAtCalls()).toBe(1);
  });

  it('uses two bounded physical passes for an atomic batch across sheets', () => {
    const fixture = makeRaw([
      [
        { row: 0, col: 0, formula: '=A2', value: numberValue(1) },
        { row: 1, col: 0, formula: null, value: numberValue(2) },
      ],
      [{ row: 0, col: 0, formula: '=A1', value: numberValue(3) }],
    ]);
    const wb = makeHandle(fixture.raw);

    const result = wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 10 }, formula: null },
      { addr: addr(0, 1, 0), value: { kind: 'number', value: 2 }, formula: null },
      { addr: addr(1, 0, 0), value: { kind: 'number', value: 20 }, formula: null },
    ]);

    expect(result.changed).toHaveLength(2);
    expect(fixture.cellCountCalls()).toBe(4);
    expect(fixture.cellAtCalls()).toBe(6);
  });

  it('keeps a first-cell atomic edit to one physical read per snapshot pass', () => {
    const cells = Array.from({ length: 1000 }, (_, row) => ({
      row,
      col: 0,
      formula: row === 0 ? '=A1' : null,
      value: numberValue(row + 1),
    }));
    const fixture = makeRaw([cells]);
    const wb = makeHandle(fixture.raw);

    wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 99 }, formula: null },
    ]);

    expect(fixture.cellAtCalls()).toBe(2);
  });

  it('keeps duplicate last-patch and no-op result semantics', () => {
    const duplicateFixture = makeRaw([[{ row: 0, col: 0, formula: null, value: numberValue(7) }]]);
    const duplicateWb = makeHandle(duplicateFixture.raw);
    const duplicate = duplicateWb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 9 }, formula: null },
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 11 }, formula: null },
    ]);
    expect(duplicate.before).toHaveLength(1);
    expect(duplicate.after).toHaveLength(1);
    expect(duplicate.changed).toEqual([addr(0, 0, 0)]);
    expect(duplicateWb.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 11 });

    const noOpFixture = makeRaw([[{ row: 0, col: 0, formula: null, value: numberValue(7) }]]);
    const noOpWb = makeHandle(noOpFixture.raw);
    const noOp = noOpWb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 7 }, formula: null },
    ]);
    expect(noOp.changed).toEqual([]);
    expect(noOp.after).toBe(noOp.before);
    expect(noOpFixture.cellCountCalls()).toBe(1);
  });

  it('returns the engine formula after a write that canonicalizes text', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: '=OLD', value: numberValue(7) }]], {
      canonicalizeFormula: () => '=CANONICAL',
    });
    const wb = makeHandle(fixture.raw);
    const result = wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 8 }, formula: '=input' },
    ]);

    expect(result.before[0]?.formula).toBe('=OLD');
    expect(result.after[0]?.formula).toBe('=CANONICAL');
  });

  it('tags changed value events with authoritative batch metadata', () => {
    const fixture = makeRaw(
      [
        [
          { row: 0, col: 0, formula: '=OLD', value: numberValue(7) },
          { row: 0, col: 1, formula: null, value: numberValue(2) },
        ],
      ],
      { canonicalizeFormula: () => '=CANONICAL' },
    );
    const wb = makeHandle(fixture.raw);
    const events: unknown[] = [];
    wb.subscribe((event) => events.push(event));

    const first = wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 8 }, formula: '=input' },
      { addr: addr(0, 0, 1), value: { kind: 'number', value: 3 }, formula: null },
    ]);
    const firstValues = events.filter(
      (event): event is { kind: 'value'; atomicBatch?: unknown } =>
        typeof event === 'object' &&
        event !== null &&
        (event as { kind?: string }).kind === 'value',
    );
    const firstMeta = firstValues.map(
      (event) =>
        (event.atomicBatch ?? null) as {
          id: number;
          index: number;
          size: number;
          formula: string | null;
        } | null,
    );
    expect(firstMeta).toEqual([
      { id: 1, index: 0, size: 2, formula: first.after[0]?.formula ?? null },
      { id: 1, index: 1, size: 2, formula: first.after[1]?.formula ?? null },
    ]);
    expect(JSON.parse(JSON.stringify(firstMeta[0]))).toEqual(firstMeta[0]);
    expect(structuredClone(firstMeta[0])).toEqual(firstMeta[0]);

    events.length = 0;
    wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 9 }, formula: null },
    ]);
    const second = events.find(
      (event): event is { kind: 'value'; atomicBatch?: { id?: number } } =>
        typeof event === 'object' &&
        event !== null &&
        (event as { kind?: string }).kind === 'value',
    );
    expect(second?.atomicBatch?.id).toBeGreaterThan(firstMeta[0]?.id ?? 0);
  });

  it('keeps manual batches authoritative without recalc or recalc events', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: '=OLD', value: numberValue(7) }]], {
      canonicalizeFormula: () => '=CANONICAL',
      manual: true,
    });
    const wb = makeHandle(fixture.raw);
    const events: Array<{ kind: string; next?: unknown }> = [];
    wb.subscribe((event) =>
      events.push(
        event.kind === 'value' ? { kind: event.kind, next: event.next } : { kind: event.kind },
      ),
    );

    const result = wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 8 }, formula: '=input' },
    ]);

    expect(fixture.recalcCalls()).toBe(0);
    expect(result.after[0]?.formula).toBe('=CANONICAL');
    expect(result.after[0]?.value).toEqual({ kind: 'number', value: 8 });
    expect(events).toEqual([{ kind: 'value', next: { kind: 'number', value: 8 } }]);
  });

  it('keeps listener value mutations separate from authoritative snapshots', () => {
    const fixture = makeRaw([[{ row: 0, col: 0, formula: null, value: numberValue(7) }]]);
    const wb = makeHandle(fixture.raw);
    wb.subscribe((event) => {
      if (event.kind === 'value' && event.next.kind === 'number') {
        Object.assign(event.next, { value: 99 });
      }
    });

    const result = wb.applyCellPatchAtomic([
      { addr: addr(0, 0, 0), value: { kind: 'number', value: 9 }, formula: null },
    ]);

    expect(result.after[0]?.value).toEqual({ kind: 'number', value: 9 });
    expect(wb.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 9 });
  });

  it('saves and reloads a successful atomic formula and literal batch through real WASM', async () => {
    const wb = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | undefined;
    try {
      expect(wb.isStub).toBe(false);
      const result = wb.applyCellPatchAtomic([
        { addr: addr(0, 0, 0), value: { kind: 'number', value: 2 }, formula: '=1+1' },
        { addr: addr(0, 0, 1), value: { kind: 'number', value: 42 }, formula: null },
      ]);
      expect(result.before[0]?.formula).toBeNull();
      expect(result.after[0]?.formula).toBe('=1+1');
      expect(result.after[0]?.value).toEqual({ kind: 'number', value: 2 });
      expect(result.after[1]?.formula).toBeNull();
      expect(result.after[1]?.value).toEqual({ kind: 'number', value: 42 });

      loaded = await WorkbookHandle.loadBytes(wb.save());
      expect(loaded.isStub).toBe(false);
      expect(loaded.cellFormula(addr(0, 0, 0))).toBe('=1+1');
      expect(loaded.getValue(addr(0, 0, 0))).toEqual({ kind: 'number', value: 2 });
      expect(loaded.cellFormula(addr(0, 0, 1))).toBeNull();
      expect(loaded.getValue(addr(0, 0, 1))).toEqual({ kind: 'number', value: 42 });
    } finally {
      loaded?.dispose();
      wb.dispose();
    }
  });
});
