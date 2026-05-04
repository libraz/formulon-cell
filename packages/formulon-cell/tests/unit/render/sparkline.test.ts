import { describe, expect, it, vi } from 'vitest';
import {
  applySparklineSnapshot,
  captureSparklineSnapshot,
  History,
  recordSparklineChange,
} from '../../../src/commands/history.js';
import { resolveNumericRangeFromCells } from '../../../src/engine/range-resolver.js';
import { paintSparkline } from '../../../src/render/painters.js';
import {
  createSpreadsheetStore,
  mutators,
  type Sparkline,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const seedNumber = (
  store: SpreadsheetStore,
  sheet: number,
  row: number,
  col: number,
  value: number,
): void => {
  store.setState((s) => {
    const cells = new Map(s.data.cells);
    cells.set(`${sheet}:${row}:${col}`, { value: { kind: 'number', value }, formula: null });
    return { ...s, data: { ...s.data, cells } };
  });
};

describe('sparkline mutators', () => {
  it('default state has an empty sparkline map', () => {
    const store = createSpreadsheetStore();
    expect(store.getState().sparkline.sparklines.size).toBe(0);
  });

  it('setSparkline writes a spec keyed by addr', () => {
    const store = createSpreadsheetStore();
    const spec: Sparkline = { kind: 'line', source: 'B2:B6' };
    mutators.setSparkline(store, { sheet: 0, row: 1, col: 4 }, spec);
    const stored = store.getState().sparkline.sparklines.get('0:1:4');
    expect(stored).toEqual(spec);
    // Stored copy is detached from caller — mutating the input must not bleed.
    spec.kind = 'column';
    expect(store.getState().sparkline.sparklines.get('0:1:4')?.kind).toBe('line');
  });

  it('setSparkline with null removes the entry', () => {
    const store = createSpreadsheetStore();
    mutators.setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'column', source: 'A1:A5' });
    mutators.setSparkline(store, { sheet: 0, row: 0, col: 0 }, null);
    expect(store.getState().sparkline.sparklines.has('0:0:0')).toBe(false);
  });

  it('clearSparkline removes the entry without affecting siblings', () => {
    const store = createSpreadsheetStore();
    mutators.setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'line', source: 'A1:A5' });
    mutators.setSparkline(store, { sheet: 0, row: 1, col: 0 }, { kind: 'column', source: 'B1:B5' });
    mutators.clearSparkline(store, { sheet: 0, row: 0, col: 0 });
    expect(store.getState().sparkline.sparklines.has('0:0:0')).toBe(false);
    expect(store.getState().sparkline.sparklines.has('0:1:0')).toBe(true);
  });
});

describe('sparkline history snapshots', () => {
  it('captureSparklineSnapshot returns a detached copy', () => {
    const store = createSpreadsheetStore();
    mutators.setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'line', source: 'A1:A5' });
    const snap = captureSparklineSnapshot(store.getState());
    mutators.setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'column', source: 'A1:A5' });
    expect(snap.get('0:0:0')?.kind).toBe('line');
  });

  it('applySparklineSnapshot restores prior state', () => {
    const store = createSpreadsheetStore();
    mutators.setSparkline(store, { sheet: 0, row: 0, col: 0 }, { kind: 'line', source: 'A1:A5' });
    const before = captureSparklineSnapshot(store.getState());
    mutators.setSparkline(
      store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'win-loss', source: 'B1:B5' },
    );
    applySparklineSnapshot(store, before);
    expect(store.getState().sparkline.sparklines.get('0:0:0')).toEqual({
      kind: 'line',
      source: 'A1:A5',
    });
  });

  it('recordSparklineChange round-trips through History', () => {
    const store = createSpreadsheetStore();
    const h = new History();
    recordSparklineChange(h, store, () => {
      mutators.setSparkline(
        store,
        { sheet: 0, row: 0, col: 0 },
        { kind: 'column', source: 'A1:A5' },
      );
    });
    expect(store.getState().sparkline.sparklines.has('0:0:0')).toBe(true);
    h.undo();
    expect(store.getState().sparkline.sparklines.has('0:0:0')).toBe(false);
    h.redo();
    expect(store.getState().sparkline.sparklines.get('0:0:0')?.kind).toBe('column');
  });
});

describe('resolveNumericRangeFromCells', () => {
  it('extracts numeric values in source order', () => {
    const store = createSpreadsheetStore();
    seedNumber(store, 0, 0, 0, 1);
    seedNumber(store, 0, 1, 0, 2);
    seedNumber(store, 0, 2, 0, 3);
    const out = resolveNumericRangeFromCells(store.getState().data.cells, 'A1:A3', 0);
    expect(out).toEqual([1, 2, 3]);
  });

  it('skips non-numeric cells', () => {
    const store = createSpreadsheetStore();
    seedNumber(store, 0, 0, 0, 10);
    store.setState((s) => {
      const cells = new Map(s.data.cells);
      cells.set('0:1:0', { value: { kind: 'text', value: 'x' }, formula: null });
      cells.set('0:2:0', { value: { kind: 'blank' }, formula: null });
      return { ...s, data: { ...s.data, cells } };
    });
    seedNumber(store, 0, 3, 0, 5);
    const out = resolveNumericRangeFromCells(store.getState().data.cells, 'A1:A4', 0);
    expect(out).toEqual([10, 5]);
  });

  it('returns [] for an unparseable ref', () => {
    const store = createSpreadsheetStore();
    expect(resolveNumericRangeFromCells(store.getState().data.cells, 'not a ref', 0)).toEqual([]);
  });

  it('honors a sheet name when sheetByName resolves it', () => {
    const store = createSpreadsheetStore();
    seedNumber(store, 1, 0, 0, 99);
    const sheetByName = (n: string): number => (n === 'Sheet2' ? 1 : -1);
    expect(
      resolveNumericRangeFromCells(store.getState().data.cells, 'Sheet2!A1:A1', 0, sheetByName),
    ).toEqual([99]);
  });

  it('skips sheet-prefixed refs when sheetByName is missing', () => {
    const store = createSpreadsheetStore();
    seedNumber(store, 1, 0, 0, 1);
    expect(resolveNumericRangeFromCells(store.getState().data.cells, 'Sheet2!A1:A1', 0)).toEqual(
      [],
    );
  });
});

/* ---- Painter exercise. We stub a minimal CanvasRenderingContext2D and
 * record draw calls so we can assert each kind issues the expected primitives.
 * Visual fidelity isn't checked — only the primitive shape. */
function makeStubCtx() {
  const calls: string[] = [];
  const ctx = {
    save: vi.fn(() => calls.push('save')),
    restore: vi.fn(() => calls.push('restore')),
    beginPath: vi.fn(() => calls.push('beginPath')),
    rect: vi.fn(() => calls.push('rect')),
    clip: vi.fn(() => calls.push('clip')),
    moveTo: vi.fn(() => calls.push('moveTo')),
    lineTo: vi.fn(() => calls.push('lineTo')),
    stroke: vi.fn(() => calls.push('stroke')),
    fillRect: vi.fn(() => calls.push('fillRect')),
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 0,
    lineJoin: '' as CanvasLineJoin,
    globalAlpha: 1,
  };
  return { ctx, calls };
}

describe('paintSparkline', () => {
  const rect = { x: 0, y: 0, w: 100, h: 30 };

  it('no-ops on empty values', () => {
    const { ctx, calls } = makeStubCtx();
    paintSparkline(
      ctx as unknown as CanvasRenderingContext2D,
      rect,
      {
        kind: 'line',
        source: 'A1:A1',
      },
      [],
    );
    expect(calls).toEqual([]);
  });

  it('line draws a stroked polyline', () => {
    const { ctx, calls } = makeStubCtx();
    paintSparkline(
      ctx as unknown as CanvasRenderingContext2D,
      rect,
      {
        kind: 'line',
        source: 'A1:A4',
      },
      [1, 2, 3, 4],
    );
    expect(calls).toContain('moveTo');
    expect(calls).toContain('lineTo');
    expect(calls).toContain('stroke');
  });

  it('column emits one fillRect per value', () => {
    const { ctx, calls } = makeStubCtx();
    paintSparkline(
      ctx as unknown as CanvasRenderingContext2D,
      rect,
      {
        kind: 'column',
        source: 'A1:A3',
      },
      [1, 2, 3],
    );
    const fills = calls.filter((c) => c === 'fillRect');
    expect(fills.length).toBe(3);
  });

  it('win-loss skips zeros, draws upper bars for positives, lower for negatives', () => {
    const { ctx, calls } = makeStubCtx();
    paintSparkline(
      ctx as unknown as CanvasRenderingContext2D,
      rect,
      {
        kind: 'win-loss',
        source: 'A1:A4',
      },
      [1, -1, 0, 2],
    );
    // 3 non-zero values → 3 fillRect calls.
    const fills = calls.filter((c) => c === 'fillRect');
    expect(fills.length).toBe(3);
  });

  it('uses negativeColor when showNegative is set', () => {
    const { ctx } = makeStubCtx();
    const seen: string[] = [];
    Object.defineProperty(ctx, 'fillStyle', {
      get: () => '',
      set: (v: string) => {
        seen.push(v);
      },
    });
    paintSparkline(
      ctx as unknown as CanvasRenderingContext2D,
      rect,
      {
        kind: 'win-loss',
        source: 'A1:A2',
        color: '#0000ff',
        showNegative: true,
        negativeColor: '#ff00aa',
      },
      [1, -1],
    );
    expect(seen).toContain('#0000ff');
    expect(seen).toContain('#ff00aa');
  });
});
