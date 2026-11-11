import { BlanksAs, SparklineAxisType, SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  applySparklineSnapshot,
  captureSparklineSnapshot,
  recordSparklineChange,
} from '../../../src/commands/slice-history.js';
import {
  defaultSparklineGroupSettings,
  groupOfCell,
  groupSparklines,
  membersOfGroup,
  rgbColorSpec,
  setSparkline,
  setSparklines,
  ungroupSparklines,
  updateSparklineGroup,
} from '../../../src/commands/sparkline.js';
import { resolveNumericRangeFromCells } from '../../../src/engine/range-resolver.js';
import type { SparklinePoint } from '../../../src/render/painters/controls.js';
import { paintSparkline } from '../../../src/render/painters.js';
import {
  createSparklineGroupExtents,
  paintCellSparkline,
  readSparklinePoints,
  sparklineColors,
} from '../../../src/render/sparkline.js';
import {
  createSpreadsheetStore,
  mutators,
  type SparklineGroupSettings,
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

const settingsOf = (
  type: SparklineType,
  patch: Partial<SparklineGroupSettings> = {},
): SparklineGroupSettings => ({ ...defaultSparklineGroupSettings(type), ...patch });
const LINE = settingsOf(SparklineType.Line);
const COLUMN = settingsOf(SparklineType.Column);
const A = { sheet: 0, row: 0, col: 0 };
const B = { sheet: 0, row: 1, col: 0 };
const C = { sheet: 0, row: 2, col: 0 };

describe('sparkline store helpers', () => {
  it('default state has empty sparkline and group maps', () => {
    const store = createSpreadsheetStore();
    expect(store.getState().sparkline.sparklines.size).toBe(0);
    expect(store.getState().sparkline.groups.size).toBe(0);
  });

  it('setSparkline creates a group entry holding detached settings', () => {
    const store = createSpreadsheetStore();
    const settings = settingsOf(SparklineType.Line, { colorSeries: rgbColorSpec('#102030') });
    setSparkline(store, { sheet: 0, row: 1, col: 4 }, 'B2:B6', settings);
    const member = store.getState().sparkline.sparklines.get('0:1:4');
    expect(member?.source).toBe('B2:B6');
    const stored = store.getState().sparkline.groups.get(member?.group as string);
    expect(stored).toEqual(settings);
    expect(stored).not.toBe(settings);
    expect(stored?.colorSeries).not.toBe(settings.colorSeries);
  });

  it('clearing the last member of a group drops the group', () => {
    const store = createSpreadsheetStore();
    setSparklines(
      store,
      [
        { addr: A, source: 'B1:B5' },
        { addr: B, source: 'C1:C5' },
      ],
      LINE,
    );
    mutators.clearSparkline(store, A);
    expect(store.getState().sparkline.groups.size).toBe(1);
    mutators.clearSparkline(store, B);
    expect(store.getState().sparkline.groups.size).toBe(0);
  });

  it('setSparkline with null removes the entry and its emptied group', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, A, 'A1:A5', COLUMN);
    mutators.setSparkline(store, A, null);
    expect(store.getState().sparkline.sparklines.has('0:0:0')).toBe(false);
    expect(store.getState().sparkline.groups.size).toBe(0);
  });

  it('updateSparklineGroup patches the settings of every member', () => {
    const store = createSpreadsheetStore();
    setSparklines(
      store,
      [
        { addr: A, source: 'B1:B5' },
        { addr: B, source: 'C1:C5' },
      ],
      LINE,
    );
    const id = groupOfCell(store.getState(), A)?.id as string;
    expect(updateSparklineGroup(store, id, { markers: true, manualMin: 3 })).toBe(true);
    expect(groupOfCell(store.getState(), B)?.settings).toMatchObject({
      markers: true,
      manualMin: 3,
      type: SparklineType.Line,
    });
    expect(updateSparklineGroup(store, 'nope', { markers: true })).toBe(false);
  });

  it('groupSparklines merges the groups of the cells under the active cell settings', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, A, 'B1:B5', LINE);
    setSparkline(store, B, 'C1:C5', settingsOf(SparklineType.Column, { markers: true }));
    setSparkline(store, C, 'D1:D5', LINE);

    const id = groupSparklines(store, [A, B], B);
    const state = store.getState();
    expect(id).not.toBeNull();
    expect(membersOfGroup(state, id as string).map((e) => e.addr)).toEqual([A, B]);
    expect(groupOfCell(state, A)?.settings).toMatchObject({
      type: SparklineType.Column,
      markers: true,
    });
    expect(groupOfCell(state, C)?.id).not.toBe(id);
    expect(state.sparkline.groups.size).toBe(2);
    expect(groupSparklines(store, [{ sheet: 0, row: 9, col: 9 }])).toBeNull();
  });

  it('ungroupSparklines gives each member its own copy of the settings', () => {
    const store = createSpreadsheetStore();
    setSparklines(
      store,
      [
        { addr: A, source: 'B1:B5' },
        { addr: B, source: 'C1:C5' },
      ],
      settingsOf(SparklineType.Line, { high: true }),
    );
    expect(ungroupSparklines(store, [A])).toBe(2);
    const a = groupOfCell(store.getState(), A);
    const b = groupOfCell(store.getState(), B);
    expect(a?.id).not.toBe(b?.id);
    expect(a?.settings).toEqual(b?.settings);
    expect(a?.settings.high).toBe(true);
    expect(store.getState().sparkline.groups.size).toBe(2);
    // Editing one no longer touches the other.
    updateSparklineGroup(store, a?.id as string, { high: false });
    expect(groupOfCell(store.getState(), B)?.settings.high).toBe(true);
  });
});

describe('sparkline history snapshots', () => {
  it('captureSparklineSnapshot returns a detached copy of members and groups', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, A, 'A1:A5', LINE);
    const snap = captureSparklineSnapshot(store.getState());
    const id = groupOfCell(store.getState(), A)?.id as string;
    updateSparklineGroup(store, id, { markers: true });
    setSparkline(store, A, 'B1:B5', COLUMN);
    expect(snap.sparklines.get('0:0:0')?.source).toBe('A1:A5');
    expect(snap.groups.get(id)?.markers).toBe(false);
  });

  it('applySparklineSnapshot restores prior members and groups', () => {
    const store = createSpreadsheetStore();
    setSparkline(store, A, 'A1:A5', LINE);
    const before = captureSparklineSnapshot(store.getState());
    setSparkline(store, A, 'B1:B5', settingsOf(SparklineType.Stacked));
    applySparklineSnapshot(store, before);
    expect(store.getState().sparkline.sparklines.get('0:0:0')?.source).toBe('A1:A5');
    expect(groupOfCell(store.getState(), A)?.settings.type).toBe(SparklineType.Line);
  });

  it('recordSparklineChange round-trips settings changes through History', () => {
    const store = createSpreadsheetStore();
    const h = new History();
    setSparkline(store, A, 'A1:A5', LINE);
    const id = groupOfCell(store.getState(), A)?.id as string;
    recordSparklineChange(h, store, () => {
      updateSparklineGroup(store, id, { lineWidthPt: 3 });
    });
    h.undo();
    expect(groupOfCell(store.getState(), A)?.settings.lineWidthPt).toBe(0.75);
    h.redo();
    expect(groupOfCell(store.getState(), A)?.settings.lineWidthPt).toBe(3);
  });

  it('recordSparklineChange skips unchanged snapshots', () => {
    const store = createSpreadsheetStore();
    const h = new History();
    setSparkline(store, A, 'A1:A5', COLUMN);
    const id = groupOfCell(store.getState(), A)?.id as string;
    recordSparklineChange(h, store, () => {
      updateSparklineGroup(store, id, { markers: false });
    });
    expect(h.canUndo()).toBe(false);
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

/* ---- Painter exercise. A minimal CanvasRenderingContext2D records the draw
 * primitives with the style each was issued under. */
interface Rec {
  fills: { x: number; y: number; w: number; h: number; style: string }[];
  moves: { x: number; y: number }[];
  lines: { x: number; y: number }[];
  strokes: { style: string; width: number }[];
}

function makeStubCtx() {
  const rec: Rec = { fills: [], moves: [], lines: [], strokes: [] };
  const ctx = {
    save: () => {},
    restore: () => {},
    beginPath: () => {},
    rect: () => {},
    clip: () => {},
    moveTo: (x: number, y: number) => rec.moves.push({ x, y }),
    lineTo: (x: number, y: number) => rec.lines.push({ x, y }),
    stroke: () => rec.strokes.push({ style: ctx.strokeStyle, width: ctx.lineWidth }),
    fillRect: (x: number, y: number, w: number, h: number) =>
      rec.fills.push({ x, y, w, h, style: ctx.fillStyle }),
    fillStyle: '',
    strokeStyle: '',
    lineWidth: 0,
    lineJoin: '' as CanvasLineJoin,
    globalAlpha: 1,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, rec };
}

const pts = (...values: (number | null)[]): SparklinePoint[] =>
  values.map((value, x) => ({ x, value }));

const rect = { x: 0, y: 0, w: 100, h: 30 };
// The painter insets the rect by 2px: the plot spans y 2..28.
const TOP = 2;
const BOTTOM = 28;

const paint = (
  settings: SparklineGroupSettings,
  points: SparklinePoint[],
  extra: { groupExtent?: { min: number; max: number } } = {},
) => {
  const { ctx, rec } = makeStubCtx();
  paintSparkline(ctx, rect, {
    settings,
    colors: sparklineColors(settings, null),
    points,
    ...extra,
  });
  return rec;
};

describe('paintSparkline', () => {
  it('no-ops on an empty series', () => {
    const rec = paint(LINE, []);
    expect(rec.moves.length + rec.lines.length + rec.fills.length + rec.strokes.length).toBe(0);
  });

  it('line draws a polyline in the series colour at the group weight', () => {
    const rec = paint(settingsOf(SparklineType.Line, { lineWidthPt: 1.5 }), pts(1, 2, 3, 4));
    expect(rec.moves).toHaveLength(1);
    expect(rec.lines).toHaveLength(3);
    expect(rec.strokes[0]).toEqual({ style: '#0078d4', width: 2 });
  });

  it('column emits one bar per value and win-loss skips zeros with equal heights', () => {
    expect(paint(COLUMN, pts(1, 2, 3)).fills).toHaveLength(3);
    const rec = paint(settingsOf(SparklineType.Stacked), pts(1, -1, 0, 2));
    expect(rec.fills).toHaveLength(3);
    expect(new Set(rec.fills.map((f) => f.h)).size).toBe(1);
    const [up, down] = rec.fills;
    expect((up?.y ?? 0) + (up?.h ?? 0)).toBeLessThanOrEqual(down?.y ?? 0);
  });

  it('colours high, low, first, last and negative points', () => {
    const settings = settingsOf(SparklineType.Column, {
      colorSeries: rgbColorSpec('#000001'),
      colorHigh: rgbColorSpec('#00ff00'),
      colorLow: rgbColorSpec('#ff0000'),
      colorFirst: rgbColorSpec('#0000ff'),
      colorLast: rgbColorSpec('#ffff00'),
      colorNegative: rgbColorSpec('#ff00ff'),
      high: true,
      low: true,
      first: true,
      last: true,
      negative: true,
    });
    // first 2, high 9, plain 4, negative -1, low -7, last 3.
    const styles = paint(settings, pts(2, 9, 4, -1, -7, 3)).fills.map((f) => f.style);
    expect(styles).toEqual(['#0000ff', '#00ff00', '#000001', '#ff00ff', '#ff0000', '#ffff00']);
  });

  it('leaves the highlight colours unused while their flags are off', () => {
    const settings = settingsOf(SparklineType.Column, {
      colorHigh: rgbColorSpec('#00ff00'),
      colorNegative: rgbColorSpec('#ff00ff'),
      negative: false,
    });
    expect(new Set(paint(settings, pts(1, 9, -3)).fills.map((f) => f.style))).toEqual(
      new Set(['#0078d4']),
    );
  });

  it('marks high and low points with dots on a line, and every point with markers', () => {
    const marked = settingsOf(SparklineType.Line, {
      colorHigh: rgbColorSpec('#00ff00'),
      colorLow: rgbColorSpec('#ff0000'),
      high: true,
      low: true,
    });
    expect(paint(marked, pts(3, 9, 1, 4)).fills.map((f) => f.style)).toEqual([
      '#00ff00',
      '#ff0000',
    ]);
    const all = settingsOf(SparklineType.Line, {
      markers: true,
      colorMarkers: rgbColorSpec('#123456'),
    });
    const fills = paint(all, pts(3, 9, 1, 4)).fills;
    expect(fills).toHaveLength(4);
    expect(new Set(fills.map((f) => f.style))).toEqual(new Set(['#123456']));
  });

  it('draws the horizontal axis only when zero is within the range', () => {
    const base = { displayXAxis: true, colorAxis: rgbColorSpec('#abcdef') };
    const axisStrokes = (settings: SparklineGroupSettings, ...values: number[]) =>
      paint(settings, pts(...values)).strokes.filter((s) => s.style === '#abcdef');
    expect(axisStrokes(settingsOf(SparklineType.Line, base), -2, 3, 1)).toHaveLength(1);
    expect(axisStrokes(settingsOf(SparklineType.Line, base), 2, 3, 1)).toHaveLength(0);
    expect(
      axisStrokes(settingsOf(SparklineType.Line, { ...base, displayXAxis: false }), -2, 3),
    ).toHaveLength(0);
    // A custom floor below the data brings zero back into range.
    const custom = { ...base, minAxisType: SparklineAxisType.Custom, manualMin: -5 };
    expect(axisStrokes(settingsOf(SparklineType.Line, custom), 2, 3, 1)).toHaveLength(1);
  });

  it('scales a group axis by the group extent', () => {
    const own = paint(LINE, pts(0, 10));
    expect(own.lines[0]?.y).toBe(TOP);
    const group = settingsOf(SparklineType.Line, {
      minAxisType: SparklineAxisType.Group,
      maxAxisType: SparklineAxisType.Group,
    });
    const shared = paint(group, pts(0, 10), { groupExtent: { min: 0, max: 100 } });
    expect(shared.moves[0]?.y).toBe(BOTTOM);
    expect(shared.lines[0]?.y).toBeCloseTo(BOTTOM - (10 / 100) * (BOTTOM - TOP), 5);
  });

  it('clamps points below a custom minimum and above a custom maximum', () => {
    const settings = settingsOf(SparklineType.Line, {
      minAxisType: SparklineAxisType.Custom,
      maxAxisType: SparklineAxisType.Custom,
      manualMin: 0,
      manualMax: 10,
    });
    const rec = paint(settings, pts(-50, 5, 90));
    expect(rec.moves[0]?.y).toBe(BOTTOM);
    expect(rec.lines[0]?.y).toBeCloseTo((TOP + BOTTOM) / 2, 5);
    expect(rec.lines[1]?.y).toBe(TOP);
  });

  it('breaks the line at an empty cell for gaps, bridges it for span, and zero-fills it', () => {
    const series = pts(1, null, 3, 4);
    const gap = paint(settingsOf(SparklineType.Line, { emptyCellsAs: BlanksAs.Gap }), series);
    expect(gap.moves).toHaveLength(2);
    expect(gap.lines).toHaveLength(1);
    const span = paint(settingsOf(SparklineType.Line, { emptyCellsAs: BlanksAs.Span }), series);
    expect(span.moves).toHaveLength(1);
    expect(span.lines).toHaveLength(2);
    const zero = paint(settingsOf(SparklineType.Line, { emptyCellsAs: BlanksAs.Zero }), series);
    expect(zero.moves).toHaveLength(1);
    expect(zero.lines).toHaveLength(3);
    // The zero-filled point sits at the bottom of the 0..4 range... but 1 is not the minimum.
    expect(zero.lines[0]?.y).toBe(BOTTOM);
  });

  it('spaces points by date values on a date axis', () => {
    const rec = paint(settingsOf(SparklineType.Line, { dateAxis: true }), [
      { x: 0, value: 1 },
      { x: 1, value: 2 },
      { x: 10, value: 3 },
    ]);
    expect(rec.moves[0]?.x).toBe(2);
    expect(rec.lines[0]?.x).toBeCloseTo(2 + (96 / 10) * 1, 5);
    expect(rec.lines[1]?.x).toBeCloseTo(98, 5);
  });
});

describe('sparkline series and cell painting', () => {
  /** Rows 0..3 of column B hold `values`; the hosts sit in column A. */
  const sheetWith = (
    values: (number | null)[],
    settings: SparklineGroupSettings,
  ): SpreadsheetStore => {
    const store = createSpreadsheetStore();
    values.forEach((v, row) => {
      if (v !== null) seedNumber(store, 0, row, 1, v);
    });
    setSparkline(store, A, `B1:B${values.length}`, settings);
    return store;
  };

  it('skips values in hidden rows and columns unless the group shows them', () => {
    const store = sheetWith([1, 2, 3, 4], LINE);
    store.setState((s) => ({ ...s, layout: { ...s.layout, hiddenRows: new Set([1]) } }));
    const state = store.getState();
    const member = state.sparkline.sparklines.get('0:0:0');
    if (!member) throw new Error('missing member');
    const read = (settings: SparklineGroupSettings) =>
      readSparklinePoints(state, null, member, 0, settings)?.map((p) => p.value);
    expect(read(LINE)).toEqual([1, 3, 4]);
    expect(read({ ...LINE, displayHidden: true })).toEqual([1, 2, 3, 4]);
    store.setState((s) => ({
      ...s,
      layout: { ...s.layout, hiddenRows: new Set(), hiddenCols: new Set([1]) },
    }));
    expect(readSparklinePoints(store.getState(), null, member, 0, LINE)).toEqual([]);
  });

  it('reads blank cells as empty points and dates from the date range', () => {
    const settings = settingsOf(SparklineType.Line, {
      dateAxis: true,
      dateAxisRange: {
        sheet: 0,
        rect: { firstRow: 0, firstCol: 2, lastRow: 2, lastCol: 2 },
      },
    });
    const store = sheetWith([5, null, 7], settings);
    [10, 20, 40].forEach((d, row) => {
      seedNumber(store, 0, row, 2, d);
    });
    const state = store.getState();
    const member = state.sparkline.sparklines.get('0:0:0');
    if (!member) throw new Error('missing member');
    expect(readSparklinePoints(state, null, member, 0, settings)).toEqual([
      { x: 10, value: 5 },
      { x: 20, value: null },
      { x: 40, value: 7 },
    ]);
  });

  it('shares min and max across every member of a group, and not across groups', () => {
    const store = createSpreadsheetStore();
    [0, 10].forEach((v, row) => {
      seedNumber(store, 0, row, 1, v);
    });
    [0, 100].forEach((v, row) => {
      seedNumber(store, 0, row, 2, v);
    });
    const group = settingsOf(SparklineType.Line, {
      minAxisType: SparklineAxisType.Group,
      maxAxisType: SparklineAxisType.Group,
    });
    setSparklines(
      store,
      [
        { addr: { sheet: 0, row: 0, col: 4 }, source: 'B1:B2' },
        { addr: { sheet: 0, row: 1, col: 4 }, source: 'C1:C2' },
      ],
      group,
    );
    setSparkline(store, { sheet: 0, row: 2, col: 4 }, 'D1:D2', group);
    seedNumber(store, 0, 0, 3, -1000);
    const state = store.getState();
    const extentOf = createSparklineGroupExtents(state, null);
    const id = groupOfCell(state, { sheet: 0, row: 0, col: 4 })?.id as string;
    expect(extentOf(id)).toEqual({ min: 0, max: 100 });
    const lone = groupOfCell(state, { sheet: 0, row: 2, col: 4 })?.id as string;
    expect(extentOf(lone)).toEqual({ min: -1000, max: -1000 });

    const { ctx, rec } = makeStubCtx();
    const spec = state.sparkline.sparklines.get('0:0:4');
    if (!spec) throw new Error('missing member');
    paintCellSparkline(ctx, rect, spec, state, null, extentOf);
    // The 0..10 member only reaches a tenth of the shared 0..100 axis.
    expect(rec.lines[0]?.y).toBeCloseTo(BOTTOM - (10 / 100) * (BOTTOM - TOP), 5);
  });

  it('paints nothing for a member whose group is gone or whose source does not resolve', () => {
    const store = sheetWith([1, 2], LINE);
    const state = store.getState();
    const { ctx, rec } = makeStubCtx();
    paintCellSparkline(ctx, rect, { source: 'B1:B2', group: 'missing' }, state, null);
    paintCellSparkline(ctx, rect, { source: 'Other!B1:B2', group: 'x' }, state, null);
    const id = groupOfCell(state, A)?.id as string;
    paintCellSparkline(ctx, rect, { source: 'Other!B1:B2', group: id }, state, null);
    expect(rec.strokes).toHaveLength(0);
  });
});
