import { SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { defaultSparklineGroupSettings } from '../../../src/commands/sparkline.js';
import type { SparklinePoint } from '../../../src/render/painters/controls.js';
import { paintSparkline } from '../../../src/render/painters.js';
import { sparklineColors } from '../../../src/render/sparkline.js';
import type { SparklineGroupSettings } from '../../../src/store/store.js';

interface Rec {
  fills: { x: number; y: number; w: number; h: number }[];
  lines: { x: number; y: number }[];
}

/** The plot area inside the 100x30 cell, after the painter's 2px inset. */
const TOP = 2;
const HEIGHT = 26;

const paint = (settings: SparklineGroupSettings, points: SparklinePoint[]): Rec => {
  const rec: Rec = { fills: [], lines: [] };
  const ctx = {
    save: () => {},
    restore: () => {},
    beginPath: () => {},
    rect: () => {},
    clip: () => {},
    moveTo: (x: number, y: number) => rec.lines.push({ x, y }),
    lineTo: (x: number, y: number) => rec.lines.push({ x, y }),
    stroke: () => {},
    fillRect: (x: number, y: number, w: number, h: number) => rec.fills.push({ x, y, w, h }),
  } as unknown as CanvasRenderingContext2D;
  paintSparkline(
    ctx,
    { x: 0, y: 0, w: 100, h: 30 },
    { settings, colors: sparklineColors(settings, null), points },
  );
  return rec;
};

const settingsOf = (type: SparklineType, patch: Partial<SparklineGroupSettings> = {}) => ({
  ...defaultSparklineGroupSettings(type),
  ...patch,
});

const series = (...values: number[]): SparklinePoint[] => values.map((value, x) => ({ x, value }));
const steps = (xs: readonly number[]): number[] => xs.slice(1).map((v, i) => v - (xs[i] ?? 0));

describe('sparklines over flat data', () => {
  it('draws the line level along the top of the plot area', () => {
    const rec = paint(settingsOf(SparklineType.Line), series(5, 5, 5));
    expect(rec.lines).toHaveLength(3);
    for (const point of rec.lines) expect(point.y).toBe(TOP);
  });

  it('draws every column at full height, evenly spaced', () => {
    const rec = paint(settingsOf(SparklineType.Column), series(5, 5, 5));
    expect(rec.fills).toHaveLength(3);
    for (const bar of rec.fills) {
      expect(bar.y).toBe(TOP);
      expect(bar.h).toBe(HEIGHT);
    }
    const [a, b] = steps(rec.fills.map((bar) => bar.x));
    expect(a).toBeCloseTo(b ?? Number.NaN);
  });
});

describe('sparklines over a single value', () => {
  it('draws nothing for a line, markers included', () => {
    for (const markers of [false, true]) {
      const rec = paint(settingsOf(SparklineType.Line, { markers }), series(7));
      expect(rec.lines).toEqual([]);
      expect(rec.fills).toEqual([]);
    }
  });

  it('draws one wide full-height column', () => {
    const rec = paint(settingsOf(SparklineType.Column), series(7));
    expect(rec.fills).toHaveLength(1);
    expect(rec.fills[0]).toMatchObject({ y: TOP, h: HEIGHT });
    expect(rec.fills[0]?.w).toBeGreaterThan(90);
  });
});

describe('sparklines on a date axis', () => {
  it('joins the line in date order, not source order', () => {
    const rec = paint(settingsOf(SparklineType.Line, { dateAxis: true }), [
      { x: 45_002, value: 1 },
      { x: 45_000, value: 3 },
      { x: 45_001, value: 2 },
    ]);
    const xs = rec.lines.map((point) => point.x);
    const ys = rec.lines.map((point) => point.y);
    expect(xs).toEqual([...xs].sort((p, q) => p - q));
    // 3, 2, 1 in date order: a straight descending line.
    const [dy1, dy2] = steps(ys);
    expect(dy1).toBeGreaterThan(0);
    expect(dy1).toBeCloseTo(dy2 ?? Number.NaN);
  });

  for (const type of [SparklineType.Line, SparklineType.Column]) {
    it(`spaces equally spaced dates uniformly (type ${type})`, () => {
      const rec = paint(settingsOf(type, { dateAxis: true }), [
        { x: 45_000, value: 1 },
        { x: 45_007, value: 3 },
        { x: 45_014, value: 2 },
      ]);
      const xs = (type === SparklineType.Line ? rec.lines : rec.fills).map((p) => p.x);
      const [a, b] = steps(xs);
      expect(a).toBeGreaterThan(0);
      expect(a).toBeCloseTo(b ?? Number.NaN);
    });
  }
});
