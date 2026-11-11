import { SparklineAxisType, SparklineType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { defaultSparklineGroupSettings } from '../../../src/commands/sparkline.js';
import type { SparklinePoint } from '../../../src/render/painters/controls.js';
import { paintSparkline } from '../../../src/render/painters.js';
import { readSparklinePoints, sparklineColors } from '../../../src/render/sparkline.js';
import { createSpreadsheetStore, type SparklineGroupSettings } from '../../../src/store/store.js';

interface Rec {
  fills: { x: number; y: number; w: number; h: number }[];
  lines: { x: number; y: number }[];
}

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
    {
      settings,
      colors: sparklineColors(settings, null),
      points,
    },
  );
  return rec;
};

const settingsOf = (type: SparklineType, patch: Partial<SparklineGroupSettings> = {}) => ({
  ...defaultSparklineGroupSettings(type),
  ...patch,
});

describe('sparkline painter with a very large series', () => {
  it('paints 200000 points without a spread overflow', () => {
    const points = Array.from({ length: 200_000 }, (_, x) => ({ x, value: x % 7 }));
    expect(() => paint(settingsOf(SparklineType.Line), points)).not.toThrow();
  });
});

describe('sparkline column and win/loss on a degenerate date axis', () => {
  for (const type of [SparklineType.Column, SparklineType.Stacked]) {
    it(`keeps every bar inside the cell for one point (type ${type})`, () => {
      const rec = paint(settingsOf(type, { dateAxis: true }), [{ x: 45_000, value: 3 }]);
      expect(rec.fills).toHaveLength(1);
      expect(rec.fills[0]?.x).toBeGreaterThanOrEqual(0);
      expect(rec.fills[0]?.x).toBeLessThan(100);
    });

    it(`spreads equal dates by position (type ${type})`, () => {
      const rec = paint(settingsOf(type, { dateAxis: true }), [
        { x: 45_000, value: 1 },
        { x: 45_000, value: 2 },
        { x: 45_000, value: 3 },
      ]);
      const xs = rec.fills.map((f) => f.x);
      expect(new Set(xs).size).toBe(3);
      expect(Math.max(...xs)).toBeLessThan(100);
    });
  }
});

describe('sparkline win/loss axis', () => {
  it('sits on the bars baseline even when the limits exclude zero', () => {
    const rec = paint(
      settingsOf(SparklineType.Stacked, {
        displayXAxis: true,
        minAxisType: SparklineAxisType.Custom,
        manualMin: 5,
        maxAxisType: SparklineAxisType.Custom,
        manualMax: 9,
      }),
      [
        { x: 0, value: 1 },
        { x: 1, value: -1 },
      ],
    );
    // Plot spans y 2..28, so the mid line is 15 (pixel-snapped to 15.5).
    expect(rec.lines[0]?.y).toBe(15.5);
    expect(rec.lines[1]?.y).toBe(15.5);
  });
});

describe('readSparklinePoints size cap', () => {
  it('reads nothing from a source larger than one full-width row', () => {
    const state = createSpreadsheetStore().getState();
    const settings = defaultSparklineGroupSettings(SparklineType.Line);
    expect(
      readSparklinePoints(state, null, { source: 'A1:A1000000', group: 'g' }, 0, settings),
    ).toBeNull();
  });
});
