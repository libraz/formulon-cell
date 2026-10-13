import { describe, expect, it } from 'vitest';
import { layoutChart } from '../../../src/charts/layout.js';
import { clipScatterSegment, scatterLineSegments } from '../../../src/charts/scatter-lines.js';
import { resolveChartStyle } from '../../../src/charts/styles.js';
import type { ChartData, ChartPlotBounds, ChartSpec } from '../../../src/charts/types.js';

const data = (
  points: readonly { id: string; x: number | null; y: number | null }[],
): ChartData => ({
  domain: 'xy',
  series: [{ id: 'series', points }],
});

const sceneFor = (spec: ChartSpec, chartData: ChartData = data([])) =>
  layoutChart({ spec, data: chartData, width: 320, height: 132 });

const lines = (scene: ReturnType<typeof sceneFor>) =>
  scene.elements.filter((element) => element.role === 'line');

const points = (scene: ReturnType<typeof sceneFor>) =>
  scene.elements.filter((element) => element.role === 'point');

const finiteAttrs = (scene: ReturnType<typeof sceneFor>): boolean =>
  scene.elements.every((element) =>
    Object.values(element.attrs).every(
      (value) => typeof value !== 'number' || Number.isFinite(value),
    ),
  );

describe('scatter line geometry', () => {
  it('renders marker, line, and lineMarker modes with line-only default markers hidden', () => {
    const chartData = data([
      { id: 'p0', x: 0, y: 0 },
      { id: 'p1', x: 5, y: 5 },
      { id: 'p2', x: 10, y: 10 },
    ]);
    expect(points(sceneFor({ kind: 'scatter' }, chartData))).toHaveLength(3);
    expect(lines(sceneFor({ kind: 'scatter' }, chartData))).toHaveLength(0);
    expect(points(sceneFor({ kind: 'scatter', scatterStyle: 'line' }, chartData))).toHaveLength(0);
    expect(lines(sceneFor({ kind: 'scatter', scatterStyle: 'line' }, chartData))).toHaveLength(2);
    expect(
      points(sceneFor({ kind: 'scatter', scatterStyle: 'lineMarker' }, chartData)),
    ).toHaveLength(3);
    expect(
      lines(sceneFor({ kind: 'scatter', scatterStyle: 'lineMarker' }, chartData)),
    ).toHaveLength(2);
  });

  it('lets explicit series and point marker visibility override line defaults', () => {
    const chartData = data([
      { id: 'p0', x: 0, y: 0 },
      { id: 'p1', x: 5, y: 5 },
      { id: 'p2', x: 10, y: 10 },
    ]);
    const seriesVisible = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        seriesStyles: [{ seriesId: 'series', marker: { visible: true } }],
      },
      chartData,
    );
    expect(points(seriesVisible)).toHaveLength(3);

    const pointVisible = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        seriesStyles: [
          {
            seriesId: 'series',
            points: [{ pointId: 'p1', marker: { visible: true } }],
          },
        ],
      },
      chartData,
    );
    expect(points(pointVisible).map((point) => point.pointId)).toEqual(['p1']);

    const resolvedDefault = resolveChartStyle(
      { kind: 'scatter', scatterStyle: 'line' },
      ['#000'],
      'series',
      undefined,
      0,
      undefined,
    );
    expect(resolvedDefault.marker.visible).toBe(false);
    const resolvedSeries = resolveChartStyle(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        seriesStyles: [{ seriesId: 'series', marker: { visible: true } }],
      },
      ['#000'],
      'series',
      undefined,
      0,
      undefined,
    );
    expect(resolvedSeries.marker.visible).toBe(true);

    const inheritedFalse = resolveChartStyle(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        seriesStyles: [
          {
            seriesId: 'series',
            marker: { visible: undefined },
            points: [{ pointId: 'p1', marker: { visible: undefined } }],
          },
        ],
      },
      ['#000'],
      'series',
      'p1',
      0,
      1,
    );
    expect(inheritedFalse.marker.visible).toBe(false);
  });

  it('uses series line styling for connectors while point line styling remains marker-only', () => {
    const scene = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'lineMarker',
        seriesStyles: [
          {
            seriesId: 'series',
            line: { color: '#123456', transparency: 20, width: 3, dash: 'dot' },
            points: [{ pointId: 'p1', marker: { line: { color: '#abcdef', width: 5 } } }],
          },
        ],
      },
      data([
        { id: 'p0', x: 0, y: 0 },
        { id: 'p1', x: 5, y: 5 },
      ]),
    );
    const connector = lines(scene)[0];
    expect(connector?.attrs.stroke).toBe('#123456');
    expect(connector?.attrs['stroke-opacity']).toBe(0.8);
    expect(connector?.attrs['stroke-width']).toBe(3);
    expect(connector?.attrs['stroke-dasharray']).toBe('1 4');
    const styledMarker = points(scene).find((point) => point.pointId === 'p1');
    expect(styledMarker?.attrs.stroke).toBe('#abcdef');
    expect(styledMarker?.attrs['stroke-width']).toBe(5);

    const suppressed = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'lineMarker',
        seriesStyles: [{ seriesId: 'series', line: { width: 0 } }],
      },
      data([
        { id: 'p0', x: 0, y: 0 },
        { id: 'p1', x: 5, y: 5 },
      ]),
    );
    expect(lines(suppressed)).toHaveLength(0);
    expect(points(suppressed)).toHaveLength(2);
  });

  it('applies gap, zero, and span blank policies and always breaks on non-finite values', () => {
    const blankData = data([
      { id: 'p0', x: 0, y: 0 },
      { id: 'blank', x: null, y: 2 },
      { id: 'p2', x: 10, y: 10 },
    ]);
    expect(
      lines(sceneFor({ kind: 'scatter', scatterStyle: 'line', emptyCells: 'gap' }, blankData)),
    ).toHaveLength(0);
    expect(
      lines(sceneFor({ kind: 'scatter', scatterStyle: 'line', emptyCells: 'zero' }, blankData)),
    ).toHaveLength(2);
    expect(
      lines(sceneFor({ kind: 'scatter', scatterStyle: 'line', emptyCells: 'span' }, blankData)),
    ).toHaveLength(1);

    const invalidData = data([
      { id: 'p0', x: 0, y: 0 },
      { id: 'invalid', x: Number.NaN, y: 2 },
      { id: 'p2', x: 10, y: 10 },
    ]);
    expect(
      lines(sceneFor({ kind: 'scatter', scatterStyle: 'line', emptyCells: 'span' }, invalidData)),
    ).toHaveLength(0);

    const mixedInvalidData = data([
      { id: 'p0', x: 0, y: 0 },
      { id: 'invalid', x: null, y: Number.POSITIVE_INFINITY },
      { id: 'p2', x: 10, y: 10 },
    ]);
    expect(
      lines(
        sceneFor({ kind: 'scatter', scatterStyle: 'line', emptyCells: 'span' }, mixedInvalidData),
      ),
    ).toHaveLength(0);
  });

  it('keeps filtered source order and stable identities for connectors and markers', () => {
    const chartData: ChartData = {
      domain: 'xy',
      series: [
        { id: 'hidden', points: [{ id: 'hidden-point', x: 0, y: 0 }] },
        {
          id: 'series',
          points: [
            { id: 'first', x: 8, y: 2 },
            { id: 'filtered', x: 2, y: 4 },
            { id: 'last', x: 6, y: 8 },
          ],
        },
      ],
    };
    const scene = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'lineMarker',
        filter: { hiddenSeriesIds: ['hidden'], hiddenCategoryIds: ['filtered'] },
        xAxis: { min: 0, max: 10 },
        yAxis: { min: 0, max: 10 },
      },
      chartData,
    );
    expect(lines(scene)).toHaveLength(1);
    expect(lines(scene)[0]?.seriesId).toBe('series');
    expect(lines(scene)[0]?.seriesIndex).toBe(1);
    expect(points(scene).map((point) => point.pointId)).toEqual(['first', 'last']);
    expect(points(scene).every((point) => point.seriesId === 'series')).toBe(true);

    const unsorted = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        xAxis: { min: 0, max: 10 },
        yAxis: { min: 0, max: 10 },
      },
      data([
        { id: 'a', x: 8, y: 2 },
        { id: 'b', x: 2, y: 4 },
        { id: 'c', x: 6, y: 8 },
      ]),
    );
    const unsortedLines = lines(unsorted);
    expect(Number(unsortedLines[0]?.attrs.x1)).toBeGreaterThan(Number(unsortedLines[0]?.attrs.x2));
    expect(Number(unsortedLines[1]?.attrs.x1)).toBeLessThan(Number(unsortedLines[1]?.attrs.x2));
  });

  it('clips crossing segments, skips same-side offscreen pairs, and handles re-entry', () => {
    const spec: ChartSpec = {
      kind: 'scatter',
      scatterStyle: 'line',
      xAxis: { min: 0, max: 10 },
      yAxis: { min: 0, max: 10 },
    };
    const crossing = sceneFor(
      spec,
      data([
        { id: 'left', x: -100, y: 5 },
        { id: 'right', x: 100, y: 5 },
      ]),
    );
    const crossingLine = lines(crossing)[0];
    expect(lines(crossing)).toHaveLength(1);
    expect(crossingLine?.attrs.x1).toBe(crossing.plot.x);
    expect(crossingLine?.attrs.x2).toBe(crossing.plot.x + crossing.plot.width);
    expect(crossingLine?.attrs.y1).toBeCloseTo(crossing.plot.y + crossing.plot.height / 2);
    expect(crossingLine?.attrs.y2).toBeCloseTo(crossing.plot.y + crossing.plot.height / 2);

    const offscreen = sceneFor(
      spec,
      data([
        { id: 'a', x: -100, y: 5 },
        { id: 'b', x: -50, y: 5 },
      ]),
    );
    expect(lines(offscreen)).toHaveLength(0);
    const extremeOffscreen = sceneFor(
      spec,
      data([
        { id: 'a', x: -Number.MAX_VALUE, y: 5 },
        { id: 'b', x: -Number.MAX_VALUE / 2, y: 5 },
      ]),
    );
    expect(lines(extremeOffscreen)).toHaveLength(0);

    const reentry = sceneFor(
      spec,
      data([
        { id: 'left', x: -100, y: 5 },
        { id: 'middle', x: 5, y: 5 },
        { id: 'right', x: 100, y: 5 },
      ]),
    );
    expect(lines(reentry)).toHaveLength(2);
    expect(lines(reentry)[0]?.attrs.x1).toBe(reentry.plot.x);
    expect(lines(reentry)[1]?.attrs.x2).toBe(reentry.plot.x + reentry.plot.width);
    expect(finiteAttrs(reentry)).toBe(true);
  });

  it('does not invent crossings when an inactive corner coordinate remains outside', () => {
    const bounds: ChartPlotBounds = { x: 20, y: 20, width: 280, height: 140 };
    const sameSideCases = [
      [
        { kind: 'point' as const, x: 817.1999286860228, y: 183.43204772099853, sourceIndex: 0 },
        { kind: 'point' as const, x: -298.63125924021006, y: -724.6508649550378, sourceIndex: 1 },
      ],
      [
        { kind: 'point' as const, x: 991.1466613411903, y: 868.6048905365169, sourceIndex: 0 },
        { kind: 'point' as const, x: -972.4437585100532, y: -474.99800054356456, sourceIndex: 1 },
      ],
    ] as const;
    for (const [first, second] of sameSideCases) {
      expect(clipScatterSegment(first, second, bounds)).toBeNull();
    }
  });

  it('tracks both active boundaries for a true corner-to-corner crossing', () => {
    const bounds: ChartPlotBounds = { x: 20, y: 20, width: 280, height: 140 };
    const segment = clipScatterSegment(
      { kind: 'point', x: 0, y: 10, sourceIndex: 0 },
      { kind: 'point', x: 320, y: 170, sourceIndex: 1 },
      bounds,
    );
    expect(segment).toMatchObject({ x1: 20, y1: 20, x2: 300, y2: 160 });
  });

  it('keeps extreme screen coordinates finite and clips an opposing MAX_VALUE crossing', () => {
    const bounds: ChartPlotBounds = { x: 20, y: 20, width: 280, height: 92 };
    const segments = scatterLineSegments(
      [
        { kind: 'point', x: -Number.MAX_VALUE, y: 66, sourceIndex: 0 },
        { kind: 'point', x: Number.MAX_VALUE, y: 66, sourceIndex: 1 },
      ],
      bounds,
    );
    expect(segments).toHaveLength(1);
    expect(segments[0]).toMatchObject({ x1: 20, x2: 300, y1: 66, y2: 66 });
    expect(segments.every((segment) => Object.values(segment).every(Number.isFinite))).toBe(true);

    const scene = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        xAxis: { min: 0, max: 10 },
        yAxis: { min: 0, max: 10 },
      },
      data([
        { id: 'left', x: -Number.MAX_VALUE, y: 5 },
        { id: 'right', x: Number.MAX_VALUE, y: 5 },
      ]),
    );
    expect(lines(scene)).toHaveLength(1);
    expect(lines(scene)[0]?.attrs.x1).toBe(scene.plot.x);
    expect(lines(scene)[0]?.attrs.x2).toBe(scene.plot.x + scene.plot.width);
    expect(lines(scene)[0]?.attrs.y1).toBeCloseTo(scene.plot.y + scene.plot.height / 2);
    expect(lines(scene)[0]?.attrs.y2).toBeCloseTo(scene.plot.y + scene.plot.height / 2);
    expect(finiteAttrs(scene)).toBe(true);

    const diagonal = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        xAxis: { min: 0, max: 10 },
        yAxis: { min: 0, max: 10 },
      },
      data([
        { id: 'bottomLeft', x: -Number.MAX_VALUE, y: -Number.MAX_VALUE },
        { id: 'topRight', x: Number.MAX_VALUE, y: Number.MAX_VALUE },
      ]),
    );
    expect(lines(diagonal)).toHaveLength(1);
    expect(lines(diagonal)[0]?.attrs.x1).toBe(diagonal.plot.x);
    expect(lines(diagonal)[0]?.attrs.y1).toBe(diagonal.plot.y + diagonal.plot.height);
    expect(lines(diagonal)[0]?.attrs.x2).toBe(diagonal.plot.x + diagonal.plot.width);
    expect(lines(diagonal)[0]?.attrs.y2).toBe(diagonal.plot.y);
    expect(finiteAttrs(diagonal)).toBe(true);
  });

  it('preserves the crossing ratio for unequal extreme source coordinates', () => {
    const values = data([
      { id: 'left', x: -Number.MAX_VALUE / 2, y: 0 },
      { id: 'right', x: Number.MAX_VALUE, y: 10 },
    ]);
    for (const xMax of [10, 1]) {
      const scene = sceneFor(
        {
          kind: 'scatter',
          scatterStyle: 'line',
          xAxis: { min: 0, max: xMax },
          yAxis: { min: 0, max: 10 },
        },
        values,
      );
      const connector = lines(scene)[0];
      const expectedY = scene.plot.y + scene.plot.height * (1 - 1 / 3);
      expect(lines(scene)).toHaveLength(1);
      expect(connector?.attrs.x1).toBe(scene.plot.x);
      expect(connector?.attrs.x2).toBe(scene.plot.x + scene.plot.width);
      expect(connector?.attrs.y1).toBeCloseTo(expectedY, 8);
      expect(connector?.attrs.y2).toBeCloseTo(expectedY, 8);
      expect(finiteAttrs(scene)).toBe(true);
    }
  });

  it('keeps narrow subnormal domains finite for extreme connectors', () => {
    const scene = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        xAxis: { min: 0, max: Number.MIN_VALUE },
        yAxis: { min: 0, max: Number.MIN_VALUE },
      },
      data([
        { id: 'left', x: -Number.MAX_VALUE, y: -Number.MAX_VALUE },
        { id: 'right', x: Number.MAX_VALUE, y: Number.MAX_VALUE },
      ]),
    );
    expect(finiteAttrs(scene)).toBe(true);
    expect(
      lines(scene).every((line) =>
        Object.values(line.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);

    const subnormalBounds: ChartPlotBounds = { x: 20, y: 20, width: 280, height: 140 };
    const subnormalSegments = scatterLineSegments(
      [
        { kind: 'point', x: -5e-309, y: 0, sourceIndex: 0 },
        { kind: 'point', x: 1e-308, y: 10, sourceIndex: 1 },
      ],
      subnormalBounds,
      false,
      { x: { min: 0, max: 5e-309 }, y: { min: 0, max: 10 } },
    );
    expect(subnormalSegments).toHaveLength(1);
    expect(subnormalSegments[0]?.y1).toBeCloseTo(20 + 140 * (1 - 1 / 3), 8);
    expect(subnormalSegments[0]?.y2).toBeCloseTo(20 + 140 * (1 - 2 / 3), 8);
  });

  it('applies reverse screen projection after raw-domain clipping', () => {
    const bounds: ChartPlotBounds = { x: 20, y: 20, width: 280, height: 140 };
    const domains = { x: { min: 0, max: 10 }, y: { min: 0, max: 10 } };
    const normal = scatterLineSegments(
      [
        { kind: 'point', x: -Number.MAX_VALUE, y: -Number.MAX_VALUE, sourceIndex: 0 },
        { kind: 'point', x: Number.MAX_VALUE, y: Number.MAX_VALUE, sourceIndex: 1 },
      ],
      bounds,
      false,
      domains,
    );
    const reversed = scatterLineSegments(
      [
        { kind: 'point', x: -Number.MAX_VALUE, y: -Number.MAX_VALUE, sourceIndex: 0 },
        { kind: 'point', x: Number.MAX_VALUE, y: Number.MAX_VALUE, sourceIndex: 1 },
      ],
      bounds,
      false,
      domains,
      { reverseX: true, reverseY: true },
    );
    expect(normal).toHaveLength(1);
    expect(reversed).toHaveLength(1);
    expect(reversed[0]).toMatchObject({
      x1: bounds.x + bounds.width,
      y1: bounds.y,
      x2: bounds.x,
      y2: bounds.y + bounds.height,
    });
    expect(reversed.every((segment) => Object.values(segment).every(Number.isFinite))).toBe(true);
  });
});
