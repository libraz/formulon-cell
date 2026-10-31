import { describe, expect, it } from 'vitest';
import { layoutChart } from '../../../src/charts/index.js';
import {
  chartLabelPositions,
  defaultChartLabelPosition,
  resolveChartLabelPosition,
} from '../../../src/charts/label-options.js';
import {
  axisTitleText,
  chartTitleText,
  layoutUnsupportedChart,
} from '../../../src/charts/layout.js';
import type {
  ChartAxisSpec,
  ChartData,
  ChartPrimitive,
  ChartSpec,
} from '../../../src/charts/types.js';

const sceneFor = (spec: ChartSpec, data: ChartData) =>
  layoutChart({ spec, data, width: 320, height: 132 });

const sceneAt = (spec: ChartSpec, data: ChartData, width: number, height: number) =>
  layoutChart({ spec, data, width, height });

const barElements = (scene: ReturnType<typeof sceneFor>) =>
  scene.elements.filter((element) => element.role === 'bar');

const pointPairs = (points: string): [number, number][] =>
  points.split(/\s+/).map((pair) => {
    const [x, y] = pair.split(',').map(Number);
    return [x as number, y as number];
  });

describe('layoutChart', () => {
  const salesData: ChartData = {
    domain: 'category',
    categories: ['Jan', 'Feb', 'Mar', 'Apr'],
    series: [
      { id: 'sales', name: 'Sales', values: [10, 15, 8, 12] },
      { id: 'costs', name: 'Costs', values: [6, 8, 3, 7] },
    ],
  };

  it('scales the value axis as Excel does for the default column chart', () => {
    const scene = sceneAt({ kind: 'column' }, salesData, 466, 274);
    const ticks = scene.elements
      .filter(
        (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
      )
      .map((element) => element.text);
    expect(ticks).toEqual(['0', '2', '4', '6', '8', '10', '12', '14', '16']);
  });

  it('centres a bottom legend under the chart', () => {
    const width = 466;
    const scene = sceneAt({ kind: 'column' }, salesData, width, 274);
    const keys = scene.elements.filter(
      (element) => element.role === 'legend' && element.tag === 'rect',
    );
    const labels = scene.elements.filter(
      (element) => element.role === 'legend' && element.tag === 'text',
    );
    const left = Number(keys[0]?.attrs.x);
    const last = labels[labels.length - 1];
    const right = Number(last?.attrs.x) + Number(String(last?.text).length) * 5.8;
    expect(Math.abs(left - (width - right))).toBeLessThan(2);
  });

  it('sets the title at 14 px above the plot', () => {
    const scene = sceneAt({ kind: 'column', title: 'Sales' }, salesData, 466, 274);
    const title = scene.elements.find((element) => element.role === 'title');
    expect(title?.attrs['font-size']).toBe(14);
    expect(title?.attrs.x).toBe(233);
    expect(scene.plot.y).toBeGreaterThan(Number(title?.attrs.y));
  });

  it('retains the default padded plot and emits readable category axes', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['one', 'two'],
      series: [{ id: 'series-0', name: 'B', values: [1, 2] }],
    };
    const scene = sceneFor({ kind: 'column' }, data);
    expect(scene.plot).toEqual({ x: 20, y: 20, width: 280, height: 92 });
    expect(
      scene.elements.some((element) => element.role === 'category-label' && element.text === 'one'),
    ).toBe(true);
    expect(scene.elements.some((element) => element.role === 'axis-tick')).toBe(true);
    expect(
      scene.elements.every((element) =>
        Object.values(element.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);
  });

  it('uses numeric x scaling for scatter and renders a full-circle pie safely', () => {
    const scatterData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'value',
          points: [
            { x: 1, y: 2 },
            { x: 10, y: 20 },
            { x: 100, y: 200 },
          ],
        },
      ],
    };
    const scatter = sceneFor({ kind: 'scatter' }, scatterData);
    const points = scatter.elements.filter((element) => element.role === 'point');
    expect(points).toHaveLength(3);
    expect(new Set(points.map((point) => point.attrs.cx)).size).toBe(3);

    const pieData: ChartData = {
      domain: 'category',
      categories: ['slice'],
      series: [{ id: 'series-0', name: 'Only', values: [1] }],
    };
    const pie = sceneFor({ kind: 'pie' }, pieData);
    expect(
      pie.elements.some((element) => element.tag === 'circle' && element.role === 'slice'),
    ).toBe(true);
    expect(
      pie.elements.every((element) =>
        Object.values(element.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);
  });

  it('separates positive and negative stacks and leaves empty points unplotted', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B', 'C'],
      series: [
        { id: 'a', values: [2, null, 3] },
        { id: 'b', values: [-1, 2, null] },
      ],
    };
    const stacked = sceneFor({ kind: 'column', grouping: 'stacked' }, data);
    const bars = stacked.elements.filter((element) => element.role === 'bar');
    expect(bars).toHaveLength(4);
    expect(bars.every((bar) => Number(bar.attrs.height) >= 0)).toBe(true);

    const gap = sceneFor({ kind: 'line' }, data);
    expect(
      gap.elements.some(
        (element) =>
          element.role === 'point' && element.seriesIndex === 0 && element.pointIndex === 1,
      ),
    ).toBe(false);
  });

  it('lays out all seven chart families with finite semantic primitives', () => {
    const categoryData: ChartData = {
      domain: 'category',
      categories: ['one', 'two', 'three'],
      series: [{ id: 'series-0', name: 'A', values: [1, 10, 100] }],
    };
    const scatterData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'A',
          points: [
            { x: 1, y: 2 },
            { x: 10, y: 20 },
            { x: 100, y: 200 },
          ],
        },
      ],
    };
    const cases: Array<{ kind: ChartSpec['kind']; data: ChartData; role: string }> = [
      { kind: 'column', data: categoryData, role: 'bar' },
      { kind: 'bar', data: categoryData, role: 'bar' },
      { kind: 'line', data: categoryData, role: 'point' },
      { kind: 'area', data: categoryData, role: 'area' },
      { kind: 'pie', data: categoryData, role: 'slice' },
      { kind: 'doughnut', data: categoryData, role: 'slice' },
      { kind: 'scatter', data: scatterData, role: 'point' },
    ];
    for (const testCase of cases) {
      const scene = sceneFor({ kind: testCase.kind }, testCase.data);
      expect(scene.elements.some((element) => element.role === testCase.role)).toBe(true);
      for (const element of scene.elements) {
        for (const value of Object.values(element.attrs)) {
          if (typeof value === 'number') expect(Number.isFinite(value)).toBe(true);
        }
      }
    }
  });

  it('uses logarithmically meaningful scatter spacing and percent endpoints', () => {
    const scatterData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'A',
          points: [
            { x: 1, y: 2 },
            { x: 10, y: 20 },
            { x: 100, y: 200 },
          ],
        },
      ],
    };
    const scatter = sceneFor({ kind: 'scatter' }, scatterData);
    const points = scatter.elements.filter((element) => element.role === 'point');
    const x = points.map((point) => Number(point.attrs.cx));
    expect((x[2] as number) - (x[1] as number)).toBeGreaterThan(
      8 * ((x[1] as number) - (x[0] as number)),
    );

    const percentData: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: [
        { id: 'positive', values: [2] },
        { id: 'negative', values: [-1] },
      ],
    };
    const percent = sceneFor({ kind: 'column', grouping: 'percent-stacked' }, percentData);
    const tickLabels = percent.elements
      .filter((element) => element.role === 'axis-label')
      .map((element) => element.text);
    expect(tickLabels).toContain('100%');
    expect(tickLabels).toContain('-100%');
  });

  it('breaks a line at gap points, joins across connect points, and plots value points', () => {
    const base = { domain: 'category', categories: ['A', 'B', 'C'] } as const;
    const lines = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements.filter((element) => element.role === 'line');
    const gap = sceneFor(
      { kind: 'line' },
      { ...base, series: [{ id: 's', values: [1, null, 3], states: ['value', 'gap', 'value'] }] },
    );
    const connect = sceneFor(
      { kind: 'line' },
      {
        ...base,
        series: [{ id: 's', values: [1, null, 3], states: ['value', 'connect', 'value'] }],
      },
    );
    const zero = sceneFor(
      { kind: 'line' },
      { ...base, series: [{ id: 's', values: [1, 0, 3], states: ['value', 'value', 'value'] }] },
    );
    expect(lines(gap)).toHaveLength(0);
    expect(gap.elements.filter((element) => element.role === 'point')).toHaveLength(2);
    expect(lines(connect)).toHaveLength(1);
    expect(pointPairs(String(lines(connect)[0]?.attrs.points))).toHaveLength(2);
    expect(connect.elements.filter((element) => element.role === 'point')).toHaveLength(2);
    expect(lines(zero)).toHaveLength(1);
    expect(zero.elements.filter((element) => element.role === 'point')).toHaveLength(3);

    const split = sceneFor(
      { kind: 'line' },
      {
        domain: 'category',
        categories: ['A', 'B', 'C', 'D', 'E'],
        series: [
          {
            id: 's',
            values: [1, 2, null, 4, 5],
            states: ['value', 'value', 'gap', 'value', 'value'],
          },
        ],
      },
    );
    expect(lines(split)).toHaveLength(2);
  });

  it('keeps the states of visible points when a filter hides series or categories', () => {
    const scene = sceneFor(
      { kind: 'line', filter: { hiddenSeriesIds: ['hidden'], hiddenCategoryIds: ['category-1'] } },
      {
        domain: 'category',
        categories: ['A', 'B', 'C', 'D'],
        series: [
          { id: 'hidden', values: [1, 1, 1, 1] },
          { id: 's', values: [1, 2, null, 4], states: ['value', 'value', 'gap', 'value'] },
        ],
      },
    );
    const lines = scene.elements.filter((element) => element.role === 'line');
    expect(lines).toHaveLength(0);
    expect(scene.elements.filter((element) => element.role === 'point')).toHaveLength(2);
  });

  it('breaks scatter lines at gap points and joins across connect points', () => {
    const points = (state: 'gap' | 'connect') => [
      { x: 1, y: 1 },
      { x: 2, y: 2 },
      { x: null, y: null, state },
      { x: 4, y: 4 },
      { x: 5, y: 5 },
    ];
    const segments = (state: 'gap' | 'connect') =>
      sceneFor(
        { kind: 'scatter', scatterStyle: 'lineMarker' },
        { domain: 'xy', series: [{ id: 's', points: points(state) }] },
      ).elements.filter((element) => element.role === 'line');
    expect(segments('gap')).toHaveLength(2);
    expect(segments('connect')).toHaveLength(3);
  });

  it('keeps stacked area envelopes', () => {
    const stackedAreaData: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [
        { id: 'base', values: [2, 2] },
        { id: 'top', values: [3, 3] },
      ],
    };
    const area = sceneFor({ kind: 'area', grouping: 'stacked' }, stackedAreaData);
    const areas = area.elements.filter((element) => element.role === 'area');
    expect(areas).toHaveLength(2);
    const baseUpper = pointPairs(String(areas[0]?.attrs.points)).slice(0, 2);
    const topPolygon = pointPairs(String(areas[1]?.attrs.points));
    const topLower = topPolygon.slice(2).reverse();
    expect(topLower[0]?.[1]).toBeCloseTo(baseUpper[0]?.[1] as number);
    expect(topLower[1]?.[1]).toBeCloseTo(baseUpper[1]?.[1] as number);
    expect(topLower.some(([, y]) => y !== area.plot.y + area.plot.height)).toBe(true);
  });

  it('keeps explicit one-sided axis bounds and never renders negative bars', () => {
    const lowerBound = sceneFor(
      { kind: 'column', yAxis: { min: 10 } },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [5] }] },
    );
    const lowerLabels = lowerBound.elements
      .filter((element) => element.role === 'axis-label')
      .map((element) => Number(element.text));
    expect(lowerLabels.length).toBeGreaterThan(0);
    expect(lowerLabels.every((value) => value >= 10)).toBe(true);
    expect(
      barElements(lowerBound).every(
        (bar) => Number(bar.attrs.height) >= 0 && Number(bar.attrs.y) >= lowerBound.plot.y,
      ),
    ).toBe(true);

    const upperBound = sceneFor(
      { kind: 'column', yAxis: { max: -10 } },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [-5] }] },
    );
    const upperLabels = upperBound.elements
      .filter((element) => element.role === 'axis-label')
      .map((element) => Number(element.text));
    expect(upperLabels.length).toBeGreaterThan(0);
    expect(upperLabels.every((value) => value <= -10)).toBe(true);
    expect(
      barElements(upperBound).every(
        (bar) => Number(bar.attrs.height) >= 0 && Number(bar.attrs.y) >= upperBound.plot.y,
      ),
    ).toBe(true);
  });

  it('supports axis visibility, independent gridlines, axis identities, and major units', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [{ id: 's', values: [3, 13] }],
    };
    const scene = sceneFor(
      {
        kind: 'column',
        xAxis: { title: 'Category' },
        yAxis: { title: 'Value', min: 3, max: 13, majorUnit: 5 },
      },
      data,
    );
    const numericLabels = scene.elements
      .filter(
        (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
      )
      .map((element) => element.text);
    expect(numericLabels).toEqual(['3', '8', '13']);
    expect(
      scene.elements
        .filter((element) =>
          ['axis', 'axis-tick', 'axis-label', 'category-label', 'axis-title'].includes(
            String(element.role),
          ),
        )
        .every(
          (element) =>
            element.attrs['data-chart-axis'] === 'x' || element.attrs['data-chart-axis'] === 'y',
        ),
    ).toBe(true);

    const hidden = sceneFor(
      {
        kind: 'column',
        xAxis: { visible: false, title: 'Hidden category' },
        yAxis: { visible: false, title: 'Hidden value' },
      },
      data,
    );
    expect(hidden.elements.filter((element) => element.role === 'axis-label')).toHaveLength(0);
    expect(hidden.elements.filter((element) => element.role === 'category-label')).toHaveLength(0);
    expect(hidden.elements.filter((element) => element.role === 'axis-title')).toHaveLength(0);
    expect(
      hidden.elements.some(
        (element) => element.role === 'axis-tick' && element.attrs['data-chart-axis'] === 'y',
      ),
    ).toBe(true);

    const noGridlines = sceneFor({ kind: 'column', yAxis: { majorGridlines: false } }, data);
    expect(
      noGridlines.elements.some(
        (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
      ),
    ).toBe(true);
    expect(
      noGridlines.elements.some(
        (element) => element.role === 'axis-tick' && element.attrs['data-chart-axis'] === 'y',
      ),
    ).toBe(false);
  });

  it('formats numeric axes through the locale callback and supports bar/scatter units', () => {
    const formatted = layoutChart({
      spec: {
        kind: 'column',
        yAxis: { min: 3, max: 13, majorUnit: 5, numberFormat: 'axis-code' },
      },
      data: {
        domain: 'category',
        categories: ['A'],
        series: [{ id: 's', values: [3] }],
      },
      width: 320,
      height: 132,
      locale: 'ja-JP',
      formatValue: (value, formatCode, locale) => `${locale}:${formatCode}:${value}`,
    });
    const formattedLabels = formatted.elements
      .filter(
        (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
      )
      .map((element) => element.text);
    expect(formattedLabels).toEqual([
      'ja-JP:axis-code:3',
      'ja-JP:axis-code:8',
      'ja-JP:axis-code:13',
    ]);
    expect(formatted.plot.x).toBeGreaterThan(20);

    const bar = sceneFor(
      { kind: 'bar', xAxis: { min: 3, max: 13, majorUnit: 5 } },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [3] }] },
    );
    expect(
      bar.elements
        .filter(
          (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'x',
        )
        .map((element) => element.text),
    ).toEqual(['3', '8', '13']);

    const scatter = sceneFor(
      {
        kind: 'scatter',
        xAxis: { min: 3, max: 13, majorUnit: 5 },
        yAxis: { min: 10, max: 30, majorUnit: 10 },
      },
      { domain: 'xy', series: [{ id: 's', points: [{ x: 3, y: 10 }] }] },
    );
    expect(
      scatter.elements
        .filter(
          (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'x',
        )
        .map((element) => element.text),
    ).toEqual(['3', '8', '13']);
    expect(
      scatter.elements
        .filter(
          (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
        )
        .map((element) => element.text),
    ).toEqual(['10', '20', '30']);
  });

  it('keeps x-axis lines horizontal and y-axis lines vertical across chart families', () => {
    const categoryData: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [{ id: 's', values: [-2, 4] }],
    };
    const scatterData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 's',
          points: [
            { x: -2, y: 1 },
            { x: 4, y: 3 },
          ],
        },
      ],
    };
    const cases = [
      ['column', categoryData],
      ['bar', categoryData],
      ['line', categoryData],
      ['area', categoryData],
      ['scatter', scatterData],
    ] as const;
    for (const [kind, data] of cases) {
      const scene = sceneFor({ kind }, data);
      const linesFor = (axis: 'x' | 'y') =>
        scene.elements.filter(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === axis,
        );
      const xLines = linesFor('x');
      const yLines = linesFor('y');
      expect(xLines.length, `${kind} x-axis line`).toBeGreaterThan(0);
      expect(yLines.length, `${kind} y-axis line`).toBeGreaterThan(0);
      expect(xLines.every((line) => line.attrs.y1 === line.attrs.y2)).toBe(true);
      expect(yLines.every((line) => line.attrs.x1 === line.attrs.x2)).toBe(true);

      const hiddenX = sceneFor({ kind, xAxis: { visible: false } }, data);
      expect(
        hiddenX.elements.filter(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
        ),
      ).toHaveLength(0);
      expect(
        hiddenX.elements.some(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
        ),
      ).toBe(true);

      const hiddenY = sceneFor({ kind, yAxis: { visible: false } }, data);
      expect(
        hiddenY.elements.filter(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
        ),
      ).toHaveLength(0);
      expect(
        hiddenY.elements.some(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
        ),
      ).toBe(true);
    }
    const bar = sceneFor({ kind: 'bar' }, categoryData);
    const barXLine = bar.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
    );
    expect(barXLine?.attrs.y1).toBe(bar.plot.y + bar.plot.height);
    const barYLine = bar.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
    );
    const zeroBaseline = Math.max(...barElements(bar).map((element) => Number(element.attrs.x)));
    expect(barYLine?.attrs.x1).toBeCloseTo(zeroBaseline);
  });

  it('keeps percent formatting and extreme invalid axis settings finite', () => {
    const percent = sceneFor(
      { kind: 'column', grouping: 'percent-stacked', yAxis: { numberFormat: '0.0%' } },
      {
        domain: 'category',
        categories: ['A'],
        series: [
          { id: 'positive', values: [2] },
          { id: 'negative', values: [-1] },
        ],
      },
    );
    expect(
      percent.elements
        .filter((element) => element.role === 'axis-label')
        .map((element) => element.text),
    ).toContain('100.0%');

    for (const spec of [
      {
        kind: 'column' as const,
        yAxis: {
          min: Number.NaN,
          max: Number.POSITIVE_INFINITY,
          majorUnit: Number.MIN_VALUE,
        },
      },
      { kind: 'scatter' as const, xAxis: { majorUnit: Number.MIN_VALUE }, yAxis: { majorUnit: 1 } },
    ]) {
      const data: ChartData =
        spec.kind === 'scatter'
          ? {
              domain: 'xy',
              series: [{ id: 's', points: [{ x: -Number.MAX_VALUE, y: Number.MAX_VALUE }] }],
            }
          : {
              domain: 'category',
              categories: ['A'],
              series: [{ id: 's', values: [Number.MAX_VALUE] }],
            };
      const scene = sceneFor(spec, data);
      expect(
        scene.elements.every((element) =>
          Object.values(element.attrs).every(
            (value) => typeof value !== 'number' || Number.isFinite(value),
          ),
        ),
      ).toBe(true);
    }

    const auto = sceneFor(
      { kind: 'column' },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [5, 15] }] },
    );
    const invalidExplicit = sceneFor(
      { kind: 'column', yAxis: { min: Number.NaN, max: Number.POSITIVE_INFINITY } },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [5, 15] }] },
    );
    const labels = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements
        .filter(
          (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
        )
        .map((element) => element.text);
    expect(labels(invalidExplicit)).toEqual(labels(auto));
  });

  it('reserves endpoints and y-title space for bounded custom axis labels', () => {
    const fullLabel = 'custom-number-label-'.repeat(40);
    const formattedX = layoutChart({
      spec: { kind: 'scatter', xAxis: { numberFormat: 'host-code' } },
      data: {
        domain: 'xy',
        series: [
          {
            id: 's',
            points: [
              { x: 0, y: 1 },
              { x: 10, y: 2 },
            ],
          },
        ],
      },
      width: 320,
      height: 132,
      formatValue: () => fullLabel,
    });
    const xLabel = formattedX.elements.find(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'x',
    );
    expect(formattedX.plot.x).toBeGreaterThan(20);
    expect(formattedX.plot.width).toBeGreaterThan(1);
    expect(xLabel?.text?.endsWith('…')).toBe(true);
    expect(xLabel?.attrs['aria-label']).toBe(fullLabel);

    const formattedY = layoutChart({
      spec: { kind: 'column', yAxis: { title: 'Value', numberFormat: 'host-code' } },
      data: {
        domain: 'category',
        categories: ['A'],
        series: [{ id: 's', values: [1, 2] }],
      },
      width: 320,
      height: 132,
      formatValue: () => fullLabel,
    });
    expect(formattedY.plot.x).toBeGreaterThan(100);
    expect(
      formattedY.elements.some(
        (element) => element.role === 'axis-title' && element.attrs['data-chart-axis'] === 'y',
      ),
    ).toBe(true);
    const formattedYTitle = formattedY.elements.find(
      (element) => element.role === 'axis-title' && element.attrs['data-chart-axis'] === 'y',
    );
    const formattedYLabel = formattedY.elements.find(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(Number(formattedYTitle?.attrs.x)).toBeLessThan(Number(formattedYLabel?.attrs.x) - 100);
  });

  it('keeps long custom scatter labels inside a compact canvas', () => {
    const fullLabel = 'host-number-label-'.repeat(10);
    const scene = layoutChart({
      spec: {
        kind: 'scatter',
        xAxis: { min: 0, max: 10, majorUnit: 5, numberFormat: 'host-code' },
        yAxis: { min: 0, max: 10, majorUnit: 5, numberFormat: 'host-code', title: 'Value' },
      },
      data: {
        domain: 'xy',
        series: [
          {
            id: 's',
            points: [
              { x: 0, y: 0 },
              { x: 10, y: 10 },
            ],
          },
        ],
      },
      width: 240,
      height: 132,
      formatValue: () => fullLabel,
    });
    expect(scene.plot.width).toBeGreaterThanOrEqual(48);

    const approximateTextWidth = (text: string): number =>
      Array.from(text).reduce((width, character) => width + (character === '…' ? 8.2 : 5.8), 0);
    const xLabels = scene.elements.filter(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'x',
    );
    expect(xLabels.length).toBeGreaterThan(0);
    expect(xLabels.every((label) => label.attrs['aria-label'] === fullLabel)).toBe(true);
    for (const label of xLabels) {
      const center = Number(label.attrs.x);
      const halfWidth = approximateTextWidth(String(label.text ?? '')) / 2;
      expect(center - halfWidth).toBeGreaterThanOrEqual(-0.5);
      expect(center + halfWidth).toBeLessThanOrEqual(240.5);
    }

    const yLabels = scene.elements.filter(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(yLabels.length).toBeGreaterThan(0);
    expect(yLabels.every((label) => label.attrs['aria-label'] === fullLabel)).toBe(true);
    for (const label of yLabels) {
      const right = Number(label.attrs.x);
      const labelWidth = approximateTextWidth(String(label.text ?? ''));
      expect(right - labelWidth).toBeGreaterThanOrEqual(-0.5);
      expect(right).toBeLessThanOrEqual(240.5);
    }
    const yTitle = scene.elements.find(
      (element) => element.role === 'axis-title' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(yTitle).toBeDefined();
    expect(Number(yTitle?.attrs.x)).toBeLessThan(Number(yLabels[0]?.attrs.x) - 100);
  });

  it('keeps rotated y-axis titles inside narrow charts in both orientations', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 's',
          points: [
            { x: 0, y: 0 },
            { x: 10, y: 10 },
          ],
        },
      ],
    };

    for (const reverseY of [false, true]) {
      const scene = sceneAt(
        {
          kind: 'scatter',
          xAxis: { min: 0, max: 10 },
          yAxis: { min: 0, max: 10, title: 'Y', reverseOrder: reverseY },
        },
        data,
        240,
        120,
      );
      const title = scene.elements.find(
        (element) => element.role === 'axis-title' && element.attrs['data-chart-axis'] === 'y',
      );
      const yTicks = scene.elements.filter(
        (element) => element.role === 'axis-tick' && element.attrs['data-chart-axis'] === 'y',
      );
      expect(title).toBeDefined();
      expect(Number(title?.attrs.x)).toBeGreaterThanOrEqual(12);
      expect(yTicks.length).toBeGreaterThan(0);
      expect(Number(title?.attrs.x)).toBeLessThan(
        Math.min(...yTicks.map((tick) => Number(tick.attrs.x1))),
      );
      expect(String(title?.attrs.transform)).toContain(`rotate(-90 ${title?.attrs.x} `);
    }
  });

  it('thins colliding numeric labels while retaining every gridline', () => {
    const denseX = layoutChart({
      spec: {
        kind: 'scatter',
        xAxis: { min: 0, max: 6, majorUnit: 1, numberFormat: 'host-code' },
        yAxis: { min: 0, max: 2, majorUnit: 1 },
      },
      data: {
        domain: 'xy',
        series: [
          {
            id: 's',
            points: [
              { x: 0, y: 0 },
              { x: 6, y: 2 },
            ],
          },
        ],
      },
      width: 320,
      height: 132,
      formatValue: () => 'long-x-axis-label',
    });
    const xLabels = denseX.elements.filter(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'x',
    );
    const xGridlines = denseX.elements.filter(
      (element) => element.role === 'axis-tick' && element.attrs['data-chart-axis'] === 'x',
    );
    expect(xGridlines).toHaveLength(7);
    expect(xLabels.length).toBeLessThan(xGridlines.length);
    expect(Number(xLabels[0]?.attrs.x)).toBeCloseTo(denseX.plot.x);
    expect(Number(xLabels.at(-1)?.attrs.x)).toBeCloseTo(denseX.plot.x + denseX.plot.width);
    for (let index = 1; index < xLabels.length; index += 1) {
      const previous = xLabels[index - 1];
      const current = xLabels[index];
      const previousWidth = (previous?.text?.length ?? 0) * 5.8;
      const currentWidth = (current?.text?.length ?? 0) * 5.8;
      expect(Number(current?.attrs.x) - Number(previous?.attrs.x)).toBeGreaterThanOrEqual(
        (previousWidth + currentWidth) / 2 + 3.9,
      );
    }

    const denseY = layoutChart({
      spec: { kind: 'column', yAxis: { min: 0, max: 1, majorUnit: 0.01 } },
      data: {
        domain: 'category',
        categories: ['A'],
        series: [{ id: 's', values: [0, 1] }],
      },
      width: 320,
      height: 132,
    });
    const yLabels = denseY.elements.filter(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
    );
    const yGridlines = denseY.elements.filter(
      (element) => element.role === 'axis-tick' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(yGridlines).toHaveLength(101);
    expect(yLabels.length).toBeLessThanOrEqual(Math.ceil(denseY.plot.height / 12) + 1);
    const yPositions = yLabels.map((label) => Number(label.attrs.y)).sort((a, b) => a - b);
    for (let index = 1; index < yPositions.length; index += 1) {
      expect((yPositions[index] ?? 0) - (yPositions[index - 1] ?? 0)).toBeGreaterThanOrEqual(13);
    }

    const barTitle = layoutChart({
      spec: { kind: 'bar', yAxis: { title: 'Category' } },
      data: {
        domain: 'category',
        categories: ['2026年10月', '2026年11月'],
        series: [{ id: 's', values: [-5, 20] }],
      },
      width: 640,
      height: 220,
    });
    const barYTitle = barTitle.elements.find(
      (element) => element.role === 'axis-title' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(Number(barYTitle?.attrs.x)).toBeGreaterThanOrEqual(12);
  });

  it('maps extreme and subnormal column values to finite full-span marks', () => {
    for (const value of [Number.MAX_VALUE, Number.MIN_VALUE]) {
      const scene = sceneFor(
        { kind: 'column' },
        { domain: 'category', categories: ['A'], series: [{ id: 's', values: [value] }] },
      );
      const mark = barElements(scene)[0];
      expect(mark).toBeDefined();
      expect(Number(mark?.attrs.y)).toBeCloseTo(scene.plot.y);
      expect(Number(mark?.attrs.height)).toBeCloseTo(scene.plot.height);
    }

    const tiny = sceneFor(
      { kind: 'column' },
      { domain: 'category', categories: ['A', 'B'], series: [{ id: 's', values: [1e-13, 2e-13] }] },
    );
    expect(barElements(tiny).every((bar) => Number(bar.attrs.height) > 0)).toBe(true);

    const huge = sceneFor(
      { kind: 'column' },
      {
        domain: 'category',
        categories: ['A', 'B'],
        series: [{ id: 's', values: [-1e308, 1e308] }],
      },
    );
    expect(
      huge.elements.every((element) =>
        Object.values(element.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);
  });

  it('derives scatter domains from finite pairs and omits points outside explicit bounds', () => {
    const scene = sceneFor(
      { kind: 'scatter', xAxis: { min: 0, max: 10 }, yAxis: { min: 0, max: 10 } },
      {
        domain: 'xy',
        series: [
          {
            id: 's',
            points: [
              { x: -1, y: 0 },
              { x: 0, y: 0 },
              { x: 10, y: 10 },
              { x: 11, y: 10 },
              { x: 1_000_000, y: null },
            ],
          },
        ],
      },
    );
    const points = scene.elements.filter((element) => element.role === 'point');
    expect(points).toHaveLength(2);
    expect(points.map((point) => point.pointIndex)).toEqual([1, 2]);
    expect(points.every((point) => Number(point.attrs.cx) >= scene.plot.x)).toBe(true);
    expect(points.every((point) => Number(point.attrs.cx) <= scene.plot.x + scene.plot.width)).toBe(
      true,
    );
  });

  it('keeps mixed-sign stack segments contiguous around zero', () => {
    const scene = sceneFor(
      { kind: 'column', grouping: 'stacked' },
      {
        domain: 'category',
        categories: ['A'],
        series: [
          { id: 'p1', values: [2] },
          { id: 'n1', values: [-1] },
          { id: 'p2', values: [3] },
          { id: 'n2', values: [-2] },
        ],
      },
    );
    const bars = barElements(scene);
    const positive = [bars[0], bars[2]];
    const negative = [bars[1], bars[3]];
    expect(Number(positive[1]?.attrs.y) + Number(positive[1]?.attrs.height)).toBeCloseTo(
      Number(positive[0]?.attrs.y),
    );
    expect(Number(negative[0]?.attrs.y) + Number(negative[0]?.attrs.height)).toBeCloseTo(
      Number(negative[1]?.attrs.y),
    );
    expect(Number(positive[0]?.attrs.y) + Number(positive[0]?.attrs.height)).toBeCloseTo(
      Number(negative[0]?.attrs.y),
    );
  });

  it('reserves y-axis space from the actual stacked domain', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: Array.from({ length: 100 }, (_, index) => ({
        id: `series-${index}`,
        values: [9],
      })),
    };
    const scene = sceneFor({ kind: 'column', grouping: 'stacked' }, data);
    expect(scene.plot.x).toBeGreaterThan(20);
    expect(
      scene.elements.some((element) => element.role === 'axis-label' && element.text === '1000'),
    ).toBe(true);
  });

  it('keeps overflowing positive and negative stacks finite and visible', () => {
    for (const value of [1e308, -1e308]) {
      for (const grouping of ['stacked', 'percent-stacked'] as const) {
        const scene = sceneFor(
          { kind: 'column', grouping },
          {
            domain: 'category',
            categories: ['A'],
            series: [
              { id: 'first', values: [value] },
              { id: 'second', values: [value] },
            ],
          },
        );
        const bars = barElements(scene);
        expect(bars).toHaveLength(2);
        expect(bars.every((bar) => Number(bar.attrs.height) > 0)).toBe(true);
        expect(
          scene.elements.every((element) =>
            Object.values(element.attrs).every(
              (attribute) => typeof attribute !== 'number' || Number.isFinite(attribute),
            ),
          ),
        ).toBe(true);
      }
    }
  });

  it('reserves no legend space for a doughnut legend that draws no entries', () => {
    const data: ChartData = { domain: 'category', categories: ['A', 'B'], series: [] };
    const shown = sceneFor(
      { kind: 'doughnut', legend: { visible: true, position: 'right' } },
      data,
    );
    const hidden = sceneFor({ kind: 'doughnut', legend: { visible: false } }, data);
    expect(shown.elements.some((element) => element.role === 'legend')).toBe(false);
    expect(shown.plot).toEqual(hidden.plot);
  });

  it('bounds compact legends and retains full labels in aria-label attributes', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: Array.from({ length: 20 }, (_, index) => ({
        id: `series-${index}`,
        name: `very-long-series-name-${index}`,
        values: [index + 1],
      })),
    };
    const right = sceneAt(
      { kind: 'column', legend: { visible: true, position: 'right' } },
      data,
      320,
      132,
    );
    const compactData = { ...data, series: data.series.slice(0, 12) };
    const top = sceneAt(
      { kind: 'column', legend: { visible: true, position: 'top' } },
      compactData,
      120,
      60,
    );
    const bottom = sceneAt(
      { kind: 'column', legend: { visible: true, position: 'bottom' }, xAxis: { title: '横軸' } },
      compactData,
      120,
      60,
    );
    for (const scene of [right, top, bottom]) {
      const legend = scene.elements.filter((element) => element.role === 'legend');
      expect(legend.length).toBeGreaterThan(0);
      for (const element of legend) {
        if (element.tag === 'rect') {
          expect(Number(element.attrs.x)).toBeGreaterThanOrEqual(0);
          expect(Number(element.attrs.x) + Number(element.attrs.width)).toBeLessThanOrEqual(
            scene.width,
          );
          expect(Number(element.attrs.y)).toBeGreaterThanOrEqual(0);
          expect(Number(element.attrs.y) + Number(element.attrs.height)).toBeLessThanOrEqual(
            scene.height,
          );
        } else {
          expect(Number(element.attrs.x)).toBeGreaterThanOrEqual(0);
          expect(Number(element.attrs.x)).toBeLessThanOrEqual(scene.width);
          expect(element.attrs['aria-label']).toBeDefined();
        }
      }
    }
    expect(
      right.elements.some(
        (element) =>
          element.tag === 'text' && element.role === 'legend' && element.text?.startsWith('+'),
      ),
    ).toBe(true);
  });

  it('keeps no-data and single-value scenes finite', () => {
    const noData: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: [{ id: 's', values: [null] }],
    };
    const scene = sceneFor({ kind: 'pie' }, noData);
    expect(scene.elements.some((element) => element.role === 'empty' && element.text === '—')).toBe(
      true,
    );
    const zero = sceneFor(
      { kind: 'column' },
      {
        domain: 'category',
        categories: ['A'],
        series: [{ id: 's', values: [0], states: ['value'] }],
      },
    );
    expect(zero.elements.some((element) => element.role === 'empty')).toBe(false);
    expect(barElements(zero)).toHaveLength(1);
    for (const kind of ['column', 'bar', 'line', 'area', 'pie'] as const) {
      const single = sceneFor(
        { kind },
        { domain: 'category', categories: ['A'], series: [{ id: 's', values: [1] }] },
      );
      expect(single.elements.length).toBeGreaterThan(0);
    }
  });

  it('uses Excel gapWidth and overlap equations for clustered and stacked slots', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: [
        { id: 'first', values: [1] },
        { id: 'second', values: [2] },
      ],
    };
    const defaultBars = barElements(sceneFor({ kind: 'column', legend: { visible: false } }, data));
    expect(defaultBars.map((bar) => Number(bar.attrs.x))).toEqual([80, 160]);
    expect(defaultBars.map((bar) => Number(bar.attrs.width))).toEqual([80, 80]);

    const noGap = barElements(
      sceneFor({ kind: 'column', gapWidth: 0, legend: { visible: false } }, data),
    );
    expect(noGap[0]?.attrs.x).toBe(20);
    expect(noGap[1]?.attrs.x).toBe(160);
    expect(noGap.map((bar) => Number(bar.attrs.width))).toEqual([140, 140]);

    const fullOverlap = barElements(
      sceneFor({ kind: 'column', gapWidth: 0, overlap: 100, legend: { visible: false } }, data),
    );
    expect(fullOverlap[0]?.attrs.x).toBe(20);
    expect(fullOverlap[1]?.attrs.x).toBe(20);
    expect(fullOverlap.map((bar) => Number(bar.attrs.width))).toEqual([280, 280]);

    const stacked = barElements(
      sceneFor({ kind: 'column', grouping: 'stacked', legend: { visible: false } }, data),
    );
    expect(stacked.map((bar) => Number(bar.attrs.x))).toEqual([104, 104]);
    expect(stacked.map((bar) => Number(bar.attrs.width))).toEqual([112, 112]);

    const clamped = barElements(
      sceneFor(
        {
          kind: 'column',
          gapWidth: Number.POSITIVE_INFINITY,
          overlap: Number.NaN,
          legend: { visible: false },
        },
        data,
      ),
    );
    expect(clamped.map((bar) => Number(bar.attrs.x))).toEqual(
      defaultBars.map((bar) => Number(bar.attrs.x)),
    );
    expect(clamped.map((bar) => Number(bar.attrs.width))).toEqual(
      defaultBars.map((bar) => Number(bar.attrs.width)),
    );
  });

  it('draws stacked and area series from the engine states without re-deriving blanks', () => {
    const data = (state: 'gap' | 'connect' | 'value'): ChartData => ({
      domain: 'category',
      categories: ['A', 'B', 'C'],
      series: [
        {
          id: 'series',
          values: [1, state === 'value' ? 0 : null, 3],
          states: ['value', state, 'value'],
        },
      ],
    });
    const points = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements.filter((element) => element.role === 'point');
    expect(points(sceneFor({ kind: 'line', grouping: 'stacked' }, data('value')))).toHaveLength(3);
    expect(points(sceneFor({ kind: 'line', grouping: 'stacked' }, data('gap')))).toHaveLength(2);
    expect(points(sceneFor({ kind: 'area' }, data('connect')))).toHaveLength(2);
    expect(barElements(sceneFor({ kind: 'column' }, data('gap')))).toHaveLength(2);
    expect(barElements(sceneFor({ kind: 'column' }, data('value')))).toHaveLength(3);
  });

  it('draws scatter points from the engine states and drops nonfinite coordinates', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series',
          points: [
            { x: 0, y: 1, state: 'value' },
            { x: 2, y: 0, state: 'value' },
            { x: Number.POSITIVE_INFINITY, y: 3 },
            { x: Number.NaN, y: 4 },
            { x: 4, y: 5 },
            { x: null, y: null, state: 'gap' },
          ],
        },
      ],
    };
    const scene = sceneFor({ kind: 'scatter' }, data);
    const points = scene.elements.filter((element) => element.role === 'point');
    expect(points.map((point) => point.pointIndex)).toEqual([0, 1, 4]);
    expect(Number(points[0]?.attrs.cx)).toBeCloseTo(scene.plot.x);
    expect(Number(points[1]?.attrs.cy)).toBeCloseTo(scene.plot.y + scene.plot.height);
    expect(scene.elements.some((element) => element.role === 'empty')).toBe(false);
  });

  it('normalizes pie magnitudes before summing and rotates from twelve o’clock', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['positive', 'large', 'negative', 'zero', 'invalid'],
      series: [
        {
          id: 'series',
          values: [1, Number.MAX_VALUE, -Number.MAX_VALUE, 0, Number.NaN],
        },
      ],
    };
    const scene = sceneFor({ kind: 'pie', firstSliceAngle: 90 }, data);
    const slices = scene.elements.filter((element) => element.role === 'slice');
    expect(slices.map((slice) => slice.pointIndex)).toEqual([0, 1, 2]);
    expect(slices).toHaveLength(3);
    expect(
      scene.elements.every((element) =>
        Object.values(element.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);
    const firstPath = slices[0]?.attrs.d;
    const cx = scene.plot.x + scene.plot.width / 2;
    const cy = scene.plot.y + scene.plot.height / 2;
    const radius = Math.min(scene.plot.width, scene.plot.height) / 2 - 2;
    expect(firstPath).toContain(`L ${cx + radius} ${cy} A`);
  });

  it('lays out doughnut rings with a clamped hole and stable full-circle paths', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [
        { id: 'inner', values: [1, 1] },
        { id: 'outer', values: [1, 0] },
      ],
    };
    const doughnut = (holeSize: number | undefined) =>
      sceneFor({ kind: 'doughnut', holeSize, legend: { visible: false } }, data);
    const normal = doughnut(50);
    const invalid = doughnut(Number.NaN);
    const low = doughnut(0);
    const high = doughnut(100);
    const normalSlices = normal.elements.filter((element) => element.role === 'slice');
    expect(normalSlices).toHaveLength(3);
    expect(invalid.elements.filter((element) => element.role === 'slice')).toEqual(normalSlices);
    expect(low.elements.filter((element) => element.role === 'slice')[0]?.attrs.d).not.toBe(
      normalSlices[0]?.attrs.d,
    );
    expect(high.elements.filter((element) => element.role === 'slice')[0]?.attrs.d).not.toBe(
      normalSlices[0]?.attrs.d,
    );
    const full = normalSlices[2];
    expect(String(full?.attrs.d).match(/A/g)).toHaveLength(4);
    expect(full?.attrs['fill-rule']).toBe('evenodd');
    expect(full?.attrs.stroke).toBe('Canvas');
    expect(full?.attrs['stroke-width']).toBe(1);
    expect(full).toMatchObject({ role: 'slice', seriesIndex: 1, pointIndex: 0 });
    expect(String(full?.attrs.d)).not.toMatch(/NaN|Infinity/);

    const identical = sceneFor(
      { kind: 'doughnut', legend: { visible: false } },
      {
        domain: 'category',
        categories: ['A', 'B'],
        series: [
          { id: 'inner', values: [1, 1] },
          { id: 'outer', values: [1, 1] },
        ],
      },
    );
    const identicalSlices = identical.elements.filter((element) => element.role === 'slice');
    expect(identicalSlices).toHaveLength(4);
    expect(
      identicalSlices.map((slice) => [slice.attrs.stroke, slice.attrs['stroke-width']]),
    ).toEqual([
      ['Canvas', 1],
      ['Canvas', 1],
      ['Canvas', 1],
      ['Canvas', 1],
    ]);
    const explicitLine = sceneFor(
      {
        kind: 'doughnut',
        legend: { visible: false },
        seriesStyles: [{ seriesId: 'inner', line: { color: '#edge', width: 0 } }],
      },
      data,
    ).elements.find((element) => element.role === 'slice' && element.seriesId === 'inner');
    expect(explicitLine?.attrs.stroke).toBe('#edge');
    expect(explicitLine?.attrs['stroke-width']).toBe(0);
  });

  it('draws per-slice pie fills keyed by category id, also under a category filter', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B', 'C'],
      categoryIds: ['category-0', 'category-1', 'category-2'],
      series: [{ id: 'series-0', values: [1, 2, 3] }],
    };
    const spec: ChartSpec = {
      kind: 'pie',
      seriesStyles: [
        {
          seriesId: 'series-0',
          points: [
            { pointId: 'category-0', fill: { color: '#aa0000' } },
            { pointId: 'category-1', fill: { color: '#00aa00' } },
            { pointId: 'category-2', fill: { color: '#0000aa' } },
          ],
        },
      ],
    };
    const fills = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements
        .filter((element) => element.role === 'slice')
        .map((element) => [element.pointId, element.attrs.fill]);

    expect(fills(sceneFor(spec, data))).toEqual([
      ['category-0', '#aa0000'],
      ['category-1', '#00aa00'],
      ['category-2', '#0000aa'],
    ]);
    // Hiding the first category leaves the others on their own styles.
    expect(
      fills(sceneFor({ ...spec, filter: { hiddenCategoryIds: ['category-0'] } }, data)),
    ).toEqual([
      ['category-1', '#00aa00'],
      ['category-2', '#0000aa'],
    ]);
  });

  it('keeps doughnut filtering, first-angle rotation, identities, ring styles, and category legends', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B', 'C'],
      categoryIds: ['a', 'b', 'c'],
      series: [
        { id: 'hidden', values: [9, 9, 9] },
        { id: 'kept', values: [1, 2, 3] },
        { id: 'outer', values: [3, 2, 1] },
      ],
    };
    const scene = sceneFor(
      {
        kind: 'doughnut',
        firstSliceAngle: 90,
        palette: ['#palette-a', '#palette-b'],
        seriesStyles: [
          {
            seriesId: 'kept',
            fill: { color: '#ring' },
            points: [{ pointId: 'c', fill: { color: '#point' } }],
          },
          { seriesId: 'outer', fill: { color: '#outer' } },
        ],
        filter: { hiddenSeriesIds: ['hidden'], hiddenCategoryIds: ['b'] },
      },
      data,
    );
    const slices = scene.elements.filter((element) => element.role === 'slice');
    expect(slices).toHaveLength(4);
    expect(
      slices.map((slice) => [slice.seriesIndex, slice.pointIndex, slice.seriesId, slice.pointId]),
    ).toEqual([
      [1, 0, 'kept', 'a'],
      [1, 2, 'kept', 'c'],
      [2, 0, 'outer', 'a'],
      [2, 2, 'outer', 'c'],
    ]);
    expect(slices.map((slice) => slice.attrs.fill)).toEqual([
      '#ring',
      '#point',
      '#outer',
      '#outer',
    ]);
    const firstPath = String(slices[0]?.attrs.d);
    const numbers = firstPath.match(/[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gi)?.map(Number);
    const cx = scene.plot.x + scene.plot.width / 2;
    const cy = scene.plot.y + scene.plot.height / 2;
    expect(numbers?.[0]).toBeGreaterThan(cx);
    expect(numbers?.[1]).toBeCloseTo(cy);
    const legend = scene.elements.filter(
      (element) => element.role === 'legend' && element.tag === 'rect',
    );
    expect(legend.map((entry) => entry.attrs.fill)).toEqual(['#ring', '#point']);
    expect(legend.map((entry) => [entry.seriesId, entry.pointId])).toEqual([
      ['kept', 'a'],
      ['kept', 'c'],
    ]);
  });

  it('renders doughnut labels per ring with signed values and per-ring percentages', () => {
    const scene = sceneFor(
      {
        kind: 'doughnut',
        legend: { visible: false },
        dataLabels: { visible: true, showValue: true, showPercentage: true },
      },
      {
        domain: 'category',
        categories: ['A', 'B'],
        series: [
          { id: 'inner', values: [2, -1] },
          { id: 'outer', values: [1, 3] },
        ],
      },
    );
    const labels = scene.elements.filter((element) => element.role === 'data-label');
    expect(labels.map((label) => label.text)).toEqual(['2, 67%', '-1, 33%', '1, 25%', '3, 75%']);
    expect(labels.map((label) => label.attrs['text-anchor'])).toEqual([
      'middle',
      'middle',
      'middle',
      'middle',
    ]);
    expect(labels.map((label) => label.attrs['aria-label'])).toEqual(
      labels.map((label) => label.text),
    );
    expect(labels.map((label) => [label.seriesIndex, label.pointIndex, label.seriesId])).toEqual([
      [0, 0, 'inner'],
      [0, 1, 'inner'],
      [1, 0, 'outer'],
      [1, 1, 'outer'],
    ]);
    const cx = scene.plot.x + scene.plot.width / 2;
    const cy = scene.plot.y + scene.plot.height / 2;
    const outerRadius = Math.min(scene.plot.width, scene.plot.height) / 2 - 2;
    const holeRadius = outerRadius / 2;
    const ringWidth = (outerRadius - holeRadius) / 2;
    for (const label of labels) {
      const ringIndex = label.seriesIndex ?? 0;
      const expectedRadius = holeRadius + ringWidth * ringIndex + ringWidth / 2;
      expect(Math.hypot(Number(label.attrs.x) - cx, Number(label.attrs.y) - cy)).toBeCloseTo(
        expectedRadius,
        5,
      );
    }
  });

  it('separates same-angle full-ring labels while retaining both percentages', () => {
    const scene = sceneFor(
      {
        kind: 'doughnut',
        legend: { visible: false },
        dataLabels: { visible: true, showValue: false, showPercentage: true },
      },
      {
        domain: 'category',
        categories: ['A'],
        series: [
          { id: 'inner', values: [1] },
          { id: 'outer', values: [1] },
        ],
      },
    );
    const labels = scene.elements.filter((element) => element.role === 'data-label');
    expect(labels.map((label) => label.text)).toEqual(['100%', '100%']);
    expect(labels).toHaveLength(2);
    const firstY = Number(labels[0]?.attrs.y);
    const secondY = Number(labels[1]?.attrs.y);
    const fontSize = Math.max(
      Number(labels[0]?.attrs['font-size']),
      Number(labels[1]?.attrs['font-size']),
    );
    expect(Math.abs(firstY - secondY)).toBeGreaterThan(fontSize + 4);
    expect(labels.map((label) => label.attrs['aria-label'])).toEqual(['100%', '100%']);
  });

  it('uses themed text for labels while keeping legend keys on the data color', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: [{ id: 's', values: [2] }],
    };
    const baseSpec: ChartSpec = {
      kind: 'doughnut',
      legend: { visible: false },
      seriesStyles: [{ seriesId: 's', fill: { color: '#slice' } }],
      dataLabels: { visible: true, showValue: true, showLegendKey: true },
    };
    const defaultScene = sceneFor(baseSpec, data);
    const defaultLabel = defaultScene.elements.find((element) => element.role === 'data-label');
    const defaultKey = defaultScene.elements.find((element) => element.role === 'data-label-key');
    expect(defaultLabel?.attrs.fill).toBe('currentColor');
    expect(defaultKey?.attrs.fill).toBe('#slice');

    const explicitScene = sceneFor(
      { ...baseSpec, dataLabels: { ...baseSpec.dataLabels, color: '#label' } },
      data,
    );
    const explicitLabel = explicitScene.elements.find((element) => element.role === 'data-label');
    const explicitKey = explicitScene.elements.find((element) => element.role === 'data-label-key');
    expect(explicitLabel?.attrs.fill).toBe('#label');
    expect(explicitKey?.attrs.fill).toBe('#slice');
  });

  it('keeps percentage labels separated across three visible rings', () => {
    const scene = sceneFor(
      {
        kind: 'doughnut',
        holeSize: 50,
        legend: { visible: false },
        dataLabels: { visible: true, showPercentage: true },
      },
      {
        domain: 'category',
        categories: ['A', 'B', 'C'],
        series: [
          { id: 'inner', values: [10, 15, 8] },
          { id: 'middle', values: [10, 15, 8] },
          { id: 'outer', values: [10, 15, 8] },
        ],
      },
    );
    const labels = scene.elements.filter((element) => element.role === 'data-label');
    expect(labels.map((label) => label.text)).toEqual([
      '30%',
      '45%',
      '24%',
      '30%',
      '45%',
      '24%',
      '30%',
      '45%',
      '24%',
    ]);
    const boxes = labels.map((label) => {
      const fontSize = Number(label.attrs['font-size']);
      const x = Number(label.attrs.x);
      const y = Number(label.attrs.y);
      const halfWidth = String(label.text).length * fontSize * 0.35 + 2;
      return {
        left: x - halfWidth,
        right: x + halfWidth,
        top: y - fontSize * 0.75,
        bottom: y + fontSize * 0.65,
      };
    });
    for (let first = 0; first < boxes.length; first += 1) {
      for (let second = first + 1; second < boxes.length; second += 1) {
        const a = boxes[first];
        const b = boxes[second];
        if (!a || !b) continue;
        expect(
          a.left < b.right && b.left < a.right && a.top < b.bottom && b.top < a.bottom,
          `${first}:${second} ${JSON.stringify({ a, b })}`,
        ).toBe(false);
      }
    }
  });

  it('uses every visible doughnut series for emptiness and omits axes and numeric callbacks', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: [
        { id: 'zero', values: [0] },
        { id: 'positive', values: [2] },
      ],
    };
    let callbackCalls = 0;
    const scene = layoutChart({
      spec: {
        kind: 'doughnut',
        legend: { visible: false },
        xAxis: { title: 'X', numberFormat: '0.0' },
        yAxis: { title: 'Y', numberFormat: '0.0' },
      },
      data,
      width: 320,
      height: 132,
      formatValue: () => {
        callbackCalls += 1;
        return 'callback';
      },
    });
    expect(scene.elements.filter((element) => element.role === 'slice')).toHaveLength(1);
    expect(scene.elements.some((element) => element.seriesIndex === 1)).toBe(true);
    expect(scene.elements.some((element) => element.role === 'empty')).toBe(false);
    expect(scene.elements.some((element) => element.role?.startsWith('axis'))).toBe(false);
    expect(callbackCalls).toBe(0);

    const axisData: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [{ id: 's', values: [1, 2] }],
    };
    for (const kind of ['pie', 'doughnut'] as const) {
      const withoutAxes = sceneFor({ kind, legend: { visible: false } }, axisData);
      const inheritedAxes = sceneFor(
        {
          kind,
          legend: { visible: false },
          xAxis: { title: 'X', numberFormat: '0.0' },
          yAxis: { title: 'Y', numberFormat: '0.0' },
        },
        axisData,
      );
      expect(inheritedAxes.plot).toEqual(withoutAxes.plot);
    }

    const filtered = sceneFor(
      { kind: 'doughnut', filter: { hiddenSeriesIds: ['zero', 'positive'] } },
      data,
    );
    expect(filtered.elements.filter((element) => element.role === 'slice')).toHaveLength(0);
    expect(filtered.elements.some((element) => element.role === 'empty')).toBe(true);
  });

  it('keeps extreme doughnut magnitudes finite in each ring', () => {
    const scene = sceneFor(
      { kind: 'doughnut', legend: { visible: false } },
      {
        domain: 'category',
        categories: ['max', 'min'],
        series: [
          { id: 'outer', values: [Number.MAX_VALUE, Number.MIN_VALUE] },
          { id: 'inner', values: [Number.MIN_VALUE, Number.MAX_VALUE] },
        ],
      },
    );
    const slices = scene.elements.filter((element) => element.role === 'slice');
    expect(slices).toHaveLength(2);
    expect(slices.every((slice) => !String(slice.attrs.d).match(/NaN|Infinity/))).toBe(true);
  });

  it('keeps large doughnut label sets complete without recursive placement', () => {
    const categories = Array.from({ length: 1000 }, (_, index) => `C${index}`);
    const scene = sceneFor(
      {
        kind: 'doughnut',
        legend: { visible: false },
        dataLabels: { visible: true, showValue: true },
      },
      {
        domain: 'category',
        categories,
        series: [{ id: 's', values: categories.map(() => 1) }],
      },
    );
    expect(scene.elements.filter((element) => element.role === 'data-label')).toHaveLength(
      categories.length,
    );
  });

  it('marks an all-zero pie empty while retaining valid zero category values', () => {
    const pie = sceneFor(
      { kind: 'pie' },
      { domain: 'category', categories: ['A', 'B'], series: [{ id: 's', values: [0, 0] }] },
    );
    expect(pie.elements.filter((element) => element.role === 'slice')).toHaveLength(0);
    expect(pie.elements.some((element) => element.role === 'empty')).toBe(true);

    const column = sceneFor(
      { kind: 'column' },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [0] }] },
    );
    expect(column.elements.some((element) => element.role === 'empty')).toBe(false);
    expect(barElements(column)).toHaveLength(1);
  });

  it.each(['bar', 'column'] as const)(
    'keeps %s stack bands aligned and applies finite geometry limits',
    (kind) => {
      const data: ChartData = {
        domain: 'category',
        categories: ['A'],
        series: [
          { id: 'a', values: [1] },
          { id: 'b', values: [2] },
        ],
      };
      const position = kind === 'bar' ? 'y' : 'x';
      const size = kind === 'bar' ? 'height' : 'width';
      for (const grouping of ['stacked', 'percent-stacked'] as const) {
        const scene = sceneFor({ kind, grouping, legend: { visible: false } }, data);
        const bars = barElements(scene);
        expect(bars).toHaveLength(2);
        expect(bars[0]?.attrs[position]).toBe(bars[1]?.attrs[position]);
        expect(bars[0]?.attrs[size]).toBe(bars[1]?.attrs[size]);
        const slot = kind === 'bar' ? scene.plot.height : scene.plot.width;
        expect(Number(bars[0]?.attrs[size])).toBeCloseTo(slot / 2.5);
      }
      const scene = sceneFor(
        { kind, gapWidth: 500, overlap: -100, legend: { visible: false } },
        data,
      );
      const bars = barElements(scene);
      const slot = kind === 'bar' ? scene.plot.height : scene.plot.width;
      expect(Number(bars[0]?.attrs[size])).toBeCloseTo(slot / 8);
      expect(Number(bars[1]?.attrs[position]) - Number(bars[0]?.attrs[position])).toBeCloseTo(
        slot / 4,
      );
      const clamped = sceneFor(
        { kind, gapWidth: 999, overlap: -999, legend: { visible: false } },
        data,
      );
      expect(barElements(clamped).map((bar) => [bar.attrs[position], bar.attrs[size]])).toEqual(
        bars.map((bar) => [bar.attrs[position], bar.attrs[size]]),
      );
    },
  );

  it('draws a stacked line through an engine zero and breaks an area at a gap', () => {
    const base = { domain: 'category', categories: ['A', 'B', 'C'] } as const;
    const scene = sceneFor(
      { kind: 'line', grouping: 'stacked' },
      { ...base, series: [{ id: 's', values: [1, 0, 3], states: ['value', 'value', 'value'] }] },
    );
    const lines = scene.elements.filter((element) => element.role === 'line');
    expect(lines).toHaveLength(1);
    const points = pointPairs(String(lines[0]?.attrs.points));
    expect(points).toHaveLength(3);
    expect(points[1]?.[1]).toBeCloseTo(scene.plot.y + scene.plot.height);
    const area = sceneFor(
      { kind: 'area' },
      { ...base, series: [{ id: 's', values: [1, null, 3], states: ['value', 'gap', 'value'] }] },
    );
    expect(area.elements.filter((element) => element.role === 'area')).toHaveLength(2);
  });

  it('preserves pie proportions for overflowing totals and negative magnitudes', () => {
    const pie = (values: number[], firstSliceAngle = 0) =>
      sceneFor(
        { kind: 'pie', firstSliceAngle, legend: { visible: false } },
        {
          domain: 'category',
          categories: ['A', 'B'],
          series: [{ id: 's', values }],
        },
      );
    const scene = pie([Number.MAX_VALUE, Number.MAX_VALUE], 90);
    const slices = scene.elements.filter((element) => element.role === 'slice');
    expect(slices).toHaveLength(2);
    const cx = scene.plot.x + scene.plot.width / 2;
    const cy = scene.plot.y + scene.plot.height / 2;
    const radius = Math.min(scene.plot.width, scene.plot.height) / 2 - 2;
    for (const [index, slice] of slices.entries()) {
      const path = String(slice.attrs.d);
      expect(path).not.toMatch(/NaN|Infinity/);
      const numbers = path.match(/[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gi)?.map(Number);
      expect(numbers).toHaveLength(11);
      const sign = index === 0 ? 1 : -1;
      expect(numbers?.[2]).toBeCloseTo(cx + sign * radius);
      expect(numbers?.[3]).toBeCloseTo(cy);
      expect(numbers?.[9]).toBeCloseTo(cx - sign * radius);
      expect(numbers?.[10]).toBeCloseTo(cy);
    }
    const positive = pie([2, 1]).elements.filter((element) => element.role === 'slice');
    const negative = pie([-2, 1]).elements.filter((element) => element.role === 'slice');
    expect(negative).toEqual(positive);
    const firstNumbers = String(negative[0]?.attrs.d)
      .match(/[-+]?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/gi)
      ?.map(Number);
    expect(firstNumbers?.[7]).toBe(1); // The first slice spans 240°, larger than a semicircle.
    expect(firstNumbers?.[9]).toBeCloseTo(cx - (radius * Math.sqrt(3)) / 2);
    expect(firstNumbers?.[10]).toBeCloseTo(cy + radius / 2);
  });

  it('filters before geometry and preserves original category/series identities', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B', 'C'],
      categoryIds: ['a', 'b', 'c'],
      series: [
        { id: 'hidden-series', values: [1, 2, 3] },
        { id: 'visible-series', values: [4, 5, 6] },
      ],
    };
    const original = structuredClone(data);
    const scene = sceneFor(
      {
        kind: 'column',
        palette: ['#first', '#second', '#third'],
        filter: { hiddenSeriesIds: ['hidden-series'], hiddenCategoryIds: ['b'] },
      },
      data,
    );
    const bars = barElements(scene);
    expect(bars).toHaveLength(2);
    expect(bars.map((bar) => [bar.seriesIndex, bar.pointIndex, bar.seriesId, bar.pointId])).toEqual(
      [
        [1, 0, 'visible-series', 'a'],
        [1, 2, 'visible-series', 'c'],
      ],
    );
    expect(bars.map((bar) => bar.attrs.fill)).toEqual(['#second', '#second']);
    expect(data).toEqual(original);
  });

  it('filters XY points by explicit or fallback identities without renumbering marks', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 's0',
          points: [
            { id: 'keep', x: 1, y: 1 },
            { x: 2, y: 2 },
            { id: 'hide', x: 3, y: 3 },
          ],
        },
        { id: 'hidden-series', points: [{ x: 4, y: 4 }] },
      ],
    };
    const scene = sceneFor(
      {
        kind: 'scatter',
        filter: { hiddenCategoryIds: ['point-1', 'hide'], hiddenSeriesIds: ['hidden-series'] },
      },
      data,
    );
    const points = scene.elements.filter((element) => element.role === 'point');
    expect(points).toHaveLength(1);
    expect(points[0]).toMatchObject({
      seriesIndex: 0,
      pointIndex: 0,
      seriesId: 's0',
      pointId: 'keep',
    });
  });

  it('uses visible-series denominators for percent labels and stable point style ids', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      categoryIds: ['a'],
      series: [
        { id: 'kept', values: [1] },
        { id: 'filtered', values: [3] },
      ],
    };
    const scene = sceneFor(
      {
        kind: 'column',
        grouping: 'percent-stacked',
        filter: { hiddenSeriesIds: ['filtered'] },
        dataLabels: { visible: true, showValue: false, showPercentage: true },
      },
      data,
    );
    expect(
      scene.elements.filter((element) => element.role === 'data-label').map((e) => e.text),
    ).toEqual(['100%']);
    expect(scene.elements.find((element) => element.role === 'data-label')?.seriesId).toBe('kept');
  });

  it('does not invent value labels when every label field is explicitly disabled', () => {
    const scene = sceneFor(
      {
        kind: 'column',
        dataLabels: {
          visible: true,
          showSeriesName: false,
          showCategoryName: false,
          showValue: false,
          showPercentage: false,
          showLegendKey: false,
        },
      },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [1] }] },
    );
    expect(scene.elements.some((element) => element.role === 'data-label')).toBe(false);
  });

  it('applies point > series > palette styles and keeps legend colors aligned', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      categoryIds: ['a', 'b'],
      series: [{ id: 'sales', name: 'Sales', values: [1, 2] }],
    };
    const scene = sceneFor(
      {
        kind: 'column',
        palette: ['#palette'],
        legend: { visible: true },
        seriesStyles: [
          {
            seriesId: 'sales',
            fill: { color: '#series', transparency: 75 },
            line: { color: '#edge', transparency: 25, width: 3, dash: 'dash-dot' },
            points: [{ pointId: 'a', fill: { color: '#point', transparency: 50 } }],
          },
        ],
      },
      data,
    );
    const bars = barElements(scene);
    expect(bars.map((bar) => bar.attrs.fill)).toEqual(['#point', '#series']);
    expect(bars[0]?.attrs['fill-opacity']).toBe(0.5);
    expect(bars[1]?.attrs['fill-opacity']).toBe(0.25);
    expect(bars[0]?.attrs.stroke).toBe('#edge');
    expect(bars[0]?.attrs['stroke-opacity']).toBe(0.75);
    expect(bars[0]?.attrs['stroke-width']).toBe(3);
    expect(bars[0]?.attrs['stroke-dasharray']).toBe('6 3 1 3');
    const legend = scene.elements.find(
      (element) => element.role === 'legend' && element.tag === 'rect',
    );
    expect(legend?.attrs.fill).toBe('#series');
  });

  it('inherits a point fill into its marker unless the point marker sets a fill', () => {
    const scene = sceneFor(
      {
        kind: 'line',
        palette: ['#palette'],
        seriesStyles: [
          {
            seriesId: 'sales',
            marker: { fill: { color: '#series-marker' } },
            points: [
              { pointId: 'a', fill: { color: '#point-fill' } },
              {
                pointId: 'b',
                fill: { color: '#point-fill-ignored' },
                marker: { fill: { color: '#point-marker' } },
              },
            ],
          },
        ],
      },
      {
        domain: 'category',
        categories: ['A', 'B'],
        categoryIds: ['a', 'b'],
        series: [{ id: 'sales', values: [1, 2] }],
      },
    );
    const markers = scene.elements.filter((element) => element.role === 'point');
    expect(markers.map((marker) => marker.attrs.fill)).toEqual(['#point-fill', '#point-marker']);
  });

  it('draws a line or scatter series colour on its line, markers and key, without the fill transparency', () => {
    const line = sceneFor(
      {
        kind: 'line',
        palette: ['#palette'],
        legend: { visible: true },
        seriesStyles: [
          {
            seriesId: 'sales',
            fill: { color: '#series', transparency: 75 },
            marker: { visible: true },
          },
          { seriesId: 'costs', line: { color: '#edge' } },
        ],
      },
      {
        domain: 'category',
        categories: ['A', 'B'],
        series: [
          { id: 'sales', values: [1, 2] },
          { id: 'costs', values: [2, 1] },
        ],
      },
    );
    const lines = line.elements.filter((element) => element.role === 'line');
    expect(lines.map((element) => element.attrs.stroke)).toEqual(['#series', '#edge']);
    const markers = line.elements.filter(
      (element) => element.role === 'point' && element.seriesIndex === 0,
    );
    expect(markers.map((marker) => marker.attrs.fill)).toEqual(['#series', '#series']);
    expect(markers[0]?.attrs['fill-opacity']).toBeUndefined();
    const keys = line.elements.filter(
      (element) => element.role === 'legend' && element.tag === 'rect',
    );
    expect(keys.map((key) => key.attrs.fill)).toEqual(['#series', '#edge']);
    const scatter = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'marker',
        palette: ['#palette'],
        legend: { visible: true },
        seriesStyles: [{ seriesId: 'y', marker: { fill: { color: '#marker' } } }],
      },
      { domain: 'xy', series: [{ id: 'y', points: [{ id: 'point-0', x: 1, y: 2 }] }] },
    );
    expect(
      scatter.elements.find((element) => element.role === 'legend' && element.tag === 'rect')?.attrs
        .fill,
    ).toBe('#marker');
  });

  it.each([
    ['x', 'path', { d: 'M15,15 L25,25 M15,25 L25,15', fill: 'none', stroke: '#edge' }],
    [
      'star',
      'path',
      { d: 'M15,15 L25,25 M15,25 L25,15 M20,15 L20,25', fill: 'none', stroke: '#edge' },
    ],
    ['plus', 'path', { d: 'M15,20 L25,20 M20,15 L20,25', fill: 'none', stroke: '#edge' }],
    ['dot', 'rect', { x: 17.5, y: 19, width: 5, height: 2, fill: '#series' }],
    ['dash', 'rect', { x: 15, y: 19, width: 10, height: 2, fill: '#series' }],
    ['picture', 'circle', { cx: 20, cy: 20, r: 5, fill: '#series' }],
  ] as const)('draws a %s marker as Excel does at its size', (shape, tag, attrs) => {
    const marker = (strokeColor?: string) =>
      sceneFor(
        {
          kind: 'scatter',
          scatterStyle: 'marker',
          palette: ['#series'],
          xAxis: { min: 0, max: 2 },
          yAxis: { min: 0, max: 2 },
          seriesStyles: [
            {
              seriesId: 's',
              marker: {
                visible: true,
                shape,
                size: 10,
                line: { width: 2, ...(strokeColor ? { color: strokeColor } : {}) },
              },
            },
          ],
        },
        { domain: 'xy', series: [{ id: 's', points: [{ id: 'point-0', x: 1, y: 1 }] }] },
      ).elements.find((element) => element.role === 'point');
    const drawn = marker('#edge');
    expect(drawn?.tag).toBe(tag);
    // Coordinates are compared relative to the marker centre, moved to (20, 20).
    const centre = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'marker',
        xAxis: { min: 0, max: 2 },
        yAxis: { min: 0, max: 2 },
      },
      { domain: 'xy', series: [{ id: 's', points: [{ id: 'point-0', x: 1, y: 1 }] }] },
    ).elements.find((element) => element.role === 'point')?.attrs;
    const dx = Number(centre?.cx) - 20;
    const dy = Number(centre?.cy) - 20;
    const moved = Object.fromEntries(
      Object.keys(attrs).map((key) => {
        const value = drawn?.attrs[key];
        if (key === 'd') {
          return [
            key,
            String(value).replace(
              /(-?[\d.]+),(-?[\d.]+)/g,
              (_match, px, py) => `${Number(px) - dx},${Number(py) - dy}`,
            ),
          ];
        }
        if (key === 'x' || key === 'cx') return [key, Number(value) - dx];
        if (key === 'y' || key === 'cy') return [key, Number(value) - dy];
        return [key, value];
      }),
    );
    expect(moved).toEqual(attrs);
    if (tag === 'path') {
      expect(drawn?.attrs['stroke-width']).toBe(2);
      // Without an outline colour the line shapes take the series colour.
      expect(marker()?.attrs.stroke).toBe('#series');
    }
  });

  it('colours series without a palette in Excel’s automatic cycle of the theme accents', () => {
    const slices = (accents?: readonly string[]) =>
      layoutChart({
        spec: { kind: 'pie' },
        data: {
          domain: 'category',
          categories: Array.from({ length: 13 }, (_, index) => `c${index}`),
          series: [{ id: 's', values: Array.from({ length: 13 }, () => 1) }],
        },
        width: 320,
        height: 132,
        ...(accents ? { accents } : {}),
      })
        .elements.filter((element) => element.role === 'slice')
        .map((slice) => slice.attrs.fill);
    // Default Office theme: accent1, ..., then point 10 is accent5 at lumMod 60000, 12 accent1 at 80000 + 20000.
    const automatic = slices();
    expect(automatic.slice(0, 6)).toEqual([
      '#156082',
      '#e97132',
      '#196b24',
      '#0f9ed5',
      '#a02b93',
      '#4ea72e',
    ]);
    expect(automatic[10]).toBe('#601a58');
    expect(automatic[12]).toBe('#1f8ebe');
    // A workbook theme's accents replace the default theme's.
    const theme = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
    expect(slices(theme).slice(0, 6)).toEqual(theme);
    expect(slices(theme)[6]).toBe('#254278');
  });

  it('draws a smoothed line as Catmull-Rom Béziers with tension 1/6', () => {
    const lineOf = (smooth: boolean) =>
      sceneFor(
        { kind: 'line', seriesStyles: [{ seriesId: 's', smooth }] },
        {
          domain: 'category',
          categories: ['A', 'B', 'C'],
          series: [{ id: 's', values: [1, 3, 2] }],
        },
      ).elements.find((element) => element.role === 'line');
    const straight = lineOf(false);
    expect(straight?.tag).toBe('polyline');
    const points = String(straight?.attrs.points)
      .split(' ')
      .map((pair) => pair.split(',').map(Number) as [number, number]);
    const [p0, p1, p2] = points as [[number, number], [number, number], [number, number]];
    const smooth = lineOf(true);
    expect(smooth?.tag).toBe('path');
    expect(smooth?.attrs.d).toBe(
      [
        `M${p0[0]},${p0[1]}`,
        `C${p0[0] + (p1[0] - p0[0]) / 6},${p0[1] + (p1[1] - p0[1]) / 6} ${p1[0] - (p2[0] - p0[0]) / 6},${p1[1] - (p2[1] - p0[1]) / 6} ${p1[0]},${p1[1]}`,
        `C${p1[0] + (p2[0] - p0[0]) / 6},${p1[1] + (p2[1] - p0[1]) / 6} ${p2[0] - (p2[0] - p1[0]) / 6},${p2[1] - (p2[1] - p1[1]) / 6} ${p2[0]},${p2[1]}`,
      ].join(' '),
    );
    const scatter = sceneFor(
      {
        kind: 'scatter',
        scatterStyle: 'line',
        seriesStyles: [{ seriesId: 's', smooth: true }],
      },
      {
        domain: 'xy',
        series: [
          {
            id: 's',
            points: [
              { id: 'point-0', x: 1, y: 1 },
              { id: 'point-1', x: 2, y: 3 },
              { id: 'point-2', x: 3, y: 2 },
            ],
          },
        ],
      },
    ).elements.filter((element) => element.role === 'line');
    // One smoothed path through the joined segments, not one line per segment.
    expect(scatter.map((element) => element.tag)).toEqual(['path']);
    expect(String(scatter[0]?.attrs.d).match(/C/g)).toHaveLength(2);
  });

  it('formats a scatter chart by its x and y number formats unless the spec sets one', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 's',
          formatCode: '0.000',
          xFormatCode: '0.00',
          points: [
            { id: 'point-0', x: 1, y: 2 },
            { id: 'point-1', x: 3, y: 4 },
          ],
        },
      ],
    };
    const texts = (spec: ChartSpec, role: 'axis-label' | 'data-label') =>
      sceneAt(spec, data, 640, 320)
        .elements.filter((element) => element.role === role)
        // A data label near the edge is truncated; its accessible name keeps the whole text.
        .map((element) => String(element.attrs['aria-label'] ?? element.text ?? ''));
    const axes = texts({ kind: 'scatter' }, 'axis-label');
    const xLabel = /^\d+\.\d\d$/;
    const yLabel = /^\d+\.\d\d\d$/;
    expect(axes.some((text) => xLabel.test(text))).toBe(true);
    expect(axes.some((text) => yLabel.test(text))).toBe(true);
    expect(axes.every((text) => xLabel.test(text) || yLabel.test(text))).toBe(true);
    // An axis format of the spec wins over the data's.
    const own = texts({ kind: 'scatter', xAxis: { numberFormat: '0' } }, 'axis-label');
    expect(own.some((text) => /^\d+$/.test(text))).toBe(true);
    expect(own.some((text) => xLabel.test(text))).toBe(false);
    const labels = { visible: true, showCategoryName: true, showValue: true };
    expect(texts({ kind: 'scatter', dataLabels: labels }, 'data-label')).toEqual([
      '1.00, 2.000',
      '3.00, 4.000',
    ]);
    expect(
      texts({ kind: 'scatter', dataLabels: { ...labels, numberFormat: '0' } }, 'data-label'),
    ).toEqual(['1, 2', '3, 4']);
  });

  it('draws an automatic title as the only series name, else the localised chart title', () => {
    const titleOf = (series: number, chartTitle?: string) =>
      layoutChart({
        spec: { kind: 'column', title: { auto: true } },
        data: {
          domain: 'category',
          categories: ['A'],
          series: Array.from({ length: series }, (_, index) => ({
            id: `s${index}`,
            name: `Series ${index}`,
            values: [1],
          })),
        },
        width: 320,
        height: 132,
        ...(chartTitle ? { chartTitle } : {}),
      }).elements.find((element) => element.role === 'title')?.text;
    expect(titleOf(1)).toBe('Series 0');
    expect(titleOf(2)).toBe('Chart Title');
    expect(titleOf(2, 'グラフ タイトル')).toBe('グラフ タイトル');
    expect(chartTitleText('', { domain: 'category', categories: [], series: [] })).toBeUndefined();
    expect(chartTitleText(undefined, { domain: 'category', categories: [], series: [] })).toBe(
      undefined,
    );
  });

  it('draws an automatic axis title as the localised default', () => {
    const titles = (axisTitle?: string) =>
      layoutChart({
        spec: { kind: 'column', xAxis: { title: { auto: true } }, yAxis: { title: 'Units' } },
        data: {
          domain: 'category',
          categories: ['A'],
          series: [{ id: 's0', name: 'S', values: [1] }],
        },
        width: 320,
        height: 132,
        ...(axisTitle ? { axisTitle } : {}),
      })
        .elements.filter((element) => element.role === 'axis-title')
        .map((element) => element.text);
    expect(titles()).toEqual(['Axis Title', 'Units']);
    expect(titles('軸ラベル')).toEqual(['軸ラベル', 'Units']);
    expect(axisTitleText('')).toBeUndefined();
    expect(axisTitleText(undefined)).toBeUndefined();
  });

  it('supports vary-colors defaults, marker shapes, and separate marker transparency', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      categoryIds: ['a', 'b'],
      series: [{ id: 's', values: [1, 2] }],
    };
    const line = sceneFor(
      {
        kind: 'line',
        palette: ['#one', '#two'],
        varyColors: true,
        seriesStyles: [
          {
            seriesId: 's',
            marker: {
              shape: 'diamond',
              size: 12,
              fill: { transparency: 20 },
              line: { color: '#m', transparency: 60 },
            },
          },
        ],
      },
      data,
    );
    const markers = line.elements.filter((element) => element.role === 'point');
    expect(markers.map((marker) => marker.tag)).toEqual(['polygon', 'polygon']);
    expect(markers.map((marker) => marker.attrs.fill)).toEqual(['#one', '#two']);
    expect(markers[0]?.attrs['fill-opacity']).toBe(0.8);
    expect(markers[0]?.attrs['stroke-opacity']).toBe(0.4);
    const pie = sceneFor(
      { kind: 'pie', palette: ['#one', '#two'], varyColors: false },
      { domain: 'category', categories: ['A', 'B'], series: [{ id: 's', values: [1, 2] }] },
    );
    expect(
      pie.elements.filter((element) => element.role === 'slice').map((slice) => slice.attrs.fill),
    ).toEqual(['#one', '#one']);
  });

  it('formats labels through the host callback and keeps negative positions finite', () => {
    const scene = sceneAt(
      {
        kind: 'bar',
        dataLabels: {
          visible: true,
          showValue: true,
          showCategoryName: true,
          showLegendKey: true,
          separator: '\n',
          numberFormat: '0.00',
          position: 'outside-end',
          fontSize: 500,
        },
      },
      {
        domain: 'category',
        categories: ['Long category name'],
        series: [{ id: 's', name: 'Sales', values: [-1.25] }],
      },
      180,
      80,
    );
    const label = scene.elements.find((element) => element.role === 'data-label');
    expect(label?.text).toContain('…');
    expect(label?.attrs['white-space']).toBe('pre');
    expect(label?.attrs['aria-label']).toBe('Long category name\n-1.25');
    expect(Number(label?.attrs['font-size'])).toBeLessThan(72);
    expect(Number(label?.attrs['font-size'])).toBeGreaterThanOrEqual(6);
    expect(scene.elements.some((element) => element.role === 'data-label-key')).toBe(true);
    expect(
      scene.elements.every((element) =>
        Object.values(element.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);

    const formatted = sceneFor(
      {
        kind: 'column',
        dataLabels: { visible: true, showValue: true, numberFormat: 'host-code' },
      },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [1234.5] }] },
    );
    const formattedLabel = formatted.elements.find((element) => element.role === 'data-label');
    expect(formattedLabel?.text).toBe('1,234.5');
    const custom = layoutChart({
      spec: {
        kind: 'column',
        dataLabels: { visible: true, showValue: true, numberFormat: 'host-code' },
      },
      data: { domain: 'category', categories: ['A'], series: [{ id: 's', values: [1234.5] }] },
      width: 320,
      height: 132,
      locale: 'ja-JP',
      formatValue: (value, formatCode, locale) => `${locale}:${formatCode}:${value}`,
    });
    expect(custom.elements.find((element) => element.role === 'data-label')?.text).toBe(
      'ja-JP:host-code:1234.5',
    );
  });

  it('keeps Excel label positions distinct for positive and negative bars', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: [{ id: 's', values: [-2] }],
    };
    const atBase = sceneFor(
      { kind: 'bar', dataLabels: { visible: true, position: 'inside-base' } },
      data,
    );
    const atEnd = sceneFor(
      { kind: 'bar', dataLabels: { visible: true, position: 'inside-end' } },
      data,
    );
    const outside = sceneFor(
      { kind: 'bar', dataLabels: { visible: true, position: 'outside-end' } },
      data,
    );
    const baseX = Number(atBase.elements.find((element) => element.role === 'data-label')?.attrs.x);
    const endX = Number(atEnd.elements.find((element) => element.role === 'data-label')?.attrs.x);
    const outsideX = Number(
      outside.elements.find((element) => element.role === 'data-label')?.attrs.x,
    );
    expect(baseX).not.toBe(endX);
    expect(outsideX).not.toBe(endX);
  });

  it('anchors stacked inside-base labels at each positive and negative segment base', () => {
    const scene = sceneFor(
      {
        kind: 'column',
        grouping: 'stacked',
        dataLabels: { visible: true, position: 'inside-base' },
      },
      {
        domain: 'category',
        categories: ['A'],
        series: [
          { id: 'positive-first', values: [2] },
          { id: 'positive-second', values: [3] },
          { id: 'negative-first', values: [-2] },
          { id: 'negative-second', values: [-3] },
        ],
      },
    );
    const labels = scene.elements.filter((element) => element.role === 'data-label');
    const labelFor = (seriesId: string) =>
      labels.find((element) => element.seriesId === seriesId) as
        | (typeof labels)[number]
        | undefined;
    const positiveFirst = Number(labelFor('positive-first')?.attrs.y);
    const positiveSecond = Number(labelFor('positive-second')?.attrs.y);
    const negativeFirst = Number(labelFor('negative-first')?.attrs.y);
    const negativeSecond = Number(labelFor('negative-second')?.attrs.y);
    expect(positiveSecond).not.toBe(positiveFirst);
    expect(negativeSecond).not.toBe(negativeFirst);
    expect(
      [positiveFirst, positiveSecond, negativeFirst, negativeSecond].every(Number.isFinite),
    ).toBe(true);
  });

  it('uses Excel data-label position matrices and chart-specific defaults', () => {
    expect(chartLabelPositions({ kind: 'pie' })).toEqual([
      'best-fit',
      'outside-end',
      'inside-end',
      'center',
    ]);
    expect(chartLabelPositions({ kind: 'column' })).toEqual([
      'inside-base',
      'inside-end',
      'outside-end',
      'center',
    ]);
    expect(chartLabelPositions({ kind: 'column', grouping: 'stacked' })).toEqual([
      'inside-base',
      'inside-end',
      'center',
    ]);
    expect(chartLabelPositions({ kind: 'line' })).toEqual([
      'left',
      'right',
      'below',
      'above',
      'center',
    ]);
    expect(chartLabelPositions({ kind: 'area' })).toEqual([]);
    expect(defaultChartLabelPosition({ kind: 'pie' })).toBe('best-fit');
    expect(chartLabelPositions({ kind: 'doughnut' })).toEqual([]);
    expect(defaultChartLabelPosition({ kind: 'doughnut' })).toBe('center');
    expect(defaultChartLabelPosition({ kind: 'column' })).toBe('outside-end');
    expect(defaultChartLabelPosition({ kind: 'column', grouping: 'percent-stacked' })).toBe(
      'center',
    );
    expect(defaultChartLabelPosition({ kind: 'line' })).toBe('right');
    expect(defaultChartLabelPosition({ kind: 'area' })).toBe('center');
    expect(resolveChartLabelPosition({ kind: 'column' }, 'best-fit')).toBe('outside-end');
    expect(resolveChartLabelPosition({ kind: 'column' }, 'left')).toBe('outside-end');
    expect(resolveChartLabelPosition({ kind: 'bar' }, 'right')).toBe('outside-end');
    expect(resolveChartLabelPosition({ kind: 'line' }, 'inside-end')).toBe('right');
    expect(resolveChartLabelPosition({ kind: 'pie' }, 'above')).toBe('best-fit');
    expect(resolveChartLabelPosition({ kind: 'area' }, 'above')).toBe('center');

    const data: ChartData = {
      domain: 'category',
      categories: ['A'],
      series: [{ id: 's', values: [2] }],
    };
    const labelAttrs = (
      kind: ChartSpec['kind'],
      position: NonNullable<ChartSpec['dataLabels']>['position'],
    ) => {
      const label = sceneFor({ kind, dataLabels: { visible: true, position } }, data).elements.find(
        (element) => element.role === 'data-label',
      );
      return {
        x: label?.attrs.x,
        y: label?.attrs.y,
        anchor: label?.attrs['text-anchor'],
      };
    };
    expect(labelAttrs('column', 'left')).toEqual(labelAttrs('column', 'outside-end'));
    expect(labelAttrs('column', 'best-fit')).toEqual(labelAttrs('column', 'outside-end'));
    expect(labelAttrs('bar', 'right')).toEqual(labelAttrs('bar', 'outside-end'));
    expect(labelAttrs('line', 'inside-end')).toEqual(labelAttrs('line', 'right'));
    expect(labelAttrs('pie', 'above')).toEqual(labelAttrs('pie', 'best-fit'));
    expect(labelAttrs('pie', 'inside-end')).not.toEqual(labelAttrs('pie', 'best-fit'));
    expect(labelAttrs('doughnut', 'outside-end')).toEqual(labelAttrs('doughnut', 'center'));
  });

  it('centers scatter labels at the point and can label formatted X values', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [{ id: 's', points: [{ id: 'p', x: 1.25, y: 2.5 }] }],
    };
    const centered = sceneFor(
      {
        kind: 'scatter',
        dataLabels: { visible: true, showCategoryName: true, position: 'center' },
      },
      data,
    );
    const point = centered.elements.find((element) => element.role === 'point');
    const centerLabel = centered.elements.find((element) => element.role === 'data-label');
    expect(centerLabel?.attrs.x).toBe(point?.attrs.cx);
    expect(centerLabel?.attrs.y).toBe(point?.attrs.cy);

    const formatted = layoutChart({
      spec: {
        kind: 'scatter',
        dataLabels: { visible: true, showCategoryName: true, numberFormat: '0.0' },
      },
      data,
      width: 320,
      height: 132,
      locale: 'ja-JP',
      formatValue: (value, formatCode, locale) => `${locale}:${formatCode}:${value}`,
    });
    expect(formatted.elements.find((element) => element.role === 'data-label')?.text).toBe(
      'ja-JP:0.0:1.25',
    );
  });

  it('keeps oversized CJK labels and swatches inside the scene bounds', () => {
    const width = 320;
    const height = 132;
    const scene = sceneAt(
      {
        kind: 'column',
        legend: { visible: false },
        dataLabels: {
          visible: true,
          showSeriesName: true,
          showCategoryName: true,
          showValue: true,
          showLegendKey: true,
          position: 'outside-end',
          fontSize: 72,
        },
      },
      {
        domain: 'category',
        categories: ['2026年10月'],
        series: [{ id: 's', name: '日本語の系列名', values: [10] }],
      },
      width,
      height,
    );
    const label = scene.elements.find((element) => element.role === 'data-label');
    const key = scene.elements.find((element) => element.role === 'data-label-key');
    expect(label?.text).toContain('…');
    expect(label?.attrs['aria-label']).toContain('日本語の系列名');
    const fontSize = Number(label?.attrs['font-size']);
    const baseline = Number(label?.attrs.y);
    expect(fontSize).toBeLessThanOrEqual(72);
    expect(baseline - fontSize).toBeGreaterThanOrEqual(0);
    expect(baseline + fontSize * 0.25).toBeLessThanOrEqual(height);
    expect(Number(label?.attrs.x)).toBeGreaterThanOrEqual(0);
    expect(Number(label?.attrs.x)).toBeLessThanOrEqual(width);
    expect(key?.tag).toBe('rect');
    const keyX = Number(key?.attrs.x);
    const keyY = Number(key?.attrs.y);
    const keyWidth = Number(key?.attrs.width);
    const keyHeight = Number(key?.attrs.height);
    expect(keyX).toBeGreaterThanOrEqual(0);
    expect(keyY).toBeGreaterThanOrEqual(0);
    expect(keyX + keyWidth).toBeLessThanOrEqual(width);
    expect(keyY + keyHeight).toBeLessThanOrEqual(height);
    expect(key?.attrs['aria-label']).toBe('s');
  });

  it.each(['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'] as const)(
    'emits finite labels for %s charts',
    (kind) => {
      const data: ChartData =
        kind === 'scatter'
          ? { domain: 'xy', series: [{ id: 's', points: [{ id: 'p', x: 1, y: 2 }] }] }
          : { domain: 'category', categories: ['A'], series: [{ id: 's', values: [2] }] };
      const scene = sceneFor({ kind, dataLabels: { visible: true, showValue: true } }, data);
      expect(scene.elements.some((element) => element.role === 'data-label')).toBe(true);
      expect(
        scene.elements
          .filter((element) => element.role === 'data-label')
          .every((element) =>
            Object.values(element.attrs).every(
              (value) => typeof value !== 'number' || Number.isFinite(value),
            ),
          ),
      ).toBe(true);
    },
  );

  it('reverses numeric projections while retaining ascending tick values', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 's',
          points: [
            { id: 'low', x: 0, y: 0 },
            { id: 'high', x: 10, y: 10 },
          ],
        },
      ],
    };
    const normal = sceneFor(
      { kind: 'scatter', xAxis: { min: 0, max: 10 }, yAxis: { min: 0, max: 10 } },
      data,
    );
    const reversed = sceneFor(
      {
        kind: 'scatter',
        xAxis: { min: 0, max: 10, reverseOrder: true },
        yAxis: { min: 0, max: 10, reverseOrder: true },
      },
      data,
    );
    const labelsFor = (scene: ReturnType<typeof sceneFor>, axis: 'x' | 'y') =>
      scene.elements
        .filter(
          (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === axis,
        )
        .map((element) => element.text);
    expect(labelsFor(reversed, 'x')).toEqual(labelsFor(normal, 'x'));
    expect(labelsFor(reversed, 'y')).toEqual(labelsFor(normal, 'y'));
    const normalPoints = normal.elements.filter((element) => element.role === 'point');
    const reversedPoints = reversed.elements.filter((element) => element.role === 'point');
    expect(Number(reversedPoints[0]?.attrs.cx)).toBeCloseTo(
      reversed.plot.x + reversed.plot.width - (Number(normalPoints[0]?.attrs.cx) - normal.plot.x),
    );
    expect(Number(reversedPoints[0]?.attrs.cy)).toBeCloseTo(
      reversed.plot.y + reversed.plot.height - (Number(normalPoints[0]?.attrs.cy) - normal.plot.y),
    );
  });

  it('reverses category bands without mirroring clustered series offsets or identities', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      categoryIds: ['a', 'b'],
      series: [
        { id: 'first', values: [1, 2] },
        { id: 'second', values: [3, 4] },
      ],
    };
    const normal = sceneFor({ kind: 'column' }, data);
    const reversed = sceneFor({ kind: 'column', xAxis: { reverseOrder: true } }, data);
    const categoryLabels = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements.filter((element) => element.role === 'category-label');
    expect(categoryLabels(reversed).map((label) => label.text)).toEqual(
      categoryLabels(normal).map((label) => label.text),
    );
    expect(categoryLabels(reversed).map((label) => label.attrs.x)).toEqual([
      categoryLabels(normal)[1]?.attrs.x,
      categoryLabels(normal)[0]?.attrs.x,
    ]);
    const normalBars = normal.elements.filter((element) => element.role === 'bar');
    const reversedBars = reversed.elements.filter((element) => element.role === 'bar');
    for (const bar of normalBars) {
      const reversedBar = reversedBars.find(
        (candidate) => candidate.seriesId === bar.seriesId && candidate.pointId === bar.pointId,
      );
      expect(reversedBar).toBeDefined();
      const normalCategory = bar.pointId === 'a' ? 0 : 1;
      const reversedCategory = 1 - normalCategory;
      const slot = normal.plot.width / 2;
      expect(Number(reversedBar?.attrs.x)).toBeCloseTo(
        reversed.plot.x +
          reversedCategory * slot +
          (Number(bar.attrs.x) - normal.plot.x - normalCategory * slot),
      );
    }
  });

  it('keeps bar and column label directions screen-relative for both signs and zero', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['positive', 'negative', 'zero'],
      categoryIds: ['positive', 'negative', 'zero'],
      series: [{ id: 's', values: [2, -2, 0] }],
    };
    const labelFor = (scene: ReturnType<typeof sceneFor>, pointId: string) =>
      scene.elements.find(
        (element) => element.role === 'data-label' && element.pointId === pointId,
      );
    const barFor = (scene: ReturnType<typeof sceneFor>, pointId: string) =>
      scene.elements.find((element) => element.role === 'bar' && element.pointId === pointId);
    for (const [kind, reverseX, reverseY] of [
      ['bar', false, false],
      ['bar', true, false],
      ['column', false, false],
      ['column', false, true],
    ] as const) {
      const scene = sceneFor(
        {
          kind,
          xAxis: { reverseOrder: reverseX },
          yAxis: { reverseOrder: reverseY },
          dataLabels: { visible: true, position: 'outside-end' },
        },
        data,
      );
      const positiveBar = barFor(scene, 'positive');
      const negativeBar = barFor(scene, 'negative');
      const zeroBar = barFor(scene, 'zero');
      const positiveLabel = labelFor(scene, 'positive');
      const negativeLabel = labelFor(scene, 'negative');
      const zeroLabel = labelFor(scene, 'zero');
      expect(positiveBar).toBeDefined();
      expect(negativeBar).toBeDefined();
      expect(zeroBar).toBeDefined();
      expect(positiveLabel).toBeDefined();
      expect(negativeLabel).toBeDefined();
      expect(zeroLabel).toBeDefined();
      if (kind === 'bar') {
        const positiveEnd = Number(positiveBar?.attrs.x) + Number(positiveBar?.attrs.width);
        const negativeEnd = Number(negativeBar?.attrs.x);
        const zeroEnd = Number(zeroBar?.attrs.x);
        const positiveGoesRight = !reverseX;
        expect(
          positiveGoesRight
            ? Number(positiveLabel?.attrs.x) >= positiveEnd
            : Number(positiveLabel?.attrs.x) <= Number(positiveBar?.attrs.x),
        ).toBe(true);
        expect(
          positiveGoesRight
            ? Number(negativeLabel?.attrs.x) <= Number(negativeBar?.attrs.x)
            : Number(negativeLabel?.attrs.x) >= negativeEnd,
        ).toBe(true);
        expect(
          positiveGoesRight
            ? Number(zeroLabel?.attrs.x) >= zeroEnd
            : Number(zeroLabel?.attrs.x) <= zeroEnd,
        ).toBe(true);
      } else {
        const positiveTop = Number(positiveBar?.attrs.y);
        const positiveBottom = positiveTop + Number(positiveBar?.attrs.height);
        const negativeTop = Number(negativeBar?.attrs.y);
        const negativeBottom = negativeTop + Number(negativeBar?.attrs.height);
        const zeroY = Number(zeroBar?.attrs.y);
        const positiveGoesUp = !reverseY;
        expect(
          positiveGoesUp
            ? Number(positiveLabel?.attrs.y) <= positiveTop
            : Number(positiveLabel?.attrs.y) >= positiveBottom,
        ).toBe(true);
        expect(
          positiveGoesUp
            ? Number(negativeLabel?.attrs.y) >= negativeBottom
            : Number(negativeLabel?.attrs.y) <= negativeTop,
        ).toBe(true);
        expect(
          positiveGoesUp
            ? Number(zeroLabel?.attrs.y) <= zeroY
            : Number(zeroLabel?.attrs.y) >= zeroY,
        ).toBe(true);
      }
    }
  });

  it('keeps inside-base and inside-end labels near the requested screen endpoint', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['positive', 'negative'],
      categoryIds: ['positive', 'negative'],
      series: [{ id: 's', values: [4, -4] }],
    };
    const barFor = (scene: ReturnType<typeof sceneAt>, pointId: string) =>
      scene.elements.find((element) => element.role === 'bar' && element.pointId === pointId);
    const labelFor = (scene: ReturnType<typeof sceneAt>, pointId: string) =>
      scene.elements.find(
        (element) => element.role === 'data-label' && element.pointId === pointId,
      );
    for (const [kind, reverse, position] of [
      ['bar', false, 'inside-base'],
      ['bar', false, 'inside-end'],
      ['bar', true, 'inside-base'],
      ['bar', true, 'inside-end'],
      ['column', false, 'inside-base'],
      ['column', false, 'inside-end'],
      ['column', true, 'inside-base'],
      ['column', true, 'inside-end'],
    ] as const) {
      const scene = sceneAt(
        {
          kind,
          xAxis: kind === 'bar' ? { min: -10, max: 10, reverseOrder: reverse } : undefined,
          yAxis: kind === 'column' ? { min: -10, max: 10, reverseOrder: reverse } : undefined,
          dataLabels: { visible: true, showValue: true, position },
        },
        data,
        640,
        240,
      );
      for (const [pointId, value] of [
        ['positive', 4],
        ['negative', -4],
      ] as const) {
        const bar = barFor(scene, pointId);
        const label = labelFor(scene, pointId);
        expect(bar).toBeDefined();
        expect(label).toBeDefined();
        const labelCoordinate = Number(label?.attrs[kind === 'bar' ? 'x' : 'y']);
        const start = kind === 'bar' ? Number(bar?.attrs.x) : Number(bar?.attrs.y);
        const size = kind === 'bar' ? Number(bar?.attrs.width) : Number(bar?.attrs.height);
        const end = start + size;
        const coordinate = labelCoordinate;
        expect(coordinate).toBeGreaterThanOrEqual(start);
        expect(coordinate).toBeLessThanOrEqual(end);
        const reverseDirection = reverse ? -1 : 1;
        const screenDirection =
          kind === 'bar' ? value * reverseDirection : -value * reverseDirection;
        const base = screenDirection > 0 ? start : end;
        const endpoint = screenDirection > 0 ? end : start;
        const baseDistance = Math.abs(coordinate - base);
        const endpointDistance = Math.abs(coordinate - endpoint);
        if (position === 'inside-base') {
          expect(baseDistance).toBeLessThan(endpointDistance);
        } else {
          expect(endpointDistance).toBeLessThan(baseDistance);
        }
      }
    }
  });

  it('accepts only literal true for reverseOrder and ignores it on circular charts', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [{ id: 's', values: [1, 2] }],
    };
    const baseline = sceneFor({ kind: 'column' }, data);
    const falseFlag = sceneFor({ kind: 'column', xAxis: { reverseOrder: false } }, data);
    const invalidFlag = sceneFor(
      { kind: 'column', xAxis: { reverseOrder: 1 as unknown as boolean } },
      data,
    );
    expect(falseFlag).toEqual(baseline);
    expect(invalidFlag).toEqual(baseline);

    const circularData: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      categoryIds: ['a', 'b'],
      series: [{ id: 's', values: [1, 2] }],
    };
    const pie = sceneFor(
      {
        kind: 'pie',
        xAxis: { reverseOrder: true },
        yAxis: { reverseOrder: true },
        filter: { hiddenCategoryIds: ['a'] },
        seriesStyles: [{ seriesId: 's', fill: { color: '#123456' } }],
      },
      circularData,
    );
    const pieBaseline = sceneFor(
      {
        kind: 'pie',
        filter: { hiddenCategoryIds: ['a'] },
        seriesStyles: [{ seriesId: 's', fill: { color: '#123456' } }],
      },
      circularData,
    );
    expect(pie).toEqual(pieBaseline);
    const doughnut = sceneFor(
      {
        kind: 'doughnut',
        xAxis: { reverseOrder: true },
        yAxis: { reverseOrder: true },
        filter: { hiddenCategoryIds: ['a'] },
        seriesStyles: [{ seriesId: 's', fill: { color: '#123456' } }],
      },
      circularData,
    );
    const doughnutBaseline = sceneFor(
      {
        kind: 'doughnut',
        filter: { hiddenCategoryIds: ['a'] },
        seriesStyles: [{ seriesId: 's', fill: { color: '#123456' } }],
      },
      circularData,
    );
    expect(doughnut).toEqual(doughnutBaseline);
  });

  it.each([
    ['line', 'percent-stacked', 'false', 'true'],
    ['area', 'stacked', 'true', 'true'],
    ['area', 'percent-stacked', 'true', 'false'],
    ['scatter', 'clustered', 'false', 'true'],
    ['line', 'stacked', 'false', 'false'],
    ['area', 'clustered', 'false', 'false'],
    ['column', 'percent-stacked', 'false', 'true'],
    ['bar', 'clustered', 'true', 'true'],
    ['column', 'stacked', 'true', 'true'],
    ['scatter', 'percent-stacked', 'true', 'false'],
    ['bar', 'stacked', 'false', 'false'],
    ['line', 'clustered', 'true', 'true'],
    ['bar', 'percent-stacked', 'true', 'false'],
    ['column', 'clustered', 'true', 'false'],
    ['scatter', 'stacked', 'false', 'true'],
  ] as const)(
    'keeps the reverse axis pairwise matrix geometric for %s/%s/%s/%s',
    (kind, grouping, reverseX, reverseY) => {
      const isScatter = kind === 'scatter';
      const data: ChartData = isScatter
        ? {
            domain: 'xy',
            series: [
              {
                id: 's',
                points: [
                  { id: 'p0', x: -2, y: -1 },
                  { id: 'p1', x: 0, y: 2 },
                  { id: 'p2', x: 3, y: 4 },
                ],
              },
            ],
          }
        : {
            domain: 'category',
            categories: ['A', 'B', 'C'],
            categoryIds: ['a', 'b', 'c'],
            series: [
              { id: 's', values: [-2, 0, 4] },
              { id: 't', values: [1, 3, -1] },
            ],
          };
      const reverseXFlag = reverseX === 'true';
      const reverseYFlag = reverseY === 'true';
      const axesFor = (xReverse: boolean, yReverse: boolean) =>
        kind === 'scatter'
          ? {
              xAxis: { min: -5, max: 5, reverseOrder: xReverse },
              yAxis: { min: -5, max: 5, reverseOrder: yReverse },
            }
          : kind === 'bar'
            ? {
                xAxis: { min: -5, max: 5, reverseOrder: xReverse },
                yAxis: { reverseOrder: yReverse },
              }
            : {
                xAxis: { reverseOrder: xReverse },
                yAxis: { min: -5, max: 5, reverseOrder: yReverse },
              };
      const normal = sceneFor({ kind, grouping, ...axesFor(false, false) }, data);
      const reversed = sceneFor({ kind, grouping, ...axesFor(reverseXFlag, reverseYFlag) }, data);
      const finite = (scene: ReturnType<typeof sceneFor>) =>
        scene.elements.every((element) =>
          Object.values(element.attrs).every(
            (value) => typeof value !== 'number' || Number.isFinite(value),
          ),
        );
      expect(finite(normal)).toBe(true);
      expect(finite(reversed)).toBe(true);

      const role = isScatter
        ? 'point'
        : kind === 'line'
          ? 'line'
          : kind === 'area'
            ? 'area'
            : 'bar';
      const normalMarks = normal.elements.filter((element) => element.role === role);
      const reversedMarks = reversed.elements.filter((element) => element.role === role);
      expect(normalMarks.length).toBeGreaterThan(0);
      expect(reversedMarks.length).toBe(normalMarks.length);
      const identity = (element: (typeof normalMarks)[number]) =>
        `${element.seriesId ?? ''}:${element.pointId ?? ''}:${element.seriesIndex ?? ''}:${element.pointIndex ?? ''}`;
      expect(reversedMarks.map(identity)).toEqual(normalMarks.map(identity));

      const mirror = (value: number, start: number, size: number): number =>
        start + size - (value - start);
      if (isScatter) {
        for (const normalPoint of normalMarks) {
          const reversedPoint = reversedMarks.find(
            (candidate) => candidate.pointId === normalPoint.pointId,
          );
          expect(reversedPoint).toBeDefined();
          const expectedX = reverseXFlag
            ? mirror(Number(normalPoint.attrs.cx), normal.plot.x, normal.plot.width)
            : Number(normalPoint.attrs.cx);
          const expectedY = reverseYFlag
            ? mirror(Number(normalPoint.attrs.cy), normal.plot.y, normal.plot.height)
            : Number(normalPoint.attrs.cy);
          expect(Number(reversedPoint?.attrs.cx)).toBeCloseTo(expectedX);
          expect(Number(reversedPoint?.attrs.cy)).toBeCloseTo(expectedY);
        }
      } else if (kind === 'bar' || kind === 'column') {
        const count = data.domain === 'category' ? data.categories.length : 1;
        for (const normalBar of normalMarks) {
          const reversedBar = reversedMarks.find(
            (candidate) =>
              candidate.seriesId === normalBar.seriesId && candidate.pointId === normalBar.pointId,
          );
          expect(reversedBar).toBeDefined();
          const category =
            data.domain === 'category'
              ? (data.categoryIds?.indexOf(normalBar.pointId ?? '') ?? -1)
              : -1;
          expect(category).toBeGreaterThanOrEqual(0);
          if (kind === 'bar') {
            const slot = normal.plot.height / count;
            const normalOffset = Number(normalBar.attrs.y) - normal.plot.y - category * slot;
            const expectedY =
              reversed.plot.y +
              (reverseYFlag ? count - 1 - category : category) * slot +
              normalOffset;
            const expectedX = reverseXFlag
              ? mirror(
                  Number(normalBar.attrs.x) + Number(normalBar.attrs.width),
                  normal.plot.x,
                  normal.plot.width,
                )
              : Number(normalBar.attrs.x);
            expect(Number(reversedBar?.attrs.x)).toBeCloseTo(expectedX);
            expect(Number(reversedBar?.attrs.y)).toBeCloseTo(expectedY);
          } else {
            const slot = normal.plot.width / count;
            const normalOffset = Number(normalBar.attrs.x) - normal.plot.x - category * slot;
            const expectedX =
              reversed.plot.x +
              (reverseXFlag ? count - 1 - category : category) * slot +
              normalOffset;
            const expectedY = reverseYFlag
              ? mirror(
                  Number(normalBar.attrs.y) + Number(normalBar.attrs.height),
                  normal.plot.y,
                  normal.plot.height,
                )
              : Number(normalBar.attrs.y);
            expect(Number(reversedBar?.attrs.x)).toBeCloseTo(expectedX);
            expect(Number(reversedBar?.attrs.y)).toBeCloseTo(expectedY);
          }
          expect(Number(reversedBar?.attrs.width)).toBeCloseTo(Number(normalBar.attrs.width));
          expect(Number(reversedBar?.attrs.height)).toBeCloseTo(Number(normalBar.attrs.height));
        }
      }

      if (kind === 'line' || kind === 'area') {
        const roles = kind === 'line' ? ['line'] : ['line', 'area'];
        for (const lineRole of roles) {
          const normalLines = normal.elements.filter((element) => element.role === lineRole);
          const reversedLines = reversed.elements.filter((element) => element.role === lineRole);
          expect(normalLines.length).toBeGreaterThan(0);
          expect(reversedLines.length).toBe(normalLines.length);
          expect(reversedLines.map(identity)).toEqual(normalLines.map(identity));
          for (let index = 0; index < normalLines.length; index += 1) {
            const normalPoints = pointPairs(String(normalLines[index]?.attrs.points));
            const reversedPoints = pointPairs(String(reversedLines[index]?.attrs.points));
            expect(reversedPoints).toHaveLength(normalPoints.length);
            for (let pointIndex = 0; pointIndex < normalPoints.length; pointIndex += 1) {
              const normalPoint = normalPoints[pointIndex] as [number, number];
              const reversedPoint = reversedPoints[pointIndex] as [number, number];
              expect(reversedPoint[0]).toBeCloseTo(
                reverseXFlag
                  ? mirror(normalPoint[0], normal.plot.x, normal.plot.width)
                  : normalPoint[0],
              );
              expect(reversedPoint[1]).toBeCloseTo(
                reverseYFlag
                  ? mirror(normalPoint[1], normal.plot.y, normal.plot.height)
                  : normalPoint[1],
              );
            }
          }
        }
      }
    },
  );

  it.each([
    ['bar', 'auto', 'custom', 'true', 'false'],
    ['column', 'custom', 'max', 'true', 'true'],
    ['area', 'max', 'auto', 'false', 'true'],
    ['line', 'min', 'min', 'true', 'true'],
    ['column', 'max', 'min', 'false', 'false'],
    ['scatter', 'min', 'auto', 'false', 'false'],
    ['area', 'custom', 'custom', 'true', 'false'],
    ['column', 'min', 'custom', 'false', 'true'],
    ['line', 'auto', 'max', 'false', 'false'],
    ['column', 'auto', 'auto', 'true', 'true'],
    ['bar', 'min', 'max', 'true', 'true'],
    ['line', 'max', 'custom', 'true', 'false'],
    ['line', 'custom', 'auto', 'false', 'true'],
    ['bar', 'custom', 'min', 'false', 'false'],
    ['scatter', 'auto', 'custom', 'true', 'true'],
    ['line', 'auto', 'min', 'false', 'true'],
    ['bar', 'max', 'auto', 'true', 'true'],
    ['area', 'min', 'custom', 'false', 'false'],
    ['scatter', 'max', 'max', 'true', 'true'],
    ['scatter', 'custom', 'min', 'false', 'false'],
    ['area', 'max', 'min', 'true', 'false'],
    ['area', 'max', 'max', 'false', 'true'],
    ['area', 'auto', 'max', 'true', 'false'],
  ] as const)(
    'positions both axis lines analytically for %s/%s/%s/%s/%s while preserving every other primitive',
    (kind, crossX, crossY, reverseX, reverseY) => {
      type CrossingMode = 'auto' | 'min' | 'max' | 'custom';
      const data: ChartData =
        kind === 'scatter'
          ? {
              domain: 'xy',
              series: [
                {
                  id: 's',
                  points: [
                    { id: 'p0', x: -3, y: -2 },
                    { id: 'p1', x: 0, y: 1 },
                    { id: 'p2', x: 4, y: 3 },
                  ],
                },
              ],
            }
          : {
              domain: 'category',
              categories: ['A', 'B', 'C'],
              categoryIds: ['a', 'b', 'c'],
              series: [
                { id: 's', values: [-3, 0, 4] },
                { id: 't', values: [2, -1, 3] },
              ],
            };
      const crossingValue = (mode: CrossingMode): ChartAxisSpec['crossing'] =>
        mode === 'custom' ? 2.4 : mode;
      const isBar = kind === 'bar';
      const isScatter = kind === 'scatter';
      const xAxis: ChartAxisSpec =
        isBar || isScatter
          ? {
              min: -5,
              max: 5,
              crossing: crossingValue(crossX),
              reverseOrder: reverseX === 'true',
            }
          : {
              crossing: crossingValue(crossX),
              reverseOrder: reverseX === 'true',
            };
      const yAxis: ChartAxisSpec =
        !isBar || isScatter
          ? {
              min: -5,
              max: 5,
              crossing: crossingValue(crossY),
              reverseOrder: reverseY === 'true',
            }
          : {
              crossing: crossingValue(crossY),
              reverseOrder: reverseY === 'true',
            };
      const spec: ChartSpec = { kind, xAxis, yAxis };
      const baseline = sceneFor(
        {
          kind,
          xAxis: { ...xAxis, crossing: undefined },
          yAxis: { ...yAxis, crossing: undefined },
        },
        data,
      );
      const scene = sceneFor(spec, data);
      const reverseXFlag = reverseX === 'true';
      const reverseYFlag = reverseY === 'true';
      const crossingNumber = (mode: CrossingMode): number =>
        mode === 'auto' ? 0 : mode === 'min' ? -5 : mode === 'max' ? 5 : 2.4;
      const numericCoordinate = (
        mode: CrossingMode,
        start: number,
        end: number,
        reverse: boolean,
      ): number => {
        const value = Math.min(5, Math.max(-5, crossingNumber(mode)));
        const from = reverse ? end : start;
        const to = reverse ? start : end;
        return from + ((value + 5) / 10) * (to - from);
      };
      const categoryCoordinate = (
        mode: CrossingMode,
        start: number,
        end: number,
        reverse: boolean,
        automatic: number,
      ): number => {
        if (mode === 'auto') return automatic;
        const count = 3;
        const edge =
          mode === 'min'
            ? 0
            : mode === 'max'
              ? count
              : Math.min(count, Math.max(1, Math.round(2.4))) - 1;
        return start + (reverse ? count - edge : edge) * ((end - start) / count);
      };
      const expectedX =
        isBar || isScatter
          ? numericCoordinate(crossX, scene.plot.x, scene.plot.x + scene.plot.width, reverseXFlag)
          : categoryCoordinate(
              crossX,
              scene.plot.x,
              scene.plot.x + scene.plot.width,
              reverseXFlag,
              scene.plot.x,
            );
      const expectedY =
        !isBar || isScatter
          ? numericCoordinate(crossY, scene.plot.y + scene.plot.height, scene.plot.y, reverseYFlag)
          : categoryCoordinate(
              crossY,
              scene.plot.y,
              scene.plot.y + scene.plot.height,
              reverseYFlag,
              scene.plot.y + scene.plot.height,
            );
      const xLine = scene.elements.find(
        (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
      );
      const yLine = scene.elements.find(
        (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
      );
      expect(xLine).toBeDefined();
      expect(yLine).toBeDefined();
      expect(Number(xLine?.attrs.y1)).toBeCloseTo(expectedY);
      expect(Number(xLine?.attrs.y2)).toBeCloseTo(expectedY);
      expect(Number(yLine?.attrs.x1)).toBeCloseTo(expectedX);
      expect(Number(yLine?.attrs.x2)).toBeCloseTo(expectedX);
      const withoutAxis = (candidate: ReturnType<typeof sceneFor>) =>
        candidate.elements.filter((element) => element.role !== 'axis');
      expect(scene.plot).toEqual(baseline.plot);
      expect(withoutAxis(scene)).toEqual(withoutAxis(baseline));
    },
  );

  it('treats omitted and auto crossings as exact legacy output, including reverse axes', () => {
    const cases: Array<{ kind: ChartSpec['kind']; data: ChartData }> = [
      {
        kind: 'column',
        data: {
          domain: 'category',
          categories: ['A', 'B'],
          series: [{ id: 's', values: [2, -3] }],
        },
      },
      {
        kind: 'bar',
        data: {
          domain: 'category',
          categories: ['A', 'B'],
          series: [{ id: 's', values: [2, -3] }],
        },
      },
      {
        kind: 'scatter',
        data: {
          domain: 'xy',
          series: [
            {
              id: 's',
              points: [
                { x: -2, y: 1 },
                { x: 3, y: -4 },
              ],
            },
          ],
        },
      },
    ];
    for (const { kind, data } of cases) {
      const axis = {
        xAxis: { min: -5, max: 5, reverseOrder: true },
        yAxis: { min: -5, max: 5, reverseOrder: true },
      };
      const baseline = sceneFor({ kind, ...axis }, data);
      const explicitAuto = sceneFor(
        {
          kind,
          xAxis: { ...axis.xAxis, crossing: 'auto' },
          yAxis: { ...axis.yAxis, crossing: 'auto' },
        },
        data,
      );
      const omitted = sceneFor(
        {
          kind,
          xAxis: { ...axis.xAxis, crossing: undefined },
          yAxis: { ...axis.yAxis, crossing: undefined },
        },
        data,
      );
      expect(explicitAuto).toEqual(baseline);
      expect(omitted).toEqual(baseline);
    }
  });

  it('clamps automatic zero and finite out-of-domain numeric crossings to resolved endpoints', () => {
    const positive = sceneFor(
      { kind: 'column', yAxis: { min: 2, max: 8, crossing: 'auto' } },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [2, 8] }] },
    );
    const negative = sceneFor(
      { kind: 'column', yAxis: { min: -8, max: -2, crossing: 'auto' } },
      { domain: 'category', categories: ['A'], series: [{ id: 's', values: [-2, -8] }] },
    );
    const positiveAxis = positive.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
    );
    const negativeAxis = negative.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
    );
    expect(Number(positiveAxis?.attrs.y1)).toBeCloseTo(positive.plot.y + positive.plot.height);
    expect(Number(negativeAxis?.attrs.y1)).toBeCloseTo(negative.plot.y);

    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 's',
          points: [
            { x: -2, y: -2 },
            { x: 3, y: 4 },
          ],
        },
      ],
    };
    const below = sceneFor(
      {
        kind: 'scatter',
        xAxis: { min: -5, max: 5, crossing: -100 },
        yAxis: { min: -5, max: 5, crossing: -100 },
      },
      data,
    );
    const above = sceneFor(
      {
        kind: 'scatter',
        xAxis: { min: -5, max: 5, crossing: 100 },
        yAxis: { min: -5, max: 5, crossing: 100 },
      },
      data,
    );
    const lineFor = (scene: ReturnType<typeof sceneFor>, axis: 'x' | 'y') =>
      scene.elements.find(
        (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === axis,
      );
    expect(Number(lineFor(below, 'y')?.attrs.x1)).toBeCloseTo(below.plot.x);
    expect(Number(lineFor(below, 'x')?.attrs.y1)).toBeCloseTo(below.plot.y + below.plot.height);
    expect(Number(lineFor(above, 'y')?.attrs.x1)).toBeCloseTo(above.plot.x + above.plot.width);
    expect(Number(lineFor(above, 'x')?.attrs.y1)).toBeCloseTo(above.plot.y);
  });

  it('uses automatic crossing for invalid runtime values and keeps hidden controlling axes effective', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 's',
          points: [
            { x: -2, y: -2 },
            { x: 3, y: 4 },
          ],
        },
      ],
    };
    const auto = sceneFor(
      {
        kind: 'scatter',
        xAxis: { min: -5, max: 5, crossing: 'auto' },
        yAxis: { min: -5, max: 5, crossing: 'auto' },
      },
      data,
    );
    for (const invalid of [
      'bogus',
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
    ]) {
      const invalidScene = sceneFor(
        {
          kind: 'scatter',
          xAxis: { min: -5, max: 5, crossing: invalid as unknown as ChartAxisSpec['crossing'] },
          yAxis: { min: -5, max: 5, crossing: invalid as unknown as ChartAxisSpec['crossing'] },
        },
        data,
      );
      expect(invalidScene).toEqual(auto);
    }

    const hiddenController = sceneFor(
      {
        kind: 'column',
        xAxis: { visible: false, crossing: 'max' },
        yAxis: { min: -5, max: 5 },
      },
      { domain: 'category', categories: ['A', 'B'], series: [{ id: 's', values: [-1, 2] }] },
    );
    const yLine = hiddenController.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(yLine).toBeDefined();
    expect(Number(yLine?.attrs.x1)).toBeCloseTo(
      hiddenController.plot.x + hiddenController.plot.width,
    );
    expect(
      hiddenController.elements.some(
        (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'x',
      ),
    ).toBe(false);
  });

  it('maps category crossings to visible bands, IDs, edge rules, and reverse order', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B', 'C', 'D'],
      categoryIds: ['a', 'b', 'c', 'd'],
      series: [{ id: 's', values: [1, 2, 3, 4] }],
    };
    const filtered = sceneFor(
      {
        kind: 'column',
        filter: { hiddenCategoryIds: ['b', 'd'] },
        xAxis: { crossing: 2.49 },
      },
      data,
    );
    const filteredReverse = sceneFor(
      {
        kind: 'column',
        filter: { hiddenCategoryIds: ['b', 'd'] },
        xAxis: { crossing: 2.49, reverseOrder: true },
      },
      data,
    );
    const yLineFor = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements.find(
        (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
      );
    expect(Number(yLineFor(filtered)?.attrs.x1)).toBeCloseTo(
      filtered.plot.x + filtered.plot.width / 2,
    );
    expect(Number(yLineFor(filteredReverse)?.attrs.x1)).toBeCloseTo(
      filteredReverse.plot.x + filteredReverse.plot.width / 2,
    );
    expect(
      filtered.elements
        .filter((element) => element.role === 'category-label')
        .map((element) => element.pointId),
    ).toEqual(['a', 'c']);
    expect(
      filtered.elements
        .filter((element) => element.role === 'bar')
        .map((element) => element.pointId),
    ).toEqual(['a', 'c']);

    for (const [crossing, expectedNormal, expectedReverse] of [
      ['min', filtered.plot.x, filtered.plot.x + filtered.plot.width],
      ['max', filtered.plot.x + filtered.plot.width, filtered.plot.x],
      [0.2, filtered.plot.x, filtered.plot.x + filtered.plot.width],
      [99, filtered.plot.x + filtered.plot.width / 2, filtered.plot.x + filtered.plot.width / 2],
    ] as const) {
      const normal = sceneFor(
        { kind: 'column', filter: { hiddenCategoryIds: ['b', 'd'] }, xAxis: { crossing } },
        data,
      );
      const reverse = sceneFor(
        {
          kind: 'column',
          filter: { hiddenCategoryIds: ['b', 'd'] },
          xAxis: { crossing, reverseOrder: true },
        },
        data,
      );
      expect(Number(yLineFor(normal)?.attrs.x1)).toBeCloseTo(expectedNormal);
      expect(Number(yLineFor(reverse)?.attrs.x1)).toBeCloseTo(expectedReverse);
    }

    const bar = sceneFor(
      {
        kind: 'bar',
        yAxis: { crossing: 2, reverseOrder: false },
      },
      {
        domain: 'category',
        categories: ['A', 'B', 'C'],
        categoryIds: ['a', 'b', 'c'],
        series: [{ id: 's', values: [1, 2, 3] }],
      },
    );
    const barXLine = bar.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
    );
    expect(Number(barXLine?.attrs.y1)).toBeCloseTo(bar.plot.y + bar.plot.height / 3);
    expect(
      bar.elements.filter((element) => element.role === 'bar').map((element) => element.pointId),
    ).toEqual(['a', 'b', 'c']);
  });

  it('falls back to legacy edges for empty visible categories and ignores circular crossings', () => {
    const emptyColumn = sceneFor(
      { kind: 'column', xAxis: { crossing: 'max' } },
      { domain: 'category', categories: [], series: [{ id: 's', values: [] }] },
    );
    const emptyColumnLine = emptyColumn.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(Number(emptyColumnLine?.attrs.x1)).toBeCloseTo(emptyColumn.plot.x);

    const emptyBar = sceneFor(
      { kind: 'bar', yAxis: { crossing: 'min' } },
      { domain: 'category', categories: [], series: [{ id: 's', values: [] }] },
    );
    const emptyBarLine = emptyBar.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
    );
    expect(Number(emptyBarLine?.attrs.y1)).toBeCloseTo(emptyBar.plot.y + emptyBar.plot.height);

    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      categoryIds: ['a', 'b'],
      series: [{ id: 's', values: [1, 2] }],
    };
    for (const kind of ['pie', 'doughnut'] as const) {
      const baseline = sceneFor({ kind }, data);
      const crossing = sceneFor(
        { kind, xAxis: { crossing: 'max' }, yAxis: { crossing: -100 } },
        data,
      );
      expect(crossing).toEqual(baseline);
    }
  });

  it('positions axis labels additively while preserving legacy and mark identities', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      categoryIds: ['category-a', 'category-b'],
      series: [{ id: 'series-a', values: [1, 2] }],
    };
    const baseline = sceneFor({ kind: 'column' }, data);
    const explicitLow = sceneFor(
      { kind: 'column', xAxis: { labelPosition: 'low' }, yAxis: { labelPosition: 'low' } },
      data,
    );
    const invalid = sceneFor(
      {
        kind: 'column',
        xAxis: { labelPosition: 'runtime-invalid' as never },
        yAxis: { labelPosition: 17 as never },
      },
      data,
    );
    expect(explicitLow).toEqual(baseline);
    expect(invalid).toEqual(baseline);

    const high = sceneFor(
      { kind: 'column', xAxis: { labelPosition: 'high' }, yAxis: { labelPosition: 'high' } },
      data,
    );
    const highXLabels = high.elements.filter(
      (element) => element.role === 'category-label' && element.attrs['data-chart-axis'] === 'x',
    );
    const highYLabels = high.elements.filter(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(highXLabels.length).toBe(2);
    expect(highXLabels.every((label) => label.attrs.y === high.plot.y - 4)).toBe(true);
    expect(highYLabels.length).toBeGreaterThan(0);
    expect(highYLabels.every((label) => label.attrs.x === high.plot.x + high.plot.width + 4)).toBe(
      true,
    );
    expect(highYLabels.every((label) => label.attrs['text-anchor'] === 'start')).toBe(true);
    expect(
      high.elements
        .filter((element) => element.role === 'bar')
        .map((element) => [element.seriesId, element.pointId]),
    ).toEqual(
      baseline.elements
        .filter((element) => element.role === 'bar')
        .map((element) => [element.seriesId, element.pointId]),
    );

    const none = sceneFor(
      { kind: 'column', xAxis: { labelPosition: 'none' }, yAxis: { labelPosition: 'none' } },
      data,
    );
    expect(none.plot).toEqual(baseline.plot);
    expect(
      none.elements.filter((element) =>
        ['axis-label', 'category-label'].includes(String(element.role)),
      ),
    ).toHaveLength(0);
    expect(
      none.elements.filter(
        (element) => !['axis-label', 'category-label'].includes(String(element.role)),
      ),
    ).toEqual(
      baseline.elements.filter(
        (element) => !['axis-label', 'category-label'].includes(String(element.role)),
      ),
    );
  });

  it('keeps next-to labels beside the actual crossing lines and retains controllers', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [{ id: 'series-a', values: [1, 2] }],
    };
    const scene = sceneFor(
      {
        kind: 'column',
        xAxis: { crossing: 'max', labelPosition: 'nextTo' },
        yAxis: { crossing: 'max', labelPosition: 'nextTo' },
      },
      data,
    );
    const xAxisLine = scene.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
    );
    const yAxisLine = scene.elements.find(
      (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(xAxisLine).toBeDefined();
    expect(yAxisLine).toBeDefined();
    const xLabels = scene.elements.filter(
      (element) => element.role === 'category-label' && element.attrs['data-chart-axis'] === 'x',
    );
    const yLabels = scene.elements.filter(
      (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(xLabels.every((label) => label.attrs.y === Number(xAxisLine?.attrs.y1) + 12)).toBe(true);
    expect(yLabels.every((label) => label.attrs.x === Number(yAxisLine?.attrs.x1) - 2)).toBe(true);

    const hiddenController = sceneFor(
      { kind: 'column', xAxis: { visible: false, labelPosition: 'nextTo' } },
      data,
    );
    expect(
      hiddenController.elements.some(
        (element) => element.role === 'axis-tick' && element.attrs['data-chart-axis'] === 'y',
      ),
    ).toBe(true);
    expect(
      hiddenController.elements.some(
        (element) => element.role === 'category-label' && element.attrs['data-chart-axis'] === 'x',
      ),
    ).toBe(false);

    const circularBaseline = sceneFor({ kind: 'pie' }, data);
    for (const labelPosition of ['nextTo', 'high', 'none'] as const) {
      expect(
        sceneFor({ kind: 'pie', xAxis: { labelPosition }, yAxis: { labelPosition } }, data),
      ).toEqual(circularBaseline);
    }
  });

  it('reserves high label lanes for titles and separates right legends from bounded labels', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['非常に長いカテゴリラベル一', '非常に長いカテゴリラベル二'],
      series: [{ id: 'series-a', name: 'Series', values: [1, 2] }],
    };
    const title = sceneAt(
      { kind: 'column', title: 'Chart title', xAxis: { labelPosition: 'high' } },
      data,
      240,
      132,
    );
    expect(title.plot.y).toBeGreaterThanOrEqual(34);
    const topLegend = sceneAt(
      {
        kind: 'column',
        legend: { visible: true, position: 'top' },
        xAxis: { labelPosition: 'high' },
      },
      data,
      240,
      132,
    );
    expect(topLegend.plot.y).toBeGreaterThanOrEqual(34);

    const rightLegend = sceneAt(
      {
        kind: 'bar',
        legend: { visible: true, position: 'right' },
        yAxis: { labelPosition: 'high' },
      },
      data,
      240,
      132,
    );
    const categoryLabels = rightLegend.elements.filter(
      (element) => element.role === 'category-label' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(categoryLabels).toHaveLength(2);
    expect(categoryLabels.every((label) => label.attrs['aria-label'])).toBe(true);
    const legendRects = rightLegend.elements.filter(
      (element) => element.role === 'legend' && element.tag === 'rect',
    );
    expect(legendRects.length).toBeGreaterThan(0);
    const labelRight = Math.max(
      ...categoryLabels.map(
        (label) =>
          Number(label.attrs.x) +
          Array.from(String(label.text ?? '')).reduce(
            (width, character) => width + (character === '…' ? 8.2 : 10),
            0,
          ),
      ),
    );
    expect(Math.min(...legendRects.map((legend) => Number(legend.attrs.x))) > labelRight).toBe(
      true,
    );
    expect(
      rightLegend.elements.every((element) =>
        Object.values(element.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);

    const mixedData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-a',
          points: [
            { x: 0, y: 0 },
            { x: 10, y: 1 },
          ],
        },
      ],
    };
    const mixedBaseline = layoutChart({
      spec: {
        kind: 'scatter',
        xAxis: { min: 0, max: 10, majorUnit: 5, numberFormat: 'long-x' },
        yAxis: { min: 0, max: 1, majorUnit: 1 },
      },
      data: mixedData,
      width: 320,
      height: 132,
      formatValue: (value, code) => `${code}:${value}:a-long-x-axis-label`,
    });
    const mixedHighY = layoutChart({
      spec: {
        kind: 'scatter',
        xAxis: { min: 0, max: 10, majorUnit: 5, numberFormat: 'long-x' },
        yAxis: { min: 0, max: 1, majorUnit: 1, labelPosition: 'high' },
      },
      data: mixedData,
      width: 320,
      height: 132,
      formatValue: (value, code) => `${code}:${value}:a-long-x-axis-label`,
    });
    const xLabelTexts = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements
        .filter(
          (element) => element.role === 'axis-label' && element.attrs['data-chart-axis'] === 'x',
        )
        .map((element) => [element.text, element.attrs['aria-label']]);
    expect(xLabelTexts(mixedHighY)).toEqual(xLabelTexts(mixedBaseline));

    const highYTitle = sceneAt(
      {
        kind: 'bar',
        legend: { visible: true, position: 'left' },
        yAxis: { labelPosition: 'high', title: 'Categories' },
      },
      data,
      240,
      132,
    );
    const yTitle = highYTitle.elements.find(
      (element) => element.role === 'axis-title' && element.attrs['data-chart-axis'] === 'y',
    );
    expect(Number(yTitle?.attrs.x)).toBeGreaterThanOrEqual(12);
    expect(Number(yTitle?.attrs.x)).toBeLessThan(highYTitle.plot.x);
  });

  it('covers the approved 31-row axis-position pairwise matrix analytically', () => {
    const rows = [
      ['area', 'low', 'default', 'max', true, false],
      ['area', 'high', 'high', 'custom', false, true],
      ['column', 'nextTo', 'low', 'max', false, false],
      ['scatter', 'none', 'none', 'min', false, false],
      ['bar', 'low', 'nextTo', 'min', false, true],
      ['column', 'high', 'none', 'min', true, true],
      ['scatter', 'default', 'low', 'custom', true, true],
      ['line', 'default', 'nextTo', 'custom', false, false],
      ['bar', 'none', 'high', 'max', true, true],
      ['line', 'nextTo', 'default', 'min', true, true],
      ['area', 'low', 'low', 'min', false, true],
      ['bar', 'nextTo', 'none', 'custom', true, true],
      ['column', 'none', 'nextTo', 'custom', true, false],
      ['line', 'none', 'low', 'custom', false, true],
      ['bar', 'high', 'default', 'custom', true, false],
      ['column', 'low', 'high', 'min', false, false],
      ['column', 'default', 'default', 'max', false, true],
      ['scatter', 'high', 'nextTo', 'max', false, true],
      ['line', 'low', 'none', 'max', false, true],
      ['scatter', 'low', 'high', 'custom', true, false],
      ['scatter', 'nextTo', 'default', 'max', false, false],
      ['area', 'none', 'default', 'min', true, true],
      ['scatter', 'default', 'high', 'max', true, false],
      ['area', 'nextTo', 'nextTo', 'max', true, true],
      ['line', 'default', 'none', 'min', false, false],
      ['line', 'high', 'high', 'custom', true, false],
      ['bar', 'default', 'default', 'max', false, true],
      ['bar', 'high', 'low', 'custom', true, true],
      ['area', 'default', 'high', 'min', true, true],
      ['area', 'high', 'none', 'custom', false, false],
      ['column', 'nextTo', 'high', 'custom', false, true],
    ] as const;
    expect(rows).toHaveLength(31);

    const categoryData: ChartData = {
      domain: 'category',
      categories: ['A', 'B', 'C'],
      categoryIds: ['a', 'b', 'c'],
      series: [{ id: 'series-a', name: 'Series', values: [-4, 10, 7] }],
    };
    const scatterData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-a',
          name: 'Series',
          points: [
            { id: 'p-a', x: -4, y: 2 },
            { id: 'p-b', x: 10, y: -3 },
            { id: 'p-c', x: 7, y: 5 },
          ],
        },
      ],
    };
    const baseAxis = (axis: ChartAxisSpec): ChartAxisSpec => {
      const { labelPosition: _labelPosition, ...withoutPosition } = axis;
      return withoutPosition;
    };
    const axisFor = (
      kind: ChartSpec['kind'],
      axis: 'x' | 'y',
      position: (typeof rows)[number][1],
      reverseOrder: boolean,
      crossingMode: (typeof rows)[number][3],
    ): ChartAxisSpec => {
      const numeric = kind === 'scatter' || (axis === 'x' ? kind === 'bar' : kind !== 'bar');
      const crossing = crossingMode === 'custom' ? (numeric ? 2.5 : 2) : crossingMode;
      return {
        min: -10,
        max: 10,
        reverseOrder,
        crossing,
        ...(position === 'default' ? {} : { labelPosition: position }),
      };
    };
    const labelsFor = (scene: ReturnType<typeof sceneFor>, axis: 'x' | 'y') =>
      scene.elements.filter(
        (element) =>
          (element.role === 'axis-label' || element.role === 'category-label') &&
          element.attrs['data-chart-axis'] === axis,
      );
    const marksFor = (scene: ReturnType<typeof sceneFor>) =>
      scene.elements.filter((element) => element.role === 'bar' || element.role === 'point');
    const normalized = (
      scene: ReturnType<typeof sceneFor>,
      mark: ChartPrimitive,
      coordinate: 'x' | 'y',
    ): number => {
      const value = Number(
        mark.attrs[
          coordinate === 'x'
            ? mark.tag === 'circle'
              ? 'cx'
              : 'x'
            : mark.tag === 'circle'
              ? 'cy'
              : 'y'
        ],
      );
      return coordinate === 'x'
        ? (value - scene.plot.x) / scene.plot.width
        : (value - scene.plot.y) / scene.plot.height;
    };

    for (const [kind, labelX, labelY, crossingMode, reverseX, reverseY] of rows) {
      const xAxis = axisFor(kind, 'x', labelX, reverseX, crossingMode);
      const yAxis = axisFor(kind, 'y', labelY, reverseY, crossingMode);
      const spec: ChartSpec = {
        kind,
        scatterStyle: 'lineMarker',
        legend: { visible: false },
        xAxis,
        yAxis,
      };
      const data = kind === 'scatter' ? scatterData : categoryData;
      const scene = sceneAt(spec, data, 420, 180);
      const baselineSpec: ChartSpec = {
        ...spec,
        xAxis: baseAxis(xAxis),
        yAxis: baseAxis(yAxis),
      };
      const baseline = sceneAt(baselineSpec, data, 420, 180);
      if (labelX === 'default' || labelX === 'low') {
        expect(sceneAt({ ...spec, xAxis: baseAxis(xAxis) }, data, 420, 180)).toEqual(scene);
      }
      if (labelY === 'default' || labelY === 'low') {
        expect(sceneAt({ ...spec, yAxis: baseAxis(yAxis) }, data, 420, 180)).toEqual(scene);
      }

      const xLine = scene.elements.find(
        (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
      );
      const yLine = scene.elements.find(
        (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
      );
      expect(xLine).toBeDefined();
      expect(yLine).toBeDefined();
      for (const [axis, position] of [
        ['x', labelX],
        ['y', labelY],
      ] as const) {
        const actual = labelsFor(scene, axis);
        const expectedCount = labelsFor(baseline, axis).length;
        if (position === 'none') {
          expect(actual).toHaveLength(0);
          continue;
        }
        expect(actual).toHaveLength(expectedCount);
        for (const label of actual) {
          if (axis === 'x') {
            const expectedY =
              position === 'high'
                ? scene.plot.y - 4
                : position === 'nextTo'
                  ? Number(xLine?.attrs.y1) + 12
                  : scene.plot.y + scene.plot.height + 12;
            expect(Number(label.attrs.y)).toBeCloseTo(expectedY);
          } else {
            const offset = kind === 'bar' ? 4 : 2;
            const expectedX =
              position === 'high'
                ? scene.plot.x + scene.plot.width + 4
                : position === 'nextTo'
                  ? Number(yLine?.attrs.x1) - offset
                  : scene.plot.x - offset;
            expect(Number(label.attrs.x)).toBeCloseTo(expectedX);
            expect(label.attrs['text-anchor']).toBe(position === 'high' ? 'start' : 'end');
          }
        }
      }

      const marks = marksFor(scene);
      const baselineMarks = marksFor(baseline);
      expect(marks).toHaveLength(baselineMarks.length);
      marks.forEach((mark, index) => {
        const baselineMark = baselineMarks[index];
        if (!baselineMark) throw new Error(`missing baseline mark at index ${index}`);
        expect([mark.seriesId, mark.pointId, mark.attrs.fill]).toEqual([
          baselineMark.seriesId,
          baselineMark.pointId,
          baselineMark.attrs.fill,
        ]);
        expect(normalized(scene, mark, 'x')).toBeCloseTo(
          normalized(baseline, baselineMark, 'x'),
          6,
        );
        expect(normalized(scene, mark, 'y')).toBeCloseTo(
          normalized(baseline, baselineMark, 'y'),
          6,
        );
      });
    }
  });

  // Category crossings intentionally address visible ordinal bands by number;
  // this module has no ID-valued axis-position API for an individual category.

  it('projects scatter points and axes in log space while keeping reversal and crossings finite', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-a',
          points: [
            { id: 'p1', x: 1, y: 1 },
            { id: 'p2', x: 10, y: 10 },
            { id: 'p3', x: 100, y: 100 },
          ],
        },
      ],
    };
    const scene = sceneAt(
      {
        kind: 'scatter',
        legend: { visible: false },
        xAxis: { logBase: 10, reverseOrder: true, crossing: 10 },
        yAxis: { logBase: 10, reverseOrder: true, crossing: 10 },
      },
      data,
      420,
      180,
    );
    const points = scene.elements.filter((element) => element.role === 'point');
    expect(points).toHaveLength(3);
    const x = points.map((point) => Number(point.attrs.cx));
    const y = points.map((point) => Number(point.attrs.cy));
    expect(x[0]).toBeCloseTo(scene.plot.x + scene.plot.width);
    expect(x[1]).toBeCloseTo(scene.plot.x + scene.plot.width / 2);
    expect(x[2]).toBeCloseTo(scene.plot.x);
    expect(y[0]).toBeCloseTo(scene.plot.y);
    expect(y[1]).toBeCloseTo(scene.plot.y + scene.plot.height / 2);
    expect(y[2]).toBeCloseTo(scene.plot.y + scene.plot.height);
    expect(
      scene.elements
        .filter((element) => element.role === 'axis-label')
        .map((element) => element.text),
    ).toEqual(expect.arrayContaining(['1', '10', '100']));
    const axes = scene.elements.filter((element) => element.role === 'axis');
    expect(
      axes.every((axis) =>
        Object.values(axis.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);
  });

  it('skips non-positive category values and uses the positive domain minimum as baseline', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['one', 'zero', 'negative', 'ten', 'nan'],
      categoryIds: ['one', 'zero', 'negative', 'ten', 'nan'],
      series: [{ id: 'series-a', values: [1, 0, -2, 10, Number.NaN] }],
    };
    const scene = sceneAt(
      {
        kind: 'column',
        legend: { visible: false },
        yAxis: { logBase: 10 },
        dataLabels: { visible: true, showValue: true },
      },
      data,
      420,
      180,
    );
    const bars = scene.elements.filter((element) => element.role === 'bar');
    expect(bars).toHaveLength(2);
    expect(bars.map((bar) => bar.pointId)).toEqual(['one', 'ten']);
    expect(Number(bars[0]?.attrs.height)).toBeCloseTo(0);
    expect(Number(bars[1]?.attrs.height)).toBeGreaterThan(0);
    expect(
      scene.elements
        .filter((element) => element.role === 'data-label')
        .map((element) => element.pointId),
    ).toEqual(['one', 'ten']);
    expect(scene.elements.some((element) => element.role === 'empty')).toBe(false);
  });

  it('keeps clustered log lines positive and breaks span runs at invalid values', () => {
    const data: ChartData = {
      domain: 'category',
      categories: ['a', 'b', 'blank', 'c', 'zero', 'd', 'nan', 'e', 'f'],
      series: [{ id: 'series-a', values: [1, 2, null, 3, 0, 4, Number.NaN, 5, 6] }],
    };
    const scene = sceneFor({ kind: 'line', emptyCells: 'span', yAxis: { logBase: 10 } }, data);
    const lines = scene.elements.filter((element) => element.role === 'line');
    const points = scene.elements.filter((element) => element.role === 'point');
    expect(lines).toHaveLength(2);
    expect(points).toHaveLength(6);
    expect(
      scene.elements.every((element) =>
        Object.values(element.attrs).every(
          (value) => typeof value !== 'number' || Number.isFinite(value),
        ),
      ),
    ).toBe(true);
  });

  it('uses pair eligibility for mixed log-linear scatter axes and hard-breaks invalid coordinates', () => {
    const data: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-a',
          points: [
            { id: 'p1', x: 1, y: 1 },
            { id: 'p2', x: 10, y: 10 },
            { id: 'bad', x: 0, y: null },
            { id: 'p3', x: 100, y: 100 },
          ],
        },
      ],
    };
    const scene = sceneAt(
      {
        kind: 'scatter',
        scatterStyle: 'lineMarker',
        emptyCells: 'span',
        legend: { visible: false },
        xAxis: { logBase: 10 },
      },
      data,
      420,
      180,
    );
    const points = scene.elements.filter((element) => element.role === 'point');
    expect(points.map((point) => point.pointId)).toEqual(['p1', 'p2', 'p3']);
    expect(scene.elements.filter((element) => element.role === 'line')).toHaveLength(1);
  });

  it('keeps mixed-axis scatter data labels aligned with marker pair eligibility', () => {
    const yLinearData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-a',
          points: [
            { id: 'inside', x: 2, y: 50 },
            { id: 'outside-y', x: 2, y: 200 },
          ],
        },
      ],
    };
    const xLogScene = sceneAt(
      {
        kind: 'scatter',
        legend: { visible: false },
        dataLabels: { visible: true, showValue: true },
        xAxis: { logBase: 10, min: 1, max: 10 },
        yAxis: { min: 0, max: 100 },
      },
      yLinearData,
      420,
      180,
    );
    expect(
      xLogScene.elements.filter((element) => element.role === 'point').map((e) => e.pointId),
    ).toEqual(['inside']);
    expect(
      xLogScene.elements.filter((element) => element.role === 'data-label').map((e) => e.pointId),
    ).toEqual(['inside']);

    const xLinearData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-a',
          points: [
            { id: 'inside', x: 50, y: 2 },
            { id: 'outside-x', x: 200, y: 2 },
          ],
        },
      ],
    };
    const yLogScene = sceneAt(
      {
        kind: 'scatter',
        legend: { visible: false },
        dataLabels: { visible: true, showValue: true },
        xAxis: { min: 0, max: 100 },
        yAxis: { logBase: 10, min: 1, max: 10 },
      },
      xLinearData,
      420,
      180,
    );
    expect(
      yLogScene.elements.filter((element) => element.role === 'point').map((e) => e.pointId),
    ).toEqual(['inside']);
    expect(
      yLogScene.elements.filter((element) => element.role === 'data-label').map((e) => e.pointId),
    ).toEqual(['inside']);
  });

  it('covers the approved 22-row clustered log matrix analytically', () => {
    const rows = [
      ['bar', 10, true, 'custom', 'nextTo'],
      ['column', 1000, false, 'custom', 'none'],
      ['line', 2, false, 'max', 'nextTo'],
      ['area', 2, true, 'auto', 'none'],
      ['area', 10, false, 'min', 'low'],
      ['scatter', 1000, true, 'min', 'nextTo'],
      ['bar', 2, false, 'min', 'high'],
      ['bar', 1000, true, 'max', 'low'],
      ['column', 10, true, 'max', 'high'],
      ['column', 2, true, 'min', 'low'],
      ['scatter', 10, false, 'min', 'none'],
      ['area', 1000, true, 'max', 'high'],
      ['line', 2, true, 'custom', 'low'],
      ['line', 10, false, 'auto', 'high'],
      ['scatter', 2, true, 'custom', 'high'],
      ['column', 10, false, 'auto', 'nextTo'],
      ['bar', 1000, true, 'auto', 'none'],
      ['scatter', 2, true, 'auto', 'low'],
      ['line', 1000, true, 'max', 'none'],
      ['area', 10, false, 'custom', 'nextTo'],
      ['line', 1000, true, 'min', 'low'],
      ['scatter', 1000, false, 'max', 'high'],
    ] as const;
    expect(rows).toHaveLength(22);

    const categoryData: ChartData = {
      domain: 'category',
      categories: ['A', 'B', 'C'],
      categoryIds: ['a', 'b', 'c'],
      series: [{ id: 'series-a', values: [1, 10, 100] }],
    };
    const scatterData: ChartData = {
      domain: 'xy',
      series: [
        {
          id: 'series-a',
          points: [
            { id: 'p-a', x: 1, y: 1 },
            { id: 'p-b', x: 10, y: 10 },
            { id: 'p-c', x: 100, y: 100 },
          ],
        },
      ],
    };
    const labelPositions = ['default', 'low', 'high', 'nextTo', 'none'] as const;
    type MatrixLabelPosition = (typeof labelPositions)[number];
    const axisFor = (
      kind: ChartSpec['kind'],
      axis: 'x' | 'y',
      base: number,
      reverse: boolean,
      crossingMode: (typeof rows)[number][3],
      labelPosition: MatrixLabelPosition,
    ): ChartAxisSpec => {
      const numeric = kind === 'scatter' || (kind === 'bar' ? axis === 'x' : axis === 'y');
      return {
        ...(numeric ? { logBase: base } : {}),
        reverseOrder: reverse,
        crossing: crossingMode === 'custom' ? 10 : crossingMode,
        ...(labelPosition === 'default' ? {} : { labelPosition }),
      };
    };
    const numericMaximum = (base: number): number => (base === 2 ? 128 : base === 10 ? 100 : base);
    const logFraction = (value: number, base: number): number =>
      Math.log(value) / Math.log(numericMaximum(base));
    const pointCoordinate = (mark: ChartPrimitive, coordinate: 'x' | 'y'): number =>
      Number(
        mark.attrs[
          coordinate === 'x'
            ? mark.tag === 'circle'
              ? 'cx'
              : 'x'
            : mark.tag === 'circle'
              ? 'cy'
              : 'y'
        ],
      );
    const withoutLog = (axis: ChartAxisSpec): ChartAxisSpec => {
      const { logBase: _logBase, ...legacyAxis } = axis;
      return legacyAxis;
    };

    for (const [kind, base, reverse, crossingMode, labelPosition] of rows) {
      const xAxis = axisFor(kind, 'x', base, reverse, crossingMode, labelPosition);
      const yAxis = axisFor(kind, 'y', base, reverse, crossingMode, labelPosition);
      const spec: ChartSpec = {
        kind,
        legend: { visible: false },
        dataLabels: { visible: true, showValue: true },
        xAxis,
        yAxis,
      };
      const data = kind === 'scatter' ? scatterData : categoryData;
      const scene = sceneAt(spec, data, 420, 180);
      const marks = scene.elements.filter(
        (element) =>
          (element.role === 'bar' || element.role === 'point') && element.pointId !== undefined,
      );
      expect(marks.map((mark) => mark.pointId)).toEqual(
        kind === 'scatter' ? ['p-a', 'p-b', 'p-c'] : ['a', 'b', 'c'],
      );
      const values = [1, 10, 100];
      for (const [index, value] of values.entries()) {
        const mark = marks[index];
        if (!mark) throw new Error(`missing matrix mark ${kind}/${base}/${index}`);
        const fraction = logFraction(value, base);
        if (kind === 'bar') {
          const end = reverse
            ? Number(mark.attrs.x)
            : Number(mark.attrs.x) + Number(mark.attrs.width);
          expect((end - scene.plot.x) / scene.plot.width).toBeCloseTo(
            reverse ? 1 - fraction : fraction,
            6,
          );
        } else if (kind === 'scatter') {
          const x = pointCoordinate(mark, 'x');
          const y = pointCoordinate(mark, 'y');
          expect((x - scene.plot.x) / scene.plot.width).toBeCloseTo(
            reverse ? 1 - fraction : fraction,
            6,
          );
          expect((y - scene.plot.y) / scene.plot.height).toBeCloseTo(
            reverse ? fraction : 1 - fraction,
            6,
          );
        } else {
          const end = reverse
            ? Number(mark.attrs.y) + Number(mark.attrs.height)
            : Number(mark.attrs.y);
          const y = mark.role === 'bar' ? end : pointCoordinate(mark, 'y');
          expect((y - scene.plot.y) / scene.plot.height).toBeCloseTo(
            reverse ? fraction : 1 - fraction,
            6,
          );
        }
      }

      const labels = scene.elements.filter(
        (element) => element.role === 'data-label' && element.pointId !== undefined,
      );
      expect(labels.map((label) => label.pointId)).toEqual(
        kind === 'scatter' ? ['p-a', 'p-b', 'p-c'] : ['a', 'b', 'c'],
      );
      const axisLabels = scene.elements.filter(
        (element) =>
          (element.role === 'axis-label' || element.role === 'category-label') &&
          element.attrs['data-chart-axis'] !== undefined,
      );
      if (labelPosition === 'none') expect(axisLabels).toHaveLength(0);
      else expect(axisLabels.length).toBeGreaterThan(0);
      expect(
        scene.elements.every((element) =>
          Object.values(element.attrs).every(
            (attribute) => typeof attribute !== 'number' || Number.isFinite(attribute),
          ),
        ),
      ).toBe(true);

      const crossingFraction =
        crossingMode === 'max' ? 1 : crossingMode === 'custom' ? logFraction(10, base) : 0;
      if (kind === 'bar') {
        const line = scene.elements.find(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
        );
        expect(Number(line?.attrs.x1)).toBeCloseTo(
          scene.plot.x + scene.plot.width * (reverse ? 1 - crossingFraction : crossingFraction),
          6,
        );
      } else if (kind === 'scatter') {
        const xLine = scene.elements.find(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'y',
        );
        const yLine = scene.elements.find(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
        );
        expect(Number(xLine?.attrs.x1)).toBeCloseTo(
          scene.plot.x + scene.plot.width * (reverse ? 1 - crossingFraction : crossingFraction),
          6,
        );
        expect(Number(yLine?.attrs.y1)).toBeCloseTo(
          scene.plot.y + scene.plot.height * (reverse ? crossingFraction : 1 - crossingFraction),
          6,
        );
      } else {
        const line = scene.elements.find(
          (element) => element.role === 'axis' && element.attrs['data-chart-axis'] === 'x',
        );
        expect(Number(line?.attrs.y1)).toBeCloseTo(
          scene.plot.y + scene.plot.height * (reverse ? crossingFraction : 1 - crossingFraction),
          6,
        );
      }

      const invalidSpec: ChartSpec = {
        ...spec,
        xAxis: { ...xAxis, ...(xAxis.logBase === undefined ? {} : { logBase: 1 }) },
        yAxis: { ...yAxis, ...(yAxis.logBase === undefined ? {} : { logBase: 1 }) },
      };
      const legacySpec: ChartSpec = {
        ...spec,
        xAxis: withoutLog(xAxis),
        yAxis: withoutLog(yAxis),
      };
      expect(sceneAt(invalidSpec, data, 420, 180)).toEqual(sceneAt(legacySpec, data, 420, 180));
    }
    expect(new Set(rows.map((row) => row[4]))).toEqual(new Set(labelPositions.slice(1)));
  });

  it('leaves stacked and invalid log settings on the exact legacy linear layout', () => {
    const category: ChartData = {
      domain: 'category',
      categories: ['A', 'B'],
      series: [
        { id: 'a', values: [1, 10] },
        { id: 'b', values: [2, 20] },
      ],
    };
    expect(
      sceneFor({ kind: 'column', grouping: 'stacked', yAxis: { logBase: 10 } }, category),
    ).toEqual(sceneFor({ kind: 'column', grouping: 'stacked' }, category));
    expect(sceneFor({ kind: 'column', yAxis: { logBase: 1 } }, category)).toEqual(
      sceneFor({ kind: 'column' }, category),
    );
    expect(
      sceneFor({ kind: 'pie', xAxis: { logBase: 10 }, yAxis: { logBase: 10 } }, category),
    ).toEqual(sceneFor({ kind: 'pie' }, category));
  });
});

describe('layoutChart series data labels', () => {
  const data: ChartData = {
    domain: 'category',
    categories: ['A', 'B'],
    series: [
      { id: 'one', values: [1, 2] },
      { id: 'two', values: [3, 4] },
    ],
  };
  const labelTexts = (scene: ReturnType<typeof sceneFor>, seriesIndex: number) =>
    scene.elements
      .filter((element) => element.role === 'data-label' && element.seriesIndex === seriesIndex)
      .map((element) => element.text);

  it('draws labels only for the series that carry them', () => {
    const scene = sceneFor(
      {
        kind: 'column',
        seriesStyles: [{ seriesId: 'two', dataLabels: { visible: true, showValue: true } }],
      },
      data,
    );
    expect(labelTexts(scene, 0)).toEqual([]);
    expect(labelTexts(scene, 1)).toEqual(['3', '4']);
  });

  it('lets a series override or suppress the chart-wide labels', () => {
    const scene = sceneFor(
      {
        kind: 'column',
        dataLabels: { visible: true, showValue: true },
        seriesStyles: [{ seriesId: 'one', dataLabels: { visible: false } }],
      },
      data,
    );
    expect(labelTexts(scene, 0)).toEqual([]);
    expect(labelTexts(scene, 1)).toEqual(['3', '4']);
  });

  it('draws per-series labels on scatter and doughnut charts', () => {
    const scatter = sceneFor(
      {
        kind: 'scatter',
        seriesStyles: [{ seriesId: 'b', dataLabels: { visible: true, showValue: true } }],
      },
      {
        domain: 'xy',
        series: [
          { id: 'a', points: [{ x: 1, y: 1 }] },
          { id: 'b', points: [{ x: 2, y: 7 }] },
        ],
      },
    );
    expect(labelTexts(scatter, 0)).toEqual([]);
    expect(labelTexts(scatter, 1)).toHaveLength(1);

    const doughnut = sceneFor(
      {
        kind: 'doughnut',
        seriesStyles: [{ seriesId: 'two', dataLabels: { visible: true, showValue: true } }],
      },
      data,
    );
    expect(labelTexts(doughnut, 0)).toEqual([]);
    expect(labelTexts(doughnut, 1).length).toBeGreaterThan(0);
  });
});

describe('layoutChart legend default', () => {
  it('places the legend at the bottom when no position is given', () => {
    const scene = layoutChart({
      spec: { kind: 'column' },
      data: {
        domain: 'category',
        categories: ['A'],
        series: [
          { id: 'one', name: 'One', values: [1] },
          { id: 'two', name: 'Two', values: [2] },
        ],
      },
      width: 320,
      height: 200,
    });
    const legend = scene.elements.filter((element) => element.role === 'legend');
    expect(legend.length).toBeGreaterThan(0);
    const textY = legend.filter((element) => element.tag === 'text').map((e) => Number(e.attrs.y));
    expect(Math.min(...textY)).toBeGreaterThan(scene.plot.y + scene.plot.height);
  });
});

describe('layoutUnsupportedChart', () => {
  it('draws a frame, the title, and the given message', () => {
    const scene = layoutUnsupportedChart({
      title: 'Sales',
      message: 'Not supported',
      width: 300,
      height: 150,
    });
    expect(scene.elements.some((element) => element.role === 'frame')).toBe(true);
    expect(scene.elements.find((element) => element.role === 'title')?.text).toBe('Sales');
    expect(scene.elements.find((element) => element.role === 'unsupported')?.text).toBe(
      'Not supported',
    );
  });

  it('omits the title element when the chart has none', () => {
    const scene = layoutUnsupportedChart({ message: 'Not supported', width: 300, height: 150 });
    expect(scene.elements.some((element) => element.role === 'title')).toBe(false);
  });
});

describe('layoutChart engine format codes', () => {
  const dateFormatter = (value: number, code: string | undefined) =>
    code === 'yyyy-mm-dd' ? `D${value}` : undefined;
  const category = (formatCode?: string): ChartData => ({
    domain: 'category',
    categories: ['45000', '45001'],
    ...(formatCode === undefined ? {} : { categoryFormatCode: 'yyyy-mm-dd' }),
    series: [{ id: 's', values: [0.25, 0.5], ...(formatCode ? { formatCode } : {}) }],
  });
  const texts = (scene: ReturnType<typeof sceneFor>, role: string) =>
    scene.elements.filter((element) => element.role === role).map((element) => element.text);

  it('renders date serial categories with the category format code', () => {
    const scene = layoutChart({
      spec: { kind: 'column' },
      data: category('0%'),
      width: 320,
      height: 200,
      formatValue: dateFormatter,
    });
    expect(texts(scene, 'category-label')).toEqual(['D45000', 'D45001']);
    const plain = sceneFor({ kind: 'column' }, category());
    expect(texts(plain, 'category-label')).toEqual(['45000', '45001']);
  });

  it('uses the first series format for value labels and the value axis', () => {
    const scene = sceneFor(
      { kind: 'column', dataLabels: { visible: true, showValue: true } },
      category('0%'),
    );
    expect(texts(scene, 'data-label')).toEqual(['25%', '50%']);
    expect(texts(scene, 'axis-label').some((text) => text?.endsWith('%'))).toBe(true);
  });

  it('lets an explicit numberFormat override the series format', () => {
    const scene = sceneFor(
      {
        kind: 'column',
        dataLabels: { visible: true, showValue: true, numberFormat: '0.00' },
        yAxis: { numberFormat: '0.0' },
      },
      category('0%'),
    );
    expect(texts(scene, 'data-label')).toEqual(['0.25', '0.50']);
    expect(texts(scene, 'axis-label').some((text) => text?.endsWith('%'))).toBe(false);
  });
});
