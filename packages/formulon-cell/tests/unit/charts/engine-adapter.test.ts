import type {
  ChartAxisInfo,
  ChartDataSource,
  ChartFillInfo,
  ChartInfo,
  ChartLabelsInfo,
  ChartLineInfo,
  ChartMarkerInfo,
  ChartPlotInfo,
  ChartSeriesInfo,
  ChartData as EngineChartData,
} from '@libraz/formulon';
import {
  ChartAxisCrosses,
  ChartAxisKind,
  ChartBlanksAs,
  ChartInsertKind,
  ChartLegendPosition,
  ChartLineDash,
  ChartMarkerSymbol,
  ChartPlotType,
  ChartTitleKind,
  ChartGrouping as EngineChartGrouping,
  ChartLabelPosition as EngineChartLabelPosition,
  ChartSeriesDirection as EngineChartSeriesDirection,
} from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import {
  argbFromCssColor,
  CHART_SPEC_FIELDS,
  type ChartEngineOp,
  type ChartNormalizeContext,
  chartAxisId,
  chartDataFromEngine,
  chartInsertOptions,
  chartSpecFromEngine,
  chartSpecOps,
  cssColorFromArgb,
  normalizeChartSpec,
  seriesColorOf,
  unsupportedChartReason,
  withoutSeriesColor,
} from '../../../src/charts/engine-adapter.js';
import type {
  ChartAxisSpec,
  ChartKind,
  ChartLineSpec,
  ChartMarkerSpec,
  ChartPointStyleSpec,
  ChartSeriesStyleSpec,
  ChartSpec,
} from '../../../src/charts/types.js';

// ---------------------------------------------------------------------------
// Hand-built engine models
// ---------------------------------------------------------------------------

const NO_SOURCE: ChartDataSource = { kind: 'none', formula: '', literal: [] };

const fillInfo = (
  color: number | null = null,
  transparency: number | null = null,
): ChartFillInfo => ({
  color,
  transparency,
});

const lineInfo = (overrides: Partial<ChartLineInfo> = {}): ChartLineInfo => ({
  color: null,
  transparency: null,
  widthPt: null,
  dash: null,
  none: false,
  ...overrides,
});

const markerInfo = (overrides: Partial<ChartMarkerInfo> = {}): ChartMarkerInfo => ({
  symbol: ChartMarkerSymbol.Circle,
  size: null,
  fill: fillInfo(),
  line: lineInfo(),
  ...overrides,
});

const labelsInfo = (overrides: Partial<ChartLabelsInfo> = {}): ChartLabelsInfo => ({
  showValue: true,
  showCategory: false,
  showSeries: false,
  showPercent: false,
  showLegendKey: false,
  separator: null,
  position: null,
  numberFormat: null,
  sourceLinked: true,
  fontColor: null,
  fontSizePt: null,
  deleted: false,
  ...overrides,
});

const seriesInfo = (idx: number, overrides: Partial<ChartSeriesInfo> = {}): ChartSeriesInfo => ({
  idx,
  order: idx,
  filtered: false,
  name: NO_SOURCE,
  categories: NO_SOURCE,
  values: NO_SOURCE,
  x: NO_SOURCE,
  y: NO_SOURCE,
  size: NO_SOURCE,
  fill: fillInfo(),
  line: lineInfo(),
  marker: null,
  smooth: null,
  points: [],
  labels: null,
  ...overrides,
});

const axisInfo = (
  id: number,
  kind: ChartAxisKind,
  position: ChartAxisInfo['position'],
  overrides: Partial<ChartAxisInfo> = {},
): ChartAxisInfo => ({
  id,
  kind,
  position,
  deleted: false,
  crossAxis: 0,
  crosses: ChartAxisCrosses.AutoZero,
  crossesAt: null,
  min: null,
  max: null,
  majorUnit: null,
  minorUnit: null,
  logBase: null,
  reversed: false,
  numberFormat: 'General',
  sourceLinked: true,
  title: { kind: ChartTitleKind.None, text: '', formula: '' },
  majorGridlines: false,
  minorGridlines: false,
  tickLabelPosition: 'nextTo',
  ...overrides,
});

const X_ID = 10;
const Y_ID = 20;

const PLOT_TYPE: Record<ChartKind, Pick<ChartPlotInfo, 'type' | 'barDirection'>> = {
  column: { type: ChartPlotType.Bar, barDirection: 'col' },
  bar: { type: ChartPlotType.Bar, barDirection: 'bar' },
  line: { type: ChartPlotType.Line, barDirection: null },
  area: { type: ChartPlotType.Area, barDirection: null },
  pie: { type: ChartPlotType.Pie, barDirection: null },
  doughnut: { type: ChartPlotType.Doughnut, barDirection: null },
  scatter: { type: ChartPlotType.Scatter, barDirection: null },
};

/** The x-role and y-role axes Excel gives each kind (none for pie and doughnut). */
const axesFor = (
  kind: ChartKind,
  x: Partial<ChartAxisInfo> = {},
  y: Partial<ChartAxisInfo> = {},
): ChartAxisInfo[] => {
  switch (kind) {
    case 'pie':
    case 'doughnut':
      return [];
    case 'scatter':
      return [
        axisInfo(X_ID, ChartAxisKind.Value, 'b', { crossAxis: Y_ID, ...x }),
        axisInfo(Y_ID, ChartAxisKind.Value, 'l', { crossAxis: X_ID, ...y }),
      ];
    case 'bar':
      return [
        axisInfo(X_ID, ChartAxisKind.Category, 'l', { crossAxis: Y_ID, ...x }),
        axisInfo(Y_ID, ChartAxisKind.Value, 'b', { crossAxis: X_ID, ...y }),
      ];
    default:
      return [
        axisInfo(X_ID, ChartAxisKind.Category, 'b', { crossAxis: Y_ID, ...x }),
        axisInfo(Y_ID, ChartAxisKind.Value, 'l', { crossAxis: X_ID, ...y }),
      ];
  }
};

const plotInfo = (kind: ChartKind, overrides: Partial<ChartPlotInfo> = {}): ChartPlotInfo => ({
  ...PLOT_TYPE[kind],
  chartexLayout: null,
  grouping: null,
  varyColors: false,
  gapWidth: null,
  overlap: null,
  holeSize: null,
  firstSliceAngle: null,
  axisIds: kind === 'pie' || kind === 'doughnut' ? [] : [X_ID, Y_ID],
  series: [seriesInfo(0), seriesInfo(1)],
  ...overrides,
});

const chartInfo = (
  kind: ChartKind,
  overrides: Partial<ChartInfo> = {},
  plot: Partial<ChartPlotInfo> = {},
): ChartInfo => ({
  objectId: 3,
  name: 'Chart 1',
  description: '',
  isChartex: false,
  isPivotChart: false,
  onChartsheet: false,
  title: { kind: ChartTitleKind.Auto, text: '', formula: '' },
  autoTitleDeleted: false,
  legend: { position: ChartLegendPosition.Bottom, overlay: false, hiddenSeries: [] },
  blanksAs: ChartBlanksAs.Gap,
  plotVisibleOnly: true,
  date1904: false,
  style: null,
  plots: [plotInfo(kind, plot)],
  axes: axesFor(kind),
  hiddenCategories: [],
  source: {
    range: 'Sheet1!$A$1:$C$4',
    direction: EngineChartSeriesDirection.Columns,
    headerRow: true,
    headerColumn: true,
  },
  ...overrides,
});

const read = (info: ChartInfo, data: EngineChartData | null = null): ChartSpec => {
  const spec = chartSpecFromEngine(info, data);
  if (!spec) throw new Error('expected a supported chart');
  return spec;
};

const KINDS: readonly ChartKind[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];

// ---------------------------------------------------------------------------
// The table itself
// ---------------------------------------------------------------------------

describe('CHART_SPEC_FIELDS', () => {
  it('has one entry per ChartSpec field with the kinds of the mapping table', () => {
    const kinds = Object.fromEntries(
      Object.entries(CHART_SPEC_FIELDS).map(([key, entry]) => [key, [...entry.kinds]]),
    );
    const all = [...KINDS];
    expect(kinds).toEqual({
      kind: all,
      grouping: ['column', 'bar', 'line', 'area'],
      scatterStyle: ['scatter'],
      seriesDirection: all,
      emptyCells: all,
      gapWidth: ['column', 'bar'],
      overlap: ['column', 'bar'],
      firstSliceAngle: ['pie', 'doughnut'],
      holeSize: ['doughnut'],
      headerRow: all,
      headerColumn: all,
      title: all,
      palette: all,
      legend: all,
      filter: all,
      dataLabels: all,
      seriesStyles: all,
      varyColors: all,
      xAxis: ['column', 'bar', 'line', 'area', 'scatter'],
      yAxis: ['column', 'bar', 'line', 'area', 'scatter'],
    });
  });

  it('converts colours between CSS hex and opaque ARGB', () => {
    expect(argbFromCssColor('#1A2b3C')).toBe(0xff1a2b3c);
    expect(argbFromCssColor('#abc')).toBe(0xffaabbcc);
    expect(argbFromCssColor('red')).toBeNull();
    expect(cssColorFromArgb(0x801a2b3c)).toBe('#1a2b3c');
    expect(cssColorFromArgb(0xff000001)).toBe('#000001');
  });
});

// ---------------------------------------------------------------------------
// Engine → ChartSpec
// ---------------------------------------------------------------------------

describe('chartSpecFromEngine', () => {
  it('reads only the fields of each kind', () => {
    // Every engine field is populated, so a field missing from the result is kind-gated.
    const marked = seriesInfo(0, {
      fill: fillInfo(0xff4472c4),
      marker: markerInfo(),
      labels: labelsInfo(),
    });
    const expected: Record<ChartKind, string[]> = {
      column: [
        'dataLabels',
        'emptyCells',
        'gapWidth',
        'grouping',
        'headerColumn',
        'headerRow',
        'kind',
        'legend',
        'overlap',
        'seriesDirection',
        'seriesStyles',
        'title',
        'varyColors',
        'xAxis',
        'yAxis',
      ],
      bar: [
        'dataLabels',
        'emptyCells',
        'gapWidth',
        'grouping',
        'headerColumn',
        'headerRow',
        'kind',
        'legend',
        'overlap',
        'seriesDirection',
        'seriesStyles',
        'title',
        'varyColors',
        'xAxis',
        'yAxis',
      ],
      line: [
        'dataLabels',
        'emptyCells',
        'grouping',
        'headerColumn',
        'headerRow',
        'kind',
        'legend',
        'seriesDirection',
        'seriesStyles',
        'title',
        'varyColors',
        'xAxis',
        'yAxis',
      ],
      area: [
        'dataLabels',
        'emptyCells',
        'grouping',
        'headerColumn',
        'headerRow',
        'kind',
        'legend',
        'seriesDirection',
        'seriesStyles',
        'title',
        'varyColors',
        'xAxis',
        'yAxis',
      ],
      pie: [
        'dataLabels',
        'emptyCells',
        'firstSliceAngle',
        'headerColumn',
        'headerRow',
        'kind',
        'legend',
        'seriesDirection',
        'seriesStyles',
        'title',
        'varyColors',
      ],
      doughnut: [
        'dataLabels',
        'emptyCells',
        'firstSliceAngle',
        'headerColumn',
        'headerRow',
        'holeSize',
        'kind',
        'legend',
        'seriesDirection',
        'seriesStyles',
        'title',
        'varyColors',
      ],
      scatter: [
        'dataLabels',
        'emptyCells',
        'headerColumn',
        'headerRow',
        'kind',
        'legend',
        'scatterStyle',
        'seriesDirection',
        'seriesStyles',
        'title',
        'varyColors',
        'xAxis',
        'yAxis',
      ],
    };
    for (const kind of KINDS) {
      const info = chartInfo(
        kind,
        {
          title: { kind: ChartTitleKind.Text, text: 'T', formula: '' },
          // Axes are present even on circular kinds to show they are not read there.
          axes: axesFor(kind === 'pie' || kind === 'doughnut' ? 'column' : kind),
        },
        {
          grouping: EngineChartGrouping.Clustered,
          gapWidth: 150,
          overlap: 0,
          holeSize: 50,
          firstSliceAngle: 0,
          axisIds: [X_ID, Y_ID],
          series: [marked],
        },
      );
      expect(Object.keys(read(info)).sort(), kind).toEqual(expected[kind]);
      const style = read(info).seriesStyles?.[0];
      expect(style?.marker === undefined, kind).toBe(kind !== 'line' && kind !== 'scatter');
    }
  });

  it('reads kind from the plot type and bar direction, grouping with standard as clustered', () => {
    const rows: [Partial<ChartPlotInfo>, ChartKind, ChartSpec['grouping']][] = [
      [
        { type: ChartPlotType.Bar, barDirection: 'col', grouping: EngineChartGrouping.Clustered },
        'column',
        'clustered',
      ],
      [
        { type: ChartPlotType.Bar, barDirection: 'bar', grouping: EngineChartGrouping.Stacked },
        'bar',
        'stacked',
      ],
      [{ type: ChartPlotType.Bar, barDirection: null, grouping: null }, 'column', 'clustered'],
      [{ type: ChartPlotType.Line, grouping: EngineChartGrouping.Standard }, 'line', 'clustered'],
      [
        { type: ChartPlotType.Line, grouping: EngineChartGrouping.PercentStacked },
        'line',
        'percent-stacked',
      ],
      [{ type: ChartPlotType.Area, grouping: EngineChartGrouping.Stacked }, 'area', 'stacked'],
      [{ type: ChartPlotType.Area, grouping: null }, 'area', 'clustered'],
      [{ type: ChartPlotType.Pie, grouping: null }, 'pie', undefined],
      [{ type: ChartPlotType.Doughnut, grouping: null }, 'doughnut', undefined],
      [{ type: ChartPlotType.Scatter, grouping: null }, 'scatter', undefined],
    ];
    for (const [plot, kind, grouping] of rows) {
      const spec = read(chartInfo('column', {}, plot));
      expect(spec.kind).toBe(kind);
      expect(spec.grouping).toBe(grouping);
    }
  });

  it('reports charts it cannot project', () => {
    const rows: [ChartInfo, string][] = [
      [chartInfo('column', { isChartex: true }), 'chartex'],
      [chartInfo('column', { plots: [plotInfo('column'), plotInfo('line')] }), 'multi-plot'],
      [chartInfo('column', { plots: [] }), 'multi-plot'],
      [chartInfo('column', {}, { type: ChartPlotType.Bar3D }), 'plot-type'],
      [chartInfo('column', {}, { type: ChartPlotType.Radar }), 'plot-type'],
      [chartInfo('column', {}, { type: ChartPlotType.Bubble }), 'plot-type'],
      [chartInfo('column', {}, { type: ChartPlotType.OfPie }), 'plot-type'],
      [chartInfo('column', {}, { type: ChartPlotType.Stock }), 'plot-type'],
    ];
    for (const [info, reason] of rows) {
      expect(unsupportedChartReason(info)).toBe(reason);
      expect(chartSpecFromEngine(info)).toBeNull();
    }
    expect(unsupportedChartReason(chartInfo('doughnut'))).toBeNull();
  });

  it('reads the source split, undefined without a source', () => {
    const spec = read(
      chartInfo('line', {
        source: {
          range: 'Sheet1!$A$1:$D$3',
          direction: EngineChartSeriesDirection.Rows,
          headerRow: false,
          headerColumn: true,
        },
      }),
    );
    expect([spec.seriesDirection, spec.headerRow, spec.headerColumn]).toEqual([
      'rows',
      false,
      true,
    ]);
    const none = read(chartInfo('line', { source: null }));
    expect([none.seriesDirection, none.headerRow, none.headerColumn]).toEqual([
      undefined,
      undefined,
      undefined,
    ]);
  });

  it('reads blanksAs per kind, absent as zero and span only where lines join', () => {
    const expected: Record<ChartKind, [string, string, string, string]> = {
      // Absent, Gap, Zero, Span
      column: ['zero', 'gap', 'zero', 'gap'],
      bar: ['zero', 'gap', 'zero', 'gap'],
      line: ['zero', 'gap', 'zero', 'span'],
      area: ['zero', 'gap', 'zero', 'span'],
      pie: ['zero', 'gap', 'zero', 'gap'],
      doughnut: ['zero', 'gap', 'zero', 'gap'],
      scatter: ['zero', 'gap', 'zero', 'span'],
    };
    const blanks = [
      ChartBlanksAs.Absent,
      ChartBlanksAs.Gap,
      ChartBlanksAs.Zero,
      ChartBlanksAs.Span,
    ];
    for (const kind of KINDS) {
      const actual = blanks.map((blanksAs) => read(chartInfo(kind, { blanksAs })).emptyCells);
      expect(actual, kind).toEqual(expected[kind]);
    }
  });

  it('reads plot options of the kinds that have them', () => {
    const plot = {
      gapWidth: 80,
      overlap: -20,
      firstSliceAngle: 90,
      holeSize: 65,
      varyColors: true,
    };
    const pick = (spec: ChartSpec) => [
      spec.gapWidth,
      spec.overlap,
      spec.firstSliceAngle,
      spec.holeSize,
      spec.varyColors,
    ];
    expect(pick(read(chartInfo('column', {}, plot)))).toEqual([
      80,
      -20,
      undefined,
      undefined,
      true,
    ]);
    expect(pick(read(chartInfo('bar', {}, plot)))).toEqual([80, -20, undefined, undefined, true]);
    expect(pick(read(chartInfo('line', {}, plot)))).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      true,
    ]);
    expect(pick(read(chartInfo('pie', {}, plot)))).toEqual([
      undefined,
      undefined,
      90,
      undefined,
      true,
    ]);
    expect(pick(read(chartInfo('doughnut', {}, plot)))).toEqual([
      undefined,
      undefined,
      90,
      65,
      true,
    ]);
    expect(pick(read(chartInfo('scatter', {}, { varyColors: false })))).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      false,
    ]);
    // Absent elements read as undefined.
    expect(pick(read(chartInfo('doughnut')))).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      false,
    ]);
  });

  it('reads the title by kind', () => {
    const data = { titleText: 'Sales 2026' } as EngineChartData;
    const title = (kind: ChartTitleKind, text = '', formula = '') => ({ kind, text, formula });
    expect(read(chartInfo('pie', { title: title(ChartTitleKind.Text, 'Hi') })).title).toBe('Hi');
    expect(
      read(chartInfo('pie', { title: title(ChartTitleKind.Ref, '', 'Sheet1!$A$1') }), data).title,
    ).toBe('Sales 2026');
    expect(
      read(chartInfo('pie', { title: title(ChartTitleKind.Ref, '', 'Sheet1!$A$1') })).title,
    ).toBeUndefined();
    expect(read(chartInfo('pie', { title: title(ChartTitleKind.Auto) })).title).toEqual({
      auto: true,
    });
    expect(read(chartInfo('pie', { title: title(ChartTitleKind.None) })).title).toBe('');
  });

  it('reads the legend position, top-right as right', () => {
    const rows: [ChartLegendPosition, ChartSpec['legend']][] = [
      [ChartLegendPosition.None, { visible: false }],
      [ChartLegendPosition.Bottom, { visible: true, position: 'bottom' }],
      [ChartLegendPosition.Top, { visible: true, position: 'top' }],
      [ChartLegendPosition.Left, { visible: true, position: 'left' }],
      [ChartLegendPosition.Right, { visible: true, position: 'right' }],
      [ChartLegendPosition.TopRight, { visible: true, position: 'right' }],
    ];
    for (const [position, legend] of rows) {
      const info = chartInfo('column', { legend: { position, overlay: true, hiddenSeries: [] } });
      expect(read(info).legend).toEqual(legend);
    }
  });

  it('reads filtered series and hidden categories as ids', () => {
    const series = [
      seriesInfo(0),
      seriesInfo(1, { filtered: true }),
      seriesInfo(4, { filtered: true }),
    ];
    expect(read(chartInfo('column', { hiddenCategories: [0, 3] }, { series })).filter).toEqual({
      hiddenSeriesIds: ['series-1', 'series-4'],
      hiddenCategoryIds: ['category-0', 'category-3'],
    });
    expect(read(chartInfo('scatter', { hiddenCategories: [2] })).filter).toEqual({
      hiddenCategoryIds: ['point-2'],
    });
    expect(read(chartInfo('pie')).filter).toBeUndefined();
  });

  it('reads labels shared by every series into dataLabels and differing ones per series', () => {
    const full = labelsInfo({
      showValue: false,
      showCategory: true,
      showSeries: true,
      showPercent: true,
      showLegendKey: true,
      separator: '; ',
      position: EngineChartLabelPosition.OutsideEnd,
      numberFormat: '0.0',
      fontColor: 0xff112233,
      fontSizePt: 9,
    });
    const fullSpec = {
      visible: true,
      showSeriesName: true,
      showCategoryName: true,
      showValue: false,
      showPercentage: true,
      showLegendKey: true,
      separator: '; ',
      position: 'outside-end',
      numberFormat: '0.0',
      color: '#112233',
      fontSize: 9,
    };
    const shared = read(
      chartInfo(
        'column',
        {},
        { series: [seriesInfo(0, { labels: full }), seriesInfo(1, { labels: full })] },
      ),
    );
    expect(shared.dataLabels).toEqual(fullSpec);
    expect(shared.seriesStyles).toBeUndefined();

    const mixed = read(
      chartInfo(
        'column',
        {},
        { series: [seriesInfo(0, { labels: labelsInfo() }), seriesInfo(1, { labels: full })] },
      ),
    );
    expect(mixed.dataLabels).toBeUndefined();
    expect(mixed.seriesStyles).toEqual([
      {
        seriesId: 'series-0',
        dataLabels: {
          visible: true,
          showSeriesName: false,
          showCategoryName: false,
          showValue: true,
          showPercentage: false,
          showLegendKey: false,
        },
      },
      { seriesId: 'series-1', dataLabels: fullSpec },
    ]);

    // Deleted labels and no labels are both hidden.
    const hidden = read(
      chartInfo(
        'pie',
        {},
        {
          series: [seriesInfo(0, { labels: labelsInfo({ deleted: true }) }), seriesInfo(1)],
        },
      ),
    );
    expect(hidden.dataLabels).toBeUndefined();
    expect(hidden.seriesStyles).toBeUndefined();

    const partly = read(
      chartInfo('line', {}, { series: [seriesInfo(0, { labels: labelsInfo() }), seriesInfo(1)] }),
    );
    expect(partly.dataLabels).toBeUndefined();
    expect(
      partly.seriesStyles?.map((style) => [style.seriesId, style.dataLabels?.visible]),
    ).toEqual([['series-0', true]]);
  });

  // Pairwise rows (coverwise, strength 2, seed 1) over colour x transparency x width x dash.
  it.each([
    ['set', 'null', 'null', 'null'],
    ['null', 'set', 'null', 'Dot'],
    ['set', 'set', 'null', 'Other'],
    ['set', 'set', 'set', 'Dash'],
    ['null', 'null', 'set', 'DashDot'],
    ['null', 'null', 'set', 'Other'],
    ['null', 'null', 'null', 'Dash'],
    ['set', 'set', 'null', 'Solid'],
    ['set', 'set', 'null', 'DashDot'],
    ['null', 'set', 'null', 'null'],
    ['set', 'set', 'null', 'Dot'],
    ['set', 'set', 'set', 'null'],
    ['null', 'null', 'set', 'Solid'],
    ['set', 'set', 'set', 'Dot'],
    ['set', 'null', 'set', 'Dot'],
  ] as const)(
    'reads a series line (colour %s, transparency %s, width %s, dash %s)',
    (color, transparency, width, dash) => {
      const dashes = {
        null: [null, undefined],
        Solid: [ChartLineDash.Solid, 'solid'],
        Dash: [ChartLineDash.Dash, 'dash'],
        Dot: [ChartLineDash.Dot, 'dot'],
        DashDot: [ChartLineDash.DashDot, 'dash-dot'],
        Other: [ChartLineDash.Other, undefined],
      } as const;
      const line = lineInfo({
        color: color === 'set' ? 0xff112233 : null,
        transparency: transparency === 'set' ? 25 : null,
        widthPt: width === 'set' ? 2.25 : null,
        dash: dashes[dash][0],
      });
      const expected: ChartLineSpec = {
        ...(color === 'set' ? { color: '#112233' } : {}),
        ...(transparency === 'set' ? { transparency: 25 } : {}),
        ...(width === 'set' ? { width: 2.25 } : {}),
        ...(dashes[dash][1] ? { dash: dashes[dash][1] } : {}),
      };
      const series = [seriesInfo(0, { fill: fillInfo(0xffed7d31, 10), line })];
      expect(read(chartInfo('area', {}, { series })).seriesStyles).toEqual([
        {
          seriesId: 'series-0',
          fill: { color: '#ed7d31', transparency: 10 },
          ...(Object.keys(expected).length > 0 ? { line: expected } : {}),
        },
      ]);
    },
  );

  it('reads markers: none as hidden, automatic without a shape', () => {
    const rows: [ChartMarkerInfo | null, ChartMarkerSpec | undefined][] = [
      [null, undefined],
      [markerInfo({ symbol: ChartMarkerSymbol.None }), { visible: false }],
      [markerInfo({ symbol: ChartMarkerSymbol.Auto }), { visible: true }],
      [markerInfo({ symbol: ChartMarkerSymbol.Picture }), { visible: true, shape: 'picture' }],
      [
        markerInfo({ symbol: ChartMarkerSymbol.Circle, size: 7 }),
        { visible: true, shape: 'circle', size: 7 },
      ],
      [markerInfo({ symbol: ChartMarkerSymbol.Square }), { visible: true, shape: 'square' }],
      [markerInfo({ symbol: ChartMarkerSymbol.Diamond }), { visible: true, shape: 'diamond' }],
      [
        markerInfo({
          symbol: ChartMarkerSymbol.Triangle,
          fill: fillInfo(0xff00ff00),
          line: lineInfo({ color: 0xff0000ff, widthPt: 0.75 }),
        }),
        {
          visible: true,
          shape: 'triangle',
          fill: { color: '#00ff00' },
          line: { color: '#0000ff', width: 0.75 },
        },
      ],
    ];
    for (const [marker, expected] of rows) {
      const series = [seriesInfo(0, { fill: fillInfo(0xff000000), marker })];
      expect(read(chartInfo('line', {}, { series })).seriesStyles?.[0]?.marker).toEqual(expected);
      expect(read(chartInfo('scatter', {}, { series })).seriesStyles?.[0]?.marker).toEqual(
        expected,
      );
      expect(read(chartInfo('column', {}, { series })).seriesStyles?.[0]?.marker).toBeUndefined();
    }
  });

  it.each([
    ['circle', ChartMarkerSymbol.Circle],
    ['square', ChartMarkerSymbol.Square],
    ['diamond', ChartMarkerSymbol.Diamond],
    ['triangle', ChartMarkerSymbol.Triangle],
    ['x', ChartMarkerSymbol.X],
    ['star', ChartMarkerSymbol.Star],
    ['plus', ChartMarkerSymbol.Plus],
    ['dot', ChartMarkerSymbol.Dot],
    ['dash', ChartMarkerSymbol.Dash],
  ] as const)('reads and writes a %s marker as its own symbol', (shape, symbol) => {
    const info = chartInfo(
      'line',
      {},
      { series: [seriesInfo(0, { marker: markerInfo({ symbol }) })] },
    );
    const spec = read(info);
    expect(spec.seriesStyles?.[0]?.marker).toEqual({ visible: true, shape });
    expect(ops(info, spec, spec)).toEqual([]);
    const hidden: ChartSpec = {
      ...spec,
      seriesStyles: [{ seriesId: 'series-0', marker: { visible: false } }],
    };
    expect(ops(info, hidden, spec)).toEqual([
      { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { marker: { symbol } } },
    ]);
  });

  it('reads and writes a smoothed line on line and scatter series only', () => {
    const series = [
      seriesInfo(0, { smooth: true }),
      seriesInfo(1, { smooth: false }),
      seriesInfo(2),
    ];
    for (const kind of ['line', 'scatter'] as const) {
      const info = chartInfo(kind, {}, { series });
      const spec = read(info);
      // An unsmoothed line, written as val="0" or not at all, reads as absent.
      expect(spec.seriesStyles).toEqual([{ seriesId: 'series-0', smooth: true }]);
      expect(ops(info, spec, spec)).toEqual([]);
      const toggled: ChartSpec = {
        ...spec,
        seriesStyles: [
          { seriesId: 'series-0', smooth: false },
          { seriesId: 'series-2', smooth: true },
        ],
      };
      expect(ops(info, spec, toggled)).toEqual([
        { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { smooth: false } },
        { op: 'setChartSeries', plot: 0, seriesIdx: 2, patch: { smooth: true } },
      ]);
    }
    const column = chartInfo('column', {}, { series });
    expect(read(column).seriesStyles).toBeUndefined();
    expect(
      ops(column, read(column), {
        ...read(column),
        seriesStyles: [{ seriesId: 'series-0', smooth: true }],
      }),
    ).toEqual([]);
  });

  it('keeps a picture marker, which the engine cannot write', () => {
    const info = chartInfo(
      'line',
      {},
      {
        series: [seriesInfo(0, { marker: markerInfo({ symbol: ChartMarkerSymbol.Picture }) })],
      },
    );
    const spec = read(info);
    expect(spec.seriesStyles?.[0]?.marker).toEqual({ visible: true, shape: 'picture' });
    const square: ChartSpec = {
      ...spec,
      seriesStyles: [{ seriesId: 'series-0', marker: { visible: true, shape: 'square' } }],
    };
    expect(ops(info, square, spec)).toEqual([]);
  });

  it('reads formatted points, with markers only on line and scatter', () => {
    const series = [
      seriesInfo(2, {
        points: [
          { idx: 1, fill: fillInfo(0xffff0000), line: lineInfo(), marker: markerInfo() },
          { idx: 4, fill: fillInfo(), line: lineInfo({ dash: ChartLineDash.Dash }), marker: null },
          { idx: 5, fill: fillInfo(), line: lineInfo(), marker: null },
        ],
      }),
    ];
    expect(read(chartInfo('line', {}, { series })).seriesStyles).toEqual([
      {
        seriesId: 'series-2',
        points: [
          {
            pointId: 'category-1',
            fill: { color: '#ff0000' },
            marker: { visible: true, shape: 'circle' },
          },
          { pointId: 'category-4', line: { dash: 'dash' } },
        ],
      },
    ]);
    expect(read(chartInfo('pie', {}, { series })).seriesStyles).toEqual([
      {
        seriesId: 'series-2',
        points: [
          { pointId: 'category-1', fill: { color: '#ff0000' } },
          { pointId: 'category-4', line: { dash: 'dash' } },
        ],
      },
    ]);
  });

  it('reads no series fill on line and scatter, whose colour is in the line and markers', () => {
    const series = [
      seriesInfo(0, {
        fill: fillInfo(0xff112233, 40),
        line: lineInfo({ color: 0xff445566, transparency: 10 }),
        marker: markerInfo({
          fill: fillInfo(0xff445566, 30),
          line: lineInfo({ color: 0xff445566 }),
        }),
      }),
    ];
    for (const kind of ['line', 'scatter'] as const) {
      expect(read(chartInfo(kind, {}, { series })).seriesStyles, kind).toEqual([
        {
          seriesId: 'series-0',
          line: { color: '#445566', transparency: 10 },
          marker: {
            visible: true,
            shape: 'circle',
            fill: { color: '#445566', transparency: 30 },
            line: { color: '#445566' },
          },
        },
      ]);
    }
    expect(read(chartInfo('area', {}, { series })).seriesStyles?.[0]?.fill).toEqual({
      color: '#112233',
      transparency: 40,
    });
  });

  it('never reads a palette', () => {
    for (const kind of KINDS) {
      const series = [
        seriesInfo(0, { fill: fillInfo(0xff4472c4) }),
        seriesInfo(1, { fill: fillInfo(0xffed7d31) }),
      ];
      expect(read(chartInfo(kind, {}, { series })).palette).toBeUndefined();
    }
  });

  it('reads the scatter style shared by every series', () => {
    const lineless = lineInfo({ none: true });
    const rows: [Partial<ChartSeriesInfo>[], ChartSpec['scatterStyle']][] = [
      [[{ line: lineless, marker: markerInfo() }], 'marker'],
      [[{ line: lineless, marker: null }], 'marker'],
      [[{ marker: markerInfo({ symbol: ChartMarkerSymbol.None }) }], 'line'],
      [[{ marker: null }], 'lineMarker'],
      [[{ marker: markerInfo({ symbol: ChartMarkerSymbol.Auto }) }], 'lineMarker'],
      [[{ line: lineless }, { marker: null }], undefined],
      [[], undefined],
    ];
    for (const [series, style] of rows) {
      const info = chartInfo(
        'scatter',
        {},
        { series: series.map((entry, idx) => seriesInfo(idx, entry)) },
      );
      expect(read(info).scatterStyle).toBe(style);
    }
  });

  it('assigns axis roles by kind', () => {
    const marked = (kind: ChartKind) =>
      read(
        chartInfo(kind, {
          axes: axesFor(
            kind,
            { title: { kind: ChartTitleKind.Text, text: 'X', formula: '' } },
            {
              title: { kind: ChartTitleKind.Text, text: 'Y', formula: '' },
            },
          ),
        }),
      );
    for (const kind of ['column', 'bar', 'line', 'area', 'scatter'] as const) {
      expect([marked(kind).xAxis?.title, marked(kind).yAxis?.title], kind).toEqual(['X', 'Y']);
    }
    // A date axis plays the category role; axes of another plot are ignored.
    const dated = chartInfo('line', {
      axes: [
        axisInfo(99, ChartAxisKind.Category, 'b', {
          title: { kind: ChartTitleKind.Text, text: 'other', formula: '' },
        }),
        axisInfo(Y_ID, ChartAxisKind.Value, 'l'),
        axisInfo(X_ID, ChartAxisKind.Date, 'b', {
          title: { kind: ChartTitleKind.Text, text: 'date', formula: '' },
        }),
      ],
    });
    expect(read(dated).xAxis?.title).toBe('date');
    expect(chartAxisId(dated, 'x')).toBe(X_ID);
    expect(chartAxisId(dated, 'y')).toBe(Y_ID);
    // A scatter chart's x axis is the horizontal value axis wherever it sits in the list.
    const scatter = chartInfo('scatter', {
      axes: [axisInfo(Y_ID, ChartAxisKind.Value, 'r'), axisInfo(X_ID, ChartAxisKind.Value, 't')],
    });
    expect(chartAxisId(scatter, 'x')).toBe(X_ID);
    expect(chartAxisId(scatter, 'y')).toBe(Y_ID);
    expect(chartAxisId(chartInfo('pie'), 'x')).toBeNull();
    expect(read(chartInfo('pie', { axes: axesFor('column') })).xAxis).toBeUndefined();
  });

  it('reads every axis field', () => {
    const data = {
      axisTitleTexts: [{ axisId: Y_ID, text: 'Units' }],
    } as unknown as EngineChartData;
    const info = chartInfo('column', {
      axes: axesFor(
        'column',
        {
          min: 1,
          max: 9,
          majorUnit: 2,
          logBase: 10,
          reversed: true,
          numberFormat: '0.00',
          sourceLinked: false,
          majorGridlines: true,
          tickLabelPosition: 'high',
          title: { kind: ChartTitleKind.Text, text: 'Month', formula: '' },
          crosses: ChartAxisCrosses.Max,
        },
        {
          deleted: true,
          crosses: ChartAxisCrosses.At,
          crossesAt: 3,
          title: { kind: ChartTitleKind.Ref, text: '', formula: 'Sheet1!$B$1' },
        },
      ),
    });
    expect(read(info, data).xAxis).toEqual({
      visible: true,
      logBase: 10,
      labelPosition: 'high',
      reverseOrder: true,
      // The y axis crosses the x axis at category 3.
      crossing: 3,
      title: 'Month',
      min: 1,
      max: 9,
      majorUnit: 2,
      numberFormat: '0.00',
      majorGridlines: true,
    });
    expect(read(info, data).yAxis).toEqual({
      visible: false,
      labelPosition: 'nextTo',
      reverseOrder: false,
      crossing: 'max',
      title: 'Units',
      majorGridlines: false,
    });
    expect(read(info).yAxis?.title).toBeUndefined();
  });

  // Pairwise rows (coverwise, strength 2, seed 1) over kind x crosses of the y-role axis x
  // tickLabelPosition x sourceLinked x deleted of the x-role axis.
  it.each([
    ['column', 'Max', 'null', 'false', 'false'],
    ['bar', 'At', 'other', 'false', 'true'],
    ['column', 'AutoZero', 'high', 'true', 'true'],
    ['line', 'Min', 'null', 'true', 'true'],
    ['area', 'At', 'low', 'true', 'false'],
    ['column', 'Absent', 'low', 'false', 'true'],
    ['scatter', 'Min', 'none', 'false', 'false'],
    ['area', 'Max', 'nextTo', 'false', 'true'],
    ['scatter', 'At', 'high', 'false', 'true'],
    ['bar', 'Max', 'none', 'true', 'false'],
    ['line', 'AutoZero', 'other', 'false', 'false'],
    ['area', 'Min', 'none', 'false', 'true'],
    ['column', 'Absent', 'nextTo', 'true', 'false'],
    ['scatter', 'Max', 'other', 'true', 'false'],
    ['line', 'Max', 'high', 'false', 'false'],
    ['scatter', 'Absent', 'null', 'false', 'true'],
    ['area', 'AutoZero', 'null', 'false', 'false'],
    ['line', 'At', 'nextTo', 'false', 'true'],
    ['bar', 'AutoZero', 'nextTo', 'true', 'false'],
    ['scatter', 'AutoZero', 'low', 'true', 'false'],
    ['column', 'Min', 'other', 'true', 'false'],
    ['scatter', 'AutoZero', 'nextTo', 'false', 'true'],
    ['scatter', 'AutoZero', 'none', 'true', 'true'],
    ['bar', 'At', 'null', 'true', 'false'],
    ['bar', 'Absent', 'high', 'false', 'true'],
    ['line', 'Min', 'low', 'true', 'false'],
    ['column', 'At', 'none', 'true', 'true'],
    ['area', 'AutoZero', 'high', 'false', 'true'],
    ['bar', 'Min', 'high', 'false', 'false'],
    ['line', 'Absent', 'none', 'false', 'true'],
    ['area', 'Absent', 'other', 'true', 'false'],
    ['bar', 'Max', 'low', 'true', 'true'],
    ['line', 'Min', 'nextTo', 'false', 'false'],
  ] as const)(
    'reads a %s x axis (crossed %s, labels %s, linked %s, deleted %s)',
    (kind, crosses, labels, linked, deleted) => {
      const crossing = {
        Absent: undefined,
        AutoZero: 'auto',
        Min: 'min',
        Max: 'max',
        At: 4,
      } as const;
      const tickLabelPosition = {
        null: null,
        nextTo: 'nextTo',
        low: 'low',
        high: 'high',
        none: 'none',
        other: 'cross',
      };
      const info = chartInfo(kind, {
        axes: axesFor(
          kind,
          {
            tickLabelPosition: tickLabelPosition[labels],
            sourceLinked: linked === 'true',
            numberFormat: '0.0',
            deleted: deleted === 'true',
          },
          { crosses: ChartAxisCrosses[crosses], crossesAt: crosses === 'At' ? 4 : null },
        ),
      });
      const expected: ChartAxisSpec = {
        visible: deleted === 'false',
        ...(labels === 'null' || labels === 'other' ? {} : { labelPosition: labels }),
        reverseOrder: false,
        ...(crossing[crosses] === undefined ? {} : { crossing: crossing[crosses] }),
        ...(linked === 'true' ? {} : { numberFormat: '0.0' }),
        majorGridlines: false,
      };
      expect(read(info).xAxis).toEqual(expected);
    },
  );

  it('reads a line that is not drawn and a hidden marker without their properties', () => {
    const hidden = seriesInfo(0, {
      line: lineInfo({ none: true, widthPt: 2, dash: ChartLineDash.Dash, color: 0xff112233 }),
      marker: markerInfo({
        symbol: ChartMarkerSymbol.None,
        size: 7,
        fill: fillInfo(0xffaa0000),
        line: lineInfo({ color: 0xff000000, widthPt: 1 }),
      }),
      points: [
        {
          idx: 1,
          fill: fillInfo(0xff00aa00),
          line: lineInfo({ none: true, widthPt: 1.5 }),
          marker: markerInfo({ symbol: ChartMarkerSymbol.None, size: 9 }),
        } as ChartSeriesInfo['points'][number],
      ],
    });
    const spec = read(chartInfo('scatter', {}, { series: [hidden] }));
    expect(spec.seriesStyles).toEqual([
      {
        seriesId: 'series-0',
        marker: { visible: false },
        points: [{ pointId: 'point-1', fill: { color: '#00aa00' }, marker: { visible: false } }],
      },
    ]);
  });
});

// ---------------------------------------------------------------------------
// ChartSpec diff → operations
// ---------------------------------------------------------------------------

const ops = (info: ChartInfo, before: ChartSpec, after: ChartSpec): ChartEngineOp[] =>
  chartSpecOps(info, before, after);

/** The normalized line of a spec's first series. */
const normalizeChartSpecLine = (spec: ChartSpec) =>
  normalizeChartSpec(spec, {
    direction: 'columns',
    headerRow: true,
    headerColumn: false,
    seriesIds: ['series-0'],
    accents: [],
    styled: false,
  }).seriesStyles?.[0]?.line;

describe('chartSpecOps', () => {
  it('emits nothing for an unchanged spec', () => {
    for (const kind of KINDS) {
      const info = chartInfo(kind, {}, { series: [seriesInfo(0, { labels: labelsInfo() })] });
      const spec = read(info);
      expect(ops(info, spec, spec), kind).toEqual([]);
    }
  });

  it('changes the type with grouping only for grouped kinds, standard for clustered line and area', () => {
    const info = chartInfo('column');
    const rows: [ChartSpec, ChartEngineOp][] = [
      [
        { kind: 'line' },
        {
          op: 'setChartType',
          plot: 0,
          kind: ChartInsertKind.Line,
          grouping: EngineChartGrouping.Standard,
        },
      ],
      [
        { kind: 'area', grouping: 'stacked' },
        {
          op: 'setChartType',
          plot: 0,
          kind: ChartInsertKind.Area,
          grouping: EngineChartGrouping.Stacked,
        },
      ],
      [
        { kind: 'bar', grouping: 'percent-stacked' },
        {
          op: 'setChartType',
          plot: 0,
          kind: ChartInsertKind.Bar,
          grouping: EngineChartGrouping.PercentStacked,
        },
      ],
      [{ kind: 'pie' }, { op: 'setChartType', plot: 0, kind: ChartInsertKind.Pie }],
      [
        { kind: 'doughnut', grouping: 'stacked' },
        { op: 'setChartType', plot: 0, kind: ChartInsertKind.Doughnut },
      ],
      [{ kind: 'scatter' }, { op: 'setChartType', plot: 0, kind: ChartInsertKind.Scatter }],
      [
        { kind: 'column', grouping: 'stacked' },
        {
          op: 'setChartType',
          plot: 0,
          kind: ChartInsertKind.Column,
          grouping: EngineChartGrouping.Stacked,
        },
      ],
    ];
    for (const [after, op] of rows) {
      expect(ops(info, { kind: 'column', grouping: 'clustered' }, after)).toEqual([op]);
    }
    // Omitted grouping is clustered, and grouping on a kind without it is not a change.
    expect(ops(info, { kind: 'column' }, { kind: 'column', grouping: 'clustered' })).toEqual([]);
    expect(ops(chartInfo('pie'), { kind: 'pie' }, { kind: 'pie', grouping: 'stacked' })).toEqual(
      [],
    );
  });

  it('rewrites the fields a type change rebuilds from the new spec', () => {
    const info = chartInfo('line');
    const before: ChartSpec = {
      kind: 'line',
      title: 'Kept',
      varyColors: false,
      xAxis: { title: 'Month' },
      seriesStyles: [{ seriesId: 'series-0', fill: { color: '#112233' } }],
    };
    const after: ChartSpec = { ...before, kind: 'column', gapWidth: 219 };
    expect(ops(info, before, after)).toEqual([
      {
        op: 'setChartType',
        plot: 0,
        kind: ChartInsertKind.Column,
        grouping: EngineChartGrouping.Clustered,
      },
      { op: 'setChartPlot', plot: 0, patch: { gapWidth: 219, varyColors: false } },
      { op: 'setChartAxis', axis: 'x', patch: { title: 'Month' } },
      { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { fill: { color: 0xff112233 } } },
    ]);
  });

  it('rebuilds the series from the source when the split changes', () => {
    const info = chartInfo('column');
    const before: ChartSpec = {
      kind: 'column',
      seriesDirection: 'columns',
      headerRow: true,
      headerColumn: true,
    };
    expect(ops(info, before, { ...before, seriesDirection: 'rows' })).toEqual([
      {
        op: 'setChartSource',
        source: 'Sheet1!$A$1:$C$4',
        direction: EngineChartSeriesDirection.Rows,
        headerRow: true,
        headerColumn: true,
      },
    ]);
    expect(
      ops(info, before, { kind: 'column', seriesDirection: 'auto', headerRow: false }),
    ).toEqual([
      {
        op: 'setChartSource',
        source: 'Sheet1!$A$1:$C$4',
        direction: EngineChartSeriesDirection.Auto,
        headerRow: false,
      },
    ]);
    expect(
      ops(chartInfo('column', { source: null }), before, { ...before, seriesDirection: 'rows' }),
    ).toEqual([]);
  });

  it('writes blanks per kind, span as gap where lines do not join', () => {
    const expected: Record<ChartKind, [ChartBlanksAs, ChartBlanksAs, ChartBlanksAs]> = {
      // gap, zero, span
      column: [ChartBlanksAs.Gap, ChartBlanksAs.Zero, ChartBlanksAs.Gap],
      bar: [ChartBlanksAs.Gap, ChartBlanksAs.Zero, ChartBlanksAs.Gap],
      line: [ChartBlanksAs.Gap, ChartBlanksAs.Zero, ChartBlanksAs.Span],
      area: [ChartBlanksAs.Gap, ChartBlanksAs.Zero, ChartBlanksAs.Span],
      pie: [ChartBlanksAs.Gap, ChartBlanksAs.Zero, ChartBlanksAs.Gap],
      doughnut: [ChartBlanksAs.Gap, ChartBlanksAs.Zero, ChartBlanksAs.Gap],
      scatter: [ChartBlanksAs.Gap, ChartBlanksAs.Zero, ChartBlanksAs.Span],
    };
    for (const kind of KINDS) {
      const info = chartInfo(kind, { plotVisibleOnly: null });
      const written = (['gap', 'zero', 'span'] as const).map((emptyCells) =>
        ops(info, { kind }, { kind, emptyCells }),
      );
      expect(written, kind).toEqual(
        expected[kind].map((blanksAs) => [
          { op: 'setChartDisplay', blanksAs, plotVisibleOnly: null },
        ]),
      );
    }
    // Span and gap are the same on a column chart; omitting the mode keeps it.
    expect(
      ops(
        chartInfo('column'),
        { kind: 'column', emptyCells: 'gap' },
        { kind: 'column', emptyCells: 'span' },
      ),
    ).toEqual([]);
    expect(ops(chartInfo('line'), { kind: 'line', emptyCells: 'gap' }, { kind: 'line' })).toEqual(
      [],
    );
  });

  it('writes plot options only to the kinds that have them', () => {
    const options = {
      gapWidth: 40,
      overlap: 25,
      firstSliceAngle: 45,
      holeSize: 70,
      varyColors: true,
    };
    const expected: Record<ChartKind, object> = {
      column: { gapWidth: 40, overlap: 25, varyColors: true },
      bar: { gapWidth: 40, overlap: 25, varyColors: true },
      line: { varyColors: true },
      area: { varyColors: true },
      pie: { firstSliceAngle: 45, varyColors: true },
      doughnut: { firstSliceAngle: 45, holeSize: 70, varyColors: true },
      scatter: { varyColors: true },
    };
    for (const kind of KINDS) {
      expect(ops(chartInfo(kind), { kind }, { kind, ...options }), kind).toEqual([
        { op: 'setChartPlot', plot: 0, patch: expected[kind] },
      ]);
    }
    // Clearing a plot option has no engine form and keeps the current value.
    expect(ops(chartInfo('column'), { kind: 'column', gapWidth: 40 }, { kind: 'column' })).toEqual(
      [],
    );
  });

  it('writes an automatic title as the automatic kind, once', () => {
    const info = chartInfo('pie');
    expect(
      ops(info, { kind: 'pie', title: 'Share' }, { kind: 'pie', title: { auto: true } }),
    ).toEqual([{ op: 'setChartTitle', title: { kind: ChartTitleKind.Auto } }]);
    expect(
      ops(info, { kind: 'pie', title: { auto: true } }, { kind: 'pie', title: { auto: true } }),
    ).toEqual([]);
    expect(ops(info, { kind: 'pie', title: { auto: true } }, { kind: 'pie', title: '' })).toEqual([
      { op: 'setChartTitle', title: null },
    ]);
  });

  it('reads and writes an automatic axis title as the automatic kind, once', () => {
    const auto = { kind: ChartTitleKind.Auto, text: '', formula: '' };
    expect(
      read(chartInfo('column', { axes: axesFor('column', { title: auto }) })).xAxis?.title,
    ).toEqual({
      auto: true,
    });
    const info = chartInfo('line');
    const patchOps = (title: string | { auto: true }, next: string | { auto: true }) =>
      ops(info, { kind: 'line', yAxis: { title } }, { kind: 'line', yAxis: { title: next } });
    expect(patchOps('Units', { auto: true })).toEqual([
      { op: 'setChartAxis', axis: 'y', patch: { title: { kind: ChartTitleKind.Auto } } },
    ]);
    expect(patchOps({ auto: true }, { auto: true })).toEqual([]);
    expect(patchOps({ auto: true }, '')).toEqual([
      { op: 'setChartAxis', axis: 'y', patch: { title: null } },
    ]);
  });

  it('writes the title as text, an empty title as none and keeps it when omitted', () => {
    const info = chartInfo('pie');
    expect(ops(info, { kind: 'pie' }, { kind: 'pie', title: 'Share' })).toEqual([
      { op: 'setChartTitle', title: 'Share' },
    ]);
    expect(ops(info, { kind: 'pie', title: 'Share' }, { kind: 'pie', title: '' })).toEqual([
      { op: 'setChartTitle', title: null },
    ]);
    expect(ops(info, { kind: 'pie', title: 'Share' }, { kind: 'pie' })).toEqual([]);
  });

  it('writes the legend, keeping the overlay and a top-right legend left alone', () => {
    const info = chartInfo('line', {
      legend: { position: ChartLegendPosition.TopRight, overlay: true, hiddenSeries: [] },
    });
    const before = read(info);
    expect(ops(info, before, { ...before, title: 'x' }).map((op) => op.op)).toEqual([
      'setChartTitle',
    ]);
    expect(ops(info, before, { ...before, legend: { visible: false } })).toEqual([
      { op: 'setChartLegend', position: ChartLegendPosition.None, overlay: true },
    ]);
    expect(ops(info, before, { ...before, legend: { visible: true, position: 'left' } })).toEqual([
      { op: 'setChartLegend', position: ChartLegendPosition.Left, overlay: true },
    ]);
    expect(
      ops(
        info,
        { kind: 'line', legend: { visible: false } },
        { kind: 'line', legend: { visible: true } },
      ),
    ).toEqual([{ op: 'setChartLegend', position: ChartLegendPosition.Bottom, overlay: true }]);
    // A hidden legend's position is not a change.
    expect(
      ops(
        info,
        { kind: 'line', legend: { visible: false, position: 'top' } },
        { kind: 'line', legend: { visible: false } },
      ),
    ).toEqual([]);
  });

  it('writes the filter by series idx and category position', () => {
    const series = [seriesInfo(0), seriesInfo(3)];
    const info = chartInfo('column', {}, { series });
    expect(
      ops(
        info,
        { kind: 'column' },
        {
          kind: 'column',
          filter: {
            hiddenSeriesIds: ['series-3', 'bogus'],
            hiddenCategoryIds: ['category-4', 'category-1', 'point-2'],
          },
        },
      ),
    ).toEqual([{ op: 'setChartFilter', hiddenSeries: [3], hiddenCategories: [1, 4] }]);
    expect(
      ops(
        chartInfo('scatter'),
        { kind: 'scatter' },
        { kind: 'scatter', filter: { hiddenCategoryIds: ['point-2', 'category-0'] } },
      ),
    ).toEqual([{ op: 'setChartFilter', hiddenSeries: [], hiddenCategories: [2] }]);
    expect(
      ops(
        info,
        { kind: 'column', filter: { hiddenSeriesIds: ['series-0'] } },
        { kind: 'column', filter: {} },
      ),
    ).toEqual([{ op: 'setChartFilter', hiddenSeries: [], hiddenCategories: [] }]);
    // A rebuild clears the engine's filter, so an unchanged filter is written again.
    const filtered: ChartSpec = { kind: 'column', filter: { hiddenSeriesIds: ['series-0'] } };
    expect(ops(info, filtered, { ...filtered, kind: 'bar' }).at(-1)).toEqual({
      op: 'setChartFilter',
      hiddenSeries: [0],
      hiddenCategories: [],
    });
  });

  it('writes chart-wide labels to every series and per-series labels to their series', () => {
    const info = chartInfo('column', {}, { series: [seriesInfo(0), seriesInfo(2)] });
    const shown = {
      showValue: true,
      showCategory: false,
      showSeries: false,
      showPercent: false,
      showLegendKey: false,
      separator: null,
      position: null,
      numberFormat: null,
      fontColor: null,
      fontSizePt: null,
    };
    expect(
      ops(info, { kind: 'column' }, { kind: 'column', dataLabels: { visible: true } }),
    ).toEqual([
      { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { labels: shown } },
      { op: 'setChartSeries', plot: 0, seriesIdx: 2, patch: { labels: shown } },
    ]);
    expect(
      ops(
        info,
        { kind: 'column', dataLabels: { visible: true } },
        { kind: 'column', dataLabels: { visible: false, showValue: true } },
      ),
    ).toEqual([
      { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { labels: null } },
      { op: 'setChartSeries', plot: 0, seriesIdx: 2, patch: { labels: null } },
    ]);
    const detailed = {
      showCategoryName: true,
      separator: ' / ',
      position: 'inside-end',
      numberFormat: '0%',
      color: '#ABCDEF',
      fontSize: 11,
    } as const;
    expect(
      ops(
        info,
        { kind: 'column' },
        {
          kind: 'column',
          seriesStyles: [{ seriesId: 'series-2', dataLabels: detailed }],
        },
      ),
    ).toEqual([
      {
        op: 'setChartSeries',
        plot: 0,
        seriesIdx: 2,
        patch: {
          labels: {
            showValue: false,
            showCategory: true,
            showSeries: false,
            showPercent: false,
            showLegendKey: false,
            separator: ' / ',
            position: EngineChartLabelPosition.InsideEnd,
            numberFormat: '0%',
            fontColor: 0xffabcdef,
            fontSizePt: 11,
          },
        },
      },
    ]);
    // A per-series override wins over a chart-wide change.
    expect(
      ops(
        info,
        {
          kind: 'column',
          seriesStyles: [{ seriesId: 'series-0', dataLabels: { visible: false } }],
        },
        {
          kind: 'column',
          dataLabels: { visible: true },
          seriesStyles: [{ seriesId: 'series-0', dataLabels: { visible: false } }],
        },
      ),
    ).toEqual([{ op: 'setChartSeries', plot: 0, seriesIdx: 2, patch: { labels: shown } }]);
  });

  it('writes the palette as each series fill colour, explicit colours first', () => {
    const info = chartInfo('pie', {}, { series: [seriesInfo(0), seriesInfo(1), seriesInfo(5)] });
    expect(ops(info, { kind: 'pie' }, { kind: 'pie', palette: ['#111111', '#222222'] })).toEqual([
      { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { fill: { color: 0xff111111 } } },
      { op: 'setChartSeries', plot: 0, seriesIdx: 1, patch: { fill: { color: 0xff222222 } } },
      { op: 'setChartSeries', plot: 0, seriesIdx: 5, patch: { fill: { color: 0xff111111 } } },
    ]);
    expect(
      ops(
        info,
        { kind: 'pie', palette: ['#111111'] },
        {
          kind: 'pie',
          palette: ['#111111'],
          seriesStyles: [{ seriesId: 'series-1', fill: { color: '#333333' } }],
        },
      ),
    ).toEqual([
      { op: 'setChartSeries', plot: 0, seriesIdx: 1, patch: { fill: { color: 0xff333333 } } },
    ]);
    // Dropping the palette clears the colours it supplied; an unstorable colour is skipped.
    expect(
      ops(info, { kind: 'pie', palette: ['#111111'] }, { kind: 'pie', palette: ['tomato'] }),
    ).toEqual([]);
    expect(
      ops(info, { kind: 'pie', palette: ['#111111'] }, { kind: 'pie' }).map(
        (op) => op.op === 'setChartSeries' && op.patch,
      ),
    ).toEqual([{ fill: { color: null } }, { fill: { color: null } }, { fill: { color: null } }]);
  });

  it('treats a scatter chart without a style as marker-only, adding no line for its colour', () => {
    // A chart retyped to scatter keeps its series colours but names no scatter style.
    const info = chartInfo(
      'scatter',
      {},
      { series: [seriesInfo(0, { line: lineInfo({ none: true }) })] },
    );
    const before = read(info);
    expect(before.scatterStyle).toBe('marker');
    const { scatterStyle: _style, ...unstyled } = before;
    const after: ChartSpec = {
      ...unstyled,
      seriesStyles: [
        {
          seriesId: 'series-0',
          fill: { color: '#123456' },
          marker: { visible: true, shape: 'circle' },
        },
      ],
    };
    const patches = ops(info, before, after).flatMap((op) =>
      op.op === 'setChartSeries' ? [op.patch] : [],
    );
    expect(patches).toEqual([
      {
        marker: {
          symbol: ChartMarkerSymbol.Circle,
          fill: { color: 0xff123456 },
          line: { color: 0xff123456 },
        },
      },
    ]);
    expect(normalizeChartSpecLine(after)).toBeUndefined();
  });

  it('writes a line or scatter point colour to its line and marker, never a fill', () => {
    const shown = { visible: true, shape: 'circle' } as const;
    const spec = (kind: 'line' | 'scatter', points?: ChartPointStyleSpec[]): ChartSpec => ({
      kind,
      ...(kind === 'scatter' ? { scatterStyle: 'marker' as const } : {}),
      seriesStyles: [{ seriesId: 'series-0', marker: shown, ...(points ? { points } : {}) }],
    });
    const colored = (pointId: string): ChartPointStyleSpec[] => [
      { pointId, fill: { color: '#00ff00', transparency: 50 } },
    ];
    const marker = { fill: { color: 0xff00ff00 }, line: { color: 0xff00ff00 } };
    expect(ops(chartInfo('line'), spec('line'), spec('line', colored('category-2')))).toEqual([
      {
        op: 'setChartPoint',
        plot: 0,
        seriesIdx: 0,
        pointIdx: 2,
        patch: { line: { color: 0xff00ff00 }, marker },
      },
    ]);
    // A marker-only scatter point takes the colour on its marker alone.
    expect(ops(chartInfo('scatter'), spec('scatter'), spec('scatter', colored('point-2')))).toEqual(
      [{ op: 'setChartPoint', plot: 0, seriesIdx: 0, pointIdx: 2, patch: { marker } }],
    );
  });

  it('writes a line or scatter series colour to its line and shown markers, never a fill', () => {
    const info = chartInfo('line', {}, { series: [seriesInfo(0), seriesInfo(1)] });
    const shown = { visible: true, shape: 'circle' } as const;
    expect(
      ops(
        info,
        { kind: 'line', seriesStyles: [{ seriesId: 'series-0', marker: shown }] },
        {
          kind: 'line',
          palette: ['#111111', '#222222'],
          seriesStyles: [
            { seriesId: 'series-0', fill: { transparency: 40 }, marker: shown },
            {
              seriesId: 'series-1',
              fill: { color: '#333333' },
              line: { color: '#444444' },
              marker: { visible: false },
            },
          ],
        },
      ),
    ).toEqual([
      {
        op: 'setChartSeries',
        plot: 0,
        seriesIdx: 0,
        patch: {
          line: { color: 0xff111111 },
          marker: { fill: { color: 0xff111111 }, line: { color: 0xff111111 } },
        },
      },
      {
        op: 'setChartSeries',
        plot: 0,
        seriesIdx: 1,
        patch: { line: { color: 0xff444444 }, marker: { symbol: ChartMarkerSymbol.None } },
      },
    ]);
    // A marker colour of its own wins; a marker-only scatter series takes the colour on its markers alone.
    expect(
      ops(
        chartInfo('scatter', {}, { series: [seriesInfo(0)] }),
        {
          kind: 'scatter',
          scatterStyle: 'marker',
          seriesStyles: [{ seriesId: 'series-0', marker: shown }],
        },
        {
          kind: 'scatter',
          scatterStyle: 'marker',
          seriesStyles: [
            {
              seriesId: 'series-0',
              fill: { color: '#555555', transparency: 20 },
              marker: { ...shown, fill: { color: '#666666', transparency: 50 } },
            },
          ],
        },
      ),
    ).toEqual([
      {
        op: 'setChartSeries',
        plot: 0,
        seriesIdx: 0,
        patch: {
          marker: { fill: { color: 0xff666666, transparency: 50 }, line: { color: 0xff555555 } },
        },
      },
    ]);
  });

  // Pairwise rows (coverwise, strength 2, seed 1) over colour x transparency x width x dash x
  // target, from a line of #112233, 10 % transparency, 1.5 pt and dot dashes.
  it.each([
    ['set', 'keep', 'keep', 'solid', 'point'],
    ['clear', 'set', 'clear', 'clear', 'series'],
    ['set', 'clear', 'clear', 'dash', 'point'],
    ['keep', 'clear', 'set', 'clear', 'series'],
    ['clear', 'clear', 'keep', 'dot', 'point'],
    ['keep', 'set', 'set', 'keep', 'point'],
    ['set', 'set', 'keep', 'dash-dot', 'series'],
    ['keep', 'keep', 'keep', 'keep', 'series'],
    ['clear', 'keep', 'set', 'dash-dot', 'point'],
    ['keep', 'keep', 'clear', 'dot', 'series'],
    ['keep', 'clear', 'clear', 'dash-dot', 'point'],
    ['keep', 'set', 'set', 'dash', 'series'],
    ['clear', 'set', 'set', 'solid', 'point'],
    ['clear', 'keep', 'clear', 'keep', 'series'],
    ['set', 'set', 'set', 'dot', 'point'],
    ['set', 'keep', 'keep', 'clear', 'point'],
    ['clear', 'set', 'clear', 'solid', 'series'],
    ['clear', 'clear', 'keep', 'dash', 'point'],
    ['set', 'clear', 'clear', 'keep', 'series'],
    ['keep', 'set', 'clear', 'solid', 'series'],
    ['keep', 'keep', 'keep', 'dash', 'series'],
    ['set', 'clear', 'clear', 'solid', 'point'],
  ] as const)(
    'writes a line (colour %s, transparency %s, width %s, dash %s) to a %s',
    (color, transparency, width, dash, target) => {
      const before: ChartLineSpec = { color: '#112233', transparency: 10, width: 1.5, dash: 'dot' };
      const pick = <T>(change: string, set: T, keep: T): T | undefined =>
        change === 'keep' ? keep : change === 'set' ? set : undefined;
      const after: ChartLineSpec = {
        color: pick(color, '#AABBCC', '#112233'),
        transparency: pick(transparency, 60, 10),
        width: pick(width, 3, 1.5),
        dash: dash === 'keep' ? 'dot' : dash === 'clear' ? undefined : dash,
      };
      const patch = {
        ...({ keep: {}, set: { color: 0xffaabbcc }, clear: { color: null } } as const)[color],
        ...({ keep: {}, set: { transparency: 60 }, clear: { transparency: null } } as const)[
          transparency
        ],
        ...({ keep: {}, set: { widthPt: 3 }, clear: { widthPt: null } } as const)[width],
        ...(
          {
            keep: {},
            solid: { dash: ChartLineDash.Solid },
            dash: { dash: ChartLineDash.Dash },
            dot: {},
            'dash-dot': { dash: ChartLineDash.DashDot },
            clear: { dash: null },
          } as const
        )[dash],
      };
      const spec = (line: ChartLineSpec): ChartSpec => ({
        kind: 'line',
        seriesStyles: [
          target === 'series'
            ? { seriesId: 'series-1', line }
            : { seriesId: 'series-1', points: [{ pointId: 'category-2', line }] },
        ],
      });
      const expected: ChartEngineOp[] =
        Object.keys(patch).length === 0
          ? []
          : target === 'series'
            ? [{ op: 'setChartSeries', plot: 0, seriesIdx: 1, patch: { line: patch } }]
            : [{ op: 'setChartPoint', plot: 0, seriesIdx: 1, pointIdx: 2, patch: { line: patch } }];
      expect(ops(chartInfo('line'), spec(before), spec(after))).toEqual(expected);
    },
  );

  it('writes series fill transparency and removes point formatting', () => {
    const info = chartInfo('bar');
    expect(
      ops(
        info,
        {
          kind: 'bar',
          seriesStyles: [{ seriesId: 'series-0', fill: { color: '#000000', transparency: 5 } }],
        },
        {
          kind: 'bar',
          seriesStyles: [{ seriesId: 'series-0', fill: { color: '#000000', transparency: 30 } }],
        },
      ),
    ).toEqual([
      { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { fill: { transparency: 30 } } },
    ]);
    expect(
      ops(
        info,
        {
          kind: 'bar',
          seriesStyles: [
            {
              seriesId: 'series-0',
              points: [{ pointId: 'category-3', fill: { color: '#ff0000' } }],
            },
          ],
        },
        { kind: 'bar', seriesStyles: [{ seriesId: 'series-0', points: [] }] },
      ),
    ).toEqual([{ op: 'setChartPoint', plot: 0, seriesIdx: 0, pointIdx: 3, patch: null }]);
    // Styles of series the chart does not have are not written.
    expect(
      ops(
        info,
        { kind: 'bar' },
        { kind: 'bar', seriesStyles: [{ seriesId: 'series-9', fill: { transparency: 1 } }] },
      ),
    ).toEqual([]);
  });

  // Pairwise rows (coverwise, strength 2, seed 1) over kind x marker before x after x size.
  it.each([
    ['scatter', 'auto', 'absent', 'keep'],
    ['line', 'auto', 'square', 'set'],
    ['scatter', 'hidden', 'circle', 'set'],
    ['scatter', 'circle', 'hidden', 'set'],
    ['column', 'absent', 'auto', 'set'],
    ['line', 'circle', 'auto', 'keep'],
    ['pie', 'hidden', 'absent', 'set'],
    ['column', 'auto', 'circle', 'keep'],
    ['pie', 'absent', 'square', 'keep'],
    ['column', 'hidden', 'absent', 'keep'],
    ['scatter', 'absent', 'square', 'set'],
    ['pie', 'circle', 'circle', 'keep'],
    ['pie', 'auto', 'hidden', 'set'],
    ['column', 'circle', 'square', 'keep'],
    ['line', 'hidden', 'hidden', 'set'],
    ['line', 'absent', 'circle', 'set'],
    ['column', 'absent', 'hidden', 'keep'],
    ['scatter', 'hidden', 'auto', 'keep'],
    ['line', 'absent', 'absent', 'set'],
    ['scatter', 'auto', 'auto', 'keep'],
    ['scatter', 'circle', 'absent', 'set'],
    ['pie', 'absent', 'auto', 'keep'],
    ['scatter', 'hidden', 'square', 'keep'],
  ] as const)('writes a %s marker from %s to %s (size %s)', (kind, from, to, size) => {
    const markers = {
      absent: undefined,
      hidden: { visible: false },
      auto: { visible: true },
      circle: { visible: true, shape: 'circle' },
      square: { visible: true, shape: 'square' },
    } as const;
    const symbols: Record<string, Record<string, ChartMarkerSymbol | undefined>> = {
      absent: {
        hidden: ChartMarkerSymbol.None,
        auto: undefined,
        circle: ChartMarkerSymbol.Circle,
        square: ChartMarkerSymbol.Square,
      },
      hidden: {
        hidden: undefined,
        auto: ChartMarkerSymbol.Auto,
        circle: ChartMarkerSymbol.Circle,
        square: ChartMarkerSymbol.Square,
      },
      auto: {
        hidden: ChartMarkerSymbol.None,
        auto: undefined,
        circle: ChartMarkerSymbol.Circle,
        square: ChartMarkerSymbol.Square,
      },
      circle: {
        hidden: ChartMarkerSymbol.None,
        auto: undefined,
        circle: undefined,
        square: ChartMarkerSymbol.Square,
      },
    };
    const target = markers[to];
    const after: ChartMarkerSpec | undefined =
      target && size === 'set' ? { ...target, size: 7 } : target;
    const spec = (marker: ChartMarkerSpec | undefined): ChartSpec => ({
      kind,
      seriesStyles: [{ seriesId: 'series-0', ...(marker ? { marker } : {}) }],
    });
    let marker: object | null | undefined;
    if (to === 'absent') {
      marker = from === 'absent' ? undefined : null;
    } else {
      const symbol = symbols[from]?.[to];
      const patch = {
        ...(symbol === undefined ? {} : { symbol }),
        ...(size === 'set' ? { size: 7 } : {}),
      };
      marker = Object.keys(patch).length > 0 ? patch : undefined;
    }
    const markerKind = kind === 'line' || kind === 'scatter';
    const expected: ChartEngineOp[] =
      markerKind && marker !== undefined
        ? [{ op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: { marker } as never }]
        : [];
    expect(ops(chartInfo(kind), spec(markers[from]), spec(after))).toEqual(expected);
  });

  it('writes marker fill and line, and point markers only on marker kinds', () => {
    const marker = (
      fill: string,
      width: number,
      pointId = 'point-1',
    ): ChartSpec['seriesStyles'] => [
      {
        seriesId: 'series-0',
        points: [{ pointId, marker: { visible: true, fill: { color: fill }, line: { width } } }],
      },
    ];
    expect(
      ops(
        chartInfo('scatter'),
        { kind: 'scatter', seriesStyles: marker('#000000', 1) },
        {
          kind: 'scatter',
          seriesStyles: marker('#ffffff', 2),
        },
      ),
    ).toEqual([
      {
        op: 'setChartPoint',
        plot: 0,
        seriesIdx: 0,
        pointIdx: 1,
        patch: { marker: { fill: { color: 0xffffffff }, line: { widthPt: 2 } } },
      },
    ]);
    expect(
      ops(
        chartInfo('area'),
        { kind: 'area', seriesStyles: marker('#000000', 1, 'category-1') },
        { kind: 'area', seriesStyles: marker('#ffffff', 2, 'category-1') },
      ),
    ).toEqual([]);
  });

  it('writes the scatter style as series lines and markers', () => {
    const info = chartInfo('scatter', {}, { series: [seriesInfo(0), seriesInfo(1)] });
    const each = (patch: object): ChartEngineOp[] => [
      { op: 'setChartSeries', plot: 0, seriesIdx: 0, patch: patch as never },
      { op: 'setChartSeries', plot: 0, seriesIdx: 1, patch: patch as never },
    ];
    expect(
      ops(
        info,
        { kind: 'scatter', scatterStyle: 'marker' },
        { kind: 'scatter', scatterStyle: 'lineMarker' },
      ),
    ).toEqual(each({ line: { none: false } }));
    expect(
      ops(
        info,
        { kind: 'scatter', scatterStyle: 'lineMarker' },
        { kind: 'scatter', scatterStyle: 'line' },
      ),
    ).toEqual(each({ marker: { symbol: ChartMarkerSymbol.None } }));
    expect(
      ops(
        info,
        { kind: 'scatter', scatterStyle: 'line' },
        { kind: 'scatter', scatterStyle: 'marker' },
      ),
    ).toEqual(each({ line: { none: true }, marker: { symbol: ChartMarkerSymbol.Auto } }));
    expect(ops(info, { kind: 'scatter' }, { kind: 'scatter', scatterStyle: 'line' })).toEqual(
      each({ line: { none: false }, marker: { symbol: ChartMarkerSymbol.None } }),
    );
    // An explicit marker shape wins over the style's automatic marker.
    expect(
      ops(
        info,
        { kind: 'scatter', scatterStyle: 'line' },
        {
          kind: 'scatter',
          scatterStyle: 'lineMarker',
          seriesStyles: [{ seriesId: 'series-1', marker: { visible: true, shape: 'diamond' } }],
        },
      ),
    ).toEqual([
      {
        op: 'setChartSeries',
        plot: 0,
        seriesIdx: 0,
        patch: { marker: { symbol: ChartMarkerSymbol.Auto } },
      },
      {
        op: 'setChartSeries',
        plot: 0,
        seriesIdx: 1,
        patch: { marker: { symbol: ChartMarkerSymbol.Diamond } },
      },
    ]);
    expect(
      ops(chartInfo('line'), { kind: 'line' }, { kind: 'line', scatterStyle: 'line' }),
    ).toEqual([]);
  });

  it('writes axis visibility, title, unit, format and gridlines', () => {
    const info = chartInfo('line');
    expect(
      ops(
        info,
        { kind: 'line', yAxis: { visible: true, title: 'Old', majorGridlines: true } },
        {
          kind: 'line',
          yAxis: {
            visible: false,
            title: '',
            majorUnit: 5,
            numberFormat: '#,##0',
            majorGridlines: false,
          },
        },
      ),
    ).toEqual([
      {
        op: 'setChartAxis',
        axis: 'y',
        patch: {
          visible: false,
          title: null,
          majorUnit: 5,
          numberFormat: '#,##0',
          majorGridlines: false,
        },
      },
    ]);
    expect(
      ops(
        info,
        {
          kind: 'line',
          xAxis: { visible: false, majorUnit: 5, numberFormat: '0', majorGridlines: false },
        },
        {
          kind: 'line',
          xAxis: { title: 'Month' },
        },
      ),
    ).toEqual([
      {
        op: 'setChartAxis',
        axis: 'x',
        patch: {
          visible: true,
          title: 'Month',
          majorUnit: null,
          numberFormat: null,
          majorGridlines: true,
        },
      },
    ]);
    for (const kind of ['pie', 'doughnut'] as const) {
      expect(
        ops(chartInfo(kind), { kind }, { kind, xAxis: { min: 1 }, yAxis: { max: 2 } }),
      ).toEqual([]);
    }
  });

  // Pairwise rows (coverwise, strength 2, seed 1) over kind x role x crossing x logBase x
  // labelPosition x reverseOrder x bounds, from an axis crossed at 5 with log base 10, high
  // labels, forward order and bounds 0..100.
  it.each([
    ['column', 'x', 'max', 'set', 'clear', 'keep', 'set'],
    ['bar', 'y', 'clear', 'clear', 'clear', 'clear', 'clear'],
    ['scatter', 'y', 'keep', 'keep', 'keep', 'true', 'set'],
    ['column', 'y', 'min', 'set', 'low', 'true', 'keep'],
    ['area', 'x', 'number', 'clear', 'low', 'true', 'clear'],
    ['line', 'x', 'clear', 'keep', 'low', 'clear', 'set'],
    ['area', 'y', 'number', 'set', 'keep', 'clear', 'keep'],
    ['bar', 'x', 'auto', 'clear', 'keep', 'keep', 'keep'],
    ['area', 'y', 'auto', 'keep', 'clear', 'true', 'set'],
    ['column', 'y', 'number', 'keep', 'keep', 'keep', 'clear'],
    ['line', 'y', 'max', 'clear', 'keep', 'true', 'keep'],
    ['scatter', 'x', 'keep', 'keep', 'clear', 'keep', 'keep'],
    ['scatter', 'y', 'max', 'set', 'low', 'clear', 'clear'],
    ['line', 'x', 'auto', 'set', 'low', 'clear', 'clear'],
    ['column', 'y', 'keep', 'clear', 'low', 'clear', 'clear'],
    ['bar', 'x', 'min', 'keep', 'low', 'true', 'keep'],
    ['column', 'y', 'auto', 'keep', 'low', 'true', 'clear'],
    ['scatter', 'y', 'min', 'keep', 'low', 'keep', 'clear'],
    ['bar', 'x', 'min', 'clear', 'keep', 'true', 'set'],
    ['line', 'x', 'min', 'set', 'clear', 'clear', 'clear'],
    ['scatter', 'x', 'number', 'clear', 'clear', 'clear', 'set'],
    ['bar', 'y', 'clear', 'set', 'low', 'true', 'keep'],
    ['scatter', 'y', 'auto', 'clear', 'clear', 'true', 'clear'],
    ['bar', 'y', 'number', 'keep', 'low', 'clear', 'clear'],
    ['line', 'y', 'keep', 'clear', 'clear', 'keep', 'keep'],
    ['line', 'y', 'number', 'clear', 'low', 'clear', 'clear'],
    ['area', 'x', 'keep', 'set', 'low', 'clear', 'clear'],
    ['area', 'x', 'min', 'clear', 'clear', 'keep', 'keep'],
    ['area', 'y', 'clear', 'keep', 'keep', 'clear', 'clear'],
    ['column', 'x', 'clear', 'clear', 'low', 'keep', 'keep'],
    ['bar', 'y', 'max', 'keep', 'low', 'keep', 'keep'],
    ['scatter', 'y', 'clear', 'keep', 'keep', 'true', 'keep'],
    ['area', 'y', 'max', 'clear', 'low', 'true', 'keep'],
    ['bar', 'y', 'keep', 'clear', 'low', 'keep', 'set'],
  ] as const)(
    'writes a %s %s axis (crossing %s, log %s, labels %s, reverse %s, bounds %s)',
    (kind, role, crossing, logBase, labelPosition, reverseOrder, bounds) => {
      const before: ChartAxisSpec = {
        crossing: 5,
        logBase: 10,
        labelPosition: 'high',
        reverseOrder: false,
        min: 0,
        max: 100,
      };
      const after: ChartAxisSpec = {
        crossing: (
          { keep: 5, auto: 'auto', min: 'min', max: 'max', number: 3, clear: undefined } as const
        )[crossing],
        logBase: ({ keep: 10, set: 2, clear: undefined } as const)[logBase],
        labelPosition: ({ keep: 'high', low: 'low', clear: undefined } as const)[labelPosition],
        reverseOrder: ({ keep: false, true: true, clear: undefined } as const)[reverseOrder],
        min: ({ keep: 0, set: -5, clear: undefined } as const)[bounds],
        max: ({ keep: 100, set: 50, clear: undefined } as const)[bounds],
      };
      const own = {
        ...({ keep: {}, set: { min: -5, max: 50 }, clear: { min: null, max: null } } as const)[
          bounds
        ],
        ...({ keep: {}, true: { reversed: true }, clear: { reversed: false } } as const)[
          reverseOrder
        ],
        ...({ keep: {}, set: { logBase: 2 }, clear: { logBase: null } } as const)[logBase],
        ...(
          {
            keep: {},
            low: { tickLabelPosition: 'low' },
            clear: { tickLabelPosition: null },
          } as const
        )[labelPosition],
      };
      // The crossing is stored on the perpendicular axis.
      const crossed = (
        {
          keep: {},
          auto: { crosses: ChartAxisCrosses.AutoZero },
          min: { crosses: ChartAxisCrosses.Min },
          max: { crosses: ChartAxisCrosses.Max },
          number: { crosses: ChartAxisCrosses.At, crossesAt: 3 },
          clear: { crosses: ChartAxisCrosses.AutoZero },
        } as const
      )[crossing];
      const other = role === 'x' ? 'y' : 'x';
      const patches = new Map<string, object>([
        [role, own],
        [other, crossed],
      ]);
      const expected = (['x', 'y'] as const)
        .filter((axis) => Object.keys(patches.get(axis) ?? {}).length > 0)
        .map((axis) => ({ op: 'setChartAxis', axis, patch: patches.get(axis) }));
      const key = role === 'x' ? 'xAxis' : 'yAxis';
      expect(ops(chartInfo(kind), { kind, [key]: before }, { kind, [key]: after })).toEqual(
        expected,
      );
    },
  );

  it('orders operations as type, source, plot, display, title, legend, axes, series, points, filter', () => {
    const info = chartInfo('column', {}, { series: [seriesInfo(0), seriesInfo(1)] });
    const after: ChartSpec = {
      kind: 'bar',
      seriesDirection: 'rows',
      gapWidth: 10,
      emptyCells: 'zero',
      title: 'T',
      legend: { visible: true, position: 'top' },
      xAxis: { min: 1 },
      yAxis: { max: 2 },
      palette: ['#010203'],
      seriesStyles: [
        { seriesId: 'series-1', points: [{ pointId: 'category-0', fill: { color: '#ffffff' } }] },
      ],
      filter: { hiddenSeriesIds: ['series-1'] },
    };
    expect(ops(info, { kind: 'column' }, after).map((op) => op.op)).toEqual([
      'setChartType',
      'setChartSource',
      'setChartPlot',
      'setChartDisplay',
      'setChartTitle',
      'setChartLegend',
      'setChartAxis',
      'setChartAxis',
      'setChartSeries',
      'setChartSeries',
      'setChartPoint',
      'setChartFilter',
    ]);
  });

  it('leaves a marker-only scatter series line to the scatter style', () => {
    const info = chartInfo('scatter');
    const line = { color: '#333333', width: 2, dash: 'dash' } as const;
    expect(
      ops(
        info,
        { kind: 'scatter', scatterStyle: 'marker' },
        {
          kind: 'scatter',
          scatterStyle: 'marker',
          seriesStyles: [{ seriesId: 'series-0', line }],
        },
      ),
    ).toEqual([]);
    expect(
      ops(
        info,
        { kind: 'scatter', scatterStyle: 'lineMarker' },
        {
          kind: 'scatter',
          scatterStyle: 'lineMarker',
          seriesStyles: [{ seriesId: 'series-0', line }],
        },
      ),
    ).toEqual([
      {
        op: 'setChartSeries',
        plot: 0,
        seriesIdx: 0,
        patch: { line: { color: 0xff333333, widthPt: 2, dash: ChartLineDash.Dash } },
      },
    ]);
  });
});

// ---------------------------------------------------------------------------
// Drawing data
// ---------------------------------------------------------------------------

describe('series colours', () => {
  it('strips the colour slots a palette replaces, per kind', () => {
    const style: ChartSeriesStyleSpec = {
      seriesId: 'series-0',
      fill: { color: '#111111', transparency: 10 },
      line: { color: '#222222', width: 2 },
      marker: { visible: true, fill: { color: '#333333' }, line: { color: '#444444' } },
    };
    expect(withoutSeriesColor('column', style)).toEqual({
      seriesId: 'series-0',
      fill: { transparency: 10 },
      line: { color: '#222222', width: 2 },
      marker: style.marker,
    });
    expect(withoutSeriesColor('line', style)).toEqual({
      seriesId: 'series-0',
      fill: { transparency: 10 },
      line: { width: 2 },
      marker: { visible: true },
    });
    expect(
      withoutSeriesColor('scatter', {
        seriesId: 'series-1',
        marker: { fill: { color: '#333333' } },
      }),
    ).toEqual({
      seriesId: 'series-1',
      marker: {},
    });
  });

  it('reads the colour a series shows from where its kind stores it', () => {
    const style: ChartSeriesStyleSpec = {
      seriesId: 'series-0',
      fill: { color: '#111111' },
      line: { color: '#222222' },
      marker: { fill: { color: '#333333' } },
    };
    expect(seriesColorOf({ kind: 'column' }, style)).toBe('#111111');
    expect(seriesColorOf({ kind: 'line' }, style)).toBe('#222222');
    expect(seriesColorOf({ kind: 'scatter', scatterStyle: 'lineMarker' }, style)).toBe('#222222');
    expect(seriesColorOf({ kind: 'scatter', scatterStyle: 'marker' }, style)).toBe('#333333');
    expect(
      seriesColorOf({ kind: 'line' }, { seriesId: 'series-0', fill: { color: '#111111' } }),
    ).toBe('#111111');
  });
});

describe('chartInsertOptions', () => {
  // Line and area write Excel's standard grouping for clustered.
  it.each([
    ['column', 'clustered', ChartInsertKind.Column, EngineChartGrouping.Clustered],
    ['column', 'stacked', ChartInsertKind.Column, EngineChartGrouping.Stacked],
    ['column', 'percent-stacked', ChartInsertKind.Column, EngineChartGrouping.PercentStacked],
    ['bar', 'clustered', ChartInsertKind.Bar, EngineChartGrouping.Clustered],
    ['bar', 'stacked', ChartInsertKind.Bar, EngineChartGrouping.Stacked],
    ['bar', 'percent-stacked', ChartInsertKind.Bar, EngineChartGrouping.PercentStacked],
    ['line', 'clustered', ChartInsertKind.Line, EngineChartGrouping.Standard],
    ['line', 'stacked', ChartInsertKind.Line, EngineChartGrouping.Stacked],
    ['line', 'percent-stacked', ChartInsertKind.Line, EngineChartGrouping.PercentStacked],
    ['area', 'clustered', ChartInsertKind.Area, EngineChartGrouping.Standard],
    ['area', 'stacked', ChartInsertKind.Area, EngineChartGrouping.Stacked],
    ['area', 'percent-stacked', ChartInsertKind.Area, EngineChartGrouping.PercentStacked],
    ['pie', undefined, ChartInsertKind.Pie, undefined],
    ['doughnut', undefined, ChartInsertKind.Doughnut, undefined],
    ['scatter', undefined, ChartInsertKind.Scatter, undefined],
  ] as const)('creates %s %s', (kind, grouping, insertKind, engineGrouping) => {
    const options = chartInsertOptions({ kind, ...(grouping ? { grouping } : {}) });
    expect(options.kind).toBe(insertKind);
    expect(options.grouping).toBe(engineGrouping);
    expect('grouping' in options).toBe(engineGrouping !== undefined);
  });

  it('carries the split and the legend, leaving omitted ones to Excel', () => {
    expect(chartInsertOptions({ kind: 'column' })).toEqual({
      kind: ChartInsertKind.Column,
      grouping: EngineChartGrouping.Clustered,
      direction: EngineChartSeriesDirection.Auto,
    });
    expect(
      chartInsertOptions({
        kind: 'pie',
        seriesDirection: 'rows',
        headerRow: false,
        headerColumn: true,
        legend: { visible: false },
      }),
    ).toEqual({
      kind: ChartInsertKind.Pie,
      direction: EngineChartSeriesDirection.Rows,
      headerRow: false,
      headerColumn: true,
      legend: ChartLegendPosition.None,
    });
    expect(chartInsertOptions({ kind: 'line', legend: {} }).legend).toBe(
      ChartLegendPosition.Bottom,
    );
    expect(chartInsertOptions({ kind: 'line', legend: { position: 'top' } }).legend).toBe(
      ChartLegendPosition.Top,
    );
  });
});

describe('chartDataFromEngine', () => {
  const ranges = { name: [], categories: [], values: [], x: [], y: [], size: [] };

  it('converts category series with their point states and format codes', () => {
    const data: EngineChartData = {
      domain: 'category',
      pointCount: 3,
      categories: {
        labels: ['Jan', 'Feb', 'Mar'],
        levels: [],
        values: null,
        formatCode: 'General',
        generated: false,
      },
      series: [
        {
          plotIndex: 0,
          idx: 4,
          name: 'North',
          formatCode: '0.0',
          xFormatCode: '',
          filtered: false,
          values: new Float64Array([1.5, Number.NaN, 0]),
          states: new Uint8Array([0, 1, 2]),
          x: null,
          sizes: null,
          ranges,
        },
        {
          plotIndex: 0,
          idx: 5,
          name: null,
          formatCode: 'General',
          xFormatCode: '',
          filtered: true,
          values: new Float64Array([2, 3, 4]),
          states: new Uint8Array([0, 0, 0]),
          x: null,
          sizes: null,
          ranges,
        },
      ],
      titleText: null,
      axisTitleTexts: [],
      unresolved: [],
    };
    expect(chartDataFromEngine(data)).toEqual({
      domain: 'category',
      categories: ['Jan', 'Feb', 'Mar'],
      categoryIds: ['category-0', 'category-1', 'category-2'],
      categoryFormatCode: 'General',
      series: [
        {
          id: 'series-4',
          name: 'North',
          values: [1.5, null, null],
          states: ['value', 'gap', 'connect'],
          formatCode: '0.0',
        },
        {
          id: 'series-5',
          values: [2, 3, 4],
          states: ['value', 'value', 'value'],
          formatCode: 'General',
        },
      ],
    });
  });

  it('converts xy series, numbering points without x values', () => {
    const series = (idx: number, x: Float64Array | null) => ({
      plotIndex: 0,
      idx,
      name: `S${idx}`,
      formatCode: 'General',
      xFormatCode: 'General',
      filtered: false,
      values: new Float64Array([10, 20]),
      states: new Uint8Array([0, 1]),
      x,
      sizes: null,
      ranges,
    });
    const data: EngineChartData = {
      domain: 'xy',
      pointCount: 2,
      categories: { labels: [], levels: [], values: null, formatCode: 'General', generated: true },
      series: [series(0, new Float64Array([0.5, Number.NaN])), series(1, null)],
      titleText: null,
      axisTitleTexts: [],
      unresolved: [],
    };
    expect(chartDataFromEngine(data)).toEqual({
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'S0',
          formatCode: 'General',
          xFormatCode: 'General',
          points: [
            { id: 'point-0', x: 0.5, y: 10, state: 'value' },
            { id: 'point-1', x: null, y: null, state: 'gap' },
          ],
        },
        {
          id: 'series-1',
          name: 'S1',
          formatCode: 'General',
          xFormatCode: 'General',
          points: [
            { id: 'point-0', x: 1, y: 10, state: 'value' },
            { id: 'point-1', x: 2, y: null, state: 'gap' },
          ],
        },
      ],
    });
  });
});

// ---------------------------------------------------------------------------
// Round-trip equivalence
// ---------------------------------------------------------------------------

describe('normalizeChartSpec', () => {
  const context: ChartNormalizeContext = {
    direction: 'columns',
    headerRow: true,
    headerColumn: false,
    seriesIds: ['series-0', 'series-1', 'series-2'],
    accents: ['#a10000', '#a20000', '#a30000', '#a40000', '#a50000', '#a60000'],
    styled: false,
  };

  it('settles a written marker: none keeps only its symbol, auto takes the series symbol, size 5', () => {
    const styles = normalizeChartSpec(
      {
        kind: 'line',
        seriesStyles: [
          {
            seriesId: 'series-0',
            marker: { visible: false, size: 7, fill: { color: '#123456' }, line: { width: 2 } },
            // An automatic point marker on a series without markers has none either.
            points: [{ pointId: 'category-1', marker: { visible: true } }],
          },
          { seriesId: 'series-1', marker: { visible: true } },
          {
            seriesId: 'series-3',
            marker: {
              visible: true,
              size: 8,
              fill: { color: '#123456' },
              line: { transparency: 20 },
            },
            points: [{ pointId: 'category-2', marker: { visible: true, shape: 'circle' } }],
          },
        ],
      },
      { ...context, seriesIds: ['series-0', 'series-1', 'series-3'] },
    ).seriesStyles;
    expect(styles).toEqual([
      {
        seriesId: 'series-0',
        marker: { visible: false },
        points: [{ pointId: 'category-1', marker: { visible: false } }],
      },
      // idx 1 is a square, idx 3 an x.
      { seriesId: 'series-1', marker: { visible: true, shape: 'square', size: 5 } },
      {
        seriesId: 'series-3',
        // An unstyled chart completes no outline width or colour.
        marker: {
          visible: true,
          shape: 'x',
          size: 8,
          fill: { color: '#123456' },
          line: { transparency: 20 },
        },
        points: [
          {
            pointId: 'category-2',
            marker: {
              visible: true,
              shape: 'circle',
              size: 8,
              fill: { color: '#123456' },
              line: { transparency: 20 },
            },
          },
        ],
      },
    ]);
  });

  it('completes a shown marker in a styled chart with a fill and outline in the series colour', () => {
    // Excel's Office theme accents; idx 10 is accent5 at lumMod 60000, idx 12 accent1 at 80000 + 20000.
    const accents = ['#156082', '#e97132', '#196b24', '#0f9ed5', '#a02b93', '#4ea72e'];
    const spec: ChartSpec = {
      kind: 'scatter',
      scatterStyle: 'marker',
      seriesStyles: [
        {
          seriesId: 'series-0',
          marker: { visible: true, shape: 'circle' },
          points: [{ pointId: 'point-1', marker: { visible: true, size: 9 } }],
        },
        {
          seriesId: 'series-10',
          marker: { visible: true, shape: 'square', fill: { color: '#123456' } },
        },
        { seriesId: 'series-12', marker: { visible: true, shape: 'diamond', line: { width: 2 } } },
      ],
    };
    const seriesIds = ['series-0', 'series-10', 'series-12'];
    expect(
      normalizeChartSpec(spec, { ...context, seriesIds, accents, styled: true }).seriesStyles,
    ).toEqual([
      {
        seriesId: 'series-0',
        marker: {
          visible: true,
          shape: 'circle',
          size: 5,
          fill: { color: '#156082' },
          line: { color: '#156082', width: 0.75 },
        },
        points: [
          {
            pointId: 'point-1',
            marker: {
              visible: true,
              shape: 'circle',
              size: 9,
              fill: { color: '#156082' },
              line: { color: '#156082', width: 0.75 },
            },
          },
        ],
      },
      {
        seriesId: 'series-10',
        marker: {
          visible: true,
          shape: 'square',
          size: 5,
          fill: { color: '#123456' },
          line: { color: '#601a58', width: 0.75 },
        },
      },
      {
        seriesId: 'series-12',
        marker: {
          visible: true,
          shape: 'diamond',
          size: 5,
          fill: { color: '#1f8ebe' },
          line: { color: '#1f8ebe', width: 2 },
        },
      },
    ]);
    // The same chart unstyled only takes its sizes.
    expect(
      normalizeChartSpec(spec, { ...context, seriesIds, accents, styled: false }).seriesStyles,
    ).toEqual([
      {
        seriesId: 'series-0',
        marker: { visible: true, shape: 'circle', size: 5 },
        points: [{ pointId: 'point-1', marker: { visible: true, shape: 'circle', size: 9 } }],
      },
      {
        seriesId: 'series-10',
        marker: { visible: true, shape: 'square', size: 5, fill: { color: '#123456' } },
      },
      {
        seriesId: 'series-12',
        marker: { visible: true, shape: 'diamond', size: 5, line: { width: 2 } },
      },
    ]);
  });

  it('keeps a point as its own format, without a copy of the series, across a type change', () => {
    const line: ChartSpec = {
      kind: 'line',
      seriesStyles: [
        {
          seriesId: 'series-0',
          line: { width: 2 },
          marker: { visible: true, shape: 'square', size: 7 },
          points: [{ pointId: 'category-1', fill: { color: '#00ff00' } }],
        },
      ],
    };
    const lineContext = { ...context, seriesIds: ['series-0'] };
    const point = (previousKind?: ChartKind) =>
      normalizeChartSpec(line, { ...lineContext, ...(previousKind ? { previousKind } : {}) })
        .seriesStyles?.[0]?.points;
    // Made on a column chart: the converted point takes the patch alone, its colour on line and marker.
    expect(point('column')).toEqual([
      {
        pointId: 'category-1',
        line: { color: '#00ff00' },
        marker: {
          visible: true,
          shape: 'square',
          size: 5,
          fill: { color: '#00ff00' },
          line: { color: '#00ff00' },
        },
      },
    ]);
    // Made on the line chart itself: a new point copies the series line and marker first.
    const copied = [
      {
        pointId: 'category-1',
        line: { color: '#00ff00', width: 2 },
        marker: {
          visible: true,
          shape: 'square',
          size: 7,
          fill: { color: '#00ff00' },
          line: { color: '#00ff00' },
        },
      },
    ];
    expect(point()).toEqual(copied);
    expect(point('line')).toEqual(copied);
    // Bar to area keeps the point unchanged rather than copying the series fill.
    const area: ChartSpec = {
      kind: 'area',
      seriesStyles: [
        {
          seriesId: 'series-0',
          fill: { color: '#111111' },
          line: { width: 3 },
          points: [{ pointId: 'category-1', fill: { transparency: 50 } }],
        },
      ],
    };
    expect(
      normalizeChartSpec(area, { ...lineContext, previousKind: 'bar' }).seriesStyles?.[0]?.points,
    ).toEqual([{ pointId: 'category-1', fill: { transparency: 50 } }]);
    expect(normalizeChartSpec(area, lineContext).seriesStyles?.[0]?.points).toEqual([
      { pointId: 'category-1', fill: { color: '#111111', transparency: 50 }, line: { width: 3 } },
    ]);
  });

  it('keeps a smoothed line on line and scatter series, an unsmoothed one as absent', () => {
    const smooth = (kind: ChartKind) =>
      normalizeChartSpec(
        {
          kind,
          seriesStyles: [
            { seriesId: 'series-0', smooth: true },
            { seriesId: 'series-1', smooth: false },
          ],
        },
        context,
      ).seriesStyles?.filter((style) => style.smooth !== undefined) ?? [];
    expect(smooth('line')).toEqual([{ seriesId: 'series-0', smooth: true }]);
    expect(smooth('scatter')).toEqual([{ seriesId: 'series-0', smooth: true }]);
    expect(smooth('column')).toEqual([]);
  });

  it('resolves the split and fills the defaults the engine writes', () => {
    expect(normalizeChartSpec({ kind: 'line' }, context)).toEqual({
      kind: 'line',
      grouping: 'clustered',
      seriesDirection: 'columns',
      headerRow: true,
      headerColumn: false,
      emptyCells: 'gap',
      title: { auto: true },
    });
    expect(
      normalizeChartSpec(
        {
          kind: 'column',
          seriesDirection: 'auto',
          headerColumn: true,
          grouping: 'stacked',
          emptyCells: 'span',
        },
        { ...context, direction: 'rows' },
      ),
    ).toEqual({
      kind: 'column',
      grouping: 'stacked',
      seriesDirection: 'rows',
      headerRow: true,
      headerColumn: true,
      emptyCells: 'gap',
      title: { auto: true },
    });
    expect(
      normalizeChartSpec(
        { kind: 'scatter', seriesDirection: 'rows', headerRow: false, emptyCells: 'span' },
        context,
      ),
    ).toEqual({
      kind: 'scatter',
      seriesDirection: 'rows',
      headerRow: false,
      headerColumn: false,
      emptyCells: 'span',
      title: { auto: true },
    });
  });

  it('drops fields the kind does not have', () => {
    const everything: Omit<ChartSpec, 'kind'> = {
      grouping: 'stacked',
      scatterStyle: 'line',
      gapWidth: 50,
      overlap: 10,
      firstSliceAngle: 30,
      holeSize: 40,
      varyColors: true,
      xAxis: { min: 1 },
      yAxis: { max: 9 },
      seriesStyles: [
        {
          seriesId: 'series-0',
          marker: { visible: false },
          points: [{ pointId: 'point-0', marker: { size: 4 } }],
        },
      ],
    };
    const base = {
      seriesDirection: 'columns',
      headerRow: true,
      headerColumn: false,
      emptyCells: 'gap',
      title: { auto: true },
    };
    expect(normalizeChartSpec({ kind: 'pie', ...everything }, context)).toEqual({
      kind: 'pie',
      ...base,
      firstSliceAngle: 30,
      varyColors: true,
    });
    expect(normalizeChartSpec({ kind: 'doughnut', ...everything }, context)).toEqual({
      kind: 'doughnut',
      ...base,
      firstSliceAngle: 30,
      holeSize: 40,
      varyColors: true,
    });
    expect(normalizeChartSpec({ kind: 'bar', ...everything }, context)).toEqual({
      kind: 'bar',
      ...base,
      grouping: 'stacked',
      gapWidth: 50,
      overlap: 10,
      varyColors: true,
      xAxis: { min: 1 },
      yAxis: { max: 9 },
    });
    expect(normalizeChartSpec({ kind: 'scatter', ...everything }, context)).toEqual({
      kind: 'scatter',
      ...base,
      scatterStyle: 'line',
      varyColors: true,
      xAxis: { min: 1 },
      yAxis: { max: 9 },
      seriesStyles: [
        {
          seriesId: 'series-0',
          marker: { visible: false },
          // The point copies the hidden series marker; a size alone does not show it.
          points: [{ pointId: 'point-0', marker: { visible: false } }],
        },
      ],
    });
  });

  it('expands the palette into series fill colours', () => {
    expect(
      normalizeChartSpec(
        {
          kind: 'column',
          palette: ['#AA0000', '#00bb00'],
          seriesStyles: [
            { seriesId: 'series-1', fill: { color: '#123', transparency: 20 } },
            { seriesId: 'series-7', fill: { color: '#999999' } },
          ],
        },
        context,
      ).seriesStyles,
    ).toEqual([
      { seriesId: 'series-0', fill: { color: '#aa0000' } },
      { seriesId: 'series-1', fill: { color: '#112233', transparency: 20 } },
      { seriesId: 'series-2', fill: { color: '#aa0000' } },
    ]);
  });

  it('stores a line or scatter series colour in its line and shown markers, with no fill', () => {
    expect(
      normalizeChartSpec(
        {
          kind: 'line',
          palette: ['#AA0000', '#00bb00'],
          seriesStyles: [
            {
              seriesId: 'series-0',
              fill: { transparency: 20 },
              line: { transparency: 10, width: 2 },
              marker: { visible: true, fill: { transparency: 30 } },
              points: [{ pointId: 'category-1', fill: { color: '#999999', transparency: 60 } }],
            },
            {
              seriesId: 'series-1',
              fill: { color: '#123' },
              line: { color: '#456' },
              marker: { visible: false },
            },
          ],
        },
        context,
      ).seriesStyles,
    ).toEqual([
      {
        seriesId: 'series-0',
        line: { color: '#aa0000', transparency: 10, width: 2 },
        // An automatic marker takes the series' symbol, a diamond for idx 0, and size 5.
        marker: {
          visible: true,
          shape: 'diamond',
          size: 5,
          fill: { color: '#aa0000', transparency: 30 },
          line: { color: '#aa0000' },
        },
        points: [
          // A point's fill colour goes to its line and marker over the copied series colour.
          {
            pointId: 'category-1',
            line: { color: '#999999', transparency: 10, width: 2 },
            marker: {
              visible: true,
              shape: 'diamond',
              size: 5,
              fill: { color: '#999999', transparency: 30 },
              line: { color: '#999999' },
            },
          },
        ],
      },
      { seriesId: 'series-1', line: { color: '#445566' }, marker: { visible: false } },
      { seriesId: 'series-2', line: { color: '#aa0000' } },
    ]);
    expect(
      normalizeChartSpec(
        {
          kind: 'scatter',
          scatterStyle: 'marker',
          seriesStyles: [
            {
              seriesId: 'series-0',
              fill: { color: '#123456', transparency: 20 },
              marker: { visible: true, shape: 'square' },
            },
          ],
        },
        context,
      ).seriesStyles,
    ).toEqual([
      {
        seriesId: 'series-0',
        marker: {
          visible: true,
          shape: 'square',
          size: 5,
          fill: { color: '#123456' },
          line: { color: '#123456' },
        },
      },
    ]);
  });

  it('stores labels in their read-back form, hoisting labels every series shares', () => {
    const shown = {
      visible: true,
      showSeriesName: false,
      showCategoryName: false,
      showValue: true,
      showPercentage: false,
      showLegendKey: false,
    };
    expect(normalizeChartSpec({ kind: 'line', dataLabels: {} }, context).dataLabels).toEqual(shown);
    expect(
      normalizeChartSpec({ kind: 'line', dataLabels: { visible: false, showValue: true } }, context)
        .dataLabels,
    ).toBeUndefined();
    expect(
      normalizeChartSpec(
        {
          kind: 'line',
          dataLabels: { showPercentage: true, color: '#ABC' },
          seriesStyles: [
            { seriesId: 'series-0', dataLabels: { showPercentage: true, color: '#aabbcc' } },
          ],
        },
        context,
      ),
    ).toMatchObject({
      dataLabels: { ...shown, showValue: false, showPercentage: true, color: '#aabbcc' },
    });
    const mixed = normalizeChartSpec(
      {
        kind: 'line',
        dataLabels: { visible: true },
        seriesStyles: [{ seriesId: 'series-2', dataLabels: { visible: false } }],
      },
      context,
    );
    expect(mixed.dataLabels).toBeUndefined();
    expect(mixed.seriesStyles).toEqual([
      { seriesId: 'series-0', dataLabels: shown },
      { seriesId: 'series-1', dataLabels: shown },
    ]);
  });

  it('stores the legend, filter and axis titles in their read-back form', () => {
    const spec = normalizeChartSpec(
      {
        kind: 'area',
        legend: { visible: false, position: 'left' },
        filter: {
          hiddenSeriesIds: ['series-2', 'series-0', 'series-9', 'series-0'],
          hiddenCategoryIds: ['category-3', 'category-1', 'point-0'],
        },
        xAxis: { title: '' },
        yAxis: { title: 'Units', crossing: 'auto' },
      },
      context,
    );
    expect(spec.legend).toEqual({ visible: false });
    expect(spec.filter).toEqual({
      hiddenSeriesIds: ['series-0', 'series-2'],
      hiddenCategoryIds: ['category-1', 'category-3'],
    });
    expect(spec.xAxis).toBeUndefined();
    expect(spec.yAxis).toEqual({ title: 'Units', crossing: 'auto' });
    expect(
      normalizeChartSpec({ kind: 'area', legend: { position: 'top' } }, context).legend,
    ).toEqual({
      visible: true,
      position: 'top',
    });
    expect(normalizeChartSpec({ kind: 'area', legend: {} }, context).legend).toEqual({
      visible: true,
      position: 'bottom',
    });
    expect(
      normalizeChartSpec({ kind: 'area', filter: { hiddenSeriesIds: [] } }, context).filter,
    ).toBeUndefined();
    expect(
      normalizeChartSpec(
        { kind: 'scatter', filter: { hiddenCategoryIds: ['point-4', 'category-1'] } },
        context,
      ).filter,
    ).toEqual({ hiddenCategoryIds: ['point-4'] });
  });

  it('keeps the title as written', () => {
    expect(normalizeChartSpec({ kind: 'pie', title: '' }, context).title).toBe('');
    expect(normalizeChartSpec({ kind: 'pie', title: 'Share' }, context).title).toBe('Share');
    // A chart created without a title text keeps the template's automatic title.
    expect(normalizeChartSpec({ kind: 'pie' }, context).title).toEqual({ auto: true });
    expect(normalizeChartSpec({ kind: 'pie', title: { auto: true } }, context).title).toEqual({
      auto: true,
    });
  });

  it('keeps an automatic axis title as written', () => {
    const spec = normalizeChartSpec({ kind: 'line', yAxis: { title: { auto: true } } }, context);
    expect(spec.yAxis?.title).toEqual({ auto: true });
  });

  it('drops the properties of a hidden marker and a marker-only series line', () => {
    const spec = normalizeChartSpec(
      {
        kind: 'scatter',
        scatterStyle: 'marker',
        seriesStyles: [
          {
            seriesId: 'series-0',
            line: { color: '#333333', width: 2 },
            marker: { visible: false, shape: 'square', size: 7, fill: { color: '#aa0000' } },
            points: [{ pointId: 'point-1', marker: { visible: false, size: 9 } }],
          },
        ],
      },
      context,
    );
    expect(spec.seriesStyles).toEqual([
      {
        seriesId: 'series-0',
        marker: { visible: false },
        points: [{ pointId: 'point-1', marker: { visible: false } }],
      },
    ]);
    expect(
      normalizeChartSpec(
        {
          kind: 'scatter',
          scatterStyle: 'lineMarker',
          seriesStyles: [{ seriesId: 'series-0', line: { width: 2 } }],
        },
        context,
      ).seriesStyles,
    ).toEqual([{ seriesId: 'series-0', line: { width: 2 } }]);
  });
});
