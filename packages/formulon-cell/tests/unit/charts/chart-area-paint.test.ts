import { ChartLineDash } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { chartSpecFromEngine } from '../../../src/charts/engine-adapter.js';
import { layoutChart } from '../../../src/charts/layout.js';
import type { ChartSpec } from '../../../src/charts/types.js';
import { History } from '../../../src/commands/history.js';
import { createSessionChart, updateSessionChart } from '../../../src/commands/session-chart.js';
import { hydrateChartsFromEngine } from '../../../src/engine/charts-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { internals } from '../../../src/engine/workbook-handle-internals.js';
import { attachSessionCharts } from '../../../src/interact/session-charts.js';
import { createSpreadsheetStore, type SessionChart } from '../../../src/store/store.js';

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
});

const DATA = {
  domain: 'category',
  categories: [{ label: 'a' }, { label: 'b' }],
  series: [{ id: 'series-0', name: 'S', points: [{ value: 1 }, { value: 2 }] }],
} as const;

/** A workbook with one column chart; `paint` writes the areas' paint into the engine first. */
async function chartWith(paint?: Record<string, unknown>) {
  const wb = await WorkbookHandle.createDefault();
  cleanup.push(() => wb.dispose());
  wb.setText({ sheet: 0, row: 0, col: 1 }, 'S');
  wb.setText({ sheet: 0, row: 1, col: 0 }, 'a');
  wb.setNumber({ sheet: 0, row: 1, col: 1 }, 1);
  wb.setText({ sheet: 0, row: 2, col: 0 }, 'b');
  wb.setNumber({ sheet: 0, row: 2, col: 1 }, 2);
  const store = createSpreadsheetStore();
  const history = new History();
  const chart = createSessionChart(
    store,
    { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
    { wb, kind: 'column' },
  );
  if (!chart) throw new Error('no chart');
  if (paint) {
    internals(wb).wb.sheet(0).chart(chart.nativeObjectId).updateDisplay(paint);
  }
  return { wb, store, history, chart };
}

/** The chart of sheet 0 of a saved and reloaded workbook, as the store holds it. */
async function reloaded(wb: WorkbookHandle): Promise<SessionChart> {
  const loaded = await WorkbookHandle.loadBytes(wb.save());
  cleanup.push(() => loaded.dispose());
  const store = createSpreadsheetStore();
  hydrateChartsFromEngine(loaded, store, 0);
  const [chart] = store.getState().charts.charts;
  if (!chart) throw new Error('no chart after reload');
  return chart;
}

const STATED = {
  chartArea: {
    fill: { color: '#ffeecc', transparencyPct: 20 },
    line: { color: '#ff0000', widthPt: 2, dash: ChartLineDash.Dash },
  },
  plotArea: { fill: { color: '#cceeff' } },
};

describe('the paint of the chart area and the plot area', () => {
  it('is read as the file states it', async () => {
    const { wb } = await chartWith(STATED);
    const chart = await reloaded(wb);

    expect(chart.chartArea).toMatchObject({
      fill: { color: '#ffeecc', transparency: 20 },
      line: { color: '#ff0000', width: 2, dash: 'dash' },
    });
    expect(chart.plotArea?.fill).toMatchObject({ color: '#cceeff' });
  });

  it('reads a stated No Fill and No Line as such', async () => {
    const { wb } = await chartWith({
      chartArea: { fill: { none: true }, line: { none: true } },
    });
    expect((await reloaded(wb)).chartArea).toEqual({ fill: 'none', line: 'none' });
  });

  it('is unset where the file states none', async () => {
    const { wb, chart } = await chartWith();
    const info = wb.getChart(0, chart.nativeObjectId);
    if (!info) throw new Error('no chart');
    const bare = { fill: null, line: null };

    const spec = chartSpecFromEngine({ ...info, chartArea: bare, plotArea: bare }, null);

    expect(spec).not.toBeNull();
    expect(spec).not.toHaveProperty('chartArea');
    expect(spec).not.toHaveProperty('plotArea');
  });

  it('survives an edit that does not touch it, and a save', async () => {
    const { wb, store, history, chart } = await chartWith(STATED);
    hydrateChartsFromEngine(wb, store, 0);

    expect(updateSessionChart(store, chart.id, { title: 'Sales' }, history, wb)).not.toBeNull();

    const after = await reloaded(wb);
    expect(after.title).toBe('Sales');
    expect(after.chartArea?.fill).toEqual({ color: '#ffeecc', transparency: 20 });
    expect(after.chartArea?.line).toMatchObject({ color: '#ff0000', width: 2 });
    expect(after.plotArea?.fill).toMatchObject({ color: '#cceeff' });
    const info = wb.getChart(0, chart.nativeObjectId);
    expect(info && chartSpecFromEngine(info, null)?.chartArea?.line).toMatchObject({ width: 2 });
  });
});

describe('the drawn plot area', () => {
  const scene = (spec: Partial<ChartSpec>) =>
    layoutChart({ spec: { kind: 'column', ...spec }, data: DATA, width: 320, height: 132 });
  const plotArea = (spec: Partial<ChartSpec>) =>
    scene(spec).elements.find((element) => element.role === 'plot-area');

  it('is painted over the plot where the file states a fill and an outline', () => {
    const rect = plotArea({
      plotArea: {
        fill: { color: '#cceeff', transparency: 50 },
        line: { color: '#112233', width: 2 },
      },
    });
    expect(rect?.attrs).toMatchObject({
      fill: '#cceeff',
      'fill-opacity': 0.5,
      stroke: '#112233',
      'stroke-width': 2,
    });
  });

  it('is not drawn where the file states none, or No Fill and No Line', () => {
    expect(plotArea({})).toBeUndefined();
    expect(plotArea({ plotArea: { fill: 'none', line: 'none' } })).toBeUndefined();
  });
});

describe('the drawn chart area', () => {
  const drawn = (spec: Partial<ChartSpec>) => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    cleanup.push(() => host.remove());
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      charts: {
        charts: [
          {
            id: 'c1',
            sheet: 0,
            nativeObjectId: 1,
            nativeAnchor: {
              kind: 1,
              editAs: 1,
              fromRow: 2,
              fromCol: 2,
              fromRowOffsetEmu: 0,
              fromColOffsetEmu: 0,
              toRow: 0,
              toCol: 0,
              toRowOffsetEmu: 0,
              toColOffsetEmu: 0,
              widthEmu: 3000000,
              heightEmu: 2000000,
            },
            data: DATA,
            kind: 'column',
            ...spec,
          } as SessionChart,
        ],
      },
    }));
    const handle = attachSessionCharts({ host, store });
    cleanup.push(() => handle.detach());
    const panel = host.querySelector<HTMLElement>('.fc-chart');
    if (!panel) throw new Error('no chart panel');
    return { style: panel.style, frame: panel.querySelector('.fc-chart__frame rect') };
  };

  it('is the style default where the file states none: a surface fill and a 1 pt grey outline', () => {
    const { style, frame } = drawn({});
    expect(style.getPropertyValue('--fc-chart-area-fill')).toBe('');
    expect(style.getPropertyValue('--fc-chart-area-line-width')).toBe('1pt');
    expect(frame?.getAttribute('stroke')).toBe('#898989');
    expect(Number(frame?.getAttribute('stroke-width'))).toBeCloseTo(96 / 72, 5);
    expect(frame?.getAttribute('stroke-dasharray')).toBeNull();
  });

  it('is the fill and outline the file states', () => {
    const { style, frame } = drawn({
      chartArea: {
        fill: { color: '#ffeecc', transparency: 20 },
        line: { color: '#ff0000', width: 3, transparency: 25 },
      },
    });
    expect(style.getPropertyValue('--fc-chart-area-fill')).toContain('#ffeecc');
    expect(style.getPropertyValue('--fc-chart-area-fill')).toContain('80%');
    expect(style.getPropertyValue('--fc-chart-area-line-width')).toBe('3pt');
    expect(frame?.getAttribute('stroke')).toBe('#ff0000');
    expect(frame?.getAttribute('stroke-width')).toBe('4');
    expect(frame?.getAttribute('stroke-opacity')).toBe('0.75');
  });

  it('draws no fill and no outline where the file states them off', () => {
    const { style, frame } = drawn({ chartArea: { fill: 'none', line: 'none' } });
    expect(style.getPropertyValue('--fc-chart-area-fill')).toBe('transparent');
    expect(style.getPropertyValue('--fc-chart-area-line-width')).toBe('0pt');
    expect(frame).toBeNull();
  });

  it.each(['squareDot', 'longDashDot', 'roundDot'] as const)(
    'dashes a %s outline as a series line of that dash is',
    (dash) => {
      // 3 pt is 4 px, the width the series line below is drawn at.
      const { frame } = drawn({ chartArea: { line: { color: '#ff0000', width: 3, dash } } });
      const line = layoutChart({
        spec: { kind: 'line', seriesStyles: [{ seriesId: 'series-0', line: { width: 4, dash } }] },
        data: DATA,
        width: 320,
        height: 132,
      }).elements.find((element) => element.role === 'line');

      expect(line?.attrs['stroke-dasharray']).toBeDefined();
      expect(frame?.getAttribute('stroke-dasharray')).toBe(String(line?.attrs['stroke-dasharray']));
    },
  );
});
