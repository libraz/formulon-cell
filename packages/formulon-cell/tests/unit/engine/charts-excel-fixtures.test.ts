// @vitest-environment node
/**
 * Charts Excel saved open in the cell. The expectations are written by hand
 * from what Excel stored in each fixture (series references, caches, plot and
 * series formatting), not from the adapter's own output.
 */
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { ChartData, ChartSpec } from '../../../src/charts/types.js';
import { History } from '../../../src/commands/history.js';
import { clearSessionChart, updateSessionChart } from '../../../src/commands/session-chart.js';
import { hydrateChartsFromEngine } from '../../../src/engine/charts-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, type SessionChart } from '../../../src/store/store.js';
import { zipEntries } from './zip-entries.js';

const DIR = join(__dirname, '../../fixtures/charts');
const bytesOf = (name: string): Uint8Array => new Uint8Array(readFileSync(join(DIR, name)));

const ACCENT = ['#156082', '#e97132', '#196b24', '#0f9ed5'];
const S1 = { id: 'series-0', name: 'S1', values: [1, 2, 3, 7] };
const S2 = { id: 'series-1', name: 'S2', values: [4, 5, 6, 8] };
const CATEGORIES = ['a', 'b', 'c', 'd'];
const SPLIT = { seriesDirection: 'columns', headerRow: true, headerColumn: true } as const;
const SHOWN = {
  emptyCells: 'gap',
  legend: { visible: true, position: 'bottom' },
} as const satisfies Partial<ChartSpec>;
const varyingPoints = ACCENT.map((color, index) => ({
  pointId: `category-${index}`,
  fill: { color },
}));

interface Expectation {
  readonly file: string;
  readonly name: string;
  readonly spec: Partial<ChartSpec>;
  /** The title text; null is Excel's automatic title (a c:title without c:tx). */
  readonly title: string | null;
  readonly data: ChartData;
}

const EXPECTED: readonly Expectation[] = [
  {
    file: 'chart_column.xlsx',
    name: 'col_clustered',
    spec: {
      kind: 'column',
      grouping: 'clustered',
      gapWidth: 219,
      varyColors: false,
      ...SHOWN,
      ...SPLIT,
      seriesStyles: [
        { seriesId: 'series-0', fill: { color: ACCENT[0] } },
        { seriesId: 'series-1', fill: { color: ACCENT[1] } },
      ],
    },
    // No c:title element: no title.
    title: '',
    data: { domain: 'category', categories: CATEGORIES, series: [S1, S2] },
  },
  {
    file: 'chart_bar.xlsx',
    name: 'bar_stacked',
    spec: {
      kind: 'bar',
      grouping: 'stacked',
      gapWidth: 182,
      overlap: 100,
      varyColors: false,
      ...SHOWN,
      ...SPLIT,
      seriesStyles: [
        { seriesId: 'series-0', fill: { color: ACCENT[0] } },
        { seriesId: 'series-1', fill: { color: ACCENT[1] } },
      ],
    },
    title: '',
    data: { domain: 'category', categories: CATEGORIES, series: [S1, S2] },
  },
  {
    file: 'chart_line.xlsx',
    name: 'line_markers',
    spec: {
      kind: 'line',
      // Excel's "standard" grouping.
      grouping: 'clustered',
      varyColors: false,
      ...SHOWN,
      ...SPLIT,
      seriesStyles: ACCENT.slice(0, 2).map((color, index) => ({
        seriesId: `series-${index}`,
        // w=28575 and w=9525 EMU.
        line: { color, width: 2.25 },
        marker: {
          visible: true,
          shape: 'circle',
          size: 5,
          fill: { color },
          line: { color, width: 0.75 },
        },
      })),
    },
    title: '',
    data: { domain: 'category', categories: CATEGORIES, series: [S1, S2] },
  },
  {
    file: 'chart_area.xlsx',
    name: 'area_percent',
    spec: {
      kind: 'area',
      grouping: 'percent-stacked',
      varyColors: false,
      ...SHOWN,
      ...SPLIT,
      seriesStyles: [
        { seriesId: 'series-0', fill: { color: ACCENT[0] } },
        { seriesId: 'series-1', fill: { color: ACCENT[1] } },
      ],
    },
    title: '',
    data: { domain: 'category', categories: CATEGORIES, series: [S1, S2] },
  },
  {
    file: 'chart_pie.xlsx',
    name: 'pie_single',
    spec: {
      kind: 'pie',
      varyColors: true,
      firstSliceAngle: 0,
      ...SHOWN,
      ...SPLIT,
      seriesStyles: [{ seriesId: 'series-0', points: varyingPoints }],
    },
    // A c:title without c:tx: Excel's automatic title.
    title: null,
    data: { domain: 'category', categories: CATEGORIES, series: [S1] },
  },
  {
    file: 'chart_doughnut.xlsx',
    name: 'doughnut',
    spec: {
      kind: 'doughnut',
      varyColors: true,
      firstSliceAngle: 0,
      holeSize: 50,
      ...SHOWN,
      ...SPLIT,
      seriesStyles: [{ seriesId: 'series-0', points: varyingPoints }],
    },
    title: null,
    data: { domain: 'category', categories: CATEGORIES, series: [S1] },
  },
  {
    file: 'chart_scatter.xlsx',
    name: 'scatter_lines',
    spec: {
      kind: 'scatter',
      scatterStyle: 'lineMarker',
      varyColors: false,
      ...SHOWN,
      // E1:G5 with the x values in the first column, which is no header.
      seriesDirection: 'columns',
      headerRow: true,
      headerColumn: false,
      seriesStyles: ACCENT.slice(0, 2).map((color, index) => ({
        seriesId: `series-${index}`,
        line: { color, width: 1.5 },
        marker: {
          visible: true,
          shape: 'circle',
          size: 5,
          fill: { color },
          line: { color, width: 0.75 },
        },
      })),
    },
    title: '',
    data: {
      domain: 'xy',
      series: [
        {
          id: 'series-0',
          name: 'Y1',
          points: [
            { id: 'point-0', x: 1, y: 2, state: 'value' },
            { id: 'point-1', x: 2, y: 4, state: 'value' },
            { id: 'point-2', x: 3, y: 1, state: 'value' },
            { id: 'point-3', x: 5, y: 3, state: 'value' },
          ],
        },
        {
          id: 'series-1',
          name: 'Y2',
          points: [
            { id: 'point-0', x: 1, y: 3, state: 'value' },
            { id: 'point-1', x: 2, y: 1, state: 'value' },
            { id: 'point-2', x: 3, y: 5, state: 'value' },
            { id: 'point-3', x: 5, y: 2, state: 'value' },
          ],
        },
      ],
    },
  },
  // The xlsb fixture stores chart_column.xlsx's chart parts byte for byte.
  {
    file: 'chart_xlsb.xlsb',
    name: 'col_clustered',
    spec: {
      kind: 'column',
      grouping: 'clustered',
      gapWidth: 219,
      ...SHOWN,
      ...SPLIT,
    },
    title: '',
    data: { domain: 'category', categories: CATEGORIES, series: [S1, S2] },
  },
];

/** Charts Excel stored that the cell draws only as a placeholder. */
const PLACEHOLDERS: Readonly<Record<string, string>> = {
  'chart_chartex.xlsx:cx1': 'chartex',
  'chart_combo.xlsx:combo': 'multi-plot',
};

const load = async (
  file: string,
): Promise<{ wb: WorkbookHandle; charts: (SessionChart & { name: string })[] }> => {
  const wb = await WorkbookHandle.loadBytes(bytesOf(file));
  const store = createSpreadsheetStore();
  const charts: (SessionChart & { name: string })[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    hydrateChartsFromEngine(wb, store, sheet);
    const names = new Map(
      wb.listDrawingObjects(sheet).map((object) => [object.objectId, object.name]),
    );
    for (const chart of store.getState().charts.charts) {
      if (chart.sheet === sheet)
        charts.push({ ...chart, name: names.get(chart.nativeObjectId) ?? '' });
    }
  }
  return { wb, charts };
};

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe.skipIf(!canLoadWasm())('charts Excel saved', () => {
  it.each(EXPECTED.map((entry) => [`${entry.file} ${entry.name}`, entry] as const))(
    'reads %s as Excel stored it',
    async (_label, expected) => {
      const { wb, charts } = await load(expected.file);
      try {
        const chart = charts.find((candidate) => candidate.name === expected.name);
        expect(chart?.kind).toBe(expected.spec.kind);
        expect(chart).toMatchObject(expected.spec);
        // A null title is Excel's automatic one.
        expect(chart?.title).toEqual(expected.title ?? { auto: true });
        expect(chart?.data).toMatchObject(expected.data);
      } finally {
        wb.dispose();
      }
    },
  );

  it('draws every single-plot chart of every fixture', async () => {
    const files = readdirSync(DIR).filter((file) => file.startsWith('chart_'));
    const seen: string[] = [];
    for (const file of files) {
      const { wb, charts } = await load(file);
      try {
        for (const chart of charts) {
          const key = `${file}:${chart.name}`;
          seen.push(key);
          if (PLACEHOLDERS[key]) {
            expect(chart, key).toMatchObject({
              kind: 'unsupported',
              unsupported: { reason: PLACEHOLDERS[key] },
            });
          } else {
            expect(chart.kind, key).not.toBe('unsupported');
          }
        }
      } finally {
        wb.dispose();
      }
    }
    expect(seen).toEqual(expect.arrayContaining(Object.keys(PLACEHOLDERS)));
  });

  it.each([
    ['chart_combo.xlsx', 'combo', 'xl/charts/chart1.xml'],
    ['chart_chartex.xlsx', 'cx1', 'xl/charts/chartEx1.xml'],
  ])('keeps %s %s byte for byte through a move and a delete', async (file, name, part) => {
    const original = bytesOf(file);
    const wb = await WorkbookHandle.loadBytes(original);
    try {
      const store = createSpreadsheetStore();
      const history = new History();
      hydrateChartsFromEngine(wb, store, 0);
      const names = new Map(
        wb.listDrawingObjects(0).map((object) => [object.objectId, object.name]),
      );
      const charts = store.getState().charts.charts;
      const placeholder = charts.find((chart) => names.get(chart.nativeObjectId) === name);
      if (!placeholder) throw new Error('placeholder chart missing');
      expect(placeholder.kind).toBe('unsupported');

      const moved = updateSessionChart(
        store,
        placeholder.id,
        { anchor: { row: placeholder.nativeAnchor.fromRow + 3, col: 1 } },
        history,
        wb,
      );
      expect(moved?.nativeAnchor.fromRow).toBe(placeholder.nativeAnchor.fromRow + 3);
      const other = charts.find((chart) => chart.id !== placeholder.id);
      if (other) expect(clearSessionChart(store, other.id, history, wb)).toBe(true);

      const saved = zipEntries(wb.save());
      const before = zipEntries(original);
      expect(saved.get(part)).toEqual(before.get(part));
    } finally {
      wb.dispose();
    }
  });
});
