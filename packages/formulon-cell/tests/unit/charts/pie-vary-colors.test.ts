import { ChartBarDirection, ChartPlotType } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { chartSpecFromEngine } from '../../../src/charts/engine-adapter.js';
import { automaticPalette } from '../../../src/charts/series-colors.js';
import { resolveChartStyle } from '../../../src/charts/styles.js';
import type { ChartSpec } from '../../../src/charts/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const palette = automaticPalette();

const pointColors = (spec: ChartSpec): string[] =>
  [0, 1, 2, 3].map(
    (i) =>
      resolveChartStyle(spec, palette, 'series-0', `category-${i}`, 0, i).fill?.color ??
      resolveChartStyle(spec, palette, 'series-0', `category-${i}`, 0, i).color,
  );

describe('pie colours with varyColors', () => {
  const series = { seriesId: 'series-0', fill: { color: '#156082' } };

  it.each(['pie', 'doughnut'] as const)('%s colours each point over a series fill', (kind) => {
    const colors = pointColors({ kind, varyColors: true, seriesStyles: [series] } as ChartSpec);
    expect(new Set(colors).size).toBe(4);
    expect(colors[1]).toBe(palette[1]);
  });

  it('keeps the series fill when varyColors is off', () => {
    const colors = pointColors({
      kind: 'pie',
      varyColors: false,
      seriesStyles: [series],
    } as ChartSpec);
    expect(new Set(colors)).toEqual(new Set(['#156082']));
  });

  it('lets an explicit point fill win over the sequence', () => {
    const colors = pointColors({
      kind: 'pie',
      varyColors: true,
      seriesStyles: [
        { ...series, points: [{ pointId: 'category-2', fill: { color: '#abcdef' } }] },
      ],
    } as ChartSpec);
    expect(colors[2]).toBe('#abcdef');
    expect(colors[1]).toBe(palette[1]);
  });
});

describe.skipIf(typeof WebAssembly === 'undefined')('pie from the engine', () => {
  const build = async () => {
    const h = await WorkbookHandle.createDefault();
    h.setText({ sheet: 0, row: 0, col: 0 }, 'k');
    h.setText({ sheet: 0, row: 0, col: 1 }, 'v');
    [120, 150, 90, 180].forEach((v, i) => {
      h.setText({ sheet: 0, row: i + 1, col: 0 }, `Q${i + 1}`);
      h.setNumber({ sheet: 0, row: i + 1, col: 1 }, v);
    });
    return h;
  };
  const colorsOf = (h: WorkbookHandle, id: number): string[] => {
    const info = h.getChart(0, id);
    const spec = info ? chartSpecFromEngine(info) : null;
    if (!spec) throw new Error('chart not projected');
    return pointColors(spec);
  };

  it('colours each slice of a pie created as a pie', async () => {
    const h = await build();
    try {
      const id = h.insertChart(0, { type: ChartPlotType.Pie, range: 'Sheet1!$A$1:$B$5' }) as number;
      expect(new Set(colorsOf(h, id)).size).toBe(4);
    } finally {
      h.dispose();
    }
  });

  it('colours each slice of a column chart changed to a pie', async () => {
    const h = await build();
    try {
      const id = h.insertChart(0, {
        type: ChartPlotType.Bar,
        barDirection: ChartBarDirection.Column,
        range: 'Sheet1!$A$1:$B$5',
      }) as number;
      h.setChartType(0, id, 0, ChartPlotType.Pie);
      expect(new Set(colorsOf(h, id)).size).toBe(4);
    } finally {
      h.dispose();
    }
  });
});
