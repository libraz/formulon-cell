import { describe, expect, it } from 'vitest';
import { layoutChart } from '../../../src/charts/layout.js';
import { dashAttrs, PX_PER_PT } from '../../../src/charts/styles.js';
import type { ChartData, ChartSpec } from '../../../src/charts/types.js';

const CATEGORY: ChartData = {
  domain: 'category',
  categories: [{ label: 'a' }, { label: 'b' }, { label: 'c' }],
  series: [{ id: 'series-0', name: 'S', points: [{ value: 1 }, { value: 3 }, { value: 2 }] }],
};

const XY: ChartData = {
  domain: 'xy',
  series: [
    {
      id: 'series-0',
      points: [
        { id: 'p0', x: 0, y: 0 },
        { id: 'p1', x: 5, y: 5 },
      ],
    },
  ],
};

const scene = (spec: ChartSpec, data: ChartData = CATEGORY) =>
  layoutChart({ spec, data, width: 320, height: 132 });

const role = (spec: ChartSpec, name: string, data?: ChartData) =>
  scene(spec, data).elements.filter((element) => element.role === name);

describe('stated line widths are points, drawn in scene pixels', () => {
  it('draws a 2.25 pt series line 3 px wide', () => {
    const [line] = role(
      { kind: 'line', seriesStyles: [{ seriesId: 'series-0', line: { width: 2.25 } }] },
      'line',
    );
    expect(line?.attrs['stroke-width']).toBe(3);
  });

  it('draws a 1 pt marker outline 4/3 px wide', () => {
    const [point] = role(
      {
        kind: 'scatter',
        scatterStyle: 'marker',
        seriesStyles: [
          {
            seriesId: 'series-0',
            marker: { shape: 'circle', line: { color: '#123456', width: 1 } },
          },
        ],
      },
      'point',
      XY,
    );
    expect(point?.attrs['stroke-width']).toBeCloseTo(4 / 3, 10);
  });

  it('scales a dash pattern with the drawn width', () => {
    const [line] = role(
      {
        kind: 'line',
        seriesStyles: [{ seriesId: 'series-0', line: { width: 2.25, dash: 'dash' } }],
      },
      'line',
    );
    expect(line?.attrs['stroke-dasharray']).toBe(dashAttrs('dash', 3)['stroke-dasharray']);
    expect(line?.attrs['stroke-dasharray']).toBe('12 9');
  });

  it('draws a stated column outline in pixels too', () => {
    const [bar] = role(
      {
        kind: 'column',
        seriesStyles: [{ seriesId: 'series-0', line: { color: '#000000', width: 3 } }],
      },
      'bar',
    );
    expect(bar?.attrs['stroke-width']).toBe(3 * PX_PER_PT);
  });

  // The spreadsheet reads a series line that states no width as 3 pt (lnw-2026-11-18).
  it('draws a line series that states no width 3 pt wide', () => {
    const [line] = role({ kind: 'line' }, 'line');
    expect(line?.attrs['stroke-width']).toBe(4);
  });

  it('draws a scatter connector that states no width 3 pt wide', () => {
    const [line] = role({ kind: 'scatter', scatterStyle: 'lineMarker' }, 'line', XY);
    expect(line?.attrs['stroke-width']).toBe(4);
  });
});

describe('widths measured in pixels are unchanged', () => {
  it('draws the separators of pie slices 1 px wide', () => {
    const slices = role({ kind: 'pie' }, 'slice');
    expect(slices.length).toBeGreaterThan(0);
    for (const slice of slices) expect(slice.attrs['stroke-width']).toBe(1);
  });

  it('leaves gridlines and axes to the default hairline', () => {
    const lines = role({ kind: 'column' }, 'axis');
    expect(lines.length).toBeGreaterThan(0);
    for (const line of lines) expect(line.attrs).not.toHaveProperty('stroke-width');
  });

  it('suppresses a series line stated 0 pt wide', () => {
    expect(
      role(
        {
          kind: 'scatter',
          scatterStyle: 'lineMarker',
          seriesStyles: [{ seriesId: 'series-0', line: { width: 0 } }],
        },
        'line',
        XY,
      ),
    ).toEqual([]);
  });
});
