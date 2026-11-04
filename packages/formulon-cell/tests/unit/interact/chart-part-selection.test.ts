import { describe, expect, it } from 'vitest';
import {
  chartPartAt,
  paintChartPart,
  partKeepsChartHandles,
} from '../../../src/interact/chart-part-selection.js';
import { SVG_NS } from '../../../src/svg-namespace.js';

const el = (tag: string, attrs: Record<string, string | number>, text?: string): SVGElement => {
  const node = document.createElementNS(SVG_NS, tag);
  for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
  if (text) node.textContent = text;
  return node as SVGElement;
};

const scene = (): SVGSVGElement => {
  const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
  for (const s of [0, 1]) {
    for (const p of [0, 1, 2]) {
      svg.append(
        el('rect', {
          x: 20 + p * 30 + s * 10,
          y: 50,
          width: 8,
          height: 40,
          'data-chart-role': 'bar',
          'data-series-index': s,
          'data-point-index': p,
        }),
      );
    }
  }
  svg.append(
    el('text', { x: 100, y: 12, 'text-anchor': 'middle', 'data-chart-role': 'title' }, 'Title'),
    el('text', { x: 10, y: 60, 'text-anchor': 'end', 'data-chart-role': 'axis-label' }, '0'),
    el('text', { x: 10, y: 20, 'text-anchor': 'end', 'data-chart-role': 'axis-label' }, '500'),
    el('rect', { x: 40, y: 140, width: 8, height: 8, 'data-chart-role': 'legend' }),
    el('line', { x1: 15, x2: 190, y1: 50, y2: 50, 'data-chart-role': 'axis-tick' }),
    el('line', { x1: 15, x2: 190, y1: 90, y2: 90, 'data-chart-role': 'axis-tick' }),
  );
  return svg;
};

const count = (svg: SVGSVGElement, selector: string): number =>
  svg.querySelectorAll(`.fc-chart__parts ${selector}`).length;

describe('chartPartAt', () => {
  it('selects the series on the first press and the point on the next press of its bar', () => {
    const svg = scene();
    const bar = svg.querySelector('[data-series-index="0"][data-point-index="1"]');
    const series = chartPartAt(bar, null);
    expect(series).toEqual({ kind: 'series', seriesIndex: 0 });
    expect(chartPartAt(bar, series)).toEqual({ kind: 'point', seriesIndex: 0, pointIndex: 1 });
    expect(chartPartAt(svg.querySelector('[data-series-index="1"]'), series)).toEqual({
      kind: 'series',
      seriesIndex: 1,
    });
  });

  it('names the title, legend, axis labels and gridlines and ignores the blank area', () => {
    const svg = scene();
    const at = (role: string) =>
      chartPartAt(svg.querySelector(`[data-chart-role="${role}"]`), null);
    expect(at('title')).toEqual({ kind: 'title' });
    expect(at('legend')).toEqual({ kind: 'legend' });
    expect(at('axis-label')).toEqual({ kind: 'axisLabels', role: 'axis-label' });
    expect(at('axis-tick')).toEqual({ kind: 'gridlines' });
    expect(chartPartAt(svg, null)).toBeNull();
  });
});

describe('paintChartPart', () => {
  it('rings every bar of a series with four circle handles and leaves the other series alone', () => {
    const svg = scene();
    paintChartPart(svg, { kind: 'series', seriesIndex: 0 });
    expect(count(svg, 'circle')).toBe(12);
    paintChartPart(svg, { kind: 'point', seriesIndex: 0, pointIndex: 1 });
    expect(count(svg, 'circle')).toBe(4);
    paintChartPart(svg, null);
    expect(svg.querySelector('.fc-chart__parts')).toBeNull();
  });

  it('frames the title and axis labels with four circles and the legend with eight squares', () => {
    const svg = scene();
    paintChartPart(svg, { kind: 'title' });
    expect(count(svg, '.fc-chart__part-frame')).toBe(1);
    expect(count(svg, 'circle')).toBe(4);
    paintChartPart(svg, { kind: 'axisLabels', role: 'axis-label' });
    expect(count(svg, 'circle')).toBe(4);
    paintChartPart(svg, { kind: 'legend' });
    expect(count(svg, '.fc-chart__part-handle--square')).toBe(8);
    expect(count(svg, 'circle')).toBe(0);
  });

  it('puts a circle at both ends of every gridline', () => {
    const svg = scene();
    paintChartPart(svg, { kind: 'gridlines' });
    expect(count(svg, 'circle')).toBe(4);
  });

  it('keeps the outer chart handles except for a title', () => {
    expect(partKeepsChartHandles({ kind: 'title' })).toBe(false);
    expect(partKeepsChartHandles({ kind: 'legend' })).toBe(true);
    expect(partKeepsChartHandles(null)).toBe(true);
  });
});
