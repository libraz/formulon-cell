import { primitive } from './scene-primitives.js';
import { chartAreaFill, chartAreaLine, dashAttrs } from './styles.js';
import type { ChartPrimitive, ChartSpec } from './types.js';

/** The plot area's rectangle where the file states a fill or an outline for it; none otherwise. */
export const plotAreaPrimitives = (
  spec: ChartSpec,
  plot: { x: number; y: number; width: number; height: number },
): ChartPrimitive[] => {
  const area = spec.plotArea;
  const fill = chartAreaFill(area);
  const line = area?.line === undefined ? null : chartAreaLine(area);
  if (!fill && !line) return [];
  return [
    primitive(
      'rect',
      {
        x: plot.x,
        y: plot.y,
        width: plot.width,
        height: plot.height,
        fill: fill ? fill.color : 'none',
        ...(fill && fill.opacity < 1 ? { 'fill-opacity': fill.opacity } : {}),
        ...(line
          ? {
              stroke: line.color,
              'stroke-width': line.widthPx,
              ...(line.opacity === undefined ? {} : { 'stroke-opacity': line.opacity }),
              ...dashAttrs(line.dash, line.widthPx),
            }
          : {}),
      },
      'plot-area',
    ),
  ];
};
