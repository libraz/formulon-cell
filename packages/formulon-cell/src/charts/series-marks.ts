import type { CategoryChartView, XYChartView } from './filters.js';
import {
  categoryIndex,
  type Domain,
  numericProjection,
  reversedAxis,
  xCategory,
} from './plot-scale.js';
import { type ScatterLineSample, scatterLineSegments } from './scatter-lines.js';
import { categoryIdAt, pointIdAt, primitive } from './scene-primitives.js';
import {
  annularPath,
  barGeometry,
  circularFrame,
  doughnutFrame,
  normalizedMagnitudes,
  segmentsFor,
} from './series-geometry.js';
import {
  dashArray,
  fillColor,
  fillOpacity,
  lineOpacity,
  lineWidth,
  markerSize,
  markerVisible,
  resolveChartStyle,
} from './styles.js';
import type { ChartPrimitive, ChartScene, ChartSpec } from './types.js';

/** Whether a line or scatter series draws Excel's smoothed line. */
const smoothed = (spec: ChartSpec, seriesId: string): boolean =>
  (spec.kind === 'line' || spec.kind === 'scatter') &&
  spec.seriesStyles?.find((style) => style.seriesId === seriesId)?.smooth === true;

/**
 * Excel's smoothed line through `points`: cubic Béziers from Catmull-Rom with
 * tension 1/6, the end points repeated as their own neighbours.
 */
const smoothPath = (points: readonly (readonly [number, number])[]): string => {
  const commands = points.slice(0, 1).map(([x, y]) => `M${x},${y}`);
  for (let index = 0; index + 1 < points.length; index += 1) {
    const p1 = points[index] as readonly [number, number];
    const p2 = points[index + 1] as readonly [number, number];
    const p0 = points[index - 1] ?? p1;
    const p3 = points[index + 2] ?? p2;
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    commands.push(`C${c1[0]},${c1[1]} ${c2[0]},${c2[1]} ${p2[0]},${p2[1]}`);
  }
  return commands.join(' ');
};

const styleAttrs = (
  attrs: Record<string, string | number>,
  style: ReturnType<typeof resolveChartStyle>,
  fallbackOpacity?: number,
): Record<string, string | number> => {
  attrs.fill = fillColor(style.fill, style.color);
  const opacity = fillOpacity(style.fill, fallbackOpacity);
  if (opacity !== undefined) attrs['fill-opacity'] = opacity;
  if (style.line) {
    attrs.stroke = style.line.color ?? style.color;
    const strokeOpacity = lineOpacity(style.line);
    if (strokeOpacity !== undefined) attrs['stroke-opacity'] = strokeOpacity;
    const width = lineWidth(style.line);
    if (width !== undefined) attrs['stroke-width'] = width;
    const dash = dashArray(style.line.dash);
    if (dash !== undefined) attrs['stroke-dasharray'] = dash;
  }
  return attrs;
};

const markerAttrs = (
  x: number,
  y: number,
  style: ReturnType<typeof resolveChartStyle>,
): { tag: ChartPrimitive['tag']; attrs: Record<string, string | number> } => {
  const marker = style.marker;
  const size = markerSize(marker);
  const half = size / 2;
  const fill = fillColor(marker.fill ?? style.fill, style.color);
  const opacity = fillOpacity(marker.fill ?? style.fill);
  const line = marker.line;
  const stroke = line === undefined ? undefined : (line.color ?? style.color);
  const strokeOpacity = lineOpacity(line);
  const strokeWidth = lineWidth(line);
  const dash = dashArray(line?.dash);
  const common: Record<string, string | number> = {
    fill,
    ...(opacity === undefined ? {} : { 'fill-opacity': opacity }),
    ...(stroke === undefined ? {} : { stroke }),
    ...(strokeOpacity === undefined ? {} : { 'stroke-opacity': strokeOpacity }),
    ...(strokeWidth === undefined ? {} : { 'stroke-width': strokeWidth }),
    ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
  };
  // Excel's x, star and plus are drawn by their outline alone, in the outline colour.
  const strokedShape = (d: string) => ({
    tag: 'path' as const,
    attrs: {
      d,
      fill: 'none',
      stroke: stroke ?? style.color,
      'stroke-width': strokeWidth ?? 1,
      ...(strokeOpacity === undefined ? {} : { 'stroke-opacity': strokeOpacity }),
      ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
    },
  });
  const cross = `M${x - half},${y - half} L${x + half},${y + half} M${x - half},${y + half} L${x + half},${y - half}`;
  const upright = `M${x},${y - half} L${x},${y + half}`;
  const bar = (width: number) => ({
    tag: 'rect' as const,
    attrs: { x: x - width / 2, y: y - size / 10, width, height: size / 5, ...common },
  });
  switch (marker.shape ?? 'circle') {
    case 'x':
      return strokedShape(cross);
    case 'star':
      return strokedShape(`${cross} ${upright}`);
    case 'plus':
      return strokedShape(`M${x - half},${y} L${x + half},${y} ${upright}`);
    case 'dot':
      return bar(size / 2);
    case 'dash':
      return bar(size);
    case 'square':
      return {
        tag: 'rect',
        attrs: { x: x - half, y: y - half, width: size, height: size, ...common },
      };
    case 'diamond':
      return {
        tag: 'polygon',
        attrs: {
          points: `${x},${y - half} ${x + half},${y} ${x},${y + half} ${x - half},${y}`,
          ...common,
        },
      };
    case 'triangle':
      return {
        tag: 'polygon',
        attrs: {
          points: `${x},${y - half} ${x + half},${y + half} ${x - half},${y + half}`,
          ...common,
        },
      };
    default:
      // A circle, and a picture marker drawn as a plain one in the series colour.
      return { tag: 'circle', attrs: { cx: x, cy: y, r: half, ...common } };
  }
};

export const addCategoryMarks = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  yDomain: Domain,
  palette: readonly string[],
  numericLogBase?: number,
): void => {
  const data = view.data;
  const grouping = spec.grouping ?? 'clustered';
  const segments = segmentsFor(data, grouping);
  const count = Math.max(1, data.categories.length);
  const isBar = spec.kind === 'bar';
  const slot = (isBar ? plot.height : plot.width) / count;
  const numericReverse = isBar ? reversedAxis(spec.xAxis) : reversedAxis(spec.yAxis);
  const categoryReverse = isBar ? reversedAxis(spec.yAxis) : reversedAxis(spec.xAxis);
  const numericScale = (value: number): number => {
    if (isBar) {
      return numericProjection(
        value,
        yDomain,
        plot.x,
        plot.x + plot.width,
        numericReverse,
        numericLogBase,
      );
    }
    return numericProjection(
      value,
      yDomain,
      plot.y + plot.height,
      plot.y,
      numericReverse,
      numericLogBase,
    );
  };
  if (spec.kind === 'column' || spec.kind === 'bar') {
    const seriesCount = Math.max(1, data.series.length);
    const geometry = barGeometry(slot, seriesCount, grouping, spec);
    for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
      const series = data.series[seriesIndex];
      if (!series) continue;
      const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
      for (let pointIndex = 0; pointIndex < data.categories.length; pointIndex += 1) {
        const segment = segments[seriesIndex]?.[pointIndex] ?? null;
        if (!segment) continue;
        if (
          numericLogBase !== undefined &&
          (!(segment.value > 0) || !Number.isFinite(segment.value))
        )
          continue;
        const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
        const pointId = categoryIdAt(view, pointIndex);
        const style = resolveChartStyle(
          spec,
          palette,
          series.id,
          pointId,
          originalSeriesIndex,
          originalPointIndex,
        );
        const valueStart = numericScale(
          numericLogBase !== undefined && !(segment.start > 0) ? yDomain.min : segment.start,
        );
        const valueEnd = numericScale(segment.end);
        const categoryPosition = categoryIndex(count, pointIndex, categoryReverse);
        if (isBar) {
          const x = Math.min(valueStart, valueEnd);
          elements.push(
            primitive(
              'rect',
              styleAttrs(
                {
                  x,
                  y:
                    plot.y +
                    categoryPosition * slot +
                    geometry.groupOffset +
                    seriesIndex * geometry.seriesStep,
                  width: Math.abs(valueEnd - valueStart),
                  height: geometry.size,
                },
                style,
              ),
              'bar',
              originalSeriesIndex,
              originalPointIndex,
              undefined,
              { seriesId: series.id, pointId },
            ),
          );
        } else {
          const x =
            plot.x +
            categoryPosition * slot +
            geometry.groupOffset +
            seriesIndex * geometry.seriesStep;
          elements.push(
            primitive(
              'rect',
              styleAttrs(
                {
                  x,
                  y: Math.min(valueStart, valueEnd),
                  width: geometry.size,
                  height: Math.abs(valueEnd - valueStart),
                },
                style,
              ),
              'bar',
              originalSeriesIndex,
              originalPointIndex,
              undefined,
              { seriesId: series.id, pointId },
            ),
          );
        }
      }
    }
    return;
  }

  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    const points = Array.from({ length: data.categories.length }, (_value, pointIndex) => {
      const segment = segments[seriesIndex]?.[pointIndex] ?? null;
      const plottableSegment =
        segment !== null &&
        (numericLogBase === undefined || (segment.value > 0 && Number.isFinite(segment.value)));
      return {
        pointIndex,
        x: xCategory(plot, data.categories.length, pointIndex, categoryReverse),
        y: plottableSegment ? numericScale(segment.end) : null,
        segment: plottableSegment ? segment : null,
      };
    });
    const runs: { index: number; value: number; base: number }[][] = [];
    let run: { index: number; value: number; base: number }[] = [];
    for (const point of points) {
      if (point.y === null || !point.segment) {
        const raw = data.series[seriesIndex]?.values[point.pointIndex] ?? null;
        const logBreak =
          numericLogBase !== undefined && raw !== null && (!Number.isFinite(raw) || raw <= 0);
        if (logBreak) {
          if (run.length > 0) runs.push(run);
          run = [];
          continue;
        }
        if (series.states?.[point.pointIndex] === 'connect') continue;
        if (run.length > 0) runs.push(run);
        run = [];
      } else {
        run.push({
          index: point.pointIndex,
          value: point.y,
          base: numericScale(
            numericLogBase !== undefined && !(point.segment.start > 0)
              ? yDomain.min
              : point.segment.start,
          ),
        });
      }
    }
    if (run.length > 0) runs.push(run);
    for (const lineRun of runs) {
      const mapped = lineRun
        .map(
          (point) =>
            `${xCategory(plot, data.categories.length, point.index, categoryReverse)},${point.value}`,
        )
        .join(' ');
      if (spec.kind === 'area' && lineRun.length > 0) {
        const lower = lineRun
          .slice()
          .reverse()
          .map(
            (point) =>
              `${xCategory(plot, data.categories.length, point.index, categoryReverse)},${point.base}`,
          )
          .join(' ');
        const style = resolveChartStyle(
          spec,
          palette,
          series.id,
          undefined,
          originalSeriesIndex,
          undefined,
        );
        elements.push(
          primitive(
            'polygon',
            styleAttrs(
              {
                points: `${mapped} ${lower}`,
              },
              style,
              0.2,
            ),
            'area',
            originalSeriesIndex,
            undefined,
            undefined,
            { seriesId: series.id },
          ),
        );
      }
      if (lineRun.length >= 2) {
        const style = resolveChartStyle(
          spec,
          palette,
          series.id,
          undefined,
          originalSeriesIndex,
          undefined,
        );
        const opacity = lineOpacity(style.line);
        const width = lineWidth(style.line, 2) ?? 2;
        const dash = dashArray(style.line?.dash);
        const smooth = spec.kind === 'line' && smoothed(spec, series.id);
        const coordinates = lineRun.map(
          (point) =>
            [
              xCategory(plot, data.categories.length, point.index, categoryReverse),
              point.value,
            ] as const,
        );
        elements.push(
          primitive(
            smooth ? 'path' : 'polyline',
            {
              ...(smooth ? { d: smoothPath(coordinates) } : { points: mapped }),
              fill: 'none',
              stroke: style.line?.color ?? style.color,
              'stroke-width': width,
              ...(opacity === undefined ? {} : { 'stroke-opacity': opacity }),
              ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
            },
            'line',
            originalSeriesIndex,
            undefined,
            undefined,
            { seriesId: series.id },
          ),
        );
      }
    }
    for (const point of points) {
      if (point.y === null) continue;
      const originalPointIndex = view.originalCategoryIndices[point.pointIndex] ?? point.pointIndex;
      const pointId = categoryIdAt(view, point.pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      if (!markerVisible(style.marker)) continue;
      const marker = markerAttrs(point.x, point.y, style);
      elements.push(
        primitive(
          marker.tag,
          marker.attrs,
          'point',
          originalSeriesIndex,
          originalPointIndex,
          undefined,
          { seriesId: series.id, pointId },
        ),
      );
    }
  }
};

export const addScatterMarks = (
  elements: ChartPrimitive[],
  view: XYChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  xDomain: Domain,
  yDomain: Domain,
  palette: readonly string[],
  xLogBase?: number,
  yLogBase?: number,
): void => {
  const data = view.data;
  const xScale = (value: number): number =>
    numericProjection(
      value,
      xDomain,
      plot.x,
      plot.x + plot.width,
      reversedAxis(spec.xAxis),
      xLogBase,
    );
  const yScale = (value: number): number =>
    numericProjection(
      value,
      yDomain,
      plot.y + plot.height,
      plot.y,
      reversedAxis(spec.yAxis),
      yLogBase,
    );
  const inDomain = (value: number, domain: Domain): boolean =>
    Number.isFinite(value) && value >= domain.min && value <= domain.max;
  const logEligible = (value: number, logBase: number | undefined): boolean =>
    logBase === undefined || value > 0;
  const lineAxis = (domain: Domain, logBase: number | undefined) => {
    if (logBase === undefined) return { domain, transform: (value: number): number => value };
    const domainMin = Math.log(domain.min);
    const domainMax = Math.log(domain.max);
    if (!(domainMax > domainMin) || !Number.isFinite(domainMin) || !Number.isFinite(domainMax)) {
      return { domain, transform: (value: number): number => value };
    }
    return {
      domain: { min: domainMin, max: domainMax },
      transform: (value: number): number => Math.log(value),
    };
  };
  const lineX = lineAxis(xDomain, xLogBase);
  const lineY = lineAxis(yDomain, yLogBase);
  const lineDomains = { x: lineX.domain, y: lineY.domain };
  const scatterStyle = spec.scatterStyle ?? 'marker';
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    const lineSamples: ScatterLineSample[] = [];
    const markerPoints: Array<{
      readonly originalPointIndex: number;
      readonly pointId: string;
      readonly x: number;
      readonly y: number;
    }> = [];
    for (let pointIndex = 0; pointIndex < series.points.length; pointIndex += 1) {
      const point = series.points[pointIndex];
      if (!point) continue;
      const originalPointIndex = view.originalPointIndices[seriesIndex]?.[pointIndex] ?? pointIndex;
      const pointId = pointIdAt(view, seriesIndex, pointIndex);
      if (point.state === 'connect') continue;
      const nonFinite =
        (point.x !== null && !Number.isFinite(point.x)) ||
        (point.y !== null && !Number.isFinite(point.y));
      const nonPositiveLog =
        (point.x !== null && Number.isFinite(point.x) && !logEligible(point.x, xLogBase)) ||
        (point.y !== null && Number.isFinite(point.y) && !logEligible(point.y, yLogBase));
      if (nonFinite || nonPositiveLog) {
        // Non-finite values are a hard run break, even next to a connected gap.
        lineSamples.push({ kind: 'break', sourceIndex: originalPointIndex });
        continue;
      }
      if (point.x === null || point.y === null) {
        lineSamples.push({ kind: 'blank', sourceIndex: originalPointIndex });
        continue;
      }
      lineSamples.push({
        kind: 'point',
        x: lineX.transform(point.x),
        y: lineY.transform(point.y),
        sourceIndex: originalPointIndex,
      });
      if (
        inDomain(point.x, xDomain) &&
        inDomain(point.y, yDomain) &&
        logEligible(point.x, xLogBase) &&
        logEligible(point.y, yLogBase)
      ) {
        markerPoints.push({ originalPointIndex, pointId, x: point.x, y: point.y });
      }
    }

    if (scatterStyle !== 'marker') {
      const lineStyle = resolveChartStyle(
        spec,
        palette,
        series.id,
        undefined,
        originalSeriesIndex,
        undefined,
      );
      const width = lineWidth(lineStyle.line, 2) ?? 2;
      // A zero-width series line is an explicit request to suppress the
      // connector while leaving point markers independently controllable.
      if (width > 0) {
        const opacity = lineOpacity(lineStyle.line);
        const dash = dashArray(lineStyle.line?.dash);
        const stroke = {
          fill: 'none',
          stroke: lineStyle.line?.color ?? lineStyle.color,
          'stroke-width': width,
          ...(opacity === undefined ? {} : { 'stroke-opacity': opacity }),
          ...(dash === undefined ? {} : { 'stroke-dasharray': dash }),
        };
        const segments = scatterLineSegments(lineSamples, plot, false, lineDomains, {
          reverseX: reversedAxis(spec.xAxis),
          reverseY: reversedAxis(spec.yAxis),
        });
        if (smoothed(spec, series.id)) {
          // A smoothed line runs through each chain of joined segments.
          const runs: [number, number][][] = [];
          for (const segment of segments) {
            const run = runs.at(-1);
            const end = run?.at(-1);
            if (run && end && end[0] === segment.x1 && end[1] === segment.y1) {
              run.push([segment.x2, segment.y2]);
            } else {
              runs.push([
                [segment.x1, segment.y1],
                [segment.x2, segment.y2],
              ]);
            }
          }
          for (const run of runs) {
            elements.push(
              primitive(
                'path',
                { d: smoothPath(run), ...stroke },
                'line',
                originalSeriesIndex,
                undefined,
                undefined,
                {
                  seriesId: series.id,
                },
              ),
            );
          }
        } else {
          for (const segment of segments) {
            elements.push(
              primitive(
                'line',
                { x1: segment.x1, y1: segment.y1, x2: segment.x2, y2: segment.y2, ...stroke },
                'line',
                originalSeriesIndex,
                undefined,
                undefined,
                { seriesId: series.id },
              ),
            );
          }
        }
      }
    }

    for (const point of markerPoints) {
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        point.pointId,
        originalSeriesIndex,
        point.originalPointIndex,
      );
      if (!markerVisible(style.marker)) continue;
      const marker = markerAttrs(xScale(point.x), yScale(point.y), style);
      elements.push(
        primitive(
          marker.tag,
          marker.attrs,
          'point',
          originalSeriesIndex,
          point.originalPointIndex,
          undefined,
          { seriesId: series.id, pointId: point.pointId },
        ),
      );
    }
  }
};

export const addPieMarks = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  plot: ChartScene['plot'],
  palette: readonly string[],
  spec: ChartSpec,
): void => {
  const data = view.data;
  const first = data.series[0];
  if (!first) return;
  const magnitudes = normalizedMagnitudes(first.values);
  if (!magnitudes) return;
  const { values: normalized, total } = magnitudes;
  const { cx, cy, radius, firstSliceStart } = circularFrame(spec, plot);
  let start = firstSliceStart;
  for (let pointIndex = 0; pointIndex < normalized.length; pointIndex += 1) {
    const value = normalized[pointIndex] ?? 0;
    if (!(value > 0)) continue;
    const angle = (value / total) * Math.PI * 2;
    const end = start + angle;
    const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
    const pointId = categoryIdAt(view, pointIndex);
    const style = resolveChartStyle(
      spec,
      palette,
      first.id,
      pointId,
      view.originalSeriesIndices[0] ?? 0,
      originalPointIndex,
    );
    if (angle >= Math.PI * 2 - 1e-9) {
      elements.push(
        primitive(
          'circle',
          styleAttrs({ cx, cy, r: radius }, style),
          'slice',
          view.originalSeriesIndices[0] ?? 0,
          originalPointIndex,
          undefined,
          { seriesId: first.id, pointId },
        ),
      );
    } else {
      const startX = cx + Math.cos(start) * radius;
      const startY = cy + Math.sin(start) * radius;
      const endX = cx + Math.cos(end) * radius;
      const endY = cy + Math.sin(end) * radius;
      elements.push(
        primitive(
          'path',
          {
            d: `M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 ${angle > Math.PI ? 1 : 0} 1 ${endX} ${endY} Z`,
            ...styleAttrs(
              {
                stroke: 'Canvas',
                'stroke-width': 1,
              },
              style,
            ),
          },
          'slice',
          view.originalSeriesIndices[0] ?? 0,
          originalPointIndex,
          undefined,
          { seriesId: first.id, pointId },
        ),
      );
    }
    start = end;
  }
};

export const addDoughnutMarks = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  plot: ChartScene['plot'],
  palette: readonly string[],
  spec: ChartSpec,
): void => {
  const data = view.data;
  if (data.series.length === 0) return;
  const { cx, cy, radius: outerRadius, holeRadius, firstSliceStart } = doughnutFrame(spec, plot);
  const ringWidth = (outerRadius - holeRadius) / data.series.length;
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const normalized = normalizedMagnitudes(series.values);
    if (!normalized) continue;
    const innerRadius = holeRadius + ringWidth * seriesIndex;
    const outerRingRadius = innerRadius + ringWidth;
    let start = firstSliceStart;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < normalized.values.length; pointIndex += 1) {
      const value = normalized.values[pointIndex] ?? 0;
      if (!(value > 0) || !Number.isFinite(value)) continue;
      const angle = (value / normalized.total) * Math.PI * 2;
      if (!(angle > 0) || !Number.isFinite(angle)) continue;
      const end = start + angle;
      const originalPointIndex = view.originalCategoryIndices[pointIndex] ?? pointIndex;
      const pointId = categoryIdAt(view, pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      elements.push(
        primitive(
          'path',
          styleAttrs(
            {
              d: annularPath(cx, cy, innerRadius, outerRingRadius, start, angle),
              'fill-rule': 'evenodd',
              stroke: 'Canvas',
              'stroke-width': 1,
            },
            style,
          ),
          'slice',
          originalSeriesIndex,
          originalPointIndex,
          undefined,
          { seriesId: series.id, pointId },
        ),
      );
      start = end;
    }
  }
};
