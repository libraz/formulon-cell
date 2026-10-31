import { formatChartValue } from './format.js';
import {
  type Domain,
  numericProjection,
  reversedAxis,
  xCategory,
  yCategory,
} from './plot-scale.js';
import { formatTick, majorTicks, powerTicks } from './scales.js';
import { primitive, textWidth } from './scene-primitives.js';
import { isCircularKind } from './series-geometry.js';
import type {
  ChartAxisSpec,
  ChartKind,
  ChartPrimitive,
  ChartScene,
  ChartSpec,
  ChartValueFormatter,
} from './types.js';

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

/** The text an axis title shows: its own text, or `axisTitle` for Excel's automatic title. */
export const axisTitleText = (
  title: ChartAxisSpec['title'],
  axisTitle = 'Axis Title',
): string | undefined => {
  if (title === undefined) return undefined;
  if (typeof title === 'string') return title.length > 0 ? title : undefined;
  return axisTitle;
};

type ResolvedAxisLabelPosition = 'nextTo' | 'low' | 'high' | 'none';

/** Invalid runtime values retain the legacy low-edge label placement. */
export const resolvedAxisLabelPosition = (
  axis: ChartAxisSpec | undefined,
): ResolvedAxisLabelPosition => {
  const position = axis?.labelPosition;
  return position === 'nextTo' || position === 'low' || position === 'high' || position === 'none'
    ? position
    : 'low';
};

/**
 * Resolve a numeric crossing through the same clamped projection used by the
 * marks. Invalid runtime values (including non-finite numbers and unknown
 * strings) retain the historic zero crossing.
 */
const numericAxisCrossing = (
  axis: ChartAxisSpec | undefined,
  domain: Domain,
  scale: (value: number) => number,
  logBase?: number,
): number => {
  const crossing = axis?.crossing;
  if (crossing === 'min') return scale(domain.min);
  if (crossing === 'max') return scale(domain.max);
  if (
    typeof crossing === 'number' &&
    Number.isFinite(crossing) &&
    (logBase === undefined || crossing > 0)
  )
    return scale(crossing);
  if (logBase !== undefined) return scale(clamp(1, domain.min, domain.max));
  return scale(0);
};

/**
 * Resolve a category-axis crossing between bands. Explicit positions are in
 * logical category order and are reflected by reverseOrder. Automatic mode
 * deliberately keeps the legacy physical edge, independent of reversal.
 */
const categoryAxisCrossing = (
  axis: ChartAxisSpec | undefined,
  count: number,
  start: number,
  end: number,
  reverse: boolean,
  automatic: number,
): number => {
  if (count <= 0) return automatic;
  const crossing = axis?.crossing;
  if (crossing === undefined || crossing === 'auto') return automatic;
  let edge: number;
  if (crossing === 'min') edge = 0;
  else if (crossing === 'max') edge = count;
  else if (typeof crossing === 'number' && Number.isFinite(crossing)) {
    const categoryNumber = clamp(Math.round(crossing), 1, count);
    edge = categoryNumber - 1;
  } else {
    return automatic;
  }
  const slot = (end - start) / count;
  const reflectedEdge = reverse ? count - edge : edge;
  return start + reflectedEdge * slot;
};

export const axisIsVisible = (axis: ChartAxisSpec | undefined): boolean => axis?.visible !== false;

const axisHasGridlines = (axis: ChartAxisSpec | undefined): boolean =>
  axis?.majorGridlines !== false;

export const axisTicks = (
  domain: Domain,
  axis: ChartAxisSpec | undefined,
  logBase?: number,
): readonly number[] =>
  logBase === undefined
    ? majorTicks(domain.min, domain.max, axis?.majorUnit ?? domain.majorUnit)
    : powerTicks(domain.min, domain.max, logBase);

export const axisTickLabel = (
  value: number,
  percent: boolean,
  axis: ChartAxisSpec | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): string => {
  const formatCode = axis?.numberFormat;
  if (formatCode !== undefined && formatCode.trim().length > 0) {
    return formatChartValue(value, formatCode, locale, formatter) ?? formatTick(value, percent);
  }
  return formatTick(value, percent);
};

const AXIS_LABEL_MAX_WIDTH = 160;

const AXIS_LABEL_MAX_CHARS = 256;

// Rotated SVG text can extend a few pixels to the left of its x coordinate
// under the default system font. Keep the baseline far enough inside the
// canvas for narrow charts while leaving plot margins unchanged.
const AXIS_TITLE_MIN_X = 12;

export const numericAxisLabelMaxWidth = (canvasWidth: number): number =>
  Math.min(AXIS_LABEL_MAX_WIDTH, Math.max(12, (canvasWidth * 0.75 - 40) / 1.5));

/**
 * Keep host-supplied format strings from expanding the scene without bound.
 * The original value remains available to assistive consumers when the
 * visible SVG text needs an ellipsis.
 */
export const boundedAxisLabel = (label: string, maxWidth = AXIS_LABEL_MAX_WIDTH): string => {
  if (label.length <= AXIS_LABEL_MAX_CHARS && textWidth(label) <= maxWidth) return label;
  let result = '';
  let characterCount = 0;
  for (const character of label) {
    if (characterCount >= AXIS_LABEL_MAX_CHARS) break;
    const candidate = `${result}${character}…`;
    if (textWidth(candidate) > maxWidth) break;
    result += character;
    characterCount += 1;
  }
  return result.length > 0 ? `${result}…` : '…';
};

const numericAxisLabel = (
  attrs: Record<string, string | number>,
  value: number,
  percent: boolean,
  axis: ChartAxisSpec | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  maxLabelWidth: number,
): ChartPrimitive => {
  const fullText = axisTickLabel(value, percent, axis, locale, formatter);
  const text = boundedAxisLabel(fullText, maxLabelWidth);
  return primitive(
    'text',
    {
      ...attrs,
      ...(text === fullText ? {} : { 'aria-label': fullText }),
    },
    'axis-label',
    undefined,
    undefined,
    text,
  );
};

const categoryAxisLabel = (
  attrs: Record<string, string | number>,
  label: string,
  position: ResolvedAxisLabelPosition,
  maxLabelWidth: number,
  categoryOriginalIndex: number,
  categoryId: string,
): ChartPrimitive => {
  const text = position === 'high' ? boundedAxisLabel(label, maxLabelWidth) : label;
  return primitive(
    'text',
    {
      ...attrs,
      ...(text === label ? {} : { 'aria-label': label }),
    },
    'category-label',
    undefined,
    categoryOriginalIndex,
    text,
    { pointId: categoryId },
  );
};

const numericAxisLabelVisibility = (
  ticks: readonly number[],
  positions: readonly number[],
  percent: boolean,
  axis: ChartAxisSpec | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  maxLabelWidth: number,
  orientation: 'horizontal' | 'vertical' = 'horizontal',
): readonly boolean[] => {
  const visible = Array.from({ length: ticks.length }, () => false);
  if (ticks.length === 0) return visible;
  if (ticks.length === 1) {
    visible[0] = true;
    return visible;
  }
  const sizes = ticks.map((tick) =>
    orientation === 'horizontal'
      ? textWidth(
          boundedAxisLabel(axisTickLabel(tick, percent, axis, locale, formatter), maxLabelWidth),
        )
      : 10,
  );
  const order = positions
    .map((_, index) => index)
    .sort((left, right) => (positions[left] ?? 0) - (positions[right] ?? 0));
  const gap = orientation === 'vertical' ? 3 : 4;
  const canSeparate = (leftIndex: number, rightIndex: number): boolean => {
    const leftPosition = positions[leftIndex] ?? 0;
    const rightPosition = positions[rightIndex] ?? 0;
    const leftWidth = sizes[leftIndex] ?? 0;
    const rightWidth = sizes[rightIndex] ?? 0;
    return rightPosition - rightWidth / 2 >= leftPosition + leftWidth / 2 + gap;
  };
  const first = order[0] ?? 0;
  const last = order[order.length - 1] ?? first;
  const kept = [first];
  visible[first] = true;
  for (let orderIndex = 1; orderIndex < order.length - 1; orderIndex += 1) {
    const index = order[orderIndex] ?? first;
    const previous = kept[kept.length - 1] ?? 0;
    if (canSeparate(previous, index)) {
      kept.push(index);
      visible[index] = true;
    }
  }
  while (kept.length > 1 && !canSeparate(kept[kept.length - 1] ?? first, last)) {
    const removed = kept.pop();
    if (removed !== undefined) visible[removed] = false;
  }
  if (canSeparate(kept[kept.length - 1] ?? first, last)) {
    visible[last] = true;
  }
  return visible;
};

export const addAxis = (
  elements: ChartPrimitive[],
  plot: ChartScene['plot'],
  kind: ChartKind,
  xDomain: Domain | null,
  yDomain: Domain | null,
  xLabels: readonly string[],
  categoryOriginalIndices: readonly number[],
  categoryIds: readonly string[],
  spec: ChartSpec,
  percent: boolean,
  canvasHeight: number,
  maxLabelWidth: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  yAxisLabelMaxWidth = maxLabelWidth,
  xLogBase?: number,
  yLogBase?: number,
  axisTitle?: string,
): void => {
  const horizontalCategory = kind === 'bar';
  if (isCircularKind(kind)) return;

  const xAxis = spec.xAxis;
  const yAxis = spec.yAxis;
  const xVisible = axisIsVisible(xAxis);
  const yVisible = axisIsVisible(yAxis);
  const xGridlines = axisHasGridlines(xAxis);
  const yGridlines = axisHasGridlines(yAxis);
  const xLabelPosition = resolvedAxisLabelPosition(xAxis);
  const yLabelPosition = resolvedAxisLabelPosition(yAxis);
  const xScale = xDomain
    ? (value: number): number =>
        numericProjection(
          value,
          xDomain,
          plot.x,
          plot.x + plot.width,
          reversedAxis(xAxis),
          xLogBase,
        )
    : null;
  const yScale = yDomain
    ? (value: number): number =>
        numericProjection(
          value,
          yDomain,
          plot.y + plot.height,
          plot.y,
          reversedAxis(yAxis),
          yLogBase,
        )
    : null;
  const categoryReverse = horizontalCategory ? reversedAxis(yAxis) : reversedAxis(xAxis);
  const categoryLabelMaxWidth = (axis: 'x' | 'y'): number => {
    if (axis === 'y') return yAxisLabelMaxWidth;
    if (xLabelPosition !== 'high') return maxLabelWidth;
    const categoryCount = Math.max(1, xLabels.length);
    return Math.min(maxLabelWidth, Math.max(12, plot.width / categoryCount - 4));
  };

  if (horizontalCategory && xDomain && xScale) {
    const zeroX = numericAxisCrossing(xAxis, xDomain, xScale, xLogBase);
    const categoryY = categoryAxisCrossing(
      yAxis,
      xLabels.length,
      plot.y,
      plot.y + plot.height,
      categoryReverse,
      plot.y + plot.height,
    );
    if (xVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: plot.x,
            x2: plot.x + plot.width,
            y1: categoryY,
            y2: categoryY,
            stroke: 'currentColor',
            'data-chart-axis': 'x',
          },
          'axis',
        ),
      );
    }
    const xTicks = axisTicks(xDomain, xAxis, xLogBase);
    const xLabelVisibility = xVisible
      ? numericAxisLabelVisibility(
          xTicks,
          xTicks.map((tick) => xScale(tick)),
          percent,
          xAxis,
          locale,
          formatter,
          maxLabelWidth,
        )
      : [];
    for (let index = 0; index < xTicks.length; index += 1) {
      const tick = xTicks[index] ?? 0;
      const x = xScale(tick);
      if (xGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: x,
              x2: x,
              y1: plot.y,
              y2: plot.y + plot.height,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'x',
            },
            'axis-tick',
          ),
        );
      }
      if (xVisible && xLabelPosition !== 'none' && xLabelVisibility[index]) {
        const labelY =
          xLabelPosition === 'high'
            ? plot.y - 4
            : xLabelPosition === 'nextTo'
              ? categoryY + 12
              : plot.y + plot.height + 12;
        elements.push(
          numericAxisLabel(
            { x, y: labelY, 'text-anchor': 'middle', 'data-chart-axis': 'x' },
            tick,
            percent,
            xAxis,
            locale,
            formatter,
            maxLabelWidth,
          ),
        );
      }
    }
    if (yVisible) {
      if (yLabelPosition !== 'none') {
        for (let index = 0; index < xLabels.length; index += 1) {
          const labelX =
            yLabelPosition === 'high'
              ? plot.x + plot.width + 4
              : yLabelPosition === 'nextTo'
                ? zeroX - 4
                : plot.x - 4;
          elements.push(
            categoryAxisLabel(
              {
                x: labelX,
                y: yCategory(plot, xLabels.length, index, categoryReverse) + 3,
                'text-anchor': yLabelPosition === 'high' ? 'start' : 'end',
                'data-chart-axis': 'y',
              },
              xLabels[index] ?? '',
              yLabelPosition,
              categoryLabelMaxWidth('y'),
              categoryOriginalIndices[index] ?? index,
              categoryIds[index] ?? `category-${categoryOriginalIndices[index] ?? index}`,
            ),
          );
        }
      }
      elements.push(
        primitive(
          'line',
          {
            x1: zeroX,
            x2: zeroX,
            y1: plot.y,
            y2: plot.y + plot.height,
            stroke: 'currentColor',
            'data-chart-axis': 'y',
          },
          'axis',
        ),
      );
    }
  } else if (!horizontalCategory && yDomain && yScale && !xDomain) {
    const zeroY = numericAxisCrossing(yAxis, yDomain, yScale, yLogBase);
    const categoryX = categoryAxisCrossing(
      xAxis,
      xLabels.length,
      plot.x,
      plot.x + plot.width,
      categoryReverse,
      plot.x,
    );
    if (xVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: plot.x,
            x2: plot.x + plot.width,
            y1: zeroY,
            y2: zeroY,
            stroke: 'currentColor',
            'data-chart-axis': 'x',
          },
          'axis',
        ),
      );
    }
    const yTicks = axisTicks(yDomain, yAxis, yLogBase);
    const yLabelVisibility = yVisible
      ? numericAxisLabelVisibility(
          yTicks,
          yTicks.map((tick) => yScale(tick)),
          percent,
          yAxis,
          locale,
          formatter,
          yAxisLabelMaxWidth,
          'vertical',
        )
      : [];
    for (let index = 0; index < yTicks.length; index += 1) {
      const tick = yTicks[index] ?? 0;
      const y = yScale(tick);
      if (yGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: plot.x,
              x2: plot.x + plot.width,
              y1: y,
              y2: y,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'y',
            },
            'axis-tick',
          ),
        );
      }
      if (yVisible && yLabelPosition !== 'none' && yLabelVisibility[index]) {
        const labelX =
          yLabelPosition === 'high'
            ? plot.x + plot.width + 4
            : yLabelPosition === 'nextTo'
              ? categoryX - 2
              : plot.x - 2;
        elements.push(
          numericAxisLabel(
            {
              x: labelX,
              y: y + 3,
              'text-anchor': yLabelPosition === 'high' ? 'start' : 'end',
              'data-chart-axis': 'y',
            },
            tick,
            percent,
            yAxis,
            locale,
            formatter,
            yAxisLabelMaxWidth,
          ),
        );
      }
    }
    if (xVisible) {
      for (let index = 0; index < xLabels.length; index += 1) {
        if (xLabelPosition === 'none') continue;
        const labelY =
          xLabelPosition === 'high'
            ? plot.y - 4
            : xLabelPosition === 'nextTo'
              ? zeroY + 12
              : plot.y + plot.height + 12;
        elements.push(
          categoryAxisLabel(
            {
              x: xCategory(plot, xLabels.length, index, categoryReverse),
              y: labelY,
              'text-anchor': 'middle',
              'data-chart-axis': 'x',
            },
            xLabels[index] ?? '',
            xLabelPosition,
            categoryLabelMaxWidth('x'),
            categoryOriginalIndices[index] ?? index,
            categoryIds[index] ?? `category-${categoryOriginalIndices[index] ?? index}`,
          ),
        );
      }
    }
    if (yVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: categoryX,
            x2: categoryX,
            y1: plot.y,
            y2: plot.y + plot.height,
            stroke: 'currentColor',
            'data-chart-axis': 'y',
          },
          'axis',
        ),
      );
    }
  } else if (xDomain && xScale && yDomain && yScale) {
    const zeroY = numericAxisCrossing(yAxis, yDomain, yScale, yLogBase);
    const zeroX = numericAxisCrossing(xAxis, xDomain, xScale, xLogBase);
    if (xVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: plot.x,
            x2: plot.x + plot.width,
            y1: zeroY,
            y2: zeroY,
            stroke: 'currentColor',
            'data-chart-axis': 'x',
          },
          'axis',
        ),
      );
    }
    if (yVisible) {
      elements.push(
        primitive(
          'line',
          {
            x1: zeroX,
            x2: zeroX,
            y1: plot.y,
            y2: plot.y + plot.height,
            stroke: 'currentColor',
            'data-chart-axis': 'y',
          },
          'axis',
        ),
      );
    }
    const yTicks = axisTicks(yDomain, yAxis, yLogBase);
    const yLabelVisibility = yVisible
      ? numericAxisLabelVisibility(
          yTicks,
          yTicks.map((tick) => yScale(tick)),
          percent,
          yAxis,
          locale,
          formatter,
          yAxisLabelMaxWidth,
          'vertical',
        )
      : [];
    for (let index = 0; index < yTicks.length; index += 1) {
      const tick = yTicks[index] ?? 0;
      const y = yScale(tick);
      if (yGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: plot.x,
              x2: plot.x + plot.width,
              y1: y,
              y2: y,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'y',
            },
            'axis-tick',
          ),
        );
      }
      if (yVisible && yLabelPosition !== 'none' && yLabelVisibility[index]) {
        const labelX =
          yLabelPosition === 'high'
            ? plot.x + plot.width + 4
            : yLabelPosition === 'nextTo'
              ? zeroX - 2
              : plot.x - 2;
        elements.push(
          numericAxisLabel(
            {
              x: labelX,
              y: y + 3,
              'text-anchor': yLabelPosition === 'high' ? 'start' : 'end',
              'data-chart-axis': 'y',
            },
            tick,
            percent,
            yAxis,
            locale,
            formatter,
            yAxisLabelMaxWidth,
          ),
        );
      }
    }
    const xTicks = axisTicks(xDomain, xAxis, xLogBase);
    const xLabelVisibility = xVisible
      ? numericAxisLabelVisibility(
          xTicks,
          xTicks.map((tick) => xScale(tick)),
          false,
          xAxis,
          locale,
          formatter,
          maxLabelWidth,
        )
      : [];
    for (let index = 0; index < xTicks.length; index += 1) {
      const tick = xTicks[index] ?? 0;
      const x = xScale(tick);
      if (xGridlines) {
        elements.push(
          primitive(
            'line',
            {
              x1: x,
              x2: x,
              y1: plot.y,
              y2: plot.y + plot.height,
              stroke: 'currentColor',
              'stroke-opacity': 0.16,
              'data-chart-axis': 'x',
            },
            'axis-tick',
          ),
        );
      }
      if (xVisible && xLabelPosition !== 'none' && xLabelVisibility[index]) {
        const labelY =
          xLabelPosition === 'high'
            ? plot.y - 4
            : xLabelPosition === 'nextTo'
              ? zeroY + 12
              : plot.y + plot.height + 12;
        elements.push(
          numericAxisLabel(
            { x, y: labelY, 'text-anchor': 'middle', 'data-chart-axis': 'x' },
            tick,
            false,
            xAxis,
            locale,
            formatter,
            maxLabelWidth,
          ),
        );
      }
    }
  }

  const xTitle = xVisible ? axisTitleText(xAxis?.title, axisTitle) : undefined;
  const yTitle = yVisible ? axisTitleText(yAxis?.title, axisTitle) : undefined;
  if (xTitle) {
    elements.push(
      primitive(
        'text',
        {
          x: plot.x + plot.width / 2,
          y: Math.min(canvasHeight - 2, plot.y + plot.height + 25),
          'text-anchor': 'middle',
          'data-chart-axis': 'x',
        },
        'axis-title',
        undefined,
        undefined,
        xTitle,
      ),
    );
  }
  if (yTitle) {
    const yLabelWidth =
      yLabelPosition === 'high'
        ? 12
        : horizontalCategory
          ? xLabels.reduce((current, label) => Math.max(current, textWidth(label)), 12)
          : yDomain
            ? axisTicks(yDomain, yAxis, yLogBase).reduce(
                (current, tick) =>
                  Math.max(
                    current,
                    textWidth(
                      boundedAxisLabel(
                        axisTickLabel(tick, percent, yAxis, locale, formatter),
                        yAxisLabelMaxWidth,
                      ),
                    ),
                  ),
                12,
              )
            : 12;
    const x = Math.max(AXIS_TITLE_MIN_X, plot.x - Math.max(28, yLabelWidth + 16));
    const y = plot.y + plot.height / 2;
    elements.push(
      primitive(
        'text',
        {
          x,
          y,
          transform: `rotate(-90 ${x} ${y})`,
          'text-anchor': 'middle',
          'data-chart-axis': 'y',
        },
        'axis-title',
        undefined,
        undefined,
        yTitle,
      ),
    );
  }
};
