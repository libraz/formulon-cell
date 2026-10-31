import type { CategoryChartView, PreparedChartData, XYChartView } from './filters.js';
import { formatChartValue } from './format.js';
import { resolveChartLabelPosition } from './label-options.js';
import {
  categoryIndex,
  type Domain,
  numericProjection,
  reversedAxis,
  xCategory,
} from './plot-scale.js';
import {
  categoryIdAt,
  pointIdAt,
  primitive,
  textWidth,
  truncateLabel,
} from './scene-primitives.js';
import {
  barGeometry,
  boundedOption,
  circularPoint,
  normalizedMagnitudes,
  segmentsFor,
} from './series-geometry.js';
import { fillColor, keyColor, resolveChartStyle } from './styles.js';
import type { ChartPrimitive, ChartScene, ChartSpec, ChartValueFormatter } from './types.js';

const finite = (value: number, fallback: number): number =>
  Number.isFinite(value) ? value : fallback;

const clamp = (value: number, min: number, max: number): number =>
  Math.min(max, Math.max(min, value));

/** Series-level labels override the chart-wide labels; hidden series labels stay hidden. */
const labelSpecFor = (
  spec: ChartSpec,
  series: { readonly id: string; readonly formatCode?: string },
): ChartSpec['dataLabels'] => {
  const labels =
    spec.seriesStyles?.find((style) => style.seriesId === series.id)?.dataLabels ?? spec.dataLabels;
  // The percentage part derives its own code from `numberFormat`, so only value labels inherit.
  if (!labels || labels.showPercentage || !series.formatCode) return labels;
  if (labels.numberFormat !== undefined && labels.numberFormat.trim().length > 0) return labels;
  return { ...labels, numberFormat: series.formatCode };
};

const hasDataLabels = (spec: ChartSpec): boolean =>
  [spec.dataLabels, ...(spec.seriesStyles?.map((style) => style.dataLabels) ?? [])].some(
    (labels) => labels !== undefined && labels.visible !== false,
  );

interface LabelAnchor {
  readonly x: number;
  readonly y: number;
  readonly textAnchor?: 'start' | 'middle' | 'end';
}

const labelFontSize = (value: number | undefined): number =>
  value !== undefined && Number.isFinite(value) ? clamp(value, 6, 72) : 10;

const labelParts = (
  spec: NonNullable<ChartSpec['dataLabels']>,
  seriesName: string,
  categoryName: string | undefined,
  value: number | undefined,
  percentage: number | undefined,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): string => {
  const hasConfiguredLabelFlag = [
    spec.showSeriesName,
    spec.showCategoryName,
    spec.showValue,
    spec.showPercentage,
    spec.showLegendKey,
  ].some((flag) => flag !== undefined);
  const showValue = spec.showValue ?? !hasConfiguredLabelFlag;
  const parts: string[] = [];
  if (spec.showSeriesName) parts.push(seriesName);
  if (spec.showCategoryName && categoryName !== undefined) parts.push(categoryName);
  if (showValue && value !== undefined) {
    const formatted = formatChartValue(value, spec.numberFormat, locale, formatter);
    if (formatted !== null) parts.push(formatted);
  }
  if (spec.showPercentage && percentage !== undefined) {
    const percentageCode = spec.numberFormat?.includes('%')
      ? spec.numberFormat
      : spec.numberFormat === undefined
        ? '0%'
        : `${spec.numberFormat}%`;
    const formatted = formatChartValue(percentage, percentageCode, locale, formatter);
    if (formatted !== null) parts.push(formatted);
  }
  return parts.join(spec.separator ?? ', ');
};

const addDataLabel = (
  elements: ChartPrimitive[],
  fullText: string,
  anchor: LabelAnchor,
  spec: NonNullable<ChartSpec['dataLabels']>,
  width: number,
  height: number,
  seriesIndex: number,
  pointIndex: number,
  seriesId: string,
  pointId: string | undefined,
  color: string,
): void => {
  if (fullText.length === 0 && !spec.showLegendKey) return;
  const horizontalPadding = 1;
  const verticalPadding = 1;
  const glyphAscentRatio = 0.99;
  const glyphDescentRatio = 0.24;
  const maxFontSize = Math.min(
    72,
    Math.max(6, (height - verticalPadding * 2) / (glyphAscentRatio + glyphDescentRatio)),
  );
  const fontSize = clamp(labelFontSize(spec.fontSize), 6, maxFontSize);
  const ascent = fontSize * glyphAscentRatio;
  const descent = fontSize * glyphDescentRatio;
  const minBaseline = verticalPadding + ascent;
  const maxBaseline = Math.max(minBaseline, height - verticalPadding - descent);
  const y = clamp(finite(anchor.y, minBaseline), minBaseline, maxBaseline);
  const textAnchor = anchor.textAnchor ?? 'middle';
  const keySize = spec.showLegendKey ? clamp(fontSize * 0.45, 4, 12) : 0;
  const keyGap = keySize > 0 ? Math.max(3, fontSize * 0.12) : 0;
  const keyReservation = keySize + keyGap;
  const minimumTextWidth = fullText.length > 0 ? textWidth('…', fontSize) : 0;
  const rawX = finite(anchor.x, width / 2);
  // Like Excel, a label that would run off the chart is moved back inside it
  // whole; only a label wider than the room left is shortened.
  const fitWidth = Math.min(
    textWidth(fullText, fontSize),
    Math.max(minimumTextWidth, width - horizontalPadding * 2 - keyReservation),
  );
  const keptWidth = fullText.length > 0 ? Math.max(minimumTextWidth, fitWidth) : 0;
  const minX =
    textAnchor === 'start'
      ? horizontalPadding + keyReservation
      : textAnchor === 'end'
        ? horizontalPadding + keptWidth
        : horizontalPadding + keyReservation + keptWidth / 2;
  const maxX =
    textAnchor === 'start'
      ? width - horizontalPadding - keptWidth
      : textAnchor === 'end'
        ? width - horizontalPadding - keyReservation
        : width - horizontalPadding - keptWidth / 2;
  const x = clamp(rawX, Math.min(minX, maxX), Math.max(minX, maxX));
  const sideRoom =
    textAnchor === 'start'
      ? width - horizontalPadding - x
      : textAnchor === 'end'
        ? x - horizontalPadding
        : Math.min(x - horizontalPadding - keyReservation, width - horizontalPadding - x) * 2;
  const displayText = truncateLabel(fullText, Math.max(minimumTextWidth, sideRoom), fontSize);
  const attrs: Record<string, string | number> = {
    x,
    y,
    'text-anchor': textAnchor,
    'font-size': fontSize,
    fill: spec.color ?? 'currentColor',
    'aria-label': fullText,
    'white-space': 'pre',
  };
  if (spec.showLegendKey) {
    const displayWidth = textWidth(displayText, fontSize);
    const keyX =
      fullText.length === 0
        ? x - keySize / 2
        : textAnchor === 'end'
          ? x + keyGap
          : textAnchor === 'start'
            ? x - keyGap - keySize
            : x - displayWidth / 2 - keyGap - keySize;
    const keyY = clamp(
      y - keySize * 0.72,
      verticalPadding,
      Math.max(verticalPadding, height - verticalPadding - keySize),
    );
    elements.push(
      primitive(
        'rect',
        {
          x: clamp(keyX, horizontalPadding, Math.max(horizontalPadding, width - keySize)),
          y: keyY,
          width: keySize,
          height: keySize,
          fill: color,
          'aria-label': seriesId,
        },
        'data-label-key',
        seriesIndex,
        pointIndex,
        undefined,
        { seriesId, pointId },
      ),
    );
  }
  if (fullText.length > 0) {
    elements.push(
      primitive('text', attrs, 'data-label', seriesIndex, pointIndex, displayText, {
        seriesId,
        pointId,
      }),
    );
  }
};

interface LabelBounds {
  readonly left: number;
  readonly right: number;
  readonly top: number;
  readonly bottom: number;
}

interface DoughnutLabelRecord {
  readonly seriesIndex: number;
  readonly pointIndex: number;
  readonly originalSeriesIndex: number;
  readonly originalPointIndex: number;
  readonly seriesId: string;
  readonly pointId: string | undefined;
  readonly start: number;
  readonly angle: number;
  readonly midpointRadius: number;
  readonly text: string;
  readonly color: string;
  readonly labelSpec: NonNullable<ChartSpec['dataLabels']>;
}

interface DoughnutLabelCandidate {
  readonly anchor: LabelAnchor;
  readonly bounds: LabelBounds | null;
}

const labelBoundsFor = (elements: readonly ChartPrimitive[]): LabelBounds | null => {
  let bounds: LabelBounds | null = null;
  const include = (next: LabelBounds): void => {
    const padded = {
      left: next.left - 2,
      right: next.right + 2,
      top: next.top - 2,
      bottom: next.bottom + 2,
    };
    bounds =
      bounds === null
        ? padded
        : {
            left: Math.min(bounds.left, padded.left),
            right: Math.max(bounds.right, padded.right),
            top: Math.min(bounds.top, padded.top),
            bottom: Math.max(bounds.bottom, padded.bottom),
          };
  };
  for (const element of elements) {
    if (element.role === 'data-label') {
      const x = Number(element.attrs.x);
      const y = Number(element.attrs.y);
      const fontSize = finite(Number(element.attrs['font-size']), 10);
      if (!Number.isFinite(x) || !Number.isFinite(y) || !(fontSize > 0)) continue;
      const text = element.text ?? '';
      const lines = text.split('\n');
      const textWidthValue = lines.reduce(
        (current, line) => Math.max(current, textWidth(line, fontSize)),
        0,
      );
      const textAnchor = element.attrs['text-anchor'];
      const left =
        textAnchor === 'start'
          ? x
          : textAnchor === 'end'
            ? x - textWidthValue
            : x - textWidthValue / 2;
      include({
        left,
        right: left + textWidthValue,
        top: y - fontSize * 0.99,
        bottom: y + fontSize * 0.24 + Math.max(0, lines.length - 1) * fontSize * 1.1,
      });
    } else if (element.role === 'data-label-key') {
      const x = Number(element.attrs.x);
      const y = Number(element.attrs.y);
      const width = Number(element.attrs.width);
      const height = Number(element.attrs.height);
      if (
        Number.isFinite(x) &&
        Number.isFinite(y) &&
        Number.isFinite(width) &&
        Number.isFinite(height)
      ) {
        include({ left: x, right: x + width, top: y, bottom: y + height });
      }
    }
  }
  return bounds;
};

const labelBoundsOverlap = (left: LabelBounds, right: LabelBounds): boolean =>
  left.left < right.right &&
  right.left < left.right &&
  left.top < right.bottom &&
  right.top < left.bottom;

const addCategoryDataLabels = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  domain: Domain,
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  numericLogBase?: number,
): void => {
  const data = view.data;
  const grouping = spec.grouping ?? 'clustered';
  const segments = segmentsFor(data, grouping);
  const isBar = spec.kind === 'bar';
  const count = Math.max(1, data.categories.length);
  const slot = (isBar ? plot.height : plot.width) / count;
  const numericReverse = isBar ? reversedAxis(spec.xAxis) : reversedAxis(spec.yAxis);
  const categoryReverse = isBar ? reversedAxis(spec.yAxis) : reversedAxis(spec.xAxis);
  const valueScale = (value: number): number =>
    isBar
      ? numericProjection(
          value,
          domain,
          plot.x,
          plot.x + plot.width,
          numericReverse,
          numericLogBase,
        )
      : numericProjection(
          value,
          domain,
          plot.y + plot.height,
          plot.y,
          numericReverse,
          numericLogBase,
        );
  const geometry = barGeometry(slot, Math.max(1, data.series.length), grouping, spec);
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const labelSpec = labelSpecFor(spec, series);
    if (!labelSpec || labelSpec.visible === false) continue;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < data.categories.length; pointIndex += 1) {
      const raw = series.values[pointIndex] ?? null;
      if (raw === null || !Number.isFinite(raw) || (numericLogBase !== undefined && !(raw > 0)))
        continue;
      const segment = segments[seriesIndex]?.[pointIndex];
      if (!segment) continue;
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
      const categoryName = data.categories[pointIndex];
      const percentage = grouping === 'percent-stacked' ? Math.abs(segment.value) : undefined;
      const fullText = labelParts(
        labelSpec,
        series.name ?? String(originalSeriesIndex + 1),
        categoryName,
        raw,
        percentage,
        locale,
        formatter,
      );
      if (fullText.length === 0 && !labelSpec.showLegendKey) continue;
      const valueStart = valueScale(
        numericLogBase !== undefined && !(segment.start > 0) ? domain.min : segment.start,
      );
      const valueEnd = valueScale(segment.end);
      const valueDelta = valueEnd - valueStart;
      const rawPositive = raw >= 0;
      const screenPositive = isBar
        ? valueDelta !== 0
          ? valueDelta > 0
          : rawPositive !== numericReverse
        : valueDelta !== 0
          ? valueDelta < 0
          : rawPositive !== numericReverse;
      let anchor: LabelAnchor;
      const requested = resolveChartLabelPosition(spec, labelSpec.position);
      if (spec.kind === 'column' || spec.kind === 'bar') {
        const categoryPosition = categoryIndex(count, pointIndex, categoryReverse);
        const x =
          spec.kind === 'bar'
            ? Math.min(valueStart, valueEnd)
            : plot.x +
              categoryPosition * slot +
              geometry.groupOffset +
              seriesIndex * geometry.seriesStep +
              geometry.size / 2;
        const y =
          spec.kind === 'bar'
            ? plot.y +
              categoryPosition * slot +
              geometry.groupOffset +
              seriesIndex * geometry.seriesStep +
              geometry.size / 2
            : Math.min(valueStart, valueEnd);
        if (spec.kind === 'bar') {
          const endX = screenPositive
            ? Math.max(valueStart, valueEnd)
            : Math.min(valueStart, valueEnd);
          // For stacked bars, the segment's start is its visible base. Using
          // zero here puts every inside-base label on the axis instead.
          const baseX = valueStart;
          if (requested === 'center') {
            anchor = {
              x: (Math.min(valueStart, valueEnd) + Math.max(valueStart, valueEnd)) / 2,
              y,
              textAnchor: 'middle',
            };
          } else if (requested === 'inside-base') {
            anchor = {
              x: screenPositive ? baseX + 4 : baseX - 4,
              y,
              textAnchor: screenPositive ? 'start' : 'end',
            };
          } else if (requested === 'inside-end' || requested === 'best-fit') {
            anchor = {
              x: screenPositive ? endX - 4 : endX + 4,
              y,
              textAnchor: screenPositive ? 'end' : 'start',
            };
          } else {
            anchor = {
              x: screenPositive ? endX + 4 : endX - 4,
              y,
              textAnchor: screenPositive ? 'start' : 'end',
            };
          }
        } else if (requested === 'center') {
          anchor = { x, y: (Math.min(valueStart, valueEnd) + Math.max(valueStart, valueEnd)) / 2 };
        } else if (requested === 'inside-base') {
          // For stacked columns, the segment's start is its visible base.
          const baseY = valueStart;
          anchor = { x, y: screenPositive ? baseY - 4 : baseY + 12 };
        } else if (requested === 'inside-end' || requested === 'best-fit') {
          const endY = screenPositive
            ? Math.min(valueStart, valueEnd)
            : Math.max(valueStart, valueEnd);
          anchor = { x, y: screenPositive ? endY + 12 : endY - 4 };
        } else if (requested === 'left') {
          anchor = { x: x - 5, y, textAnchor: 'end' };
        } else if (requested === 'right') {
          anchor = { x: x + 5, y, textAnchor: 'start' };
        } else if (requested === 'below') {
          anchor = { x, y: Math.max(valueStart, valueEnd) + 12 };
        } else {
          const endY = screenPositive
            ? Math.min(valueStart, valueEnd)
            : Math.max(valueStart, valueEnd);
          anchor = { x, y: screenPositive ? endY - 4 : endY + 12 };
        }
      } else {
        anchor = {
          x: xCategory(plot, data.categories.length, pointIndex, categoryReverse),
          y: valueEnd - 5,
        };
        if (requested === 'below') anchor = { ...anchor, y: valueEnd + 12 };
        if (requested === 'left') anchor = { ...anchor, x: anchor.x - 5, textAnchor: 'end' };
        if (requested === 'right') anchor = { ...anchor, x: anchor.x + 5, textAnchor: 'start' };
        if (requested === 'center') anchor = { ...anchor, y: valueEnd };
      }
      addDataLabel(
        elements,
        fullText,
        anchor,
        labelSpec,
        width,
        height,
        originalSeriesIndex,
        originalPointIndex,
        series.id,
        pointId,
        keyColor(spec, style),
      );
    }
  }
};

const addScatterDataLabels = (
  elements: ChartPrimitive[],
  view: XYChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  xDomain: Domain,
  yDomain: Domain,
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
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
  const anyLogAxis = xLogBase !== undefined || yLogBase !== undefined;
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const labelSpec = labelSpecFor(spec, series);
    if (!labelSpec || labelSpec.visible === false) continue;
    // The X part follows the x values' format unless the labels set their own.
    const ownCode = labelSpecFor(spec, { id: series.id })?.numberFormat;
    const xLabelCode =
      ownCode !== undefined && ownCode.trim().length > 0 ? ownCode : series.xFormatCode;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < series.points.length; pointIndex += 1) {
      const point = series.points[pointIndex];
      if (!point) continue;
      if (
        point.x === null ||
        point.y === null ||
        !Number.isFinite(point.x) ||
        !Number.isFinite(point.y) ||
        (xLogBase !== undefined && !(point.x > 0)) ||
        (yLogBase !== undefined && !(point.y > 0)) ||
        (anyLogAxis &&
          (point.x < xDomain.min ||
            point.x > xDomain.max ||
            point.y < yDomain.min ||
            point.y > yDomain.max))
      )
        continue;
      const originalPointIndex = view.originalPointIndices[seriesIndex]?.[pointIndex] ?? pointIndex;
      const pointId = pointIdAt(view, seriesIndex, pointIndex);
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        pointId,
        originalSeriesIndex,
        originalPointIndex,
      );
      const fullText = labelParts(
        labelSpec,
        series.name ?? String(originalSeriesIndex + 1),
        formatChartValue(point.x, xLabelCode, locale, formatter) ?? undefined,
        point.y,
        undefined,
        locale,
        formatter,
      );
      const x = xScale(point.x);
      const y = yScale(point.y);
      const requested = resolveChartLabelPosition(spec, labelSpec.position);
      const anchor: LabelAnchor =
        requested === 'center'
          ? { x, y }
          : requested === 'below'
            ? { x, y: y + 12 }
            : requested === 'left'
              ? { x: x - 5, y, textAnchor: 'end' }
              : requested === 'right'
                ? { x: x + 5, y, textAnchor: 'start' }
                : { x, y: y - 5 };
      addDataLabel(
        elements,
        fullText,
        anchor,
        labelSpec,
        width,
        height,
        originalSeriesIndex,
        originalPointIndex,
        series.id,
        pointId,
        keyColor(spec, style),
      );
    }
  }
};

const addPieDataLabels = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): void => {
  const data = view.data;
  const first = data.series[0];
  if (!first) return;
  const labelSpec = labelSpecFor(spec, first);
  if (!labelSpec || labelSpec.visible === false) return;
  const magnitudes = first.values.map((value) =>
    value !== null && Number.isFinite(value) ? Math.abs(value) : 0,
  );
  const scale = magnitudes.reduce((maximum, value) => Math.max(maximum, value), 0);
  if (!(scale > 0)) return;
  const normalized = magnitudes.map((value) => value / scale);
  const total = normalized.reduce((sum, value) => sum + value, 0);
  if (!(total > 0)) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const radius = Math.max(0, Math.min(plot.width, plot.height) / 2 - 2);
  let start = -Math.PI / 2 + (boundedOption(spec.firstSliceAngle, 0, 360, 0) * Math.PI) / 180;
  for (let pointIndex = 0; pointIndex < normalized.length; pointIndex += 1) {
    const normalizedValue = normalized[pointIndex] ?? 0;
    if (!(normalizedValue > 0)) continue;
    const angle = (normalizedValue / total) * Math.PI * 2;
    const middle = start + angle / 2;
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
    const raw = first.values[pointIndex];
    if (raw === null || !Number.isFinite(raw)) {
      start += angle;
      continue;
    }
    const fullText = labelParts(
      labelSpec,
      first.name ?? String((view.originalSeriesIndices[0] ?? 0) + 1),
      data.categories[pointIndex],
      raw,
      normalizedValue / total,
      locale,
      formatter,
    );
    const requested = resolveChartLabelPosition(spec, labelSpec.position);
    const distance =
      requested === 'outside-end'
        ? radius + 8
        : requested === 'inside-end'
          ? radius * 0.85
          : radius * 0.58;
    const rawAnchor = {
      x: cx + Math.cos(middle) * distance,
      y: cy + Math.sin(middle) * distance,
    };
    const anchor: LabelAnchor =
      requested === 'above'
        ? { ...rawAnchor, y: rawAnchor.y - radius * 0.2 }
        : requested === 'below'
          ? { ...rawAnchor, y: rawAnchor.y + radius * 0.2 }
          : requested === 'left'
            ? { ...rawAnchor, x: rawAnchor.x - radius * 0.2, textAnchor: 'end' }
            : requested === 'right'
              ? { ...rawAnchor, x: rawAnchor.x + radius * 0.2, textAnchor: 'start' }
              : {
                  ...rawAnchor,
                  textAnchor:
                    requested === 'outside-end'
                      ? Math.cos(middle) > 0.2
                        ? 'start'
                        : Math.cos(middle) < -0.2
                          ? 'end'
                          : 'middle'
                      : 'middle',
                };
    addDataLabel(
      elements,
      fullText,
      anchor,
      labelSpec,
      width,
      height,
      view.originalSeriesIndices[0] ?? 0,
      originalPointIndex,
      first.id,
      pointId,
      fillColor(style.fill, style.color),
    );
    start += angle;
  }
};

const addDoughnutDataLabels = (
  elements: ChartPrimitive[],
  view: CategoryChartView,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
): void => {
  const data = view.data;
  if (data.series.length === 0) return;
  const cx = plot.x + plot.width / 2;
  const cy = plot.y + plot.height / 2;
  const outerRadius = Math.max(0, Math.min(plot.width, plot.height) / 2 - 2);
  const holePercent = boundedOption(spec.holeSize, 10, 90, 50);
  const holeRadius = (outerRadius * holePercent) / 100;
  const ringWidth = (outerRadius - holeRadius) / data.series.length;
  const firstSliceAngle = boundedOption(spec.firstSliceAngle, 0, 360, 0);
  const labelAngles = (start: number, angle: number): readonly number[] => {
    const middle = start + angle / 2;
    if (angle >= Math.PI * 2 - 1e-9) {
      return [
        middle,
        middle + Math.PI,
        middle + Math.PI / 2,
        middle - Math.PI / 2,
        middle + Math.PI / 4,
        middle - Math.PI / 4,
        middle + (Math.PI * 3) / 4,
        middle - (Math.PI * 3) / 4,
      ];
    }
    const fractions = [0, -0.15, 0.15, -0.3, 0.3, -0.45, 0.45, -0.49, 0.49];
    const minimum = start + angle * 0.02;
    const maximum = start + angle * 0.98;
    return fractions.map((fraction) => clamp(middle + fraction * angle, minimum, maximum));
  };
  const records: DoughnutLabelRecord[] = [];
  for (let seriesIndex = 0; seriesIndex < data.series.length; seriesIndex += 1) {
    const series = data.series[seriesIndex];
    if (!series) continue;
    const labelSpec = labelSpecFor(spec, series);
    if (!labelSpec || labelSpec.visible === false) continue;
    const normalized = normalizedMagnitudes(series.values);
    if (!normalized) continue;
    const innerRadius = holeRadius + ringWidth * seriesIndex;
    const outerRingRadius = innerRadius + ringWidth;
    const midpointRadius = (innerRadius + outerRingRadius) / 2;
    let start = -Math.PI / 2 + (firstSliceAngle * Math.PI) / 180;
    const originalSeriesIndex = view.originalSeriesIndices[seriesIndex] ?? seriesIndex;
    for (let pointIndex = 0; pointIndex < normalized.values.length; pointIndex += 1) {
      const normalizedValue = normalized.values[pointIndex] ?? 0;
      if (!(normalizedValue > 0) || !Number.isFinite(normalizedValue)) continue;
      const angle = (normalizedValue / normalized.total) * Math.PI * 2;
      if (!(angle > 0) || !Number.isFinite(angle)) continue;
      const raw = series.values[pointIndex];
      if (raw === null || !Number.isFinite(raw)) {
        start += angle;
        continue;
      }
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
      const fullText = labelParts(
        labelSpec,
        series.name ?? String(originalSeriesIndex + 1),
        data.categories[pointIndex],
        raw,
        normalizedValue / normalized.total,
        locale,
        formatter,
      );
      records.push({
        seriesIndex,
        pointIndex,
        originalSeriesIndex,
        originalPointIndex,
        seriesId: series.id,
        pointId,
        start,
        angle,
        midpointRadius,
        text: fullText,
        color: fillColor(style.fill, style.color),
        labelSpec,
      });
      start += angle;
    }
  }

  const candidatesFor = new Map<DoughnutLabelRecord, readonly DoughnutLabelCandidate[]>();
  for (const record of records) {
    const candidates = labelAngles(record.start, record.angle).map((candidateAngle) => {
      const candidateAnchor: LabelAnchor = {
        ...circularPoint(cx, cy, record.midpointRadius, candidateAngle),
        textAnchor: 'middle',
      };
      const candidateElements: ChartPrimitive[] = [];
      addDataLabel(
        candidateElements,
        record.text,
        candidateAnchor,
        record.labelSpec,
        width,
        height,
        record.originalSeriesIndex,
        record.originalPointIndex,
        record.seriesId,
        record.pointId,
        record.color,
      );
      return { anchor: candidateAnchor, bounds: labelBoundsFor(candidateElements) };
    });
    candidatesFor.set(record, candidates);
  }

  const placementOrder = records
    .slice()
    .sort((left, right) =>
      left.pointIndex === right.pointIndex
        ? left.seriesIndex - right.seriesIndex
        : left.pointIndex - right.pointIndex,
    );
  const selections = new Map<DoughnutLabelRecord, DoughnutLabelCandidate>();
  const placedLabels: LabelBounds[] = [];
  const fitsPlacedLabels = (candidate: DoughnutLabelCandidate): boolean => {
    const bounds = candidate.bounds;
    if (bounds === null) return true;
    return placedLabels.every((placed) => !labelBoundsOverlap(bounds, placed));
  };
  const maxPlacementStates = 4096;
  let placementStates = 0;
  const tryPlace = (index: number): boolean => {
    if (index >= placementOrder.length) return true;
    placementStates += 1;
    if (placementStates > maxPlacementStates) return false;
    const record = placementOrder[index];
    if (!record) return false;
    const candidates = candidatesFor.get(record) ?? [];
    for (const candidate of candidates) {
      if (!fitsPlacedLabels(candidate)) continue;
      selections.set(record, candidate);
      if (candidate.bounds !== null) placedLabels.push(candidate.bounds);
      if (tryPlace(index + 1)) return true;
      if (candidate.bounds !== null) placedLabels.pop();
      selections.delete(record);
    }
    return false;
  };
  const placedWithoutBacktracking = placementOrder.length <= 64 && tryPlace(0);
  if (!placedWithoutBacktracking) {
    selections.clear();
    placedLabels.length = 0;
    for (const record of placementOrder) {
      const candidates = candidatesFor.get(record) ?? [];
      const selected = candidates.find((candidate) => fitsPlacedLabels(candidate)) ?? candidates[0];
      if (!selected) continue;
      selections.set(record, selected);
      if (selected.bounds !== null) placedLabels.push(selected.bounds);
    }
  }

  for (const record of records) {
    const selection = selections.get(record);
    if (!selection) continue;
    addDataLabel(
      elements,
      record.text,
      selection.anchor,
      record.labelSpec,
      width,
      height,
      record.originalSeriesIndex,
      record.originalPointIndex,
      record.seriesId,
      record.pointId,
      record.color,
    );
  }
};

export const addDataLabels = (
  elements: ChartPrimitive[],
  view: PreparedChartData,
  spec: ChartSpec,
  plot: ChartScene['plot'],
  xDomain: Domain | null,
  yDomain: Domain | null,
  palette: readonly string[],
  width: number,
  height: number,
  locale: string,
  formatter: ChartValueFormatter | undefined,
  categoryLogBase?: number,
  scatterLogBases?: { readonly x?: number; readonly y?: number },
): void => {
  if (!hasDataLabels(spec)) return;
  if (view.data.domain === 'category') {
    if (spec.kind === 'pie')
      addPieDataLabels(
        elements,
        view as CategoryChartView,
        spec,
        plot,
        palette,
        width,
        height,
        locale,
        formatter,
      );
    else if (spec.kind === 'doughnut')
      addDoughnutDataLabels(
        elements,
        view as CategoryChartView,
        spec,
        plot,
        palette,
        width,
        height,
        locale,
        formatter,
      );
    else if (yDomain || spec.kind === 'bar')
      addCategoryDataLabels(
        elements,
        view as CategoryChartView,
        spec,
        plot,
        (spec.kind === 'bar' ? xDomain : yDomain) as Domain,
        palette,
        width,
        height,
        locale,
        formatter,
        categoryLogBase,
      );
  } else if (xDomain && yDomain) {
    addScatterDataLabels(
      elements,
      view as XYChartView,
      spec,
      plot,
      xDomain,
      yDomain,
      palette,
      width,
      height,
      locale,
      formatter,
      scatterLogBases?.x,
      scatterLogBases?.y,
    );
  }
};
