import type { CategoryChartView, PreparedChartData } from './filters.js';
import {
  categoryIdAt,
  primitive,
  TITLE_RESERVE,
  textWidth,
  truncateLabel,
} from './scene-primitives.js';
import { isCircularKind } from './series-geometry.js';
import { fillColor, keyColor, resolveChartStyle } from './styles.js';
import type { ChartPrimitive, ChartScene, ChartSpec } from './types.js';

const colorFor = (palette: readonly string[], index: number): string =>
  palette[((index % palette.length) + palette.length) % palette.length] ?? '#156082';

export const addLegend = (
  elements: ChartPrimitive[],
  view: PreparedChartData,
  spec: ChartSpec,
  width: number,
  height: number,
  plot: ChartScene['plot'],
  palette: readonly string[],
  rightLabelGutter = 0,
): void => {
  const data = view.data;
  const entries: {
    label: string;
    color: string;
    seriesIndex?: number;
    pointIndex?: number;
    seriesId?: string;
    pointId?: string;
  }[] = [];
  if (data.domain === 'category' && isCircularKind(spec.kind)) {
    const categoryView = view as CategoryChartView;
    if (spec.kind === 'doughnut' && data.series.length === 0) return;
    for (let index = 0; index < data.categories.length; index += 1) {
      const originalPointIndex = categoryView.originalCategoryIndices[index] ?? index;
      const pointId = categoryIdAt(categoryView, index);
      const series = data.series[0];
      const style = series
        ? resolveChartStyle(
            spec,
            palette,
            series.id,
            pointId,
            categoryView.originalSeriesIndices[0] ?? 0,
            originalPointIndex,
          )
        : undefined;
      entries.push({
        label: data.categories[index] ?? `${index + 1}`,
        color: style ? fillColor(style.fill, style.color) : colorFor(palette, originalPointIndex),
        seriesIndex: categoryView.originalSeriesIndices[0] ?? 0,
        pointIndex: originalPointIndex,
        seriesId: series?.id,
        pointId,
      });
    }
  } else if (data.domain === 'category') {
    const categoryView = view as CategoryChartView;
    data.series.forEach((series, index) => {
      const originalSeriesIndex = categoryView.originalSeriesIndices[index] ?? index;
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        undefined,
        originalSeriesIndex,
        undefined,
      );
      entries.push({
        label: series.name ?? `${originalSeriesIndex + 1}`,
        color: keyColor(spec, style),
        seriesIndex: originalSeriesIndex,
        seriesId: series.id,
      });
    });
  } else {
    data.series.forEach((series, index) => {
      const originalSeriesIndex = view.originalSeriesIndices[index] ?? index;
      const style = resolveChartStyle(
        spec,
        palette,
        series.id,
        undefined,
        originalSeriesIndex,
        undefined,
      );
      entries.push({
        label: series.name ?? `${originalSeriesIndex + 1}`,
        color: keyColor(spec, style),
        seriesIndex: originalSeriesIndex,
        seriesId: series.id,
      });
    });
  }
  if (entries.length === 0) return;
  const position = spec.legend?.position ?? 'bottom';
  let x = plot.x;
  let y = 12;
  let horizontal = false;
  if (position === 'right') {
    x = plot.x + plot.width + 8 + rightLabelGutter;
    y = plot.y + 8;
  } else if (position === 'left') {
    x = 4;
    y = plot.y + 8;
  } else if (position === 'bottom') {
    x = 4;
    y = height - 5;
    horizontal = true;
  } else {
    x = 4;
    y = spec.title ? TITLE_RESERVE + 5 : 11;
    horizontal = true;
  }
  const availableWidth = Math.max(
    12,
    position === 'right' ? width - x - 4 : position === 'left' ? x + plot.x - 4 : width - x - 4,
  );
  const availableHeight = Math.max(14, position === 'bottom' ? height - 4 : height - y - 4);
  const preferredWidth = entries.reduce(
    (current, entry) => Math.max(current, textWidth(entry.label) + 20),
    56,
  );
  const columnWidth = Math.max(20, Math.min(preferredWidth, availableWidth));
  const verticalRows = Math.max(1, Math.floor(availableHeight / 14));
  const capacity = horizontal
    ? Math.max(1, Math.floor(availableWidth / 56))
    : Math.max(1, Math.floor(availableWidth / columnWidth)) * verticalRows;
  const visibleEntries =
    entries.length > capacity
      ? [
          ...entries.slice(0, Math.max(1, capacity - 1)),
          {
            label: `+${entries.length - Math.max(1, capacity - 1)}`,
            color: 'currentColor',
            seriesIndex: undefined,
            pointIndex: undefined,
            seriesId: undefined,
            pointId: undefined,
          },
        ]
      : entries;
  const items = visibleEntries.map((entry) => {
    const itemMaxWidth = horizontal
      ? Math.max(12, availableWidth / Math.max(1, visibleEntries.length) - 20)
      : Math.max(12, columnWidth - 11);
    const displayLabel = truncateLabel(entry.label, itemMaxWidth);
    return { entry, displayLabel, width: 20 + Math.min(itemMaxWidth, textWidth(displayLabel)) };
  });
  // A horizontal legend wraps into rows, each centred on the chart as Excel does.
  const rows: (typeof items)[] = [];
  if (horizontal) {
    let rowWidth = 0;
    for (const item of items) {
      const current = rows[rows.length - 1];
      if (current && rowWidth + item.width <= width - 8) {
        current.push(item);
        rowWidth += item.width;
      } else {
        rows.push([item]);
        rowWidth = item.width;
      }
    }
  }
  const placed = horizontal
    ? rows.flatMap((rowItems, row) => {
        // The trailing gap after the last label is not part of the row.
        const rowWidth = rowItems.reduce((sum, item) => sum + item.width, 0) - 9;
        let cursor = Math.max(4, (width - rowWidth) / 2);
        // A bottom legend grows upwards so its first row stays nearest the plot.
        const itemY = position === 'bottom' ? y - (rows.length - 1 - row) * 14 : y + row * 14;
        return rowItems.map((item) => {
          const itemX = cursor;
          cursor += item.width;
          return { ...item, itemX, itemY };
        });
      })
    : items.map((item, index) => ({
        ...item,
        itemX: x + Math.floor(index / verticalRows) * columnWidth,
        itemY: y + (index % verticalRows) * 14,
      }));
  for (const { entry, displayLabel, itemX, itemY } of placed) {
    elements.push(
      primitive(
        'rect',
        { x: itemX, y: itemY - 8, width: 8, height: 8, fill: entry.color },
        'legend',
        entry.seriesIndex,
        entry.pointIndex,
        undefined,
        { seriesId: entry.seriesId, pointId: entry.pointId },
      ),
      primitive(
        'text',
        { x: itemX + 11, y: itemY - 1, 'aria-label': entry.label },
        'legend',
        entry.seriesIndex,
        entry.pointIndex,
        displayLabel,
        { seriesId: entry.seriesId, pointId: entry.pointId },
      ),
    );
  }
};
