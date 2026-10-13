import type {
  CategoryChartData,
  ChartData,
  ChartSeriesDirection,
  ChartSpec,
  ChartTableCell,
  XYChartData,
  XYChartSeries,
} from './types.js';

const isFiniteNumber = (value: ChartTableCell | undefined): value is number =>
  typeof value === 'number' && Number.isFinite(value);

const isText = (value: ChartTableCell | undefined): value is string =>
  typeof value === 'string' && value.trim().length > 0;

const asNumber = (value: ChartTableCell | undefined): number | null =>
  isFiniteNumber(value) ? value : null;

const asLabel = (value: ChartTableCell | undefined, fallback: string): string => {
  if (typeof value === 'string' && value.trim().length > 0) return value;
  if (isFiniteNumber(value)) return String(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return fallback;
};

const asHeaderName = (value: ChartTableCell | undefined): string | undefined => {
  const label = asLabel(value, '');
  return label.length > 0 ? label : undefined;
};

const widthOf = (table: readonly (readonly ChartTableCell[])[]): number =>
  table.reduce((width, row) => Math.max(width, row.length), 0);

const cellAt = (
  table: readonly (readonly ChartTableCell[])[],
  row: number,
  column: number,
): ChartTableCell | undefined => table[row]?.[column];

/** A textual first row names series when the table has at least one value row.
 * A leading label followed only by blanks is also a useful partial header;
 * numeric/date-looking headers remain opt-in through `headerRow`. */
const hasHeaderRow = (table: readonly (readonly ChartTableCell[])[], width: number): boolean => {
  if (table.length < 2 || width === 0) return false;
  const firstRow = table[0] ?? [];
  const dataStart = firstRow.length > 1 && isText(firstRow[0]) ? 1 : 0;
  const cells = firstRow.slice(dataStart);
  if (cells.some(isText)) return true;
  const firstColumnHasLabels = table.slice(1).some((row) => isText(row[0]));
  return (
    isText(firstRow[0]) &&
    cells.length > 0 &&
    !firstColumnHasLabels &&
    cells.every(
      (value) => value === null || (typeof value === 'string' && value.trim().length === 0),
    )
  );
};

/** A textual first column names categories or series, depending on direction. */
const hasHeaderColumn = (table: readonly (readonly ChartTableCell[])[], width: number): boolean => {
  if (table.length < 2 || width < 2) return false;
  return table.slice(1).some((row) => isText(row[0]));
};

const hasSingleRowSeriesLabel = (
  table: readonly (readonly ChartTableCell[])[],
  width: number,
): boolean =>
  table.length === 1 &&
  width > 1 &&
  isText(table[0]?.[0]) &&
  table[0]?.slice(1).some(isFiniteNumber) === true;

const hasSingleColumnSeriesLabel = (
  table: readonly (readonly ChartTableCell[])[],
  width: number,
): boolean =>
  width === 1 &&
  table.length > 1 &&
  isText(table[0]?.[0]) &&
  table.slice(1).some((row) => isFiniteNumber(row[0]));

const resolveDirection = (
  table: readonly (readonly ChartTableCell[])[],
  direction: ChartSeriesDirection | undefined,
): Exclude<ChartSeriesDirection, 'auto'> => {
  if (direction === 'columns' || direction === 'rows') return direction;
  const rows = table.length;
  const columns = widthOf(table);
  // A one-dimensional selection is naturally one series. This also avoids
  // interpreting a single row of values as one series per cell.
  if (rows <= 1) return 'rows';
  if (columns <= 1) return 'columns';
  return rows >= columns ? 'columns' : 'rows';
};

const emptyCategoryData = (): CategoryChartData => ({
  domain: 'category',
  categories: [],
  series: [],
});

const makeCategoryData = (
  table: readonly (readonly ChartTableCell[])[],
  direction: Exclude<ChartSeriesDirection, 'auto'>,
  headerRowOverride: boolean | undefined,
  headerColumnOverride: boolean | undefined,
): CategoryChartData => {
  const width = widthOf(table);
  if (table.length === 0 || width === 0) return emptyCategoryData();

  const headerRow = headerRowOverride ?? hasHeaderRow(table, width);
  const headerColumn = headerColumnOverride ?? hasHeaderColumn(table, width);
  // A one-dimensional selection has no separate header axis to infer from.
  // Treat its leading label as the corresponding header only while that
  // header's automatic inference remains enabled. Explicit `false` must keep
  // the cell in the data grid as a null value instead of silently naming the
  // series.
  const singleRowSeriesLabel =
    direction === 'rows' && headerColumnOverride !== false && hasSingleRowSeriesLabel(table, width);
  const singleColumnSeriesLabel =
    direction === 'columns' &&
    headerRowOverride !== false &&
    hasSingleColumnSeriesLabel(table, width);
  const categories: string[] = [];
  const series: CategoryChartData['series'][number][] = [];

  if (direction === 'columns') {
    const firstDataRow = headerRow || singleColumnSeriesLabel ? 1 : 0;
    const firstDataColumn = headerColumn ? 1 : 0;
    const categoryCount = Math.max(0, table.length - firstDataRow);
    const seriesCount = Math.max(0, width - firstDataColumn);
    for (let pointIndex = 0; pointIndex < categoryCount; pointIndex += 1) {
      categories.push(
        headerColumn
          ? asLabel(cellAt(table, firstDataRow + pointIndex, 0), String(pointIndex + 1))
          : String(pointIndex + 1),
      );
    }
    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
      const column = firstDataColumn + seriesIndex;
      const header = singleColumnSeriesLabel
        ? cellAt(table, 0, 0)
        : headerRow
          ? cellAt(table, 0, column)
          : undefined;
      const name = asHeaderName(header);
      const values = Array.from({ length: categoryCount }, (_, pointIndex) =>
        asNumber(cellAt(table, firstDataRow + pointIndex, column)),
      );
      series.push({ id: `series-${seriesIndex}`, ...(name ? { name } : {}), values });
    }
  } else {
    const firstDataRow = headerRow ? 1 : 0;
    const firstDataColumn = headerColumn || singleRowSeriesLabel ? 1 : 0;
    const categoryCount = Math.max(0, width - firstDataColumn);
    const seriesCount = Math.max(0, table.length - firstDataRow);
    for (let pointIndex = 0; pointIndex < categoryCount; pointIndex += 1) {
      categories.push(
        headerRow
          ? asLabel(cellAt(table, 0, firstDataColumn + pointIndex), String(pointIndex + 1))
          : String(pointIndex + 1),
      );
    }
    for (let seriesIndex = 0; seriesIndex < seriesCount; seriesIndex += 1) {
      const row = firstDataRow + seriesIndex;
      const header = headerColumn || singleRowSeriesLabel ? cellAt(table, row, 0) : undefined;
      const name = asHeaderName(header);
      const values = Array.from({ length: categoryCount }, (_, pointIndex) =>
        asNumber(cellAt(table, row, firstDataColumn + pointIndex)),
      );
      series.push({ id: `series-${seriesIndex}`, ...(name ? { name } : {}), values });
    }
  }

  return { domain: 'category', categories, series };
};

const transposeTable = (table: readonly (readonly ChartTableCell[])[]): ChartTableCell[][] => {
  const width = widthOf(table);
  return Array.from({ length: width }, (_, row) =>
    Array.from({ length: table.length }, (_, column) => cellAt(table, column, row) ?? null),
  );
};

const makeScatterData = (
  table: readonly (readonly ChartTableCell[])[],
  headerRowOverride: boolean | undefined,
): XYChartData => {
  if (table.length === 0) return { domain: 'xy', series: [] };
  const width = widthOf(table);
  if (width === 0) return { domain: 'xy', series: [] };
  const headerRow = headerRowOverride ?? hasHeaderRow(table, width);
  const firstDataRow = headerRow ? 1 : 0;
  const numericColumns: number[] = [];
  for (let column = 0; column < width; column += 1) {
    if (table.slice(firstDataRow).some((row) => isFiniteNumber(row[column]))) {
      numericColumns.push(column);
    }
  }
  if (numericColumns.length < 2) return { domain: 'xy', series: [] };

  const xColumn = numericColumns[0] as number;
  const series: XYChartSeries[] = numericColumns.slice(1).map((column, seriesIndex) => {
    const header = headerRow ? cellAt(table, 0, column) : undefined;
    const name = asHeaderName(header);
    const points = table.slice(firstDataRow).map((row) => ({
      x: asNumber(row[xColumn]),
      y: asNumber(row[column]),
    }));
    return {
      id: `series-${seriesIndex}`,
      ...(name ? { name } : {}),
      points,
    };
  });
  return { domain: 'xy', series };
};

/**
 * Interpret a rectangular spreadsheet selection without depending on a sheet,
 * workbook, DOM, or formula engine. Header detection is intentionally limited
 * to textual rows/columns, so numeric/date-looking headers remain chart data
 * unless `headerRow` or `headerColumn` explicitly opts in.
 */
export function chartDataFromTable(
  table: readonly (readonly ChartTableCell[])[],
  spec: Pick<ChartSpec, 'kind' | 'seriesDirection' | 'headerRow' | 'headerColumn'>,
): ChartData {
  if (spec.kind === 'scatter') {
    const orientedTable = spec.seriesDirection === 'rows' ? transposeTable(table) : table;
    const orientedHeaderRow =
      spec.seriesDirection === 'rows' ? (spec.headerColumn ?? spec.headerRow) : spec.headerRow;
    return makeScatterData(orientedTable, orientedHeaderRow);
  }
  return makeCategoryData(
    table,
    resolveDirection(table, spec.seriesDirection),
    spec.headerRow,
    spec.headerColumn,
  );
}
