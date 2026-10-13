/** The seven chart families understood by the session chart renderer. */
export type ChartKind = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';

/** How multiple category series share a bar/line plot. */
export type ChartGrouping = 'clustered' | 'stacked' | 'percent-stacked';

/** Which marks a scatter series paints for each finite point. */
export type ChartScatterStyle = 'marker' | 'line' | 'lineMarker';

/** How an empty cell is treated when laying out a chart. */
export type ChartEmptyCellMode = 'gap' | 'zero' | 'span';

/** Direction in which a rectangular table is read as a category chart. */
export type ChartSeriesDirection = 'auto' | 'columns' | 'rows';

/** Values accepted by the table adapter. Invalid numeric values are null in the model. */
export type ChartTableCell = number | string | boolean | null;

export interface ChartAxisSpec {
  /** Hide the axis line, labels, and title while allowing gridlines independently. */
  readonly visible?: boolean;
  /** Reverse the visual order of this axis while retaining ascending numeric ticks. */
  readonly reverseOrder?: boolean;
  /**
   * Physical position where the perpendicular axis crosses this axis. X-axis
   * crossing moves the y-axis line; y-axis crossing moves the x-axis line.
   * Numeric values are clamped to the resolved axis domain at layout time.
   */
  readonly crossing?: 'auto' | 'min' | 'max' | number;
  readonly title?: string;
  readonly min?: number;
  readonly max?: number;
  /** Explicit finite positive major tick interval. */
  readonly majorUnit?: number;
  /** Excel-compatible number format for numeric tick labels. */
  readonly numberFormat?: string;
  /** Numeric-axis major gridlines; enabled by default. */
  readonly majorGridlines?: boolean;
}

export interface ChartLegendSpec {
  readonly visible?: boolean;
  readonly position?: 'top' | 'right' | 'bottom' | 'left';
}

/** Data identities hidden by an Excel-style chart filter. */
export interface ChartFilterSpec {
  readonly hiddenSeriesIds?: readonly string[];
  readonly hiddenCategoryIds?: readonly string[];
}

export interface ChartFillSpec {
  readonly color?: string;
  /** Excel transparency percentage: 0 is opaque and 100 is invisible. */
  readonly transparency?: number;
}

export interface ChartLineSpec {
  readonly color?: string;
  readonly transparency?: number;
  /** Line width in points. */
  readonly width?: number;
  readonly dash?: 'solid' | 'dash' | 'dot' | 'dash-dot';
}

export interface ChartMarkerSpec {
  readonly visible?: boolean;
  readonly shape?: 'circle' | 'square' | 'diamond' | 'triangle';
  /** Marker size in points. */
  readonly size?: number;
  readonly fill?: ChartFillSpec;
  readonly line?: ChartLineSpec;
}

export interface ChartPointStyleSpec {
  readonly pointId: string;
  readonly fill?: ChartFillSpec;
  readonly line?: ChartLineSpec;
  readonly marker?: ChartMarkerSpec;
}

export interface ChartSeriesStyleSpec {
  readonly seriesId: string;
  readonly fill?: ChartFillSpec;
  readonly line?: ChartLineSpec;
  readonly marker?: ChartMarkerSpec;
  readonly points?: readonly ChartPointStyleSpec[];
}

export type ChartLabelPosition =
  | 'best-fit'
  | 'center'
  | 'inside-base'
  | 'inside-end'
  | 'outside-end'
  | 'above'
  | 'below'
  | 'left'
  | 'right';

export interface ChartDataLabelsSpec {
  readonly visible?: boolean;
  readonly showSeriesName?: boolean;
  readonly showCategoryName?: boolean;
  readonly showValue?: boolean;
  readonly showPercentage?: boolean;
  readonly showLegendKey?: boolean;
  readonly separator?: string;
  readonly position?: ChartLabelPosition;
  readonly numberFormat?: string;
  readonly color?: string;
  readonly fontSize?: number;
}

/** Backwards-compatible singular spelling. Prefer ChartDataLabelsSpec. */
export type ChartDataLabelSpec = ChartDataLabelsSpec;
/** Short spelling used by the standalone chart package API. */
export type DataLabelSpec = ChartDataLabelsSpec;

export type ChartValueFormatter = (
  value: number,
  formatCode: string | undefined,
  locale: string,
) => string | null | undefined;

export interface ChartSpec {
  readonly kind: ChartKind;
  readonly grouping?: ChartGrouping;
  /** Scatter series rendering mode; defaults to markers only. */
  readonly scatterStyle?: ChartScatterStyle;
  readonly seriesDirection?: ChartSeriesDirection;
  readonly emptyCells?: ChartEmptyCellMode;
  /** Excel chart gap width as a percentage of one category slot (0–500). */
  readonly gapWidth?: number;
  /** Excel chart overlap as a percentage for clustered series (-100–100). */
  readonly overlap?: number;
  /** Pie and doughnut first slice angle in degrees clockwise from 12 o'clock (0–360). */
  readonly firstSliceAngle?: number;
  /**
   * Doughnut hole size as a percentage of the outer radius (10–90).
   * Defaults to 50; invalid values also use 50.
   */
  readonly holeSize?: number;
  /** Override automatic first-row header inference for ambiguous tables. */
  readonly headerRow?: boolean;
  /** Override automatic first-column header inference for ambiguous tables. */
  readonly headerColumn?: boolean;
  readonly title?: string;
  readonly palette?: readonly string[];
  readonly legend?: ChartLegendSpec;
  readonly filter?: ChartFilterSpec;
  readonly dataLabels?: ChartDataLabelsSpec;
  readonly seriesStyles?: readonly ChartSeriesStyleSpec[];
  /** Excel "Vary colors by point". Pie and doughnut charts default to true. */
  readonly varyColors?: boolean;
  readonly xAxis?: ChartAxisSpec;
  readonly yAxis?: ChartAxisSpec;
}

export interface CategoryChartSeries {
  readonly id: string;
  readonly name?: string;
  readonly values: readonly (number | null)[];
}

export interface CategoryChartData {
  readonly domain: 'category';
  readonly categories: readonly string[];
  /** Stable category identities used by filters and point styles. */
  readonly categoryIds?: readonly string[];
  readonly series: readonly CategoryChartSeries[];
}

export interface XYChartPoint {
  /** Stable point identity used by filters and point styles. */
  readonly id?: string;
  readonly x: number | null;
  readonly y: number | null;
}

export interface XYChartSeries {
  readonly id: string;
  readonly name?: string;
  readonly points: readonly XYChartPoint[];
}

export interface XYChartData {
  readonly domain: 'xy';
  readonly series: readonly XYChartSeries[];
}

export type ChartData = CategoryChartData | XYChartData;

export type ChartPrimitiveTag =
  | 'rect'
  | 'line'
  | 'polyline'
  | 'polygon'
  | 'circle'
  | 'path'
  | 'text';

export type ChartPrimitiveRole =
  | 'background'
  | 'title'
  | 'axis'
  | 'axis-tick'
  | 'axis-label'
  | 'axis-title'
  | 'category-label'
  | 'legend'
  | 'bar'
  | 'line'
  | 'area'
  | 'point'
  | 'slice'
  | 'data-label'
  | 'data-label-key'
  | 'empty';

/** A browser-independent SVG-like element description. */
export interface ChartPrimitive {
  readonly tag: ChartPrimitiveTag;
  readonly attrs: Readonly<Record<string, string | number>>;
  readonly text?: string;
  readonly role?: ChartPrimitiveRole | string;
  readonly seriesIndex?: number;
  readonly pointIndex?: number;
  readonly seriesId?: string;
  readonly pointId?: string;
}

export interface ChartPlotBounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** Result of pure chart layout. Consumers can map each element to any SVG host. */
export interface ChartScene {
  readonly width: number;
  readonly height: number;
  readonly plot: ChartPlotBounds;
  readonly elements: readonly ChartPrimitive[];
}
