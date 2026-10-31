/** Label strings and English defaults for the chart settings dialog. */

export interface SessionChartSettingsLabels {
  columnChart: string;
  barChart: string;
  lineChart: string;
  areaChart: string;
  pieChart: string;
  doughnutChart: string;
  scatterChart: string;
  chartType: string;
  scatterStyle: string;
  scatterStyleMarker: string;
  scatterStyleLine: string;
  scatterStyleLineMarker: string;
  holeSize: string;
  settings: string;
  elementsTab: string;
  stylesTab: string;
  filtersTab: string;
  closeSettings: string;
  apply: string;
  cancel: string;
  dataLabels: string;
  showDataLabels: string;
  showSeriesName: string;
  showCategoryName: string;
  showValue: string;
  showPercentage: string;
  showLegendKey: string;
  labelSeparator: string;
  labelPosition: string;
  positionBestFit: string;
  positionAutomatic: string;
  positionCenter: string;
  positionInsideBase: string;
  positionInsideEnd: string;
  positionOutsideEnd: string;
  positionAbove: string;
  positionBelow: string;
  positionLeft: string;
  positionRight: string;
  numberFormat: string;
  textColor: string;
  fontSize: string;
  series: string;
  point: string;
  allPoints: string;
  fillColor: string;
  lineColor: string;
  lineWidth: string;
  transparency: string;
  lineDash: string;
  dashSolid: string;
  dashDash: string;
  dashDot: string;
  dashDashDot: string;
  marker: string;
  markerVisible: string;
  markerShape: string;
  shapeCircle: string;
  shapeSquare: string;
  shapeDiamond: string;
  shapeTriangle: string;
  shapeX: string;
  shapeStar: string;
  shapePlus: string;
  shapeDot: string;
  shapeDash: string;
  smoothLine: string;
  titleAutomatic: string;
  markerSize: string;
  categories: string;
  selectAll: string;
  clearAll: string;
  noSeries: string;
  noCategories: string;
  pivotChartLocked: string;
  addSeries: string;
  removeSeries: string;
  addSeriesConfirm: string;
  seriesNameRef: string;
  seriesValuesRef: string;
  seriesXValuesRef: string;
  seriesYValuesRef: string;
  axisLabelsRef: string;
  invalidReference: string;
  lastSeriesLocked: string;
  palette: string;
  paletteOffice: string;
  paletteBlue: string;
  paletteMonochrome: string;
  varyColors: string;
  title: string;
  legend: string;
  legendPosition: string;
  grouping: string;
  groupingClustered: string;
  groupingStacked: string;
  groupingPercentStacked: string;
  seriesDirection: string;
  seriesDirectionAuto: string;
  seriesDirectionColumns: string;
  seriesDirectionRows: string;
  emptyCells: string;
  emptyCellsGap: string;
  emptyCellsZero: string;
  emptyCellsSpan: string;
  axisTitles: string;
  xAxis: string;
  yAxis: string;
  xAxisTitle: string;
  yAxisTitle: string;
  showAxis: string;
  reverseOrder: string;
  verticalAxisCrosses: string;
  horizontalAxisCrosses: string;
  labelPositionNextTo: string;
  labelPositionLow: string;
  labelPositionHigh: string;
  labelPositionNone: string;
  logarithmicScale: string;
  logBase: string;
  custom: string;
  customValue: string;
  minimum: string;
  maximum: string;
  majorUnit: string;
  axisNumberFormat: string;
  majorGridlines: string;
  automatic: string;
  gapWidth: string;
  overlap: string;
  firstSliceAngle: string;
  positionTop: string;
  positionBottom: string;
}

export const DEFAULT_LABELS: SessionChartSettingsLabels = {
  columnChart: 'Column chart',
  barChart: 'Bar chart',
  lineChart: 'Line chart',
  areaChart: 'Area chart',
  pieChart: 'Pie chart',
  doughnutChart: 'Doughnut chart',
  scatterChart: 'Scatter chart',
  chartType: 'Chart type',
  scatterStyle: 'Scatter style',
  scatterStyleMarker: 'Markers',
  scatterStyleLine: 'Straight lines',
  scatterStyleLineMarker: 'Straight lines with markers',
  holeSize: 'Hole size (%)',
  settings: 'Chart settings',
  elementsTab: 'Elements',
  stylesTab: 'Styles',
  filtersTab: 'Filters',
  closeSettings: 'Close chart settings',
  apply: 'Apply',
  cancel: 'Cancel',
  dataLabels: 'Data labels',
  showDataLabels: 'Show data labels',
  showSeriesName: 'Series name',
  showCategoryName: 'Category name',
  showValue: 'Value',
  showPercentage: 'Percentage',
  showLegendKey: 'Legend key',
  labelSeparator: 'Separator',
  labelPosition: 'Label position',
  positionBestFit: 'Best fit',
  positionAutomatic: 'Automatic',
  positionCenter: 'Center',
  positionInsideBase: 'Inside base',
  positionInsideEnd: 'Inside end',
  positionOutsideEnd: 'Outside end',
  positionAbove: 'Above',
  positionBelow: 'Below',
  positionLeft: 'Left',
  positionRight: 'Right',
  numberFormat: 'Number format',
  textColor: 'Text color',
  fontSize: 'Font size',
  series: 'Series',
  point: 'Point',
  allPoints: 'All points',
  fillColor: 'Fill color',
  lineColor: 'Line color',
  lineWidth: 'Line width',
  transparency: 'Transparency',
  lineDash: 'Line dash',
  dashSolid: 'Solid',
  dashDash: 'Dash',
  dashDot: 'Dot',
  dashDashDot: 'Dash-dot',
  marker: 'Marker',
  markerVisible: 'Visible',
  markerShape: 'Shape',
  shapeCircle: 'Circle',
  shapeSquare: 'Square',
  shapeDiamond: 'Diamond',
  shapeTriangle: 'Triangle',
  shapeX: 'X',
  shapeStar: 'Star',
  shapePlus: 'Plus',
  shapeDot: 'Dot',
  shapeDash: 'Dash',
  smoothLine: 'Smoothed line',
  titleAutomatic: 'Automatic title',
  markerSize: 'Size',
  categories: 'Categories',
  selectAll: 'Select all',
  clearAll: 'Clear all',
  noSeries: 'No series',
  noCategories: 'No categories',
  pivotChartLocked: 'The series and data range of a PivotChart follow its PivotTable.',
  addSeries: 'Add series',
  removeSeries: 'Remove series',
  addSeriesConfirm: 'Add',
  seriesNameRef: 'Series name',
  seriesValuesRef: 'Series values',
  seriesXValuesRef: 'Series X values',
  seriesYValuesRef: 'Series Y values',
  axisLabelsRef: 'Axis label range',
  invalidReference: 'Enter a valid cell reference.',
  lastSeriesLocked: 'A chart needs at least one series.',
  palette: 'Palette',
  paletteOffice: 'Office',
  paletteBlue: 'Blue',
  paletteMonochrome: 'Monochrome',
  varyColors: 'Vary colors by point',
  title: 'Title',
  legend: 'Legend',
  legendPosition: 'Legend position',
  grouping: 'Grouping',
  groupingClustered: 'Clustered',
  groupingStacked: 'Stacked',
  groupingPercentStacked: '100% stacked',
  seriesDirection: 'Series direction',
  seriesDirectionAuto: 'Automatic',
  seriesDirectionColumns: 'Columns',
  seriesDirectionRows: 'Rows',
  emptyCells: 'Empty cells',
  emptyCellsGap: 'Gaps',
  emptyCellsZero: 'Zero',
  emptyCellsSpan: 'Connect data points',
  axisTitles: 'Axis titles',
  xAxis: 'X-axis',
  yAxis: 'Y-axis',
  xAxisTitle: 'X-axis title',
  yAxisTitle: 'Y-axis title',
  showAxis: 'Show axis',
  reverseOrder: 'Reverse order',
  verticalAxisCrosses: 'Vertical axis crosses',
  horizontalAxisCrosses: 'Horizontal axis crosses',
  labelPositionNextTo: 'Next to axis',
  labelPositionLow: 'Low (bottom/left)',
  labelPositionHigh: 'High (top/right)',
  labelPositionNone: 'None',
  logarithmicScale: 'Logarithmic scale',
  logBase: 'Log base',
  custom: 'Custom',
  customValue: 'Custom value',
  minimum: 'Minimum',
  maximum: 'Maximum',
  majorUnit: 'Major unit',
  axisNumberFormat: 'Number format',
  majorGridlines: 'Major gridlines',
  automatic: 'Automatic',
  gapWidth: 'Gap width (%)',
  overlap: 'Overlap (%)',
  firstSliceAngle: 'First slice angle (°)',
  positionTop: 'Top',
  positionBottom: 'Bottom',
};
