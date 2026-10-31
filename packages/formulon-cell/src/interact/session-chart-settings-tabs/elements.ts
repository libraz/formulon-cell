/** Elements tab of the chart settings dialog: titles, labels, axes and legend. */

import { chartLabelPositions } from '../../charts/label-options.js';
import type {
  ChartData,
  ChartEmptyCellMode,
  ChartGrouping,
  ChartLabelPosition,
  ChartScatterStyle,
  ChartSeriesDirection,
} from '../../charts/types.js';
import { projectDisabledState } from '../../toolbar/menu-a11y.js';
import {
  element,
  makeCheckbox,
  makeColorField,
  makeSelect,
  makeTextInput,
  textNode,
} from '../session-chart-settings-dom.js';
import type { SessionChartSettingsLabels } from '../session-chart-settings-labels.js';
import {
  type AxisCrossingMode,
  type AxisLabelPositionMode,
  type AxisName,
  axisIsNumeric,
  type ChartDraft,
  categoryCountForDraft,
  effectiveLegendDefault,
  isCircularKind,
  isDoughnutKind,
  type MutableDataLabels,
  normalizedCrossingCustom,
  type SessionChartSettingsShape,
  updateAxisCrossingCustom,
  updateAxisCrossingMode,
  updateAxisField,
  updateAxisLabelPosition,
  updateAxisLogBase,
  updateAxisLogEnabled,
  updateDataLabel,
} from '../session-chart-settings-model.js';

const parseOptionalNumber = (value: string): number | undefined => {
  if (value.trim() === '') return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
};

const POSITION_LABELS: Record<ChartLabelPosition, keyof SessionChartSettingsLabels> = {
  'best-fit': 'positionBestFit',
  center: 'positionCenter',
  'inside-base': 'positionInsideBase',
  'inside-end': 'positionInsideEnd',
  'outside-end': 'positionOutsideEnd',
  above: 'positionAbove',
  below: 'positionBelow',
  left: 'positionLeft',
  right: 'positionRight',
};

const allowedPositions = (
  chart: Pick<SessionChartSettingsShape, 'kind' | 'grouping'>,
  current: ChartLabelPosition,
): readonly {
  value: ChartLabelPosition;
  label: keyof SessionChartSettingsLabels;
  disabled?: boolean;
}[] => {
  const supported = chartLabelPositions(chart);
  const options: {
    value: ChartLabelPosition;
    label: keyof SessionChartSettingsLabels;
    disabled?: boolean;
  }[] = isCircularKind(chart.kind)
    ? supported.map((value) => ({ value, label: POSITION_LABELS[value] }))
    : [
        {
          value: 'best-fit',
          label: 'positionAutomatic',
          disabled: chart.kind === 'area',
        },
        ...supported.map((value) => ({ value, label: POSITION_LABELS[value] })),
      ];
  // A workbook may already contain a position that is invalid for its current
  // chart grouping. Keep that value in the draft until the user chooses a
  // supported replacement, while leaving it unavailable as a new choice.
  if (!options.some((option) => option.value === current)) {
    options.push({ value: current, label: POSITION_LABELS[current], disabled: true });
  }
  return options;
};

const supportsPercentage = (
  chart: Pick<SessionChartSettingsShape, 'kind'>,
  grouping: ChartGrouping,
): boolean => isCircularKind(chart.kind) || grouping === 'percent-stacked';

const groupingSupported = (chart: Pick<SessionChartSettingsShape, 'kind'>): boolean =>
  chart.kind === 'bar' || chart.kind === 'column' || chart.kind === 'line' || chart.kind === 'area';

const axisLogEligible = (
  chart: Pick<SessionChartSettingsShape, 'kind' | 'grouping'>,
  axis: AxisName,
): boolean => {
  if (!axisIsNumeric(chart, axis)) return false;
  if (chart.kind === 'scatter') return true;
  return (chart.grouping ?? 'clustered') === 'clustered';
};

const emptyCellOptionDisabled = (
  chart: Pick<SessionChartSettingsShape, 'kind' | 'scatterStyle'>,
  grouping: ChartGrouping,
  value: ChartEmptyCellMode,
): boolean => {
  if (value !== 'span') return false;
  if (chart.kind === 'scatter') {
    return chart.scatterStyle !== 'line' && chart.scatterStyle !== 'lineMarker';
  }
  return chart.kind !== 'line' || grouping !== 'clustered';
};

const emptyCellsForcedZero = (
  chart: Pick<SessionChartSettingsShape, 'kind'>,
  grouping: ChartGrouping,
): boolean => (chart.kind === 'line' || chart.kind === 'area') && grouping !== 'clustered';

const renderAxisControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  axisName: AxisName,
  axisLabel: string,
  rerender: () => void,
): void => {
  const axis = draft[axisName];
  const enabled = !isCircularKind(chart.kind);
  const numeric = axisIsNumeric(chart, axisName);
  const logEligible = axisLogEligible(chart, axisName);
  const categoryCount = categoryCountForDraft(data, draft);
  const group = element('fieldset', 'fc-chart-settings__group');
  group.appendChild(textNode('legend', axisLabel, 'fc-chart-settings__legend'));
  makeCheckbox(
    group,
    labels.showAxis,
    axis.visible,
    (checked) => updateAxisField(draft, axisName, 'visible', checked),
    !enabled,
  );
  makeCheckbox(
    group,
    labels.reverseOrder,
    axis.reverseOrder,
    (checked) => updateAxisField(draft, axisName, 'reverseOrder', checked),
    !enabled,
  );
  makeCheckbox(
    group,
    labels.logarithmicScale,
    axis.logEnabled,
    (checked) => {
      updateAxisLogEnabled(draft, axisName, checked);
      rerender();
    },
    !enabled || !logEligible,
    'fc-chart-settings__log-scale',
  );
  const logBaseInput = makeTextInput(
    group,
    labels.logBase,
    String(axis.logBaseDraft),
    (value) => updateAxisLogBase(draft, axisName, value),
    'number',
    {
      min: 2,
      max: 1000,
      step: 'any',
      disabled: !enabled || !logEligible || !axis.logEnabled,
    },
  );
  logBaseInput.parentElement?.classList.add('fc-chart-settings__log-base');
  const crossingSelect = makeSelect(
    group,
    axisName === 'xAxis' ? labels.verticalAxisCrosses : labels.horizontalAxisCrosses,
    axis.crossingMode,
    [
      { value: 'auto', label: labels.automatic },
      { value: 'min', label: labels.minimum },
      { value: 'max', label: labels.maximum },
      { value: 'custom', label: labels.custom },
    ],
    (value) => {
      updateAxisCrossingMode(draft, axisName, value as AxisCrossingMode, numeric, categoryCount);
      rerender();
    },
    !enabled,
  );
  // Identify the mode row separately because its options reuse the bound
  // labels "Minimum" and "Maximum".
  crossingSelect.parentElement?.classList.add('fc-chart-settings__crossing-mode');
  makeTextInput(
    group,
    labels.customValue,
    String(normalizedCrossingCustom(axis.crossingCustom, numeric, categoryCount)),
    (value) => updateAxisCrossingCustom(draft, axisName, value, numeric, categoryCount),
    'number',
    {
      allowNegative: numeric,
      min: numeric ? undefined : 1,
      step: numeric ? 'any' : '1',
      disabled: !enabled || axis.crossingMode !== 'custom',
    },
  );
  const labelPositionSelect = makeSelect(
    group,
    labels.labelPosition,
    axis.labelPositionMode,
    [
      { value: 'automatic', label: labels.automatic },
      { value: 'nextTo', label: labels.labelPositionNextTo },
      { value: 'low', label: labels.labelPositionLow },
      { value: 'high', label: labels.labelPositionHigh },
      { value: 'none', label: labels.labelPositionNone },
    ],
    (value) => {
      updateAxisLabelPosition(draft, axisName, value as AxisLabelPositionMode);
    },
    !enabled,
  );
  labelPositionSelect.parentElement?.classList.add('fc-chart-settings__label-position');
  // Typing a title replaces the automatic one; the checkbox returns to it.
  let automatic: HTMLInputElement | undefined;
  makeTextInput(
    group,
    axisName === 'xAxis' ? labels.xAxisTitle : labels.yAxisTitle,
    axis.title,
    (value) => {
      updateAxisField(draft, axisName, 'title', value);
      if (value.trim().length > 0 && draft[axisName].titleAuto) {
        updateAxisField(draft, axisName, 'titleAuto', false);
        if (automatic) automatic.checked = false;
      }
    },
    'text',
    { disabled: !enabled },
  );
  automatic = makeCheckbox(
    group,
    labels.titleAutomatic,
    axis.titleAuto,
    (checked) => updateAxisField(draft, axisName, 'titleAuto', checked),
    !enabled,
  );
  makeTextInput(
    group,
    labels.minimum,
    axis.min === undefined ? '' : String(axis.min),
    (value) => updateAxisField(draft, axisName, 'min', parseOptionalNumber(value)),
    'number',
    {
      allowNegative: true,
      step: 'any',
      placeholder: labels.automatic,
      disabled: !enabled || !numeric,
    },
  );
  makeTextInput(
    group,
    labels.maximum,
    axis.max === undefined ? '' : String(axis.max),
    (value) => updateAxisField(draft, axisName, 'max', parseOptionalNumber(value)),
    'number',
    {
      allowNegative: true,
      step: 'any',
      placeholder: labels.automatic,
      disabled: !enabled || !numeric,
    },
  );
  makeTextInput(
    group,
    labels.majorUnit,
    axis.majorUnit === undefined ? '' : String(axis.majorUnit),
    (value) => {
      const parsed = parseOptionalNumber(value);
      updateAxisField(
        draft,
        axisName,
        'majorUnit',
        parsed !== undefined && parsed > 0 ? parsed : undefined,
      );
    },
    'number',
    {
      min: 0,
      step: 'any',
      placeholder: labels.automatic,
      disabled: !enabled || !numeric || (logEligible && axis.logEnabled),
    },
  );
  makeTextInput(
    group,
    labels.axisNumberFormat,
    axis.numberFormat,
    (value) => updateAxisField(draft, axisName, 'numberFormat', value),
    'text',
    { disabled: !enabled || !numeric },
  );
  makeCheckbox(
    group,
    labels.majorGridlines,
    axis.majorGridlines,
    (checked) => updateAxisField(draft, axisName, 'majorGridlines', checked),
    !enabled || !numeric,
  );
  body.appendChild(group);
};

export const renderElements = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  rerender: () => void,
): void => {
  const general = element('fieldset', 'fc-chart-settings__group');
  general.appendChild(textNode('legend', labels.title, 'fc-chart-settings__legend'));
  // Typing a title replaces the automatic one; the checkbox returns to it.
  let automatic: HTMLInputElement | undefined;
  makeTextInput(general, labels.title, draft.title, (value) => {
    draft.title = value;
    if (value.trim().length > 0 && draft.titleAuto) {
      draft.titleAuto = false;
      if (automatic) automatic.checked = false;
    }
  });
  automatic = makeCheckbox(general, labels.titleAutomatic, draft.titleAuto, (checked) => {
    draft.titleAuto = checked;
  });
  makeSelect(
    general,
    labels.chartType,
    String(draft.kind),
    [
      { value: 'column', label: labels.columnChart },
      { value: 'bar', label: labels.barChart },
      { value: 'line', label: labels.lineChart },
      { value: 'area', label: labels.areaChart },
      { value: 'pie', label: labels.pieChart },
      { value: 'doughnut', label: labels.doughnutChart },
      { value: 'scatter', label: labels.scatterChart },
    ],
    (value) => {
      draft.kind = value as ChartDraft['kind'];
      draft.kindDirty = draft.kind !== draft.kindBaseline;
      draft.kindChanged = draft.kindDirty || draft.seriesDirectionDirty;
      draft.legendKindChanged = draft.kindChanged;
      draft.varyColorsKindChanged = draft.kindChanged;
      rerender();
    },
  );
  if (chart.kind === 'scatter') {
    makeSelect(
      general,
      labels.scatterStyle,
      draft.scatterStyle,
      [
        { value: 'marker', label: labels.scatterStyleMarker },
        { value: 'line', label: labels.scatterStyleLine },
        { value: 'lineMarker', label: labels.scatterStyleLineMarker },
      ],
      (value) => {
        const next = value as ChartScatterStyle;
        draft.scatterStyle = next;
        // Like Excel's subtype change, the style rewrites any explicit marker visibility.
        draft.seriesStyles = draft.seriesStyles.map((style) =>
          style.marker?.visible === undefined
            ? style
            : { ...style, marker: { ...style.marker, visible: next !== 'line' } },
        );
        draft.scatterStyleChanged = next !== draft.scatterStyleBaseline;
        draft.scatterStyleOverride =
          next === draft.scatterStyleBaseline ? draft.scatterStyleOriginalOverride : next;
        rerender();
      },
    );
  }
  makeCheckbox(general, labels.legend, draft.legendVisible, (checked) => {
    draft.legendVisible = checked;
    draft.legendManuallyChanged = true;
    draft.legendOverride =
      checked === (draft.legendOriginalOverride ?? effectiveLegendDefault(chart, draft, data))
        ? draft.legendOriginalOverride
        : checked;
  });
  makeSelect(
    general,
    labels.legendPosition,
    draft.legendPosition,
    [
      { value: 'top', label: labels.positionTop },
      { value: 'right', label: labels.positionRight },
      { value: 'bottom', label: labels.positionBottom },
      { value: 'left', label: labels.positionLeft },
    ],
    (value) => {
      draft.legendPosition = value as ChartDraft['legendPosition'];
    },
  );
  makeSelect(
    general,
    labels.grouping,
    draft.grouping,
    [
      { value: 'clustered', label: labels.groupingClustered },
      { value: 'stacked', label: labels.groupingStacked },
      { value: 'percent-stacked', label: labels.groupingPercentStacked },
    ],
    (value) => {
      draft.grouping = value as ChartGrouping;
      rerender();
    },
    !groupingSupported(chart),
  );
  const direction = makeSelect(
    general,
    labels.seriesDirection,
    draft.seriesDirection,
    [
      { value: 'auto', label: labels.seriesDirectionAuto },
      { value: 'columns', label: labels.seriesDirectionColumns },
      { value: 'rows', label: labels.seriesDirectionRows },
    ],
    (value) => {
      draft.seriesDirection = value as ChartSeriesDirection;
      draft.seriesDirectionDirty = draft.seriesDirection !== draft.seriesDirectionBaseline;
      draft.kindChanged = draft.kindDirty || draft.seriesDirectionDirty;
      draft.legendKindChanged = draft.kindChanged;
      draft.varyColorsKindChanged = draft.kindChanged;
      rerender();
    },
  );
  // A PivotChart's series follow its PivotTable.
  if (chart.pivotChart) projectDisabledState(direction, true, labels.pivotChartLocked);
  makeSelect(
    general,
    labels.emptyCells,
    emptyCellsForcedZero(chart, draft.grouping) ? 'zero' : draft.emptyCells,
    [
      { value: 'gap', label: labels.emptyCellsGap },
      { value: 'zero', label: labels.emptyCellsZero },
      {
        value: 'span',
        label: labels.emptyCellsSpan,
        disabled: emptyCellOptionDisabled(chart, draft.grouping, 'span'),
      },
    ],
    (value) => {
      draft.emptyCells = value as ChartEmptyCellMode;
    },
    emptyCellsForcedZero(chart, draft.grouping),
  );
  body.appendChild(general);

  const axes = element('fieldset', 'fc-chart-settings__group');
  axes.appendChild(textNode('legend', labels.axisTitles, 'fc-chart-settings__legend'));
  renderAxisControls(axes, labels, chart, data, draft, 'xAxis', labels.xAxis, rerender);
  renderAxisControls(axes, labels, chart, data, draft, 'yAxis', labels.yAxis, rerender);
  body.appendChild(axes);

  const dataLabels = element('fieldset', 'fc-chart-settings__group');
  dataLabels.appendChild(textNode('legend', labels.dataLabels, 'fc-chart-settings__legend'));
  const value = draft.dataLabels;
  makeCheckbox(dataLabels, labels.showDataLabels, value.visible, (checked) => {
    updateDataLabel(draft, 'visible', checked);
  });
  makeCheckbox(dataLabels, labels.showSeriesName, value.showSeriesName, (checked) => {
    updateDataLabel(draft, 'showSeriesName', checked);
  });
  makeCheckbox(dataLabels, labels.showCategoryName, value.showCategoryName, (checked) => {
    updateDataLabel(draft, 'showCategoryName', checked);
  });
  makeCheckbox(dataLabels, labels.showValue, value.showValue, (checked) => {
    updateDataLabel(draft, 'showValue', checked);
  });
  makeCheckbox(
    dataLabels,
    labels.showPercentage,
    value.showPercentage,
    (checked) => {
      updateDataLabel(draft, 'showPercentage', checked);
    },
    !supportsPercentage(chart, draft.grouping),
  );
  makeCheckbox(dataLabels, labels.showLegendKey, value.showLegendKey, (checked) => {
    updateDataLabel(draft, 'showLegendKey', checked);
  });
  makeTextInput(dataLabels, labels.labelSeparator, value.separator, (next) => {
    updateDataLabel(draft, 'separator', next);
  });
  makeSelect(
    dataLabels,
    labels.labelPosition,
    value.position,
    allowedPositions({ ...chart, grouping: draft.grouping }, value.position).map((item) => ({
      value: item.value,
      label: labels[item.label],
      disabled: item.disabled,
    })),
    (next) => {
      updateDataLabel(draft, 'position', next as MutableDataLabels['position']);
    },
    chart.kind === 'area' || isDoughnutKind(chart.kind),
  );
  makeTextInput(dataLabels, labels.numberFormat, value.numberFormat, (next) => {
    updateDataLabel(draft, 'numberFormat', next);
  });
  makeColorField(dataLabels, labels.textColor, value.color, (next) => {
    updateDataLabel(draft, 'color', next);
  });
  makeTextInput(
    dataLabels,
    labels.fontSize,
    String(value.fontSize),
    (next) => {
      const size = Number(next);
      updateDataLabel(
        draft,
        'fontSize',
        Number.isFinite(size) ? Math.max(6, Math.min(72, size)) : 10,
      );
    },
    'number',
  );
  body.appendChild(dataLabels);
};
