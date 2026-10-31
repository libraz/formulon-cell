/** Styles tab of the chart settings dialog: palettes, series and point styling. */

import {
  automaticPalette,
  CHART_PALETTE_PRESETS,
  DEFAULT_THEME_ACCENTS,
} from '../../charts/series-colors.js';
import { resolveChartStyle } from '../../charts/styles.js';
import type { ChartData, ChartKind, ChartLineSpec, ChartMarkerSpec } from '../../charts/types.js';
import { parseRangeRef } from '../../engine/range-resolver.js';
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
  type ChartDraft,
  defaultVaryColors,
  isCircularKind,
  isDoughnutKind,
  patchFill,
  patchLine,
  releaseSeriesColors,
  type SessionChartSeriesRefs,
  type SessionChartSettingsShape,
  seriesEntries,
  styleForSeries,
  stylePointEntries,
  styleTarget,
  updateSeriesStyle,
} from '../session-chart-settings-model.js';

const ALL_POINTS = '__all__';
/** Fixed presets; `office` is the theme's accents, the default theme's here. */
export const PALETTE_PRESETS: Record<string, readonly string[]> = CHART_PALETTE_PRESETS;

const PALETTE_LABEL_KEYS: Record<string, keyof SessionChartSettingsLabels> = {
  office: 'paletteOffice',
  blue: 'paletteBlue',
  mono: 'paletteMonochrome',
};

const clampPercent = (value: string): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? Math.max(0, Math.min(100, parsed)) : 0;
};

const pointIdForData = (data: ChartData, seriesIndex: number, pointIndex: number): string => {
  if (data.domain === 'category') {
    return data.categoryIds?.[pointIndex] ?? `category-${pointIndex}`;
  }
  return data.series[seriesIndex]?.points[pointIndex]?.id ?? `point-${pointIndex}`;
};

const resolvedStyleForSelection = (
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
): ReturnType<typeof resolveChartStyle> => {
  const seriesIndex = Math.max(
    0,
    data.series.findIndex((series) => series.id === seriesId),
  );
  const selectedSeries = data.series[seriesIndex];
  let pointIndex: number | undefined;
  if (pointId !== null && selectedSeries) {
    const index =
      data.domain === 'category'
        ? data.categories.findIndex(
            (_, index) => pointIdForData(data, seriesIndex, index) === pointId,
          )
        : (data.series[seriesIndex]?.points.findIndex(
            (_point, index) => pointIdForData(data, seriesIndex, index) === pointId,
          ) ?? -1);
    pointIndex = index < 0 ? undefined : index;
  }
  return resolveChartStyle(
    {
      ...chart,
      scatterStyle: draft.scatterStyle,
      palette: draft.palette,
      seriesStyles: draft.seriesStyles,
      varyColors: draft.varyColors,
    },
    draft.palette,
    seriesId,
    pointId ?? undefined,
    seriesIndex,
    pointIndex,
  );
};

const renderColorAndLineControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
): void => {
  const style = styleForSeries(draft, seriesId);
  const point =
    pointId === null ? undefined : style.points?.find((entry) => entry.pointId === pointId);
  const resolved = resolvedStyleForSelection(chart, data, draft, seriesId, pointId);
  let currentFill = { ...(pointId === null ? (style.fill ?? {}) : (point?.fill ?? {})) };
  let currentLine = { ...(pointId === null ? (style.line ?? {}) : (point?.line ?? {})) };
  // Excel keeps no fill on a line or scatter series or point; its colour is the line's and markers'.
  if (chart.kind !== 'line' && chart.kind !== 'scatter') {
    makeColorField(body, labels.fillColor, resolved.fill?.color ?? resolved.color, (color) => {
      currentFill = patchFill(currentFill, { color });
      styleTarget(draft, seriesId, pointId, { fill: currentFill });
    });
    makeTextInput(
      body,
      `${labels.fillColor} ${labels.transparency}`,
      String(resolved.fill?.transparency ?? 0),
      (value) => {
        // A workbook fill carries its colour with its transparency, so the inherited colour is kept.
        currentFill = patchFill(currentFill, {
          color: currentFill.color ?? resolved.fill?.color ?? resolved.color,
          transparency: clampPercent(value),
        });
        styleTarget(draft, seriesId, pointId, { fill: currentFill });
      },
      'number',
    );
  }
  makeColorField(body, labels.lineColor, resolved.line?.color ?? resolved.color, (color) => {
    currentLine = patchLine(currentLine, { color });
    styleTarget(draft, seriesId, pointId, { line: currentLine });
  });
  makeTextInput(
    body,
    `${labels.lineColor} ${labels.transparency}`,
    String(resolved.line?.transparency ?? 0),
    (value) => {
      currentLine = patchLine(currentLine, {
        color: currentLine.color ?? resolved.line?.color ?? resolved.color,
        transparency: clampPercent(value),
      });
      styleTarget(draft, seriesId, pointId, { line: currentLine });
    },
    'number',
  );
  makeTextInput(
    body,
    labels.lineWidth,
    String(resolved.line?.width ?? 1.5),
    (value) => {
      const width = Number(value);
      currentLine = patchLine(currentLine, {
        width: Number.isFinite(width) ? Math.max(0, Math.min(12, width)) : 1.5,
      });
      styleTarget(draft, seriesId, pointId, {
        line: currentLine,
      });
    },
    'number',
  );
  makeSelect(
    body,
    labels.lineDash,
    resolved.line?.dash ?? 'solid',
    [
      { value: 'solid', label: labels.dashSolid },
      { value: 'dash', label: labels.dashDash },
      { value: 'dot', label: labels.dashDot },
      { value: 'dash-dot', label: labels.dashDashDot },
    ],
    (value) => {
      currentLine = patchLine(currentLine, { dash: value as ChartLineSpec['dash'] });
      styleTarget(draft, seriesId, pointId, { line: currentLine });
    },
  );
  // Smoothing belongs to a series' line, so it shows where a line or scatter series draws one.
  const drawsLine =
    chart.kind === 'line' ||
    (chart.kind === 'scatter' && (chart.scatterStyle ?? 'marker') !== 'marker');
  if (pointId === null && drawsLine) {
    makeCheckbox(body, labels.smoothLine, style.smooth === true, (checked) => {
      updateSeriesStyle(draft, seriesId, { smooth: checked });
    });
  }
};

const renderMarkerControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  seriesId: string,
  pointId: string | null,
): void => {
  const style = styleForSeries(draft, seriesId);
  const point =
    pointId === null ? undefined : style.points?.find((entry) => entry.pointId === pointId);
  const resolved = resolvedStyleForSelection(chart, data, draft, seriesId, pointId);
  let marker = { ...(pointId === null ? (style.marker ?? {}) : (point?.marker ?? {})) };
  const patchMarker = (patch: Partial<ChartMarkerSpec>): void => {
    marker = { ...marker, ...patch };
    styleTarget(draft, seriesId, pointId, { marker });
  };
  makeCheckbox(
    body,
    `${labels.marker} ${labels.markerVisible}`,
    resolved.marker.visible !== false,
    (checked) => {
      patchMarker({ visible: checked });
    },
  );
  makeSelect(
    body,
    labels.markerShape,
    resolved.marker.shape ?? 'circle',
    [
      { value: 'circle', label: labels.shapeCircle },
      { value: 'square', label: labels.shapeSquare },
      { value: 'diamond', label: labels.shapeDiamond },
      { value: 'triangle', label: labels.shapeTriangle },
      { value: 'x', label: labels.shapeX },
      { value: 'star', label: labels.shapeStar },
      { value: 'plus', label: labels.shapePlus },
      { value: 'dot', label: labels.shapeDot },
      { value: 'dash', label: labels.shapeDash },
    ],
    (value) => patchMarker({ shape: value as ChartMarkerSpec['shape'] }),
  );
  makeTextInput(
    body,
    labels.markerSize,
    String(resolved.marker.size ?? 5),
    (value) => {
      const size = Number(value);
      patchMarker({ size: Number.isFinite(size) ? Math.max(2, Math.min(24, size)) : 6 });
    },
    'number',
  );
  makeColorField(
    body,
    `${labels.marker} ${labels.fillColor}`,
    resolved.marker.fill?.color ?? resolved.fill?.color ?? resolved.color,
    (color) => {
      patchMarker({ fill: patchFill(marker.fill, { color }) });
    },
  );
  makeColorField(
    body,
    `${labels.marker} ${labels.lineColor}`,
    resolved.marker.line?.color ?? resolved.line?.color ?? resolved.color,
    (color) => {
      patchMarker({ line: patchLine(marker.line, { color }) });
    },
  );
};

const renderChartGeometryControls = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  draft: ChartDraft,
): void => {
  if (chart.kind === 'bar' || chart.kind === 'column') {
    const geometry = element('fieldset', 'fc-chart-settings__group');
    geometry.appendChild(textNode('legend', labels.gapWidth, 'fc-chart-settings__legend'));
    makeTextInput(
      geometry,
      labels.gapWidth,
      String(draft.gapWidth),
      (value) => {
        const parsed = Number(value);
        draft.gapWidth = Number.isFinite(parsed) ? Math.max(0, Math.min(500, parsed)) : 150;
      },
      'number',
      { min: 0, max: 500, step: '1' },
    );
    makeTextInput(
      geometry,
      labels.overlap,
      String(draft.overlap),
      (value) => {
        const parsed = Number(value);
        draft.overlap = Number.isFinite(parsed) ? Math.max(-100, Math.min(100, parsed)) : 0;
      },
      'number',
      {
        allowNegative: true,
        min: -100,
        max: 100,
        step: '1',
        disabled: draft.grouping !== 'clustered',
      },
    );
    body.appendChild(geometry);
  }
  if (isCircularKind(chart.kind)) {
    const geometry = element('fieldset', 'fc-chart-settings__group');
    geometry.appendChild(textNode('legend', labels.firstSliceAngle, 'fc-chart-settings__legend'));
    makeTextInput(
      geometry,
      labels.firstSliceAngle,
      String(draft.firstSliceAngle),
      (value) => {
        const parsed = Number(value);
        draft.firstSliceAngle = Number.isFinite(parsed) ? Math.max(0, Math.min(360, parsed)) : 0;
      },
      'number',
      { min: 0, max: 360, step: '1' },
    );
    body.appendChild(geometry);
  }
  if (isDoughnutKind(chart.kind)) {
    const geometry = element('fieldset', 'fc-chart-settings__group');
    geometry.appendChild(textNode('legend', labels.holeSize, 'fc-chart-settings__legend'));
    makeTextInput(
      geometry,
      labels.holeSize,
      String(draft.holeSize),
      (value) => {
        const parsed = Number(value);
        draft.holeSize = Number.isFinite(parsed) ? Math.max(10, Math.min(90, parsed)) : 50;
        draft.holeSizeChanged = draft.holeSize !== draft.holeSizeBaseline;
      },
      'number',
      { min: 10, max: 90, step: '1' },
    );
    body.appendChild(geometry);
  }
};

type NewSeriesField = keyof SessionChartSeriesRefs;

/** The add-series form as typed, kept across re-renders. */
export interface NewSeriesDraft {
  name: string;
  values: string;
  categories: string;
  x: string;
  invalid: NewSeriesField[];
}

export interface StylesState {
  seriesId: string;
  pointId: string | null;
  newSeries: NewSeriesDraft | null;
}

/** Series add and remove, which change the workbook chart at once rather than the draft. */
export interface SeriesActions {
  /** The workbook chart's kind, which decides the references a new series takes. */
  readonly kind: ChartKind;
  /** Why both actions are unavailable; null when they are available. */
  readonly locked: string | null;
  readonly disabled: boolean;
  add(refs: SessionChartSeriesRefs): string | null;
  remove(seriesId: string): boolean;
}

/** The typed references, or the fields that are not references. Only `values` is required. */
const newSeriesRefs = (
  form: NewSeriesDraft,
  scatter: boolean,
): SessionChartSeriesRefs | NewSeriesField[] => {
  const fields: readonly NewSeriesField[] = ['name', 'values', scatter ? 'x' : 'categories'];
  const invalid = fields.filter((field) => {
    const text = form[field].trim();
    return text ? parseRangeRef(text) === null : field === 'values';
  });
  if (invalid.length > 0) return invalid;
  const refs: SessionChartSeriesRefs = { values: form.values.trim() };
  for (const field of fields) {
    const text = form[field].trim();
    if (field !== 'values' && text) refs[field] = text;
  }
  return refs;
};

const renderSeriesActions = (
  target: HTMLElement,
  labels: SessionChartSettingsLabels,
  seriesCount: number,
  state: StylesState,
  rerender: () => void,
  actions: SeriesActions,
): void => {
  const row = element('div', 'fc-chart-settings__series-actions');
  const add = element('button', 'fc-chart-settings__add-series');
  add.type = 'button';
  add.textContent = labels.addSeries;
  add.setAttribute('aria-expanded', String(state.newSeries !== null));
  add.addEventListener('click', () => {
    state.newSeries = state.newSeries
      ? null
      : { name: '', values: '', categories: '', x: '', invalid: [] };
    rerender();
  });
  const remove = element('button', 'fc-chart-settings__remove-series');
  remove.type = 'button';
  remove.textContent = labels.removeSeries;
  remove.addEventListener('click', () => {
    if (actions.remove(state.seriesId)) rerender();
  });
  row.append(add, remove);
  target.appendChild(row);
  const unavailable = actions.locked !== null || actions.disabled;
  projectDisabledState(add, unavailable, actions.locked);
  projectDisabledState(
    remove,
    unavailable || seriesCount < 2,
    actions.locked ?? (seriesCount < 2 ? labels.lastSeriesLocked : null),
  );
  const form = state.newSeries;
  if (!form || unavailable) return;

  const scatter = actions.kind === 'scatter';
  const group = element('div', 'fc-chart-settings__new-series');
  const fields: readonly [NewSeriesField, string][] = [
    ['name', labels.seriesNameRef],
    ['values', scatter ? labels.seriesYValuesRef : labels.seriesValuesRef],
    scatter ? ['x', labels.seriesXValuesRef] : ['categories', labels.axisLabelsRef],
  ];
  for (const [field, label] of fields) {
    const input = makeTextInput(group, label, form[field], (value) => {
      form[field] = value;
    });
    input.dataset.seriesRef = field;
    if (form.invalid.includes(field)) input.setAttribute('aria-invalid', 'true');
  }
  if (form.invalid.length > 0) {
    const error = textNode('div', labels.invalidReference, 'fc-chart-settings__error');
    error.setAttribute('role', 'alert');
    group.appendChild(error);
  }
  const buttons = element('div', 'fc-chart-settings__new-series-actions');
  const confirm = element('button', 'fc-chart-settings__new-series-confirm');
  confirm.type = 'button';
  confirm.textContent = labels.addSeriesConfirm;
  confirm.addEventListener('click', () => {
    const refs = newSeriesRefs(form, scatter);
    if (Array.isArray(refs)) {
      form.invalid = refs;
      rerender();
      return;
    }
    const added = actions.add(refs);
    if (added === null) {
      // The engine resolves what the syntax check cannot, such as an unknown sheet.
      form.invalid = ['values'];
    } else {
      state.seriesId = added;
      state.pointId = null;
      state.newSeries = null;
    }
    rerender();
  });
  const dismiss = element('button', 'fc-chart-settings__new-series-cancel');
  dismiss.type = 'button';
  dismiss.textContent = labels.cancel;
  dismiss.addEventListener('click', () => {
    state.newSeries = null;
    rerender();
  });
  buttons.append(confirm, dismiss);
  group.appendChild(buttons);
  target.appendChild(group);
};

export const renderStyles = (
  body: HTMLElement,
  labels: SessionChartSettingsLabels,
  chart: SessionChartSettingsShape,
  data: ChartData,
  draft: ChartDraft,
  state: StylesState,
  rerender: () => void,
  accents: readonly string[] | undefined,
  seriesActions: SeriesActions,
): void => {
  const automatic = automaticPalette(accents);
  const presets = { ...PALETTE_PRESETS, office: automatic.slice(0, 6) };
  const palette = element('fieldset', 'fc-chart-settings__group');
  palette.appendChild(textNode('legend', labels.palette, 'fc-chart-settings__legend'));
  const presetRow = element('div', 'fc-chart-settings__palette-presets');
  for (const [id, colors] of Object.entries(presets)) {
    const button = element('button');
    button.type = 'button';
    button.className = 'fc-chart-settings__palette-button';
    button.dataset.palette = id;
    const paletteLabel = PALETTE_LABEL_KEYS[id];
    button.setAttribute('aria-label', paletteLabel ? labels[paletteLabel] : labels.palette);
    const swatches = element('span', 'fc-chart-settings__palette-swatches');
    for (const color of colors) {
      const swatch = element('span', 'fc-chart-settings__palette-swatch');
      swatch.style.backgroundColor = color;
      swatches.appendChild(swatch);
    }
    button.appendChild(swatches);
    button.addEventListener('click', () => {
      draft.palette = [...colors];
      releaseSeriesColors(draft);
      rerender();
    });
    presetRow.appendChild(button);
  }
  palette.appendChild(presetRow);
  for (let index = 0; index < Math.max(6, draft.palette.length); index += 1) {
    const color =
      draft.palette[index] ?? automatic[index % automatic.length] ?? DEFAULT_THEME_ACCENTS[0] ?? '';
    makeColorField(palette, `${labels.palette} ${index + 1}`, color, (next) => {
      draft.palette[index] = next;
      releaseSeriesColors(draft);
    });
  }
  makeCheckbox(palette, labels.varyColors, draft.varyColors, (checked) => {
    draft.varyColors = checked;
    draft.varyColorsManuallyChanged = true;
    draft.varyColorsOverride =
      checked === (draft.varyColorsOriginalOverride ?? defaultVaryColors(String(draft.kind)))
        ? draft.varyColorsOriginalOverride
        : checked;
  });
  body.appendChild(palette);

  const target = element('fieldset', 'fc-chart-settings__group');
  target.appendChild(textNode('legend', labels.series, 'fc-chart-settings__legend'));
  const entries = seriesEntries(data, labels);
  const selectedSeries = entries.some((entry) => entry.id === state.seriesId)
    ? state.seriesId
    : (entries[0]?.id ?? 'series-0');
  state.seriesId = selectedSeries;
  const points = stylePointEntries(data, selectedSeries, labels);
  if (state.pointId !== null && !points.some((point) => point.id === state.pointId)) {
    state.pointId = null;
  }
  makeSelect(
    target,
    labels.series,
    selectedSeries,
    entries.map((entry) => ({ value: entry.id, label: entry.label })),
    (value) => {
      state.seriesId = value;
      state.pointId = null;
      rerender();
    },
  );
  renderSeriesActions(target, labels, entries.length, state, rerender, seriesActions);
  makeSelect(
    target,
    labels.point,
    state.pointId ?? ALL_POINTS,
    [
      { value: ALL_POINTS, label: labels.allPoints },
      ...points.map((point) => ({ value: point.id, label: point.label })),
    ],
    (value) => {
      state.pointId = value === ALL_POINTS ? null : value;
      rerender();
    },
  );
  renderColorAndLineControls(target, labels, chart, data, draft, selectedSeries, state.pointId);
  renderMarkerControls(target, labels, chart, data, draft, selectedSeries, state.pointId);
  body.appendChild(target);
  renderChartGeometryControls(body, labels, chart, draft);
};
