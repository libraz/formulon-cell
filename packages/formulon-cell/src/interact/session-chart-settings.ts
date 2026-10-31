import { automaticPalette } from '../charts/series-colors.js';
import type { ChartData, ChartValueFormatter as CoreChartValueFormatter } from '../charts/types.js';
import type { History } from '../commands/history.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  captureSettingsView,
  element,
  restoreSettingsView,
  textNode,
} from './session-chart-settings-dom.js';
import {
  DEFAULT_LABELS,
  type SessionChartSettingsLabels,
} from './session-chart-settings-labels.js';
import {
  type ChartDraft,
  draftFromChart,
  effectiveDraftChart,
  reconcileSeriesDraft,
  refreshDraftDefaults,
  type SessionChartSeriesRefs,
  type SessionChartSettingsPatch,
  type SessionChartSettingsShape,
  seriesEntries,
  stylePointEntries,
  syncUntouchedOptionalOverrides,
  toPatch,
} from './session-chart-settings-model.js';
import { renderElements } from './session-chart-settings-tabs/elements.js';
import { renderFilters } from './session-chart-settings-tabs/filters.js';
import {
  type NewSeriesDraft,
  PALETTE_PRESETS,
  renderStyles,
  type SeriesActions,
  type StylesState,
} from './session-chart-settings-tabs/styles.js';

/** The host may use this hook to format values according to its workbook
 * profile. The renderer keeps the hook optional so the chart package remains
 * usable without a workbook engine. */
export type ChartValueFormatter = CoreChartValueFormatter;

export interface SessionChartSettingsHandle {
  open(
    chart: SessionChartSettingsShape,
    data: ChartData,
    anchor: HTMLElement,
    tab?: SettingsTab,
  ): void;
  sync(chart: SessionChartSettingsShape, data: ChartData, anchor: HTMLElement): void;
  close(discard?: boolean): void;
  setLabels(next: Partial<SessionChartSettingsLabels>, locale?: string): void;
  detach(): void;
  isOpen(): boolean;
  chartId(): string | null;
}

export type SettingsTab = 'elements' | 'styles' | 'filters';

export interface SessionChartSettingsDeps {
  readonly host: HTMLElement;
  readonly store: SpreadsheetStore;
  readonly history?: History | null;
  readonly locale?: string;
  readonly formatValue?: ChartValueFormatter;
  readonly labels?: Partial<SessionChartSettingsLabels>;
  readonly canEdit?: (chart: SessionChartSettingsShape) => boolean;
  /** Rebuild source data when a draft changes chart family or orientation. */
  readonly readData?: (chart: SessionChartSettingsShape) => ChartData;
  readonly onApply: (id: string, patch: SessionChartSettingsPatch, history: History | null) => void;
  /** Append a series to the workbook chart; its id, or null when refused. */
  readonly onAddSeries: (
    id: string,
    refs: SessionChartSeriesRefs,
    history: History | null,
  ) => string | null;
  /** Remove a series from the workbook chart; false when refused. */
  readonly onRemoveSeries: (id: string, seriesId: string, history: History | null) => boolean;
  /** The workbook theme's accent1..accent6; the default theme's when omitted. */
  readonly accents?: () => readonly string[] | undefined;
}

const labelsWithDefaults = (
  labels?: Partial<SessionChartSettingsLabels>,
): SessionChartSettingsLabels => ({ ...DEFAULT_LABELS, ...labels });

const dataSignature = (data: ChartData): string =>
  JSON.stringify({
    domain: data.domain,
    series:
      data.domain === 'xy'
        ? data.series.map((series) => ({
            id: series.id,
            name: series.name,
            points: series.points.map((point) => point.id ?? ''),
          }))
        : data.series.map((series) => ({ id: series.id, name: series.name })),
    categories:
      data.domain === 'category'
        ? data.categories.map((category, index) => ({
            id: data.categoryIds?.[index] ?? `category-${index}`,
            name: category,
          }))
        : undefined,
  });

/**
 * Create the non-modal settings surface used by session chart overlays. The
 * surface is intentionally independent from the chart root: chart redraws can
 * replace SVG panels without destroying an open form or its draft.
 */
export function createSessionChartSettings(
  deps: SessionChartSettingsDeps,
): SessionChartSettingsHandle {
  const { host, history = null } = deps;
  let labels = labelsWithDefaults(deps.labels);
  let locale = deps.locale ?? 'en-US';
  let chart: SessionChartSettingsShape | null = null;
  let data: ChartData | null = null;
  let anchor: HTMLElement | null = null;
  let draft: ChartDraft | null = null;
  let activeTab: SettingsTab = 'elements';
  let styleSeriesId = 'series-0';
  let stylePointId: string | null = null;
  let newSeries: NewSeriesDraft | null = null;
  let dataKey = '';
  let editable = true;
  let disposed = false;

  const surface = element('div', 'fc-chart-settings');
  surface.hidden = true;
  surface.setAttribute('role', 'dialog');
  surface.setAttribute('aria-modal', 'false');
  surface.tabIndex = -1;
  host.appendChild(surface);

  const removeEscape = (): void => {
    document.removeEventListener('keydown', onDocumentKeydown);
  };

  const closeInternal = (returnFocus: boolean): void => {
    if (surface.hidden) return;
    const returnTarget = anchor;
    surface.hidden = true;
    removeEscape();
    chart = null;
    data = null;
    anchor = null;
    draft = null;
    newSeries = null;
    dataKey = '';
    if (returnFocus && returnTarget?.isConnected) returnTarget.focus({ preventScroll: true });
  };

  const reposition = (): void => {
    if (!anchor) return;
    const hostRect = host.getBoundingClientRect();
    const anchorRect = anchor.getBoundingClientRect();
    const hostScrollLeft = host.scrollLeft;
    const hostScrollTop = host.scrollTop;
    const hostWidth = host.clientWidth || hostRect.width || 320;
    const hostHeight = host.clientHeight || hostRect.height || 520;
    const surfaceWidth = surface.offsetWidth || 320;
    const surfaceHeight = surface.offsetHeight || 360;
    const anchorWidth = anchorRect.width || anchor.offsetWidth || 0;
    const anchorHeight = anchorRect.height || anchor.offsetHeight || 0;
    const hasAnchorRect =
      anchorRect.width !== 0 ||
      anchorRect.height !== 0 ||
      anchorRect.left !== 0 ||
      anchorRect.top !== 0 ||
      anchorRect.right !== 0 ||
      anchorRect.bottom !== 0;
    const anchorLeft = hasAnchorRect
      ? anchorRect.left - hostRect.left + hostScrollLeft
      : anchor.offsetLeft || 0;
    const anchorTop = hasAnchorRect
      ? anchorRect.top - hostRect.top + hostScrollTop
      : anchor.offsetTop || 0;
    const anchorRight = anchorLeft + anchorWidth;
    const anchorBottom = anchorTop + anchorHeight;
    const minLeft = hostScrollLeft + 4;
    const maxLeft = Math.max(minLeft, hostScrollLeft + hostWidth - surfaceWidth - 4);
    const minTop = hostScrollTop + 4;
    const maxTop = Math.max(minTop, hostScrollTop + hostHeight - surfaceHeight - 4);
    let left = anchorRight + 8;
    if (left > maxLeft) left = anchorLeft - surfaceWidth - 8;
    let top = anchorTop;
    if (top > maxTop) top = anchorBottom - surfaceHeight;
    surface.style.left = `${Math.max(minLeft, Math.min(left, maxLeft))}px`;
    surface.style.top = `${Math.max(minTop, Math.min(top, maxTop))}px`;
  };

  const refreshDraftData = (): void => {
    if (!chart || !draft) return;
    if (deps.readData) data = deps.readData(effectiveDraftChart(chart, draft));
    if (data) refreshDraftDefaults(draft, chart, data);
  };

  const reconcileStyleSelection = (): void => {
    if (!data) return;
    const entries = seriesEntries(data);
    if (!entries.some((entry) => entry.id === styleSeriesId)) {
      styleSeriesId = entries[0]?.id ?? 'series-0';
    }
    if (stylePointId !== null) {
      const points = stylePointEntries(data, styleSeriesId);
      if (!points.some((point) => point.id === stylePointId)) stylePointId = null;
    }
  };

  const renderPanel = (): void => {
    if (!chart || !data || !draft || disposed) return;
    refreshDraftDefaults(draft, chart, data);
    surface.replaceChildren();
    surface.setAttribute('aria-label', labels.settings);
    const heading = element('div', 'fc-chart-settings__header');
    const headingText = textNode('h3', labels.settings, 'fc-chart-settings__title');
    const close = element('button');
    close.type = 'button';
    close.className = 'fc-chart-settings__close';
    close.setAttribute('aria-label', labels.closeSettings);
    close.textContent = '×';
    close.addEventListener('click', () => closeInternal(true));
    heading.append(headingText, close);

    const tabs = element('div', 'fc-chart-settings__tabs');
    tabs.setAttribute('role', 'tablist');
    const tabDefinitions: readonly { id: SettingsTab; label: string }[] = [
      { id: 'elements', label: labels.elementsTab },
      { id: 'styles', label: labels.stylesTab },
      { id: 'filters', label: labels.filtersTab },
    ];
    for (const [index, item] of tabDefinitions.entries()) {
      const button = element('button');
      button.type = 'button';
      button.className = 'fc-chart-settings__tab';
      button.id = `fc-chart-settings-tab-${item.id}`;
      button.setAttribute('role', 'tab');
      button.setAttribute('aria-selected', String(activeTab === item.id));
      button.tabIndex = activeTab === item.id ? 0 : -1;
      button.textContent = item.label;
      button.addEventListener('click', () => {
        activeTab = item.id;
        renderPanel();
        surface.querySelector<HTMLButtonElement>(`#fc-chart-settings-tab-${item.id}`)?.focus();
      });
      button.addEventListener('keydown', (event) => {
        if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
        event.preventDefault();
        const nextIndex =
          (index + (event.key === 'ArrowRight' ? 1 : -1) + tabDefinitions.length) %
          tabDefinitions.length;
        const next = tabDefinitions[nextIndex];
        if (!next) return;
        activeTab = next.id;
        renderPanel();
        surface.querySelector<HTMLButtonElement>(`#fc-chart-settings-tab-${next.id}`)?.focus();
      });
      tabs.appendChild(button);
    }

    const body = element('div', 'fc-chart-settings__body');
    const renderAgain = (): void => {
      const view = captureSettingsView(surface);
      refreshDraftData();
      reconcileStyleSelection();
      renderPanel();
      restoreSettingsView(surface, view);
    };
    const effective = effectiveDraftChart(chart, draft);
    if (activeTab === 'elements') renderElements(body, labels, effective, data, draft, renderAgain);
    else if (activeTab === 'styles') {
      const state: StylesState = { seriesId: styleSeriesId, pointId: stylePointId, newSeries };
      const committed = chart;
      // The series of a draft that re-splits them are not the workbook's yet.
      const seriesActions: SeriesActions = {
        kind: committed.kind,
        locked: committed.pivotChart ? labels.pivotChartLocked : null,
        disabled: draft.kindChanged,
        add: (refs) => {
          const added = deps.onAddSeries(committed.id, refs, history);
          if (added !== null && chart && draft) reconcileSeriesDraft(draft, chart);
          return added;
        },
        remove: (seriesId) => {
          const removed = deps.onRemoveSeries(committed.id, seriesId, history);
          if (removed && chart && draft) reconcileSeriesDraft(draft, chart);
          return removed;
        },
      };
      renderStyles(
        body,
        labels,
        effective,
        data,
        draft,
        state,
        () => {
          styleSeriesId = state.seriesId;
          stylePointId = state.pointId;
          newSeries = state.newSeries;
          renderAgain();
        },
        deps.accents?.(),
        seriesActions,
      );
    } else {
      renderFilters(body, labels, data, draft);
      if (chart.pivotChart) {
        for (const control of Array.from(
          body.querySelectorAll<HTMLInputElement | HTMLButtonElement>('input, button'),
        )) {
          projectDisabledState(control, true, labels.pivotChartLocked);
        }
      }
    }

    const footer = element('div', 'fc-chart-settings__footer');
    const cancel = element('button');
    cancel.type = 'button';
    cancel.className = 'fc-chart-settings__cancel';
    cancel.textContent = labels.cancel;
    cancel.addEventListener('click', () => closeInternal(true));
    const apply = element('button');
    apply.type = 'button';
    apply.className = 'fc-chart-settings__apply';
    apply.textContent = labels.apply;
    apply.addEventListener('click', () => {
      if (!chart || !draft) return;
      const id = chart.id;
      const patch = toPatch(draft, chart);
      // The callback is deliberately invoked once per Apply click. The
      // command owns protection and history; this form only owns the draft.
      if (Object.keys(patch).length > 0) deps.onApply(id, patch, history);
      closeInternal(true);
    });
    footer.append(cancel, apply);
    surface.append(heading, tabs, body, footer);
    if (!editable) {
      for (const control of Array.from(
        body.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLButtonElement>(
          'input, select, button',
        ),
      )) {
        projectDisabledState(control, true, null);
      }
      projectDisabledState(apply, true, null);
    }
    reposition();
  };

  function onDocumentKeydown(event: KeyboardEvent): void {
    if (surface.hidden || event.key !== 'Escape') return;
    event.preventDefault();
    closeInternal(true);
  }

  const open = (
    nextChart: SessionChartSettingsShape,
    nextData: ChartData,
    nextAnchor: HTMLElement,
    tab?: SettingsTab,
  ): void => {
    if (disposed) return;
    chart = nextChart;
    data = nextData;
    anchor = nextAnchor;
    draft = draftFromChart(nextChart, nextData, automaticPalette(deps.accents?.()));
    activeTab = tab ?? 'elements';
    styleSeriesId = seriesEntries(nextData)[0]?.id ?? 'series-0';
    stylePointId = null;
    newSeries = null;
    dataKey = dataSignature(nextData);
    editable = deps.canEdit?.(nextChart) !== false;
    surface.hidden = false;
    surface.lang = locale;
    document.removeEventListener('keydown', onDocumentKeydown);
    document.addEventListener('keydown', onDocumentKeydown);
    renderPanel();
    surface
      .querySelector<HTMLElement>('[role="tab"][aria-selected="true"]')
      ?.focus({ preventScroll: true });
  };

  return {
    open,
    sync(nextChart, nextData, nextAnchor) {
      if (surface.hidden || !chart || chart.id !== nextChart.id || !draft) return;
      const previousKind = draft.kind;
      const previousSeriesDirection = draft.seriesDirection;
      const previousLegendVisible = draft.legendVisible;
      const previousVaryColors = draft.varyColors;
      const previousScatterStyle = draft.scatterStyle;
      const previousXReverseOrder = draft.xAxis.reverseOrder;
      const previousYReverseOrder = draft.yAxis.reverseOrder;
      const previousXCrossing = draft.xAxis.crossing;
      const previousYCrossing = draft.yAxis.crossing;
      const previousXCrossingMode = draft.xAxis.crossingMode;
      const previousYCrossingMode = draft.yAxis.crossingMode;
      const previousXCrossingCustom = draft.xAxis.crossingCustom;
      const previousYCrossingCustom = draft.yAxis.crossingCustom;
      const previousXLabelPosition = draft.xAxis.labelPosition;
      const previousYLabelPosition = draft.yAxis.labelPosition;
      const previousXLabelPositionMode = draft.xAxis.labelPositionMode;
      const previousYLabelPositionMode = draft.yAxis.labelPositionMode;
      const previousXLogBase = draft.xAxis.logBase;
      const previousYLogBase = draft.yAxis.logBase;
      const previousXLogEnabled = draft.xAxis.logEnabled;
      const previousYLogEnabled = draft.yAxis.logEnabled;
      const previousXLogBaseDraft = draft.xAxis.logBaseDraft;
      const previousYLogBaseDraft = draft.yAxis.logBaseDraft;
      chart = nextChart;
      if (!draft.kindDirty) {
        draft.kind = nextChart.kind;
        draft.kindBaseline = nextChart.kind;
      }
      if (!draft.seriesDirectionDirty) {
        draft.seriesDirection = nextChart.seriesDirection ?? 'auto';
        draft.seriesDirectionBaseline = draft.seriesDirection;
      }
      if (deps.readData) data = deps.readData(effectiveDraftChart(nextChart, draft));
      else data = nextData;
      syncUntouchedOptionalOverrides(draft, chart, data);
      refreshDraftDefaults(draft, chart, data);
      reconcileStyleSelection();
      const nextEditable = deps.canEdit?.(nextChart) !== false;
      if (!anchor?.isConnected) {
        anchor =
          nextAnchor.querySelector<HTMLButtonElement>(`[data-chart-settings-tab="${activeTab}"]`) ??
          nextAnchor;
      }
      const nextKey = dataSignature(data);
      const permissionChanged = nextEditable !== editable;
      const chartFamilyChanged =
        draft.kind !== previousKind || draft.seriesDirection !== previousSeriesDirection;
      const optionalDefaultsChanged =
        draft.legendVisible !== previousLegendVisible || draft.varyColors !== previousVaryColors;
      const scatterStyleChangedExternally = draft.scatterStyle !== previousScatterStyle;
      const axisReverseOrderChanged =
        draft.xAxis.reverseOrder !== previousXReverseOrder ||
        draft.yAxis.reverseOrder !== previousYReverseOrder;
      const axisCrossingChanged =
        draft.xAxis.crossingMode !== previousXCrossingMode ||
        draft.yAxis.crossingMode !== previousYCrossingMode ||
        !Object.is(draft.xAxis.crossing, previousXCrossing) ||
        !Object.is(draft.yAxis.crossing, previousYCrossing) ||
        !Object.is(draft.xAxis.crossingCustom, previousXCrossingCustom) ||
        !Object.is(draft.yAxis.crossingCustom, previousYCrossingCustom);
      const axisLabelPositionChanged =
        draft.xAxis.labelPositionMode !== previousXLabelPositionMode ||
        draft.yAxis.labelPositionMode !== previousYLabelPositionMode ||
        !Object.is(draft.xAxis.labelPosition, previousXLabelPosition) ||
        !Object.is(draft.yAxis.labelPosition, previousYLabelPosition);
      const axisLogChanged =
        draft.xAxis.logEnabled !== previousXLogEnabled ||
        draft.yAxis.logEnabled !== previousYLogEnabled ||
        !Object.is(draft.xAxis.logBase, previousXLogBase) ||
        !Object.is(draft.yAxis.logBase, previousYLogBase) ||
        !Object.is(draft.xAxis.logBaseDraft, previousXLogBaseDraft) ||
        !Object.is(draft.yAxis.logBaseDraft, previousYLogBaseDraft);
      if (
        permissionChanged ||
        chartFamilyChanged ||
        optionalDefaultsChanged ||
        scatterStyleChangedExternally ||
        axisReverseOrderChanged ||
        axisCrossingChanged ||
        axisLabelPositionChanged ||
        axisLogChanged ||
        nextKey !== dataKey
      ) {
        const view = captureSettingsView(surface);
        editable = nextEditable;
        dataKey = nextKey;
        renderPanel();
        restoreSettingsView(surface, view);
      } else reposition();
    },
    close(discard = true) {
      closeInternal(discard);
    },
    setLabels(next, nextLocale) {
      labels = { ...labels, ...next };
      if (nextLocale) locale = nextLocale;
      if (!surface.hidden) {
        const view = captureSettingsView(surface);
        surface.lang = locale;
        renderPanel();
        restoreSettingsView(surface, view);
      }
    },
    detach() {
      if (disposed) return;
      disposed = true;
      removeEscape();
      surface.remove();
      chart = null;
      data = null;
      anchor = null;
      draft = null;
    },
    isOpen() {
      return !surface.hidden;
    },
    chartId() {
      return chart?.id ?? null;
    },
  };
}

export const chartSettingsDefaults = DEFAULT_LABELS;
export const chartSettingsPalettePresets = PALETTE_PRESETS;
