import { type ChartDataSource, ChartSourceKind } from '@libraz/formulon';
import type { ChartEmptyCellMode } from '../charts/types.js';
import type { History, HistoryTransaction } from '../commands/history.js';
import { parseMacRange } from '../commands/mac-data-tools.js';
import {
  addSessionChartSeries,
  removeSessionChartSeries,
  sessionChartById,
  setSessionChartSeries,
  updateSessionChart,
} from '../commands/session-chart.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SessionChart, SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { formatSheetAbsoluteRange } from '../wrappers/toolbar-a1.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';
import { createRangePickFormatter } from './range-pick-format.js';
import { attachRangePickerButton } from './range-picker-control.js';

/** Labels of the Select Data Source dialog. */
export interface SelectDataLabels {
  selectDataTitle: string;
  selectDataRangeGroup: string;
  selectDataRange: string;
  selectDataSwitch: string;
  selectDataSeries: string;
  selectDataSeriesList: string;
  selectDataAddSeries: string;
  selectDataRemoveSeries: string;
  selectDataMoveUp: string;
  selectDataMoveDown: string;
  selectDataName: string;
  selectDataValues: string;
  selectDataXValues: string;
  selectDataCategories: string;
  selectDataHiddenGroup: string;
  selectDataEmptyCells: string;
  selectDataEmptyGap: string;
  selectDataEmptyZero: string;
  selectDataEmptySpan: string;
  selectDataNaAsEmpty: string;
  selectDataShowHidden: string;
  selectDataPicker: string;
  /** `{n}` is the number of a series added in the dialog. */
  selectDataNewSeries: string;
  selectDataUnavailable: string;
  selectDataOk: string;
  cancel: string;
  invalidReference: string;
  pivotChartLocked: string;
  lastSeriesLocked: string;
}

/** English labels the dialog falls back to. */
export const DEFAULT_SELECT_DATA_LABELS: Omit<
  SelectDataLabels,
  'cancel' | 'invalidReference' | 'pivotChartLocked' | 'lastSeriesLocked'
> = {
  selectDataTitle: 'Select Data Source',
  selectDataRangeGroup: 'Range details',
  selectDataRange: 'Chart data range:',
  selectDataSwitch: 'Switch Row/Column',
  selectDataSeries: 'Legend entries (Series):',
  selectDataSeriesList: 'Chart series list',
  selectDataAddSeries: 'Add a new series',
  selectDataRemoveSeries: 'Remove series',
  selectDataMoveUp: 'Move series up',
  selectDataMoveDown: 'Move series down',
  selectDataName: 'Name:',
  selectDataValues: 'Y values:',
  selectDataXValues: 'X values:',
  selectDataCategories: 'Horizontal (Category) axis labels:',
  selectDataHiddenGroup: 'Hidden and Empty Cells',
  selectDataEmptyCells: 'Show empty cells as:',
  selectDataEmptyGap: 'Gaps',
  selectDataEmptyZero: 'Zero',
  selectDataEmptySpan: 'Connect data points with line',
  selectDataNaAsEmpty: 'Show #N/A as an empty cell',
  selectDataShowHidden: 'Show data in hidden rows and columns',
  selectDataPicker: 'Collapse dialog',
  selectDataNewSeries: 'Series{n}',
  selectDataUnavailable: 'This setting is not available yet.',
  selectDataOk: 'OK',
};

export interface MacChartSelectDataDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history: History | null;
  getWb: () => WorkbookHandle | null;
  getLabels: () => SelectDataLabels;
}

export interface MacChartSelectDataHandle {
  /** Open on chart `chartId`; false when the chart cannot be edited here. */
  open(chartId: string): boolean;
  close(): void;
  detach(): void;
}

/** One row of the series list: an engine series, or one added but not yet given values. */
interface SeriesRow {
  idx: number | null;
  label: string;
  name: string;
  values: string;
  x: string;
}

/** The dialog edits the chart as Chart Design > Select Data. */
const AUTH = { origin: 'ribbon' as const, commandId: 'chartSelectData' };

const refText = (source: ChartDataSource | undefined): string =>
  source && source.kind === ChartSourceKind.Ref && source.formula ? `=${source.formula}` : '';

/**
 * The Mac Select Data Source dialog. Edits apply to the chart as they are
 * made, inside one history transaction: OK keeps them as a single undo step
 * and Cancel rolls them all back.
 */
export function attachMacChartSelectDataDialog(
  deps: MacChartSelectDataDeps,
): MacChartSelectDataHandle {
  const { host, store, history, getWb } = deps;
  let t = deps.getLabels();

  const shell = createDialogShell({
    host,
    className: 'fc-mac-chartdata',
    ariaLabel: t.selectDataTitle,
    onDismiss: () => cancel(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.selectDataTitle,
    panelClasses: ['fc-fmtdlg__panel', 'fc-mac-chartdata__panel'],
    bodyClass: 'fc-fmtdlg__body fc-mac-chartdata__body',
    footerClass: 'fc-fmtdlg__footer fc-mac-chartdata__footer',
  });

  const group = (cls: string): { root: HTMLElement; legend: HTMLElement } => {
    const root = document.createElement('div');
    root.className = `fc-mac-chartdata__group ${cls}`;
    root.setAttribute('role', 'group');
    const legend = document.createElement('div');
    legend.className = 'fc-mac-chartdata__legend';
    root.appendChild(legend);
    body.appendChild(root);
    return { root, legend };
  };
  const pickFormat = createRangePickFormatter({
    store,
    sheetName: (index) => getWb()?.sheetName(index) ?? '',
  });
  const refField = (
    parent: HTMLElement,
    cls: string,
  ): { row: HTMLElement; text: HTMLSpanElement; input: HTMLInputElement } => {
    const row = document.createElement('label');
    row.className = `fc-mac-chartdata__field ${cls}`;
    const text = document.createElement('span');
    text.className = 'fc-mac-chartdata__label';
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'fc-fmtdlg__input fc-mac-chartdata__ref';
    input.autocomplete = 'off';
    input.spellcheck = false;
    row.append(text, input);
    parent.appendChild(row);
    attachRangePickerButton(input, {
      label: t.selectDataPicker,
      // Chart series refs are always sheet-qualified, like the text they replace.
      getValue: () => `=${pickFormat.range({ prefix: 'always' })}`,
      onPicked: () => input.dispatchEvent(new Event('change', { bubbles: true })),
      subscribeToRangeChanges: pickFormat.subscribe,
      kind: cls,
    });
    return { row, text, input };
  };

  const rangeGroup = group('fc-mac-chartdata__range-group');
  const rangeField = refField(rangeGroup.root, 'fc-mac-chartdata__range');

  const seriesHead = document.createElement('div');
  seriesHead.className = 'fc-mac-chartdata__series-head';
  const seriesText = document.createElement('span');
  seriesText.className = 'fc-mac-chartdata__label';
  const moveUp = appendDialogButton(seriesHead, {
    label: '',
    baseClass: 'fc-mac-chartdata__icon fc-mac-chartdata__move-up',
  });
  const moveDown = appendDialogButton(seriesHead, {
    label: '',
    baseClass: 'fc-mac-chartdata__icon fc-mac-chartdata__move-down',
  });
  seriesHead.prepend(seriesText);
  body.appendChild(seriesHead);

  const seriesArea = document.createElement('div');
  seriesArea.className = 'fc-mac-chartdata__series';
  const list = document.createElement('div');
  list.className = 'fc-mac-chartdata__list';
  list.setAttribute('role', 'listbox');
  list.tabIndex = 0;
  const detail = document.createElement('div');
  detail.className = 'fc-mac-chartdata__detail';
  seriesArea.append(list, detail);
  body.appendChild(seriesArea);
  const nameField = refField(detail, 'fc-mac-chartdata__name');
  const valuesField = refField(detail, 'fc-mac-chartdata__values');
  const xField = refField(detail, 'fc-mac-chartdata__x');

  const seriesTools = document.createElement('div');
  seriesTools.className = 'fc-mac-chartdata__tools';
  const addBtn = appendDialogButton(seriesTools, {
    label: '',
    baseClass: 'fc-mac-chartdata__icon fc-mac-chartdata__add',
  });
  const removeBtn = appendDialogButton(seriesTools, {
    label: '',
    baseClass: 'fc-mac-chartdata__icon fc-mac-chartdata__remove',
  });
  const switchBtn = appendDialogButton(seriesTools, {
    label: '',
    baseClass: 'fc-fmtdlg__btn fc-mac-chartdata__switch',
  });
  body.appendChild(seriesTools);

  const categoriesField = refField(body, 'fc-mac-chartdata__categories');

  const hiddenGroup = group('fc-mac-chartdata__hidden-group');
  const emptyRow = document.createElement('label');
  emptyRow.className = 'fc-mac-chartdata__field fc-mac-chartdata__empty';
  const emptyText = document.createElement('span');
  emptyText.className = 'fc-mac-chartdata__label';
  const emptySelect = document.createElement('select');
  emptySelect.className = 'fc-fmtdlg__select fc-mac-chartdata__empty-select';
  emptyRow.append(emptyText, emptySelect);
  const checkbox = (
    cls: string,
  ): { row: HTMLLabelElement; box: HTMLInputElement; text: HTMLSpanElement } => {
    const row = document.createElement('label');
    row.className = `fc-mac-chartdata__check ${cls}`;
    const box = document.createElement('input');
    box.type = 'checkbox';
    const text = document.createElement('span');
    row.append(box, text);
    return { row, box, text };
  };
  const naCheck = checkbox('fc-mac-chartdata__na');
  const hiddenCheck = checkbox('fc-mac-chartdata__show-hidden');
  hiddenGroup.root.append(emptyRow, naCheck.row, hiddenCheck.row);

  const error = document.createElement('div');
  error.className = 'fc-fmtdlg__error fc-mac-chartdata__error';
  error.setAttribute('role', 'alert');
  error.hidden = true;
  body.appendChild(error);

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.selectDataOk,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-chartdata__btn',
  });

  let chartId: string | null = null;
  let txn: HistoryTransaction | null = null;
  let rows: SeriesRow[] = [];
  let selected = 0;

  const applyLabels = (): void => {
    t = deps.getLabels();
    shell.setAriaLabel(t.selectDataTitle);
    header.textContent = t.selectDataTitle;
    rangeGroup.legend.textContent = t.selectDataRangeGroup;
    rangeGroup.root.setAttribute('aria-label', t.selectDataRangeGroup);
    rangeField.text.textContent = t.selectDataRange;
    seriesText.textContent = t.selectDataSeries;
    list.setAttribute('aria-label', t.selectDataSeriesList);
    nameField.text.textContent = t.selectDataName;
    valuesField.text.textContent = t.selectDataValues;
    xField.text.textContent = t.selectDataXValues;
    categoriesField.text.textContent = t.selectDataCategories;
    for (const [button, label] of [
      [moveUp, t.selectDataMoveUp],
      [moveDown, t.selectDataMoveDown],
      [addBtn, t.selectDataAddSeries],
      [removeBtn, t.selectDataRemoveSeries],
    ] as const) {
      button.setAttribute('aria-label', label);
      button.title = label;
    }
    switchBtn.textContent = t.selectDataSwitch;
    hiddenGroup.legend.textContent = t.selectDataHiddenGroup;
    hiddenGroup.root.setAttribute('aria-label', t.selectDataHiddenGroup);
    emptyText.textContent = t.selectDataEmptyCells;
    naCheck.text.textContent = t.selectDataNaAsEmpty;
    hiddenCheck.text.textContent = t.selectDataShowHidden;
    cancelBtn.textContent = t.cancel;
    okBtn.textContent = t.selectDataOk;
  };
  applyLabels();

  const chart = (): SessionChart | null =>
    chartId === null ? null : sessionChartById(store.getState(), chartId);

  const showError = (message: string | null): void => {
    error.hidden = message === null;
    error.textContent = message ?? '';
  };

  const parseRef = (text: string): Range | null => {
    const wb = getWb();
    const current = chart();
    if (!wb || !current) return null;
    const raw = text.trim().replace(/^=/, '');
    return raw ? (parseMacRange(wb, raw, current.sheet)?.range ?? null) : null;
  };

  /** Re-read the chart and project it onto every control. */
  const render = (): void => {
    const current = chart();
    const wb = getWb();
    if (!current || !wb) return;
    const info = wb.getChart(current.sheet, current.nativeObjectId);
    const engineSeries = [...(info?.plots[0]?.series ?? [])].sort((a, b) => a.order - b.order);
    const scatter = current.kind === 'scatter';
    const pending = rows.filter((row) => row.idx === null);
    rows = [
      ...engineSeries.map((entry, i) => ({
        idx: entry.seriesId,
        label: current.data.series[i]?.name ?? t.selectDataNewSeries.replace('{n}', String(i + 1)),
        name: refText(entry.name),
        values: refText(scatter ? entry.y : entry.values),
        x: refText(entry.x),
      })),
      ...pending,
    ];
    selected = Math.max(0, Math.min(selected, rows.length - 1));
    list.replaceChildren(
      ...rows.map((row, i) => {
        const option = document.createElement('div');
        option.className = 'fc-mac-chartdata__option';
        option.setAttribute('role', 'option');
        option.setAttribute('aria-selected', String(i === selected));
        option.textContent = row.label;
        option.addEventListener('click', () => {
          selected = i;
          render();
        });
        return option;
      }),
    );
    const row = rows[selected];
    nameField.input.value = row?.name ?? '';
    valuesField.input.value = row?.values ?? '';
    xField.input.value = row?.x ?? '';
    xField.row.hidden = !scatter;
    const firstSeries = engineSeries[0];
    categoriesField.input.value = scatter ? '' : refText(firstSeries?.categories);
    rangeField.input.value = current.source
      ? `=${formatSheetAbsoluteRange(wb.sheetName(current.source.sheet), current.source)}`
      : '';
    emptySelect.value = current.emptyCells ?? 'gap';
    hiddenCheck.box.checked = current.showHiddenData === true;
    naCheck.box.checked = true;

    const locked = current.pivotChart === true ? t.pivotChartLocked : null;
    const lockable = [
      rangeField.input,
      nameField.input,
      valuesField.input,
      xField.input,
      emptySelect,
      hiddenCheck.box,
      addBtn,
    ];
    for (const control of lockable) {
      projectDisabledState(control, locked !== null, locked, { datasetKey: 'disabledReason' });
    }
    projectDisabledState(categoriesField.input, locked !== null || scatter, locked, {
      datasetKey: 'disabledReason',
    });
    projectDisabledState(switchBtn, locked !== null || !current.source, locked, {
      datasetKey: 'disabledReason',
    });
    const lastSeries = engineSeries.length < 2 && row?.idx !== null;
    projectDisabledState(
      removeBtn,
      locked !== null || lastSeries,
      locked ?? (lastSeries ? t.lastSeriesLocked : null),
      { datasetKey: 'disabledReason' },
    );
    for (const control of [moveUp, moveDown, naCheck.box]) {
      projectDisabledState(control, true, t.selectDataUnavailable, {
        datasetKey: 'disabledReason',
      });
    }
  };

  /** Run one edit; a refused or invalid edit leaves the chart and reports why. */
  const edit = (run: (wb: WorkbookHandle, current: SessionChart) => boolean): void => {
    const current = chart();
    const wb = getWb();
    if (!current || !wb) return;
    showError(run(wb, current) ? null : t.invalidReference);
    render();
  };

  rangeField.input.addEventListener('change', () =>
    edit((wb, current) => {
      const source = parseRef(rangeField.input.value);
      return (
        source !== null &&
        updateSessionChart(store, current.id, { source }, history, wb, AUTH) !== null
      );
    }),
  );
  switchBtn.addEventListener('click', () =>
    edit(
      (wb, current) =>
        updateSessionChart(
          store,
          current.id,
          { seriesDirection: current.seriesDirection === 'rows' ? 'columns' : 'rows' },
          history,
          wb,
          AUTH,
        ) !== null,
    ),
  );
  const editSeries = (field: 'name' | 'values' | 'x', input: HTMLInputElement): void =>
    edit((wb, current) => {
      const row = rows[selected];
      if (!row) return false;
      const range = parseRef(input.value);
      if (row.idx === null) {
        row[field] = input.value;
        const values = parseRef(row.values);
        if (!values) return field !== 'values' || input.value.trim() === '';
        const name = parseRef(row.name) ?? undefined;
        const x = parseRef(row.x) ?? undefined;
        const categories = parseRef(categoriesField.input.value) ?? undefined;
        const added = addSessionChartSeries(
          store,
          current.id,
          {
            values,
            ...(name ? { name } : {}),
            ...(x ? { x } : {}),
            ...(categories ? { categories } : {}),
          },
          history,
          wb,
          AUTH,
        );
        if (!added) return false;
        rows = rows.filter((entry) => entry !== row);
        return true;
      }
      if (!range) return false;
      return (
        setSessionChartSeries(
          store,
          current.id,
          `series-${row.idx}`,
          { [field]: range },
          history,
          wb,
          AUTH,
        ) !== null
      );
    });
  nameField.input.addEventListener('change', () => editSeries('name', nameField.input));
  valuesField.input.addEventListener('change', () => editSeries('values', valuesField.input));
  xField.input.addEventListener('change', () => editSeries('x', xField.input));
  categoriesField.input.addEventListener('change', () =>
    edit((wb, current) => {
      const categories = parseRef(categoriesField.input.value);
      return (
        categories !== null &&
        setSessionChartSeries(store, current.id, null, { categories }, history, wb, AUTH) !== null
      );
    }),
  );
  addBtn.addEventListener('click', () => {
    rows.push({
      idx: null,
      label: t.selectDataNewSeries.replace('{n}', String(rows.length + 1)),
      name: '',
      values: '',
      x: '',
    });
    selected = rows.length - 1;
    render();
    valuesField.input.focus();
  });
  removeBtn.addEventListener('click', () => {
    const row = rows[selected];
    if (!row) return;
    if (row.idx === null) {
      rows = rows.filter((entry) => entry !== row);
      render();
      return;
    }
    const idx = row.idx;
    edit(
      (wb, current) =>
        removeSessionChartSeries(store, current.id, `series-${idx}`, history, wb, AUTH) !== null,
    );
  });
  emptySelect.addEventListener('change', () =>
    edit(
      (wb, current) =>
        updateSessionChart(
          store,
          current.id,
          { emptyCells: emptySelect.value as ChartEmptyCellMode },
          history,
          wb,
          AUTH,
        ) !== null,
    ),
  );
  hiddenCheck.box.addEventListener('change', () =>
    edit(
      (wb, current) =>
        updateSessionChart(
          store,
          current.id,
          { showHiddenData: hiddenCheck.box.checked },
          history,
          wb,
          AUTH,
        ) !== null,
    ),
  );

  const finish = (keep: boolean): void => {
    if (txn !== null && history) {
      if (keep) history.end(txn);
      else history.abort(txn);
    }
    txn = null;
    chartId = null;
    rows = [];
    shell.close();
  };
  const cancel = (): void => {
    if (chartId !== null) finish(false);
  };
  okBtn.addEventListener('click', () => finish(true));
  cancelBtn.addEventListener('click', cancel);

  return {
    open(id) {
      const current = sessionChartById(store.getState(), id);
      if (!current || current.kind === 'unsupported' || !getWb()?.capabilities.charts) return false;
      if (chartId !== null) finish(false);
      applyLabels();
      emptySelect.replaceChildren();
      appendDialogSelectOptions(emptySelect, [
        { value: 'gap', label: t.selectDataEmptyGap },
        { value: 'zero', label: t.selectDataEmptyZero },
        ...(current.kind === 'line' || current.kind === 'area' || current.kind === 'scatter'
          ? [{ value: 'span', label: t.selectDataEmptySpan }]
          : []),
      ]);
      chartId = id;
      pickFormat.begin();
      rows = [];
      selected = 0;
      txn = history ? history.begin() : null;
      showError(null);
      render();
      shell.open();
      deferInitialFocus(shell.panel, () => rangeField.input.focus());
      return true;
    },
    close: cancel,
    detach() {
      cancel();
      shell.dispose();
    },
  };
}
