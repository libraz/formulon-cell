import { PivotFilterKind } from '../engine/types.js';
import type { Strings } from '../i18n/strings.js';
import { createDialogSelect } from '../toolbar/dialogs/form-controls.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';
import type { PivotFilterConditionState } from './pivot-field-settings.js';
import {
  PIVOT_FILTER_COMPARISONS,
  type PivotFilterCategory,
  type PivotTop10Basis,
  type PivotTop10Direction,
  pivotFilterCategory,
  pivotFilterOperand,
  pivotTop10Kind,
  pivotTop10Parts,
} from './pivot-filter-kinds.js';

export interface PivotFilterDialogOptions {
  host: HTMLElement;
  strings: Strings['pivotTableDialog'];
  fieldName: string;
  /** The condition to edit; its kind picks the dialog (Top 10 kinds open the Top 10 dialog). */
  condition?: PivotFilterConditionState;
  okLabel: string;
  cancelLabel: string;
  /** Data field captions a value or Top 10 filter can score; the index is the data field index. */
  dataFields?: readonly string[];
  /** The desktop's Mac dialogs carry no field name in their title. */
  mac: boolean;
}

const RANGE_SEPARATOR = '..';

const splitRange = (value: string): [string, string] => {
  const at = value.indexOf(RANGE_SEPARATOR);
  return at < 0
    ? [value.trim(), '']
    : [value.slice(0, at).trim(), value.slice(at + RANGE_SEPARATOR.length).trim()];
};

const settingsRow = (label: string | null, ...controls: HTMLElement[]): HTMLLabelElement => {
  const row = document.createElement('label');
  row.className = 'fc-pivotdlg__settings-field';
  if (label !== null) {
    const text = document.createElement('span');
    text.textContent = label;
    row.append(text);
  }
  row.append(...controls);
  return row;
};

const hintRow = (text: string): HTMLElement => {
  const hint = document.createElement('div');
  hint.className = 'fc-pivotdlg__settings-field';
  const span = document.createElement('span');
  span.textContent = text;
  hint.append(span);
  return hint;
};

const textInput = (type: 'text' | 'number', value: string, dataKey: string): HTMLInputElement => {
  const input = document.createElement('input');
  input.type = type;
  input.className = 'fc-namedlg__input';
  input.value = value;
  input.dataset[dataKey] = 'true';
  return input;
};

const dataFieldSelect = (names: readonly string[], selected: number): HTMLSelectElement => {
  const select = createDialogSelect(
    names.map((name, index) => ({ value: String(index), label: name })),
    String(selected),
    { className: 'fc-fmtdlg__select' },
  );
  select.dataset.pivotFilterDataField = 'true';
  return select;
};

const dataFieldOf = (select: HTMLSelectElement | null): { dataField?: number } => {
  const index = select ? Number(select.value) : 0;
  return index > 0 ? { dataField: index } : {};
};

/** Opens the filter dialog of the condition's category (Label, Value, Date or Top 10) and resolves
 *  with the edited condition, or null when dismissed. */
export const showPivotFilterDialog = (
  options: PivotFilterDialogOptions,
): Promise<PivotFilterConditionState | null> =>
  new Promise<PivotFilterConditionState | null>((resolve) => {
    const { strings: t, dataFields = [] } = options;
    const d = t.filterDialogs;
    const initial = options.condition ?? { kind: null, value: '' };
    const initialKind = initial.kind ?? PivotFilterKind.LabelEqual;
    const top10Parts = pivotTop10Parts(initialKind);
    const category: PivotFilterCategory | 'top10' = top10Parts
      ? 'top10'
      : pivotFilterCategory(initialKind);
    const strs = category === 'top10' ? d.top10 : d[category];
    const title = options.mac ? strs.macTitle : strs.title.replace('{field}', options.fieldName);

    let done = false;
    const finish = (result: PivotFilterConditionState | null): void => {
      if (done) return;
      done = true;
      shell.dispose();
      resolve(result);
    };
    const shell = createDialogShell({
      host: options.host,
      className: 'fc-pivotdlg',
      ariaLabel: title,
      onDismiss: () => finish(null),
    });
    shell.overlay.classList.add('fc-fmtdlg');
    shell.overlay.classList.add('fc-pivotdlg--filter');
    const { body, footer } = appendDialogFrame(shell, {
      title,
      panelClasses: ['fc-fmtdlg__panel', 'fc-pivotdlg__panel'],
      bodyClass: 'fc-fmtdlg__body fc-pivotdlg__body',
    });
    const grid = document.createElement('div');
    grid.className = 'fc-pivotdlg__settings-grid';
    body.appendChild(grid);

    let read: () => PivotFilterConditionState;

    if (category === 'top10') {
      const parts = top10Parts as { direction: PivotTop10Direction; basis: PivotTop10Basis };
      const direction = createDialogSelect(
        [
          { value: 'top', label: d.top10.top },
          { value: 'bottom', label: d.top10.bottom },
        ],
        parts.direction,
        { className: 'fc-fmtdlg__select' },
      );
      direction.dataset.pivotFilterTop10Direction = 'true';
      const count = textInput('number', initial.value.trim() || '10', 'pivotFilterTop10Count');
      count.min = '1';
      count.step = '1';
      const basis = createDialogSelect(
        [
          { value: 'items', label: d.top10.items },
          { value: 'percent', label: d.top10.percent },
          { value: 'sum', label: d.top10.sum },
        ],
        parts.basis,
        { className: 'fc-fmtdlg__select' },
      );
      basis.dataset.pivotFilterTop10Basis = 'true';
      grid.append(settingsRow(d.top10.show, direction, count, basis));
      const by = dataFields.length > 0 ? dataFieldSelect(dataFields, initial.dataField ?? 0) : null;
      if (by) grid.append(settingsRow(d.top10.by, by));
      read = () => ({
        kind: pivotTop10Kind(
          direction.value as PivotTop10Direction,
          basis.value as PivotTop10Basis,
        ),
        value: count.value,
        ...dataFieldOf(by),
      });
    } else {
      const comparisons = PIVOT_FILTER_COMPARISONS[category];
      const operator = createDialogSelect(
        comparisons.map(([kind, key]) => ({ value: String(kind), label: d.operators[key] })),
        String(initialKind),
        { className: 'fc-fmtdlg__select' },
      );
      if (operator.value !== String(initialKind)) operator.value = String(comparisons[0]?.[0]);
      operator.dataset.pivotFilterOperator = 'true';
      const selectedKind = (): PivotFilterKind => Number(operator.value) as PivotFilterKind;
      const isRange = (): boolean => pivotFilterOperand(selectedKind()).endsWith('Range');
      const rangeInitial = pivotFilterOperand(initialKind).endsWith('Range');
      const [lowText, highText] = rangeInitial ? splitRange(initial.value) : [initial.value, ''];
      const inputType = category === 'value' ? 'number' : 'text';
      const low = textInput(inputType, lowText, 'pivotFilterValue');
      const high = textInput(inputType, highText, 'pivotFilterValueHigh');
      const andRow = settingsRow(d[category].and, high);
      const syncRange = (): void => {
        andRow.hidden = !isRange();
      };
      operator.addEventListener('change', syncRange);
      syncRange();

      const by =
        category === 'value' && dataFields.length > 0
          ? dataFieldSelect(dataFields, initial.dataField ?? 0)
          : null;
      grid.append(settingsRow(d[category].prompt));
      if (by) grid.append(settingsRow(null, by));
      grid.append(settingsRow(null, operator), settingsRow(null, low), andRow);
      if (category === 'label') {
        grid.append(hintRow(d.label.wildcardOne), hintRow(d.label.wildcardMany));
      }
      read = () => ({
        kind: selectedKind(),
        value: isRange() ? `${low.value}${RANGE_SEPARATOR}${high.value}` : low.value,
        ...dataFieldOf(by),
      });
    }

    const { cancelBtn: cancel, okBtn: ok } = appendDialogActions(footer, {
      cancelLabel: options.cancelLabel,
      okLabel: options.okLabel,
    });
    shell.on(cancel, 'click', () => finish(null));
    shell.on(ok, 'click', () => finish(read()));
    shell.open();
    deferInitialFocus(shell.overlay, () => {
      const first = grid.querySelector<HTMLSelectElement>('select');
      const button =
        first?.closest('.fc-select')?.querySelector<HTMLButtonElement>('.fc-select__button') ??
        first;
      button?.focus({ preventScroll: true });
    });
  });
