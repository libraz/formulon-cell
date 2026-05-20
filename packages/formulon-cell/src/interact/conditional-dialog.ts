import { type History, recordConditionalRulesChange } from '../commands/history.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  type ConditionalIconSet,
  type ConditionalRule,
  type ConditionalScalePoint,
  mutators,
  type SpreadsheetStore,
} from '../store/store.js';
import {
  type AverageMode,
  type CellValueOp,
  type DatePeriod,
  type FormatPreset,
  formatPresetPatch,
  formatRange,
  parseRange,
  type RuleKind,
} from './conditional-dialog-spec.js';
import { createDialogSelect, type DialogSelectOption } from '../toolbar/dialogs/form-controls.js';
import {
  appendConditionalApplyFormatControls,
  applyPatchToConditionalApplyControls,
  applyPresetPatchToConditionalApplyControls,
  collectConditionalApplyPatch,
} from './conditional-apply-controls.js';
import { appendDialogButton, createDialogShell } from './dialog-shell.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface ConditionalDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
  strings?: Strings;
}

export interface ConditionalDialogOpenOptions {
  mode?: 'manage' | 'new' | 'edit';
  editIndex?: number;
  kind?: ConditionalRule['kind'];
  cellValueOp?: CellValueOp;
  topBottomMode?: 'top' | 'bottom';
  topBottomPercent?: boolean;
  averageMode?: AverageMode;
  text?: string;
  datePeriod?: DatePeriod;
}

export interface ConditionalDialogHandle {
  open(options?: ConditionalDialogOpenOptions): void;
  close(): void;
  detach(): void;
}

/**
 * Manage conditional formatting rules: list / add / remove.
 * Spreadsheet parity is intentionally narrow — three rule kinds (cell-value,
 * color-scale, data-bar) and the renderer respects whichever fields apply.
 */
export function attachConditionalDialog(deps: ConditionalDialogDeps): ConditionalDialogHandle {
  const { host, store } = deps;
  const history = deps.history ?? null;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.conditionalDialog;
  const makeSelect = (
    options: readonly DialogSelectOption[],
    initial = options[0]?.value ?? '',
  ): HTMLSelectElement => createDialogSelect(options, initial, { className: '' });

  const shell = createDialogShell({
    host,
    className: 'fc-conddlg',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-conddlg__panel');
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body fc-conddlg__body';
  panel.appendChild(body);

  // ── Existing rules list ────────────────────────────────────────────────
  const rulesLegend = document.createElement('div');
  rulesLegend.className = 'fc-conddlg__legend';
  rulesLegend.textContent = t.title;
  body.appendChild(rulesLegend);
  const rulesList = document.createElement('div');
  rulesList.className = 'fc-conddlg__list';
  body.appendChild(rulesList);

  const clearAllBtn = appendDialogButton(body, {
    label: t.clearAll,
    baseClass: 'fc-fmtdlg__btn',
    secondaryClass: 'fc-conddlg__clear',
    variant: 'secondary',
  });

  // ── Add-rule form ──────────────────────────────────────────────────────
  const formLegend = document.createElement('div');
  formLegend.className = 'fc-conddlg__legend fc-conddlg__form-legend';
  formLegend.textContent = t.addRule;
  body.appendChild(formLegend);

  const form = document.createElement('div');
  form.className = 'fc-conddlg__form';
  body.appendChild(form);

  const ruleStyleRow = document.createElement('label');
  ruleStyleRow.className = 'fc-fmtdlg__row fc-conddlg__style-row';
  const styleLabel = document.createElement('span');
  styleLabel.textContent = t.styleLabel;
  const styleSelect = makeSelect([{ value: 'classic', label: t.styleClassic }]);
  ruleStyleRow.append(styleLabel, styleSelect);
  form.appendChild(ruleStyleRow);

  // Range
  const rangeRow = document.createElement('label');
  rangeRow.className = 'fc-fmtdlg__row';
  const rangeLabel = document.createElement('span');
  rangeLabel.textContent = t.rangeLabel;
  const rangeInput = document.createElement('input');
  rangeInput.type = 'text';
  rangeInput.spellcheck = false;
  rangeInput.autocomplete = 'off';
  rangeRow.append(rangeLabel, rangeInput);
  attachRangePickerButton(rangeInput, {
    label: strings.pivotTableDialog.rangePickerSelect,
    getValue: () => formatRange(store.getState().selection.range),
    subscribeToRangeChanges: (listener) => store.subscribe(listener),
    kind: 'conditional-format-range',
  });
  form.appendChild(rangeRow);

  // Kind
  const kindRow = document.createElement('label');
  kindRow.className = 'fc-fmtdlg__row';
  const kindLabel = document.createElement('span');
  kindLabel.textContent = t.kindLabel;
  const kindOptions: { id: RuleKind; label: string }[] = [
    { id: 'cell-value', label: t.kindCellValue },
    { id: 'color-scale', label: t.kindColorScale },
    { id: 'data-bar', label: t.kindDataBar },
    { id: 'icon-set', label: t.kindIconSet },
    { id: 'top-bottom', label: t.kindTopBottom },
    { id: 'average', label: t.kindAverage },
    { id: 'formula', label: t.kindFormula },
    { id: 'text-contains', label: t.kindTextContains },
    { id: 'date-occurring', label: t.kindDateOccurring },
    { id: 'duplicates', label: t.kindDuplicates },
    { id: 'unique', label: t.kindUnique },
    { id: 'blanks', label: t.kindBlanks },
    { id: 'non-blanks', label: t.kindNonBlanks },
    { id: 'errors', label: t.kindErrors },
    { id: 'no-errors', label: t.kindNoErrors },
  ];
  const kindSelect = makeSelect(kindOptions.map((o) => ({ value: o.id, label: o.label })));
  kindRow.append(kindLabel, kindSelect);
  form.appendChild(kindRow);

  // ── Cell-value subform ─────────────────────────────────────────────────
  const cellValueGroup = document.createElement('div');
  cellValueGroup.className = 'fc-conddlg__sub';
  form.appendChild(cellValueGroup);

  const opRow = document.createElement('label');
  opRow.className = 'fc-fmtdlg__row';
  const opLabel = document.createElement('span');
  opLabel.textContent = t.opLabel;
  const opOptions: { id: CellValueOp; label: string }[] = [
    { id: '>', label: t.opGt },
    { id: '<', label: t.opLt },
    { id: '>=', label: t.opGte },
    { id: '<=', label: t.opLte },
    { id: '=', label: t.opEq },
    { id: '<>', label: t.opNeq },
    { id: 'between', label: t.opBetween },
    { id: 'not-between', label: t.opNotBetween },
  ];
  const opSelect = makeSelect(opOptions.map((o) => ({ value: o.id, label: o.label })));
  opRow.append(opLabel, opSelect);
  cellValueGroup.appendChild(opRow);

  const valueARow = document.createElement('label');
  valueARow.className = 'fc-fmtdlg__row';
  const valueALabel = document.createElement('span');
  valueALabel.textContent = t.valueA;
  const valueAInput = document.createElement('input');
  valueAInput.type = 'number';
  valueAInput.step = 'any';
  valueAInput.value = '0';
  valueARow.append(valueALabel, valueAInput);
  cellValueGroup.appendChild(valueARow);

  const valueBRow = document.createElement('label');
  valueBRow.className = 'fc-fmtdlg__row';
  const valueBLabel = document.createElement('span');
  valueBLabel.textContent = t.valueB;
  const valueBInput = document.createElement('input');
  valueBInput.type = 'number';
  valueBInput.step = 'any';
  valueBInput.value = '0';
  valueBRow.append(valueBLabel, valueBInput);
  cellValueGroup.appendChild(valueBRow);

  // Apply: fill, color, bold, italic, underline, strike
  const cellValueApplyControls = appendConditionalApplyFormatControls(cellValueGroup, t);

  const cellPresetRow = document.createElement('label');
  cellPresetRow.className = 'fc-fmtdlg__row fc-conddlg__format-row';
  const cellPresetLabel = document.createElement('span');
  cellPresetLabel.textContent = t.formatLabel;
  const formatPresetOptions: { id: FormatPreset; label: string }[] = [
    { id: 'red-fill', label: t.formatRedFill },
    { id: 'yellow-fill', label: t.formatYellowFill },
    { id: 'green-fill', label: t.formatGreenFill },
    { id: 'red-text', label: t.formatRedText },
    { id: 'plain', label: t.formatPlain },
  ];
  const cellPresetSelect = makeSelect(
    formatPresetOptions.map((o) => ({ value: o.id, label: o.label })),
  );
  const cellPresetPreview = document.createElement('span');
  cellPresetPreview.className = 'fc-conddlg__preview';
  cellPresetPreview.textContent = t.previewText;
  const cellPresetWrap = document.createElement('span');
  cellPresetWrap.className = 'fc-conddlg__format-picker';
  cellPresetWrap.append(cellPresetSelect, cellPresetPreview);
  cellPresetRow.append(cellPresetLabel, cellPresetWrap);
  cellValueGroup.appendChild(cellPresetRow);

  // ── Color scale subform ────────────────────────────────────────────────
  const colorScaleGroup = document.createElement('div');
  colorScaleGroup.className = 'fc-conddlg__sub';
  form.appendChild(colorScaleGroup);

  const useThreeRow = document.createElement('label');
  useThreeRow.className = 'fc-fmtdlg__check';
  const useThreeCk = document.createElement('input');
  useThreeCk.type = 'checkbox';
  const useThreeText = document.createElement('span');
  useThreeText.textContent = t.useThreeStops;
  useThreeRow.append(useThreeCk, useThreeText);
  colorScaleGroup.appendChild(useThreeRow);

  const stopMinRow = document.createElement('label');
  stopMinRow.className = 'fc-fmtdlg__row';
  const stopMinLabel = document.createElement('span');
  stopMinLabel.textContent = t.stopMin;
  const stopMinInput = document.createElement('input');
  stopMinInput.type = 'color';
  stopMinInput.value = '#f8696b';
  stopMinInput.setAttribute('aria-label', t.stopMin);
  stopMinRow.append(stopMinLabel, stopMinInput);
  colorScaleGroup.appendChild(stopMinRow);

  const stopMidRow = document.createElement('label');
  stopMidRow.className = 'fc-fmtdlg__row';
  const stopMidLabel = document.createElement('span');
  stopMidLabel.textContent = t.stopMid;
  const stopMidInput = document.createElement('input');
  stopMidInput.type = 'color';
  stopMidInput.value = '#ffeb84';
  stopMidInput.setAttribute('aria-label', t.stopMid);
  stopMidRow.append(stopMidLabel, stopMidInput);
  stopMidRow.hidden = true;
  colorScaleGroup.appendChild(stopMidRow);

  const stopMaxRow = document.createElement('label');
  stopMaxRow.className = 'fc-fmtdlg__row';
  const stopMaxLabel = document.createElement('span');
  stopMaxLabel.textContent = t.stopMax;
  const stopMaxInput = document.createElement('input');
  stopMaxInput.type = 'color';
  stopMaxInput.value = '#63be7b';
  stopMaxInput.setAttribute('aria-label', t.stopMax);
  stopMaxRow.append(stopMaxLabel, stopMaxInput);
  colorScaleGroup.appendChild(stopMaxRow);

  const scaleTypeOptions = [
    { id: 'min', label: t.scaleTypeMin },
    { id: 'max', label: t.scaleTypeMax },
    { id: 'number', label: t.scaleTypeNumber },
    { id: 'percent', label: t.scaleTypePercent },
    { id: 'percentile', label: t.scaleTypePercentile },
  ] as const;
  const makeScalePointRow = (
    label: string,
    defaultType: ConditionalScalePoint['kind'],
    defaultValue: string,
  ): { row: HTMLLabelElement; type: HTMLSelectElement; value: HTMLInputElement } => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row';
    const span = document.createElement('span');
    span.textContent = `${label} ${t.scaleType}`;
    const type = makeSelect(
      scaleTypeOptions.map((option) => ({ value: option.id, label: option.label })),
      defaultType,
    );
    const value = document.createElement('input');
    value.type = 'number';
    value.value = defaultValue;
    value.setAttribute('aria-label', `${label} ${t.scaleValue}`);
    const syncValue = (): void => {
      value.hidden = type.value === 'min' || type.value === 'max';
    };
    type.addEventListener('change', syncValue);
    syncValue();
    row.append(span, type, value);
    colorScaleGroup.appendChild(row);
    return { row, type, value };
  };
  const scaleMin = makeScalePointRow(t.stopMin, 'min', '0');
  const scaleMid = makeScalePointRow(t.stopMid, 'percentile', '50');
  scaleMid.row.hidden = true;
  const scaleMax = makeScalePointRow(t.stopMax, 'max', '100');

  // ── Data bar subform ───────────────────────────────────────────────────
  const dataBarGroup = document.createElement('div');
  dataBarGroup.className = 'fc-conddlg__sub';
  form.appendChild(dataBarGroup);

  const barFillStyleRow = document.createElement('label');
  barFillStyleRow.className = 'fc-fmtdlg__row';
  const barFillStyleLabel = document.createElement('span');
  barFillStyleLabel.textContent = t.barFillStyle;
  const barFillStyleSelect = makeSelect([
    { value: 'gradient', label: t.gradientFill },
    { value: 'solid', label: t.solidFill },
  ]);
  barFillStyleRow.append(barFillStyleLabel, barFillStyleSelect);
  dataBarGroup.appendChild(barFillStyleRow);

  const barColorRow = document.createElement('label');
  barColorRow.className = 'fc-fmtdlg__row';
  const barColorLabel = document.createElement('span');
  barColorLabel.textContent = t.barColor;
  const barColorInput = document.createElement('input');
  barColorInput.type = 'color';
  barColorInput.value = '#638ec6';
  barColorInput.setAttribute('aria-label', t.barColor);
  barColorRow.append(barColorLabel, barColorInput);
  dataBarGroup.appendChild(barColorRow);

  const showValueRow = document.createElement('label');
  showValueRow.className = 'fc-fmtdlg__check';
  const showValueCk = document.createElement('input');
  showValueCk.type = 'checkbox';
  showValueCk.checked = true;
  const showValueText = document.createElement('span');
  showValueText.textContent = t.showValue;
  showValueRow.append(showValueCk, showValueText);
  dataBarGroup.appendChild(showValueRow);

  // ── Icon-set subform ───────────────────────────────────────────────────
  const iconSetGroup = document.createElement('div');
  iconSetGroup.className = 'fc-conddlg__sub';
  form.appendChild(iconSetGroup);

  const iconSetRow = document.createElement('label');
  iconSetRow.className = 'fc-fmtdlg__row';
  const iconSetLabel = document.createElement('span');
  iconSetLabel.textContent = t.kindIconSet;
  const iconSetOptions: { id: ConditionalIconSet; label: string }[] = [
    { id: 'arrows3', label: t.iconSetArrows3 },
    { id: 'arrows5', label: t.iconSetArrows5 },
    { id: 'triangles3', label: t.iconSetTriangles3 },
    { id: 'traffic3', label: t.iconSetTraffic3 },
    { id: 'trafficRim3', label: t.iconSetTrafficRim3 },
    { id: 'symbols3', label: t.iconSetSymbols3 },
    { id: 'flags3', label: t.iconSetFlags3 },
    { id: 'stars3', label: t.iconSetStars3 },
    { id: 'quarters5', label: t.iconSetQuarters5 },
    { id: 'ratings5', label: t.iconSetRatings5 },
    { id: 'bars5', label: t.iconSetBars5 },
    { id: 'boxes5', label: t.iconSetBoxes5 },
  ];
  const iconSetSelect = makeSelect(iconSetOptions.map((o) => ({ value: o.id, label: o.label })));
  const iconSetLabelFor = (id: ConditionalIconSet): string =>
    iconSetOptions.find((option) => option.id === id)?.label ?? id;
  iconSetRow.append(iconSetLabel, iconSetSelect);
  iconSetGroup.appendChild(iconSetRow);

  const iconReverseRow = document.createElement('label');
  iconReverseRow.className = 'fc-fmtdlg__check';
  const iconReverseCk = document.createElement('input');
  iconReverseCk.type = 'checkbox';
  const iconReverseText = document.createElement('span');
  iconReverseText.textContent = t.reverseOrder;
  iconReverseRow.append(iconReverseCk, iconReverseText);
  iconSetGroup.appendChild(iconReverseRow);

  const iconOnlyRow = document.createElement('label');
  iconOnlyRow.className = 'fc-fmtdlg__check';
  const iconOnlyCk = document.createElement('input');
  iconOnlyCk.type = 'checkbox';
  const iconOnlyText = document.createElement('span');
  iconOnlyText.textContent = t.showIconOnly;
  iconOnlyRow.append(iconOnlyCk, iconOnlyText);
  iconSetGroup.appendChild(iconOnlyRow);

  const makeIconThresholdRow = (
    index: number,
  ): { row: HTMLLabelElement; type: HTMLSelectElement; value: HTMLInputElement } => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row';
    const span = document.createElement('span');
    span.textContent = `${t.iconThreshold} ${index + 1}`;
    const type = makeSelect(
      scaleTypeOptions.map((option) => ({ value: option.id, label: option.label })),
      'percent',
    );
    const value = document.createElement('input');
    value.type = 'number';
    value.setAttribute('aria-label', `${t.iconThreshold} ${index + 1} ${t.scaleValue}`);
    const syncValue = (): void => {
      value.hidden = type.value === 'min' || type.value === 'max';
    };
    type.addEventListener('change', syncValue);
    syncValue();
    row.append(span, type, value);
    iconSetGroup.appendChild(row);
    return { row, type, value };
  };
  const iconThresholdControls = [0, 1, 2, 3].map((index) => makeIconThresholdRow(index));

  // ── Top/Bottom subform ─────────────────────────────────────────────────
  const topBottomGroup = document.createElement('div');
  topBottomGroup.className = 'fc-conddlg__sub';
  form.appendChild(topBottomGroup);

  const tbModeRow = document.createElement('label');
  tbModeRow.className = 'fc-fmtdlg__row';
  const tbModeLabel = document.createElement('span');
  tbModeLabel.textContent = t.topBottomMode;
  const tbModeSelect = makeSelect([
    { value: 'top', label: t.topMode },
    { value: 'bottom', label: t.bottomMode },
  ]);
  tbModeRow.append(tbModeLabel, tbModeSelect);
  topBottomGroup.appendChild(tbModeRow);

  const tbNRow = document.createElement('label');
  tbNRow.className = 'fc-fmtdlg__row';
  const tbNLabel = document.createElement('span');
  tbNLabel.textContent = t.topN;
  const tbNInput = document.createElement('input');
  tbNInput.type = 'number';
  tbNInput.min = '1';
  tbNInput.step = '1';
  tbNInput.value = '10';
  tbNRow.append(tbNLabel, tbNInput);
  topBottomGroup.appendChild(tbNRow);

  const tbPercentRow = document.createElement('label');
  tbPercentRow.className = 'fc-fmtdlg__check';
  const tbPercentCk = document.createElement('input');
  tbPercentCk.type = 'checkbox';
  const tbPercentText = document.createElement('span');
  tbPercentText.textContent = t.usePercent;
  tbPercentRow.append(tbPercentCk, tbPercentText);
  topBottomGroup.appendChild(tbPercentRow);

  // ── Above/below average subform ────────────────────────────────────────
  const averageGroup = document.createElement('div');
  averageGroup.className = 'fc-conddlg__sub';
  form.appendChild(averageGroup);

  const averageModeRow = document.createElement('label');
  averageModeRow.className = 'fc-fmtdlg__row';
  const averageModeLabel = document.createElement('span');
  averageModeLabel.textContent = t.averageModeLabel;
  const averageModeOptions: { id: AverageMode; label: string }[] = [
    { id: 'above', label: t.averageAbove },
    { id: 'below', label: t.averageBelow },
    { id: 'equal-or-above', label: t.averageEqualOrAbove },
    { id: 'equal-or-below', label: t.averageEqualOrBelow },
  ];
  const averageModeSelect = makeSelect(
    averageModeOptions.map((o) => ({ value: o.id, label: o.label })),
  );
  averageModeRow.append(averageModeLabel, averageModeSelect);
  averageGroup.appendChild(averageModeRow);
  const averageModeLabelFor = (id: AverageMode): string =>
    averageModeOptions.find((option) => option.id === id)?.label ?? id;

  // ── Formula subform ────────────────────────────────────────────────────
  const formulaGroup = document.createElement('div');
  formulaGroup.className = 'fc-conddlg__sub';
  form.appendChild(formulaGroup);

  const formulaRow = document.createElement('label');
  formulaRow.className = 'fc-fmtdlg__row';
  const formulaLabelEl = document.createElement('span');
  formulaLabelEl.textContent = t.kindFormula;
  const formulaInput = document.createElement('input');
  formulaInput.type = 'text';
  formulaInput.spellcheck = false;
  formulaInput.autocomplete = 'off';
  formulaInput.placeholder = t.formulaPlaceholder;
  formulaRow.append(formulaLabelEl, formulaInput);
  formulaGroup.appendChild(formulaRow);

  // ── Text-containing subform ────────────────────────────────────────────
  const textContainsGroup = document.createElement('div');
  textContainsGroup.className = 'fc-conddlg__sub';
  form.appendChild(textContainsGroup);

  const textContainsRow = document.createElement('label');
  textContainsRow.className = 'fc-fmtdlg__row';
  const textContainsLabel = document.createElement('span');
  textContainsLabel.textContent = t.textContainsLabel;
  const textContainsInput = document.createElement('input');
  textContainsInput.type = 'text';
  textContainsInput.spellcheck = false;
  textContainsInput.autocomplete = 'off';
  textContainsInput.placeholder = t.textContainsPlaceholder;
  textContainsRow.append(textContainsLabel, textContainsInput);
  textContainsGroup.appendChild(textContainsRow);

  const caseSensitiveRow = document.createElement('label');
  caseSensitiveRow.className = 'fc-fmtdlg__check';
  const caseSensitiveCk = document.createElement('input');
  caseSensitiveCk.type = 'checkbox';
  const caseSensitiveText = document.createElement('span');
  caseSensitiveText.textContent = t.caseSensitive;
  caseSensitiveRow.append(caseSensitiveCk, caseSensitiveText);
  textContainsGroup.appendChild(caseSensitiveRow);

  // ── Date-occurring subform ─────────────────────────────────────────────
  const dateOccurringGroup = document.createElement('div');
  dateOccurringGroup.className = 'fc-conddlg__sub';
  form.appendChild(dateOccurringGroup);

  const datePeriodRow = document.createElement('label');
  datePeriodRow.className = 'fc-fmtdlg__row';
  const datePeriodLabel = document.createElement('span');
  datePeriodLabel.textContent = t.datePeriodLabel;
  const datePeriodOptions: { id: DatePeriod; label: string }[] = [
    { id: 'yesterday', label: t.dateYesterday },
    { id: 'today', label: t.dateToday },
    { id: 'tomorrow', label: t.dateTomorrow },
    { id: 'last7', label: t.dateLast7 },
    { id: 'last-week', label: t.dateLastWeek },
    { id: 'this-week', label: t.dateThisWeek },
    { id: 'next-week', label: t.dateNextWeek },
    { id: 'last-month', label: t.dateLastMonth },
    { id: 'this-month', label: t.dateThisMonth },
    { id: 'next-month', label: t.dateNextMonth },
  ];
  const datePeriodSelect = makeSelect(
    datePeriodOptions.map((o) => ({ value: o.id, label: o.label })),
  );
  const datePeriodLabelFor = (id: DatePeriod): string =>
    datePeriodOptions.find((option) => option.id === id)?.label ?? id;
  datePeriodRow.append(datePeriodLabel, datePeriodSelect);
  dateOccurringGroup.appendChild(datePeriodRow);

  // ── Apply-format shared by top-bottom / formula / dups / unique /
  //    blanks / non-blanks / errors / no-errors. We re-use the same
  //    fill/font/style controls from the cell-value subform so the
  //    "apply when matched" surface stays consistent.
  const applyGroup = document.createElement('div');
  applyGroup.className = 'fc-conddlg__sub';
  form.appendChild(applyGroup);

  const sharedApplyControls = appendConditionalApplyFormatControls(applyGroup, t);

  const sharedPresetRow = document.createElement('label');
  sharedPresetRow.className = 'fc-fmtdlg__row fc-conddlg__format-row';
  const sharedPresetLabel = document.createElement('span');
  sharedPresetLabel.textContent = t.formatLabel;
  const sharedPresetSelect = cellPresetSelect.cloneNode(true) as HTMLSelectElement;
  const sharedPresetPreview = document.createElement('span');
  sharedPresetPreview.className = 'fc-conddlg__preview';
  sharedPresetPreview.textContent = t.previewText;
  const sharedPresetWrap = document.createElement('span');
  sharedPresetWrap.className = 'fc-conddlg__format-picker';
  sharedPresetWrap.append(sharedPresetSelect, sharedPresetPreview);
  sharedPresetRow.append(sharedPresetLabel, sharedPresetWrap);
  applyGroup.appendChild(sharedPresetRow);

  // Add button
  const addRow = document.createElement('div');
  addRow.className = 'fc-fmtdlg__row fc-conddlg__addrow';
  const addBtn = appendDialogButton(addRow, { label: t.addRule, variant: 'primary' });
  form.appendChild(addRow);

  // Footer
  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);
  const closeBtn = appendDialogButton(footer, { label: t.close });

  // ── Behaviour ──────────────────────────────────────────────────────────
  /** Kinds that re-use the shared `applyGroup` (fill/font/style) controls
   *  for their "apply when matched" format. cell-value carries its own
   *  controls inside `cellValueGroup` and so is excluded here. */
  const APPLY_KINDS: ReadonlySet<RuleKind> = new Set([
    'top-bottom',
    'average',
    'formula',
    'text-contains',
    'date-occurring',
    'duplicates',
    'unique',
    'blanks',
    'non-blanks',
    'errors',
    'no-errors',
  ]);
  const syncSubforms = (): void => {
    const kind = kindSelect.value as RuleKind;
    cellValueGroup.hidden = kind !== 'cell-value';
    colorScaleGroup.hidden = kind !== 'color-scale';
    dataBarGroup.hidden = kind !== 'data-bar';
    iconSetGroup.hidden = kind !== 'icon-set';
    topBottomGroup.hidden = kind !== 'top-bottom';
    averageGroup.hidden = kind !== 'average';
    formulaGroup.hidden = kind !== 'formula';
    textContainsGroup.hidden = kind !== 'text-contains';
    dateOccurringGroup.hidden = kind !== 'date-occurring';
    applyGroup.hidden = !APPLY_KINDS.has(kind);
  };
  const syncCellValueOp = (): void => {
    const op = opSelect.value as CellValueOp;
    valueBRow.hidden = op !== 'between' && op !== 'not-between';
  };
  const syncThreeStops = (): void => {
    stopMidRow.hidden = !useThreeCk.checked;
    scaleMid.row.hidden = !useThreeCk.checked;
  };
  const syncIconThresholds = (): void => {
    const slots = iconSetSelect.value.endsWith('5') ? 5 : 3;
    for (let index = 0; index < iconThresholdControls.length; index += 1) {
      const control = iconThresholdControls[index];
      if (!control) continue;
      control.row.hidden = index >= slots - 1;
      if (control.value.value === '') {
        control.value.value = String(Math.round(((index + 1) * 100) / slots));
      }
    }
  };
  const syncPresetPreview = (preview: HTMLElement, preset: FormatPreset): void => {
    const patch = formatPresetPatch(preset);
    preview.style.color = patch.color ?? '#201f1e';
    preview.style.background = patch.fill ?? 'transparent';
  };
  const syncCellPreset = (): void => {
    const patch = formatPresetPatch(cellPresetSelect.value as FormatPreset);
    applyPresetPatchToConditionalApplyControls(cellValueApplyControls, patch);
    syncPresetPreview(cellPresetPreview, cellPresetSelect.value as FormatPreset);
  };
  const syncSharedPreset = (): void => {
    const patch = formatPresetPatch(sharedPresetSelect.value as FormatPreset);
    applyPresetPatchToConditionalApplyControls(sharedApplyControls, patch);
    syncPresetPreview(sharedPresetPreview, sharedPresetSelect.value as FormatPreset);
  };

  let currentMode: 'manage' | 'new' | 'edit' = 'manage';
  let currentEditIndex: number | null = null;
  const syncDialogMode = (): void => {
    const isNew = currentMode === 'new';
    const isEdit = currentMode === 'edit';
    const title = isEdit ? t.editRuleTitle : isNew ? t.newRuleTitle : t.title;
    header.textContent = title;
    overlay.setAttribute('aria-label', title);
    shell.panel.classList.toggle('fc-conddlg__panel--new', isNew);
    body.classList.toggle('fc-conddlg__body--new', isNew);
    shell.panel.classList.toggle('fc-conddlg__panel--edit', isEdit);
    body.classList.toggle('fc-conddlg__body--edit', isEdit);
    rulesLegend.hidden = isNew || isEdit;
    rulesList.hidden = isNew || isEdit;
    clearAllBtn.hidden = isNew || isEdit;
    formLegend.hidden = isNew || isEdit;
    addBtn.textContent = isEdit ? t.saveRule : isNew ? t.ok : t.addRule;
    closeBtn.textContent = isNew || isEdit ? t.cancel : t.close;
  };

  const renderRules = (): void => {
    rulesList.replaceChildren();
    const rules = store.getState().conditional.rules;
    if (rules.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'fc-conddlg__empty';
      empty.textContent = t.empty;
      rulesList.appendChild(empty);
      return;
    }
    rules.forEach((rule, idx) => {
      const item = document.createElement('div');
      item.className = 'fc-conddlg__item';
      const summary = document.createElement('span');
      summary.textContent = describeRule(rule);
      const removeBtn = appendDialogButton(item, { label: t.removeRule });
      removeBtn.addEventListener('click', () => {
        recordConditionalRulesChange(history, store, () => {
          mutators.removeConditionalRuleAt(store, idx);
        });
        renderRules();
      });
      item.prepend(summary);
      rulesList.appendChild(item);
    });
  };

  const describeRule = (rule: ConditionalRule): string => {
    const range = formatRange(rule.range);
    switch (rule.kind) {
      case 'cell-value': {
        const opLabel = opOptions.find((o) => o.id === rule.op)?.label ?? rule.op;
        const tail =
          rule.op === 'between' || rule.op === 'not-between'
            ? `${rule.a} – ${rule.b ?? rule.a}`
            : `${rule.a}`;
        return `${range} · ${t.kindCellValue} (${opLabel} ${tail})`;
      }
      case 'color-scale':
        return `${range} · ${t.kindColorScale} (${rule.stops.length} ${t.stopsLabel})`;
      case 'data-bar':
        return `${range} · ${t.kindDataBar} (${rule.gradient ? t.gradientFill : t.solidFill})`;
      case 'icon-set':
        return `${range} · ${t.kindIconSet} (${iconSetLabelFor(rule.icons)}${
          rule.showValue === false ? `, ${t.showIconOnly}` : ''
        })`;
      case 'top-bottom': {
        const pct = rule.percent ? '%' : '';
        const modeLabel = rule.mode === 'top' ? t.topMode : t.bottomMode;
        return `${range} · ${t.kindTopBottom} (${modeLabel} ${rule.n}${pct})`;
      }
      case 'average':
        return `${range} · ${t.kindAverage} (${averageModeLabelFor(rule.mode)})`;
      case 'text-contains':
        return `${range} · ${t.kindTextContains} ("${rule.text}")`;
      case 'date-occurring':
        return `${range} · ${t.kindDateOccurring} (${datePeriodLabelFor(rule.period)})`;
      case 'formula':
        return `${range} · ${t.kindFormula} (${rule.formula})`;
      case 'duplicates':
        return `${range} · ${t.kindDuplicates}`;
      case 'unique':
        return `${range} · ${t.kindUnique}`;
      case 'blanks':
        return `${range} · ${t.kindBlanks}`;
      case 'non-blanks':
        return `${range} · ${t.kindNonBlanks}`;
      case 'errors':
        return `${range} · ${t.kindErrors}`;
      case 'no-errors':
        return `${range} · ${t.kindNoErrors}`;
    }
  };

  const collectScalePoint = (input: {
    type: HTMLSelectElement;
    value: HTMLInputElement;
  }): ConditionalScalePoint | null => {
    const kind = input.type.value as ConditionalScalePoint['kind'];
    if (kind === 'min' || kind === 'max') return { kind };
    const value = Number.parseFloat(input.value.value);
    if (!Number.isFinite(value)) return null;
    return { kind, value };
  };

  const populateRuleForm = (rule: ConditionalRule): void => {
    rangeInput.value = formatRange(rule.range);
    kindSelect.value = rule.kind;
    if (rule.kind === 'cell-value') {
      opSelect.value = rule.op;
      valueAInput.value = String(rule.a);
      valueBInput.value = String(rule.b ?? rule.a);
      applyPatchToConditionalApplyControls(cellValueApplyControls, rule.apply);
    } else if (rule.kind === 'color-scale') {
      useThreeCk.checked = rule.stops.length === 3;
      stopMinInput.value = rule.stops[0] ?? '#f8696b';
      stopMidInput.value = rule.stops.length === 3 ? (rule.stops[1] ?? '#ffeb84') : '#ffeb84';
      stopMaxInput.value = rule.stops.at(-1) ?? '#63be7b';
      const thresholds = rule.thresholds ?? [];
      const min = thresholds[0];
      const mid = rule.stops.length === 3 ? thresholds[1] : undefined;
      const max = rule.stops.length === 3 ? thresholds[2] : thresholds[1];
      if (min) {
        scaleMin.type.value = min.kind;
        scaleMin.value.value = 'value' in min ? String(min.value) : '0';
      }
      if (mid) {
        scaleMid.type.value = mid.kind;
        scaleMid.value.value = 'value' in mid ? String(mid.value) : '50';
      }
      if (max) {
        scaleMax.type.value = max.kind;
        scaleMax.value.value = 'value' in max ? String(max.value) : '100';
      }
    } else if (rule.kind === 'data-bar') {
      barFillStyleSelect.value = rule.gradient === false ? 'solid' : 'gradient';
      barColorInput.value = rule.color;
      showValueCk.checked = rule.showValue !== false;
    } else if (rule.kind === 'icon-set') {
      iconSetSelect.value = rule.icons;
      iconReverseCk.checked = rule.reverseOrder === true;
      iconOnlyCk.checked = rule.showValue === false;
      for (const [index, point] of (rule.thresholds ?? []).entries()) {
        const control = iconThresholdControls[index];
        if (!control) continue;
        control.type.value = point.kind;
        control.value.value = 'value' in point ? String(point.value) : '';
      }
    } else if (rule.kind === 'top-bottom') {
      tbModeSelect.value = rule.mode;
      tbNInput.value = String(rule.n);
      tbPercentCk.checked = rule.percent === true;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'average') {
      averageModeSelect.value = rule.mode;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'formula') {
      formulaInput.value = rule.formula;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'text-contains') {
      textContainsInput.value = rule.text;
      caseSensitiveCk.checked = rule.caseSensitive === true;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'date-occurring') {
      datePeriodSelect.value = rule.period;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else {
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    }
    syncSubforms();
    syncCellValueOp();
    syncThreeStops();
    syncIconThresholds();
    for (const control of [scaleMin, scaleMid, scaleMax, ...iconThresholdControls]) {
      control.type.dispatchEvent(new Event('change'));
    }
  };

  const onAdd = (): void => {
    const fallback = store.getState().selection.range;
    const range = parseRange(rangeInput.value, fallback);
    const kind = kindSelect.value as RuleKind;
    let rule: ConditionalRule | null = null;
    if (kind === 'cell-value') {
      const op = opSelect.value as CellValueOp;
      const a = Number.parseFloat(valueAInput.value);
      const b = Number.parseFloat(valueBInput.value);
      if (!Number.isFinite(a)) return;
      const applyPatch = collectConditionalApplyPatch(cellValueApplyControls);
      rule = {
        kind: 'cell-value',
        range,
        op,
        a,
        ...(op === 'between' || op === 'not-between' ? { b } : {}),
        apply: applyPatch,
      };
    } else if (kind === 'color-scale') {
      const stops: [string, string] | [string, string, string] = useThreeCk.checked
        ? [stopMinInput.value, stopMidInput.value, stopMaxInput.value]
        : [stopMinInput.value, stopMaxInput.value];
      const minPoint = collectScalePoint(scaleMin);
      const maxPoint = collectScalePoint(scaleMax);
      if (!minPoint || !maxPoint) return;
      if (useThreeCk.checked) {
        const midPoint = collectScalePoint(scaleMid);
        if (!midPoint) return;
        rule = { kind: 'color-scale', range, stops, thresholds: [minPoint, midPoint, maxPoint] };
      } else {
        rule = { kind: 'color-scale', range, stops, thresholds: [minPoint, maxPoint] };
      }
    } else if (kind === 'data-bar') {
      rule = {
        kind: 'data-bar',
        range,
        color: barColorInput.value,
        gradient: barFillStyleSelect.value === 'gradient',
        showValue: showValueCk.checked,
      };
    } else if (kind === 'icon-set') {
      const iconThresholds = iconThresholdControls
        .filter((control) => !control.row.hidden)
        .map((control) => collectScalePoint(control));
      if (iconThresholds.some((point) => point === null)) return;
      rule = {
        kind: 'icon-set',
        range,
        icons: iconSetSelect.value as ConditionalIconSet,
        showValue: !iconOnlyCk.checked,
        thresholds: iconThresholds as ConditionalScalePoint[],
        reverseOrder: iconReverseCk.checked,
      };
    } else if (kind === 'top-bottom') {
      const n = Number.parseInt(tbNInput.value, 10);
      if (!Number.isFinite(n) || n <= 0) return;
      rule = {
        kind: 'top-bottom',
        range,
        mode: tbModeSelect.value as 'top' | 'bottom',
        n,
        percent: tbPercentCk.checked,
        apply: collectConditionalApplyPatch(sharedApplyControls),
      };
    } else if (kind === 'average') {
      rule = {
        kind: 'average',
        range,
        mode: averageModeSelect.value as AverageMode,
        apply: collectConditionalApplyPatch(sharedApplyControls),
      };
    } else if (kind === 'formula') {
      const f = formulaInput.value.trim();
      if (f === '') return;
      rule = {
        kind: 'formula',
        range,
        formula: f,
        apply: collectConditionalApplyPatch(sharedApplyControls),
      };
    } else if (kind === 'text-contains') {
      const text = textContainsInput.value.trim();
      if (text === '') return;
      rule = {
        kind: 'text-contains',
        range,
        text,
        caseSensitive: caseSensitiveCk.checked,
        apply: collectConditionalApplyPatch(sharedApplyControls),
      };
    } else if (kind === 'date-occurring') {
      rule = {
        kind: 'date-occurring',
        range,
        period: datePeriodSelect.value as DatePeriod,
        apply: collectConditionalApplyPatch(sharedApplyControls),
      };
    } else if (
      kind === 'duplicates' ||
      kind === 'unique' ||
      kind === 'blanks' ||
      kind === 'non-blanks' ||
      kind === 'errors' ||
      kind === 'no-errors'
    ) {
      rule = {
        kind,
        range,
        apply: collectConditionalApplyPatch(sharedApplyControls),
      };
    }
    if (!rule) return;
    const newRule = rule;
    const editIndex = currentEditIndex;
    recordConditionalRulesChange(history, store, () => {
      if (currentMode === 'edit' && editIndex !== null) {
        store.setState((state) => {
          if (!state.conditional.rules[editIndex]) return state;
          const rules = [...state.conditional.rules];
          rules[editIndex] = newRule;
          return { ...state, conditional: { rules } };
        });
      } else {
        mutators.addConditionalRule(store, newRule);
      }
    });
    renderRules();
    if (currentMode === 'new' || currentMode === 'edit') api.close();
  };

  const onClearAll = (): void => {
    recordConditionalRulesChange(history, store, () => {
      mutators.clearConditionalRules(store);
    });
    renderRules();
  };
  const onClose = (): void => api.close();

  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      onAdd();
    }
  };

  shell.on(kindSelect, 'change', syncSubforms);
  shell.on(opSelect, 'change', syncCellValueOp);
  shell.on(useThreeCk, 'change', syncThreeStops);
  shell.on(iconSetSelect, 'change', syncIconThresholds);
  shell.on(cellPresetSelect, 'change', syncCellPreset);
  shell.on(sharedPresetSelect, 'change', syncSharedPreset);
  shell.on(addBtn, 'click', onAdd);
  shell.on(clearAllBtn, 'click', onClearAll);
  shell.on(closeBtn, 'click', onClose);
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  const api: ConditionalDialogHandle = {
    open(options = {}): void {
      currentMode = options.mode ?? 'manage';
      currentEditIndex = currentMode === 'edit' ? (options.editIndex ?? null) : null;
      const sel = store.getState().selection.range;
      rangeInput.value = formatRange(sel);
      kindSelect.value = options.kind ?? 'cell-value';
      opSelect.value = options.cellValueOp ?? '>';
      valueAInput.value = '0';
      valueBInput.value = '0';
      useThreeCk.checked = false;
      scaleMin.type.value = 'min';
      scaleMin.value.value = '0';
      scaleMax.type.value = 'max';
      scaleMax.value.value = '100';
      scaleMid.type.value = 'percentile';
      scaleMid.value.value = '50';
      iconSetSelect.value = 'arrows3';
      iconReverseCk.checked = false;
      iconOnlyCk.checked = false;
      for (const control of iconThresholdControls) {
        control.type.value = 'percent';
        control.value.value = '';
        control.type.dispatchEvent(new Event('change'));
      }
      tbModeSelect.value = options.topBottomMode ?? 'top';
      tbPercentCk.checked = options.topBottomPercent ?? false;
      averageModeSelect.value = options.averageMode ?? 'above';
      textContainsInput.value = options.text ?? '';
      caseSensitiveCk.checked = false;
      datePeriodSelect.value = options.datePeriod ?? 'today';
      cellPresetSelect.value = 'red-fill';
      sharedPresetSelect.value = 'red-fill';
      syncSubforms();
      syncCellValueOp();
      syncThreeStops();
      syncIconThresholds();
      scaleMin.type.dispatchEvent(new Event('change'));
      scaleMid.type.dispatchEvent(new Event('change'));
      scaleMax.type.dispatchEvent(new Event('change'));
      syncCellPreset();
      syncSharedPreset();
      if (currentMode === 'edit' && currentEditIndex !== null) {
        const rule = store.getState().conditional.rules[currentEditIndex];
        if (rule) {
          populateRuleForm(rule);
        } else {
          currentMode = 'new';
          currentEditIndex = null;
        }
      }
      syncDialogMode();
      renderRules();
      shell.open();
      requestAnimationFrame(() => {
        rangeInput.focus();
      });
    },
    close(): void {
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };

  return api;
}
