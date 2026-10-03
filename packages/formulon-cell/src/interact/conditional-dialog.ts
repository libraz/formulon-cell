import { type History, recordConditionalRulesChange } from '../commands/history.js';
import { formatA1Range } from '../engine/address.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  type CellFormat,
  type ConditionalRule,
  mutators,
  type SpreadsheetStore,
} from '../store/store.js';
import {
  appendConditionalApplyFormatControls,
  applyPatchToConditionalApplyControls,
  applyPresetPatchToConditionalApplyControls,
  collectConditionalApplyPatch,
} from './conditional-apply-controls.js';
import { appendColorScaleForm } from './conditional-color-scale-form.js';
import { appendDataBarForm } from './conditional-data-bar-form.js';
import {
  type AverageMode,
  type CellValueOp,
  type DatePeriod,
  type FormatPreset,
  formatPresetPatch,
  parseRange,
  type RuleKind,
} from './conditional-dialog-spec.js';
import { conditionalSelect } from './conditional-form-controls.js';
import { appendIconSetForm } from './conditional-icon-set-form.js';
import { appendDialogButton, createDialogShell } from './dialog-shell.js';
import { attachFormatDialog, type FormatDialogHandle } from './format-dialog.js';
import { attachRangePickerButton } from './range-picker-control.js';

export interface ConditionalDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
  strings?: Strings;
  onChanged?: () => void;
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
  const styleSelect = conditionalSelect(
    [
      { value: 'two-color-scale', label: t.styleTwoColorScale },
      { value: 'three-color-scale', label: t.styleThreeColorScale },
      { value: 'data-bar', label: t.kindDataBar },
      { value: 'icon-set', label: t.kindIconSet },
      { value: 'classic', label: t.styleClassic },
    ],
    'classic',
  );
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
    getValue: () => formatA1Range(store.getState().selection.range, { collapse: false }),
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
  const kindSelect = conditionalSelect(kindOptions.map((o) => ({ value: o.id, label: o.label })));
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
  const opSelect = conditionalSelect(opOptions.map((o) => ({ value: o.id, label: o.label })));
  opRow.append(opLabel, opSelect);
  cellValueGroup.appendChild(opRow);

  const valueARow = document.createElement('label');
  valueARow.className = 'fc-fmtdlg__row';
  const valueALabel = document.createElement('span');
  valueALabel.textContent = t.valueA;
  const valueAInput = document.createElement('input');
  valueAInput.type = 'text';
  valueAInput.value = '0';
  valueARow.append(valueALabel, valueAInput);
  cellValueGroup.appendChild(valueARow);

  const valueBRow = document.createElement('label');
  valueBRow.className = 'fc-fmtdlg__row';
  const valueBLabel = document.createElement('span');
  valueBLabel.textContent = t.valueB;
  const valueBInput = document.createElement('input');
  valueBInput.type = 'text';
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
    { id: 'light-red-fill', label: t.formatLightRedFill },
    { id: 'red-text', label: t.formatRedText },
    { id: 'red-border', label: t.formatRedBorder },
    { id: 'custom', label: t.formatCustom },
  ];
  const cellPresetSelect = conditionalSelect(
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

  const colorScale = appendColorScaleForm(form, t);
  const dataBar = appendDataBarForm(form, t, () => currentMode === 'edit');
  const iconSet = appendIconSetForm(form, t, () => currentMode === 'edit');

  // ── Top/Bottom subform ─────────────────────────────────────────────────
  const topBottomGroup = document.createElement('div');
  topBottomGroup.className = 'fc-conddlg__sub';
  form.appendChild(topBottomGroup);

  const tbModeRow = document.createElement('label');
  tbModeRow.className = 'fc-fmtdlg__row';
  const tbModeLabel = document.createElement('span');
  tbModeLabel.textContent = t.topBottomMode;
  const tbModeSelect = conditionalSelect([
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
    { id: 'above-std-dev', label: t.averageAboveStdDev },
    { id: 'below-std-dev', label: t.averageBelowStdDev },
  ];
  const averageModeSelect = conditionalSelect(
    averageModeOptions.map((o) => ({ value: o.id, label: o.label })),
  );
  averageModeRow.append(averageModeLabel, averageModeSelect);
  averageGroup.appendChild(averageModeRow);
  const averageModeLabelFor = (id: AverageMode): string =>
    averageModeOptions.find((option) => option.id === id)?.label ?? id;
  const averageStdDevRow = document.createElement('label');
  averageStdDevRow.className = 'fc-fmtdlg__row';
  const averageStdDevLabel = document.createElement('span');
  averageStdDevLabel.textContent = t.averageStdDevTier;
  const averageStdDevSelect = conditionalSelect([
    { value: '1', label: '1' },
    { value: '2', label: '2' },
    { value: '3', label: '3' },
  ]);
  averageStdDevRow.append(averageStdDevLabel, averageStdDevSelect);
  averageGroup.appendChild(averageStdDevRow);

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

  const textContainsModeRow = document.createElement('label');
  textContainsModeRow.className = 'fc-fmtdlg__row';
  const textContainsModeLabel = document.createElement('span');
  textContainsModeLabel.textContent = t.textContainsMode;
  const textContainsModeSelect = conditionalSelect([
    { value: 'contains', label: t.textContainsContains },
    { value: 'not-contains', label: t.textContainsNotContains },
    { value: 'begins-with', label: t.textContainsBeginsWith },
    { value: 'ends-with', label: t.textContainsEndsWith },
  ]);
  textContainsModeRow.append(textContainsModeLabel, textContainsModeSelect);
  textContainsGroup.appendChild(textContainsModeRow);

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
  const datePeriodSelect = conditionalSelect(
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
    colorScale.group.hidden = kind !== 'color-scale';
    dataBar.group.hidden = kind !== 'data-bar';
    iconSet.group.hidden = kind !== 'icon-set';
    topBottomGroup.hidden = kind !== 'top-bottom';
    averageGroup.hidden = kind !== 'average';
    formulaGroup.hidden = kind !== 'formula';
    textContainsGroup.hidden = kind !== 'text-contains';
    dateOccurringGroup.hidden = kind !== 'date-occurring';
    applyGroup.hidden = !APPLY_KINDS.has(kind);
    averageStdDevRow.hidden =
      kind !== 'average' ||
      (averageModeSelect.value !== 'above-std-dev' && averageModeSelect.value !== 'below-std-dev');
  };
  const syncRuleStyle = (): void => {
    const style = styleSelect.value;
    if (style === 'two-color-scale' || style === 'three-color-scale') {
      kindSelect.value = 'color-scale';
      colorScale.useThreeCk.checked = style === 'three-color-scale';
    } else if (style === 'data-bar' || style === 'icon-set') {
      kindSelect.value = style;
    }
    kindRow.hidden = style !== 'classic';
    syncSubforms();
  };
  const syncCellValueOp = (): void => {
    const op = opSelect.value as CellValueOp;
    valueBRow.hidden = op !== 'between' && op !== 'not-between';
  };
  averageModeSelect.addEventListener('change', syncSubforms);
  let dxfFormatDialog: FormatDialogHandle | null = null;
  const getDxfFormatDialog = (): FormatDialogHandle => {
    if (!dxfFormatDialog) dxfFormatDialog = attachFormatDialog({ host, store, strings, history });
    return dxfFormatDialog;
  };
  let cellCustomStyle: Partial<CellFormat> | null = null;
  let sharedCustomStyle: Partial<CellFormat> | null = null;
  const syncPresetPreview = (
    preview: HTMLElement,
    preset: FormatPreset,
    customStyle: Partial<CellFormat> | null,
  ): void => {
    const patch = preset === 'custom' && customStyle ? customStyle : formatPresetPatch(preset);
    preview.style.color = patch.color ?? '#201f1e';
    preview.style.background = patch.fill ?? 'transparent';
  };
  const syncCellPreset = (): void => {
    const preset = cellPresetSelect.value as FormatPreset;
    const patch =
      preset === 'custom' && cellCustomStyle ? cellCustomStyle : formatPresetPatch(preset);
    if (preset === 'custom') applyPatchToConditionalApplyControls(cellValueApplyControls, patch);
    else applyPresetPatchToConditionalApplyControls(cellValueApplyControls, patch);
    syncPresetPreview(cellPresetPreview, preset, cellCustomStyle);
  };
  const syncSharedPreset = (): void => {
    const preset = sharedPresetSelect.value as FormatPreset;
    const patch =
      preset === 'custom' && sharedCustomStyle ? sharedCustomStyle : formatPresetPatch(preset);
    if (preset === 'custom') applyPatchToConditionalApplyControls(sharedApplyControls, patch);
    else applyPresetPatchToConditionalApplyControls(sharedApplyControls, patch);
    syncPresetPreview(sharedPresetPreview, preset, sharedCustomStyle);
  };
  const editCustomPreset = (
    controls: Parameters<typeof collectConditionalApplyPatch>[0],
    customStyle: Partial<CellFormat> | null,
    setCustomStyle: (style: Partial<CellFormat>) => void,
    sync: () => void,
  ): void => {
    getDxfFormatDialog().open('number', {
      mode: 'dxf',
      initialFormat: { ...collectConditionalApplyPatch(controls), ...customStyle },
      onApplyDxf: (style) => {
        setCustomStyle(style);
        sync();
      },
    });
  };
  const collectSharedApplyPatch = (): Partial<CellFormat> =>
    sharedPresetSelect.value === 'custom'
      ? { ...sharedCustomStyle, ...collectConditionalApplyPatch(sharedApplyControls) }
      : collectConditionalApplyPatch(sharedApplyControls);

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
        deps.onChanged?.();
        renderRules();
      });
      item.prepend(summary);
      rulesList.appendChild(item);
    });
  };

  const describeRule = (rule: ConditionalRule): string => {
    const range = formatA1Range(rule.range, { collapse: false });
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
        return `${range} · ${t.kindIconSet} (${iconSet.labelFor(rule.icons)}${
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

  const populateRuleForm = (rule: ConditionalRule): void => {
    rangeInput.value = formatA1Range(rule.range, { collapse: false });
    styleSelect.value =
      rule.kind === 'color-scale'
        ? rule.stops.length === 3
          ? 'three-color-scale'
          : 'two-color-scale'
        : rule.kind === 'data-bar' || rule.kind === 'icon-set'
          ? rule.kind
          : 'classic';
    kindSelect.value = rule.kind;
    if (rule.kind === 'cell-value') {
      opSelect.value = rule.op;
      valueAInput.value = String(rule.a);
      valueBInput.value = String(rule.b ?? rule.a);
      applyPatchToConditionalApplyControls(cellValueApplyControls, rule.apply);
    } else if (rule.kind === 'color-scale') {
      colorScale.populate(rule);
    } else if (rule.kind === 'data-bar') {
      dataBar.populate(rule);
    } else if (rule.kind === 'icon-set') {
      iconSet.populate(rule);
    } else if (rule.kind === 'top-bottom') {
      tbModeSelect.value = rule.mode;
      tbNInput.value = String(rule.n);
      tbPercentCk.checked = rule.percent === true;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'average') {
      averageModeSelect.value = rule.mode;
      averageStdDevSelect.value = String(rule.stdDev ?? 1);
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'formula') {
      formulaInput.value = rule.formula;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'text-contains') {
      textContainsModeSelect.value = rule.mode ?? 'contains';
      textContainsInput.value = rule.text;
      caseSensitiveCk.checked = rule.caseSensitive === true;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else if (rule.kind === 'date-occurring') {
      datePeriodSelect.value = rule.period;
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    } else {
      applyPatchToConditionalApplyControls(sharedApplyControls, rule.apply);
    }
    syncRuleStyle();
    syncCellValueOp();
    colorScale.syncThreeStops();
    iconSet.syncThresholds();
    colorScale.refreshScaleTypes();
    iconSet.refreshThresholdTypes();
  };

  const onAdd = (): void => {
    const fallback = store.getState().selection.range;
    const range = parseRange(rangeInput.value, fallback);
    const kind = kindSelect.value as RuleKind;
    let rule: ConditionalRule | null = null;
    if (kind === 'cell-value') {
      const op = opSelect.value as CellValueOp;
      const parseCellValueBoundary = (raw: string): number | string | null => {
        const text = raw.trim();
        if (text === '') return null;
        const num = Number(text);
        return Number.isFinite(num) ? num : text;
      };
      const a = parseCellValueBoundary(valueAInput.value);
      const b = parseCellValueBoundary(valueBInput.value);
      if (a === null) return;
      if ((op === 'between' || op === 'not-between') && b === null) return;
      const preset = cellPresetSelect.value as FormatPreset;
      const applyPatch =
        preset === 'custom'
          ? { ...cellCustomStyle, ...collectConditionalApplyPatch(cellValueApplyControls) }
          : {
              ...collectConditionalApplyPatch(cellValueApplyControls),
              ...formatPresetPatch(preset),
            };
      rule = {
        kind: 'cell-value',
        range,
        op,
        a,
        ...(op === 'between' || op === 'not-between' ? { b: b as number | string } : {}),
        apply: applyPatch,
      };
    } else if (kind === 'color-scale') {
      rule = colorScale.collect(range);
    } else if (kind === 'data-bar') {
      rule = dataBar.collect(range);
    } else if (kind === 'icon-set') {
      rule = iconSet.collect(range);
    } else if (kind === 'top-bottom') {
      const n = Number.parseInt(tbNInput.value, 10);
      if (!Number.isFinite(n) || n <= 0) return;
      rule = {
        kind: 'top-bottom',
        range,
        mode: tbModeSelect.value as 'top' | 'bottom',
        n,
        percent: tbPercentCk.checked,
        apply: collectSharedApplyPatch(),
      };
    } else if (kind === 'average') {
      const averageMode = averageModeSelect.value as AverageMode;
      rule = {
        kind: 'average',
        range,
        mode: averageMode,
        ...(averageMode === 'above-std-dev' || averageMode === 'below-std-dev'
          ? { stdDev: Number(averageStdDevSelect.value) as 1 | 2 | 3 }
          : {}),
        apply: collectSharedApplyPatch(),
      };
    } else if (kind === 'formula') {
      const f = formulaInput.value.trim();
      if (f === '') return;
      rule = {
        kind: 'formula',
        range,
        formula: f,
        apply: collectSharedApplyPatch(),
      };
    } else if (kind === 'text-contains') {
      const text = textContainsInput.value.trim();
      if (text === '') return;
      rule = {
        kind: 'text-contains',
        range,
        text,
        mode: textContainsModeSelect.value as
          | 'contains'
          | 'not-contains'
          | 'begins-with'
          | 'ends-with',
        caseSensitive: caseSensitiveCk.checked,
        apply: collectSharedApplyPatch(),
      };
    } else if (kind === 'date-occurring') {
      rule = {
        kind: 'date-occurring',
        range,
        period: datePeriodSelect.value as DatePeriod,
        apply: collectSharedApplyPatch(),
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
        apply: collectSharedApplyPatch(),
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
    deps.onChanged?.();
    renderRules();
    if (currentMode === 'new' || currentMode === 'edit') api.close();
  };

  const onClearAll = (): void => {
    recordConditionalRulesChange(history, store, () => {
      mutators.clearConditionalRules(store);
    });
    deps.onChanged?.();
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
  shell.on(styleSelect, 'change', syncRuleStyle);
  shell.on(opSelect, 'change', syncCellValueOp);
  shell.on(colorScale.useThreeCk, 'change', colorScale.syncThreeStops);
  shell.on(iconSet.select, 'change', iconSet.syncThresholds);
  shell.on(cellPresetSelect, 'change', () => {
    syncCellPreset();
    if (cellPresetSelect.value === 'custom') {
      editCustomPreset(
        cellValueApplyControls,
        cellCustomStyle,
        (style) => {
          cellCustomStyle = style;
        },
        syncCellPreset,
      );
    }
  });
  shell.on(sharedPresetSelect, 'change', () => {
    syncSharedPreset();
    if (sharedPresetSelect.value === 'custom') {
      editCustomPreset(
        sharedApplyControls,
        sharedCustomStyle,
        (style) => {
          sharedCustomStyle = style;
        },
        syncSharedPreset,
      );
    }
  });
  shell.on(addBtn, 'click', onAdd);
  shell.on(clearAllBtn, 'click', onClearAll);
  shell.on(closeBtn, 'click', onClose);
  shell.on(overlay, 'keydown', onOverlayKey as EventListener);

  const api: ConditionalDialogHandle = {
    open(options = {}): void {
      currentMode = options.mode ?? 'manage';
      currentEditIndex = currentMode === 'edit' ? (options.editIndex ?? null) : null;
      const sel = store.getState().selection.range;
      rangeInput.value = formatA1Range(sel, { collapse: false });
      kindSelect.value = options.kind ?? 'cell-value';
      styleSelect.value =
        kindSelect.value === 'color-scale'
          ? 'two-color-scale'
          : kindSelect.value === 'data-bar' || kindSelect.value === 'icon-set'
            ? kindSelect.value
            : 'classic';
      opSelect.value = options.cellValueOp ?? '>';
      valueAInput.value = '0';
      valueBInput.value = '0';
      colorScale.reset();
      dataBar.reset();
      iconSet.reset();
      tbModeSelect.value = options.topBottomMode ?? 'top';
      tbPercentCk.checked = options.topBottomPercent ?? false;
      averageModeSelect.value = options.averageMode ?? 'above';
      averageStdDevSelect.value = '1';
      textContainsModeSelect.value = 'contains';
      textContainsInput.value = options.text ?? '';
      caseSensitiveCk.checked = false;
      datePeriodSelect.value = options.datePeriod ?? 'today';
      cellPresetSelect.value = 'red-fill';
      sharedPresetSelect.value = 'red-fill';
      syncRuleStyle();
      syncCellValueOp();
      colorScale.syncThreeStops();
      iconSet.syncThresholds();
      colorScale.refreshScaleTypes();
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
      dxfFormatDialog?.detach();
      shell.dispose();
    },
  };

  return api;
}
