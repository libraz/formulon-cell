import type { Range } from '../engine/types.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import {
  type CellFormat,
  type ConditionalIconSet,
  type ConditionalRule,
  mutators,
  type SpreadsheetStore,
} from '../store/store.js';

export interface ConditionalDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  strings?: Strings;
}

export interface ConditionalDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

type RuleKind = ConditionalRule['kind'];
type CellValueOp = '>' | '<' | '>=' | '<=' | '=' | '<>' | 'between' | 'not-between';

const colLetters = (col: number): string => {
  let n = col;
  let s = '';
  while (true) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
    if (n < 0) break;
  }
  return s;
};

const formatRange = (r: Range): string =>
  `${colLetters(r.c0)}${r.r0 + 1}:${colLetters(r.c1)}${r.r1 + 1}`;

const parseRange = (raw: string, fallback: Range): Range => {
  const m = raw
    .trim()
    .toUpperCase()
    .match(/^\$?([A-Z]+)\$?([1-9][0-9]*):\$?([A-Z]+)\$?([1-9][0-9]*)$/);
  if (!m) return fallback;
  const lettersToCol = (letters: string): number => {
    let c = 0;
    for (let i = 0; i < letters.length; i += 1) c = c * 26 + (letters.charCodeAt(i) - 64);
    return c - 1;
  };
  const c0 = lettersToCol(m[1] ?? '');
  const r0 = Number.parseInt(m[2] ?? '', 10) - 1;
  const c1 = lettersToCol(m[3] ?? '');
  const r1 = Number.parseInt(m[4] ?? '', 10) - 1;
  if (c0 < 0 || r0 < 0 || c1 < 0 || r1 < 0) return fallback;
  return {
    sheet: fallback.sheet,
    r0: Math.min(r0, r1),
    c0: Math.min(c0, c1),
    r1: Math.max(r0, r1),
    c1: Math.max(c0, c1),
  };
};

/**
 * Manage conditional formatting rules: list / add / remove.
 * Spreadsheet parity is intentionally narrow — three rule kinds (cell-value,
 * color-scale, data-bar) and the renderer respects whichever fields apply.
 */
export function attachConditionalDialog(deps: ConditionalDialogDeps): ConditionalDialogHandle {
  const { host, store } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.conditionalDialog;

  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg fc-conddlg';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel fc-conddlg__panel';
  overlay.appendChild(panel);

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

  const clearAllBtn = document.createElement('button');
  clearAllBtn.type = 'button';
  clearAllBtn.className = 'fc-fmtdlg__btn fc-conddlg__clear';
  clearAllBtn.textContent = t.clearAll;
  body.appendChild(clearAllBtn);

  // ── Add-rule form ──────────────────────────────────────────────────────
  const formLegend = document.createElement('div');
  formLegend.className = 'fc-conddlg__legend fc-conddlg__form-legend';
  formLegend.textContent = t.addRule;
  body.appendChild(formLegend);

  const form = document.createElement('div');
  form.className = 'fc-conddlg__form';
  body.appendChild(form);

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
  form.appendChild(rangeRow);

  // Kind
  const kindRow = document.createElement('label');
  kindRow.className = 'fc-fmtdlg__row';
  const kindLabel = document.createElement('span');
  kindLabel.textContent = t.kindLabel;
  const kindSelect = document.createElement('select');
  const kindOptions: { id: RuleKind; label: string }[] = [
    { id: 'cell-value', label: t.kindCellValue },
    { id: 'color-scale', label: t.kindColorScale },
    { id: 'data-bar', label: t.kindDataBar },
    { id: 'icon-set', label: t.kindIconSet },
    { id: 'top-bottom', label: t.kindTopBottom },
    { id: 'formula', label: t.kindFormula },
    { id: 'duplicates', label: t.kindDuplicates },
    { id: 'unique', label: t.kindUnique },
    { id: 'blanks', label: t.kindBlanks },
    { id: 'non-blanks', label: t.kindNonBlanks },
    { id: 'errors', label: t.kindErrors },
    { id: 'no-errors', label: t.kindNoErrors },
  ];
  for (const o of kindOptions) {
    const opt = document.createElement('option');
    opt.value = o.id;
    opt.textContent = o.label;
    kindSelect.appendChild(opt);
  }
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
  const opSelect = document.createElement('select');
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
  for (const o of opOptions) {
    const opt = document.createElement('option');
    opt.value = o.id;
    opt.textContent = o.label;
    opSelect.appendChild(opt);
  }
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
  const applyRow1 = document.createElement('div');
  applyRow1.className = 'fc-fmtdlg__row';
  cellValueGroup.appendChild(applyRow1);
  const fillLabel = document.createElement('span');
  fillLabel.textContent = t.fillColor;
  const fillInput = document.createElement('input');
  fillInput.type = 'color';
  fillInput.value = '#ffeb3b';
  const fillToggle = document.createElement('input');
  fillToggle.type = 'checkbox';
  fillToggle.checked = true;
  applyRow1.append(fillToggle, fillLabel, fillInput);

  const applyRow2 = document.createElement('div');
  applyRow2.className = 'fc-fmtdlg__row';
  cellValueGroup.appendChild(applyRow2);
  const fontLabel = document.createElement('span');
  fontLabel.textContent = t.fontColor;
  const fontInput = document.createElement('input');
  fontInput.type = 'color';
  fontInput.value = '#000000';
  const fontToggle = document.createElement('input');
  fontToggle.type = 'checkbox';
  applyRow2.append(fontToggle, fontLabel, fontInput);

  const styleRow = document.createElement('div');
  styleRow.className = 'fc-fmtdlg__row';
  cellValueGroup.appendChild(styleRow);
  const makeApplyCheckbox = (label: string): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__check';
    const ck = document.createElement('input');
    ck.type = 'checkbox';
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(ck, span);
    styleRow.appendChild(wrap);
    return ck;
  };
  const cvBoldCk = makeApplyCheckbox(t.bold);
  const cvItalicCk = makeApplyCheckbox(t.italic);
  const cvUnderlineCk = makeApplyCheckbox(t.underline);
  const cvStrikeCk = makeApplyCheckbox(t.strike);

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
  stopMinRow.append(stopMinLabel, stopMinInput);
  colorScaleGroup.appendChild(stopMinRow);

  const stopMidRow = document.createElement('label');
  stopMidRow.className = 'fc-fmtdlg__row';
  const stopMidLabel = document.createElement('span');
  stopMidLabel.textContent = t.stopMid;
  const stopMidInput = document.createElement('input');
  stopMidInput.type = 'color';
  stopMidInput.value = '#ffeb84';
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
  stopMaxRow.append(stopMaxLabel, stopMaxInput);
  colorScaleGroup.appendChild(stopMaxRow);

  // ── Data bar subform ───────────────────────────────────────────────────
  const dataBarGroup = document.createElement('div');
  dataBarGroup.className = 'fc-conddlg__sub';
  form.appendChild(dataBarGroup);

  const barColorRow = document.createElement('label');
  barColorRow.className = 'fc-fmtdlg__row';
  const barColorLabel = document.createElement('span');
  barColorLabel.textContent = t.barColor;
  const barColorInput = document.createElement('input');
  barColorInput.type = 'color';
  barColorInput.value = '#638ec6';
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
  const iconSetSelect = document.createElement('select');
  const iconSetOptions: { id: ConditionalIconSet; label: string }[] = [
    { id: 'arrows3', label: t.iconSetArrows3 },
    { id: 'arrows5', label: t.iconSetArrows5 },
    { id: 'traffic3', label: t.iconSetTraffic3 },
    { id: 'stars3', label: t.iconSetStars3 },
  ];
  for (const o of iconSetOptions) {
    const opt = document.createElement('option');
    opt.value = o.id;
    opt.textContent = o.label;
    iconSetSelect.appendChild(opt);
  }
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

  // ── Top/Bottom subform ─────────────────────────────────────────────────
  const topBottomGroup = document.createElement('div');
  topBottomGroup.className = 'fc-conddlg__sub';
  form.appendChild(topBottomGroup);

  const tbModeRow = document.createElement('label');
  tbModeRow.className = 'fc-fmtdlg__row';
  const tbModeLabel = document.createElement('span');
  tbModeLabel.textContent = t.topBottomMode;
  const tbModeSelect = document.createElement('select');
  for (const o of [
    { id: 'top', label: `${t.kindTopBottom} ↑` },
    { id: 'bottom', label: `${t.kindTopBottom} ↓` },
  ] as const) {
    const opt = document.createElement('option');
    opt.value = o.id;
    opt.textContent = o.label;
    tbModeSelect.appendChild(opt);
  }
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

  // ── Apply-format shared by top-bottom / formula / dups / unique /
  //    blanks / non-blanks / errors / no-errors. We re-use the same
  //    fill/font/style controls from the cell-value subform so the
  //    "apply when matched" surface stays consistent.
  const applyGroup = document.createElement('div');
  applyGroup.className = 'fc-conddlg__sub';
  form.appendChild(applyGroup);

  const applyFillRow = document.createElement('div');
  applyFillRow.className = 'fc-fmtdlg__row';
  applyGroup.appendChild(applyFillRow);
  const applyFillToggle = document.createElement('input');
  applyFillToggle.type = 'checkbox';
  applyFillToggle.checked = true;
  const applyFillLabel = document.createElement('span');
  applyFillLabel.textContent = t.fillColor;
  const applyFillInput = document.createElement('input');
  applyFillInput.type = 'color';
  applyFillInput.value = '#ffeb3b';
  applyFillRow.append(applyFillToggle, applyFillLabel, applyFillInput);

  const applyFontRow = document.createElement('div');
  applyFontRow.className = 'fc-fmtdlg__row';
  applyGroup.appendChild(applyFontRow);
  const applyFontToggle = document.createElement('input');
  applyFontToggle.type = 'checkbox';
  const applyFontLabel = document.createElement('span');
  applyFontLabel.textContent = t.fontColor;
  const applyFontInput = document.createElement('input');
  applyFontInput.type = 'color';
  applyFontInput.value = '#000000';
  applyFontRow.append(applyFontToggle, applyFontLabel, applyFontInput);

  const applyStyleRow = document.createElement('div');
  applyStyleRow.className = 'fc-fmtdlg__row';
  applyGroup.appendChild(applyStyleRow);
  const makeApplyCheckboxIn = (parent: HTMLElement, label: string): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-fmtdlg__check';
    const ck = document.createElement('input');
    ck.type = 'checkbox';
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(ck, span);
    parent.appendChild(wrap);
    return ck;
  };
  const applyBoldCk = makeApplyCheckboxIn(applyStyleRow, t.bold);
  const applyItalicCk = makeApplyCheckboxIn(applyStyleRow, t.italic);
  const applyUnderlineCk = makeApplyCheckboxIn(applyStyleRow, t.underline);
  const applyStrikeCk = makeApplyCheckboxIn(applyStyleRow, t.strike);

  // Add button
  const addRow = document.createElement('div');
  addRow.className = 'fc-fmtdlg__row fc-conddlg__addrow';
  const addBtn = document.createElement('button');
  addBtn.type = 'button';
  addBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
  addBtn.textContent = t.addRule;
  addRow.appendChild(addBtn);
  form.appendChild(addRow);

  // Footer
  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);
  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.className = 'fc-fmtdlg__btn';
  closeBtn.textContent = t.close;
  footer.appendChild(closeBtn);

  host.appendChild(overlay);

  // ── Behaviour ──────────────────────────────────────────────────────────
  /** Kinds that re-use the shared `applyGroup` (fill/font/style) controls
   *  for their "apply when matched" format. cell-value carries its own
   *  controls inside `cellValueGroup` and so is excluded here. */
  const APPLY_KINDS: ReadonlySet<RuleKind> = new Set([
    'top-bottom',
    'formula',
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
    formulaGroup.hidden = kind !== 'formula';
    applyGroup.hidden = !APPLY_KINDS.has(kind);
  };
  const syncCellValueOp = (): void => {
    const op = opSelect.value as CellValueOp;
    valueBRow.hidden = op !== 'between' && op !== 'not-between';
  };
  const syncThreeStops = (): void => {
    stopMidRow.hidden = !useThreeCk.checked;
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
      const removeBtn = document.createElement('button');
      removeBtn.type = 'button';
      removeBtn.className = 'fc-fmtdlg__btn';
      removeBtn.textContent = t.removeRule;
      removeBtn.addEventListener('click', () => {
        mutators.removeConditionalRuleAt(store, idx);
        renderRules();
      });
      item.append(summary, removeBtn);
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
        return `${range} · ${t.kindColorScale} (${rule.stops.length} stop)`;
      case 'data-bar':
        return `${range} · ${t.kindDataBar}`;
      case 'icon-set':
        return `${range} · ${t.kindIconSet} (${rule.icons})`;
      case 'top-bottom': {
        const pct = rule.percent ? '%' : '';
        return `${range} · ${t.kindTopBottom} (${rule.mode} ${rule.n}${pct})`;
      }
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

  const collectApplyPatch = (): Partial<CellFormat> => {
    const apply: Partial<CellFormat> = {};
    if (applyFillToggle.checked) apply.fill = applyFillInput.value;
    if (applyFontToggle.checked) apply.color = applyFontInput.value;
    if (applyBoldCk.checked) apply.bold = true;
    if (applyItalicCk.checked) apply.italic = true;
    if (applyUnderlineCk.checked) apply.underline = true;
    if (applyStrikeCk.checked) apply.strike = true;
    return apply;
  };

  const onAdd = (): void => {
    const fallback = store.getState().selection.range;
    const range = parseRange(rangeInput.value, fallback);
    const kind = kindSelect.value as RuleKind;
    if (kind === 'cell-value') {
      const op = opSelect.value as CellValueOp;
      const a = Number.parseFloat(valueAInput.value);
      const b = Number.parseFloat(valueBInput.value);
      if (!Number.isFinite(a)) return;
      const applyPatch: Partial<CellFormat> = {};
      if (fillToggle.checked) applyPatch.fill = fillInput.value;
      if (fontToggle.checked) applyPatch.color = fontInput.value;
      if (cvBoldCk.checked) applyPatch.bold = true;
      if (cvItalicCk.checked) applyPatch.italic = true;
      if (cvUnderlineCk.checked) applyPatch.underline = true;
      if (cvStrikeCk.checked) applyPatch.strike = true;
      const rule: ConditionalRule = {
        kind: 'cell-value',
        range,
        op,
        a,
        ...(op === 'between' || op === 'not-between' ? { b } : {}),
        apply: applyPatch,
      };
      mutators.addConditionalRule(store, rule);
    } else if (kind === 'color-scale') {
      const stops: [string, string] | [string, string, string] = useThreeCk.checked
        ? [stopMinInput.value, stopMidInput.value, stopMaxInput.value]
        : [stopMinInput.value, stopMaxInput.value];
      mutators.addConditionalRule(store, { kind: 'color-scale', range, stops });
    } else if (kind === 'data-bar') {
      mutators.addConditionalRule(store, {
        kind: 'data-bar',
        range,
        color: barColorInput.value,
        showValue: showValueCk.checked,
      });
    } else if (kind === 'icon-set') {
      mutators.addConditionalRule(store, {
        kind: 'icon-set',
        range,
        icons: iconSetSelect.value as ConditionalIconSet,
        reverseOrder: iconReverseCk.checked,
      });
    } else if (kind === 'top-bottom') {
      const n = Number.parseInt(tbNInput.value, 10);
      if (!Number.isFinite(n) || n <= 0) return;
      mutators.addConditionalRule(store, {
        kind: 'top-bottom',
        range,
        mode: tbModeSelect.value as 'top' | 'bottom',
        n,
        percent: tbPercentCk.checked,
        apply: collectApplyPatch(),
      });
    } else if (kind === 'formula') {
      const f = formulaInput.value.trim();
      if (f === '') return;
      mutators.addConditionalRule(store, {
        kind: 'formula',
        range,
        formula: f,
        apply: collectApplyPatch(),
      });
    } else if (
      kind === 'duplicates' ||
      kind === 'unique' ||
      kind === 'blanks' ||
      kind === 'non-blanks' ||
      kind === 'errors' ||
      kind === 'no-errors'
    ) {
      mutators.addConditionalRule(store, {
        kind,
        range,
        apply: collectApplyPatch(),
      });
    }
    renderRules();
  };

  const onClearAll = (): void => {
    mutators.clearConditionalRules(store);
    renderRules();
  };
  const onClose = (): void => api.close();

  const onOverlayClick = (e: MouseEvent): void => {
    if (e.target === overlay) api.close();
  };
  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  kindSelect.addEventListener('change', syncSubforms);
  opSelect.addEventListener('change', syncCellValueOp);
  useThreeCk.addEventListener('change', syncThreeStops);
  addBtn.addEventListener('click', onAdd);
  clearAllBtn.addEventListener('click', onClearAll);
  closeBtn.addEventListener('click', onClose);
  overlay.addEventListener('click', onOverlayClick);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: ConditionalDialogHandle = {
    open(): void {
      const sel = store.getState().selection.range;
      rangeInput.value = formatRange(sel);
      kindSelect.value = 'cell-value';
      opSelect.value = '>';
      valueAInput.value = '0';
      valueBInput.value = '0';
      useThreeCk.checked = false;
      syncSubforms();
      syncCellValueOp();
      syncThreeStops();
      renderRules();
      overlay.hidden = false;
      requestAnimationFrame(() => {
        rangeInput.focus();
      });
    },
    close(): void {
      overlay.hidden = true;
      host.focus();
    },
    detach(): void {
      kindSelect.removeEventListener('change', syncSubforms);
      opSelect.removeEventListener('change', syncCellValueOp);
      useThreeCk.removeEventListener('change', syncThreeStops);
      addBtn.removeEventListener('click', onAdd);
      clearAllBtn.removeEventListener('click', onClearAll);
      closeBtn.removeEventListener('click', onClose);
      overlay.removeEventListener('click', onOverlayClick);
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };

  return api;
}
