import type { Range } from '../engine/types.js';
import { type Strings, defaultStrings } from '../i18n/strings.js';
import {
  type CellFormat,
  type ConditionalRule,
  type SpreadsheetStore,
  mutators,
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
 * Excel parity is intentionally narrow — three rule kinds (cell-value,
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
  body.className = 'fc-fmtdlg__body';
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
  clearAllBtn.className = 'fc-fmtdlg__btn';
  clearAllBtn.textContent = t.clearAll;
  body.appendChild(clearAllBtn);

  // ── Add-rule form ──────────────────────────────────────────────────────
  const formLegend = document.createElement('div');
  formLegend.className = 'fc-conddlg__legend';
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
  const syncSubforms = (): void => {
    const kind = kindSelect.value as RuleKind;
    cellValueGroup.hidden = kind !== 'cell-value';
    colorScaleGroup.hidden = kind !== 'color-scale';
    dataBarGroup.hidden = kind !== 'data-bar';
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
    if (rule.kind === 'cell-value') {
      const opLabel = opOptions.find((o) => o.id === rule.op)?.label ?? rule.op;
      const tail =
        rule.op === 'between' || rule.op === 'not-between'
          ? `${rule.a} – ${rule.b ?? rule.a}`
          : `${rule.a}`;
      return `${range} · ${t.kindCellValue} (${opLabel} ${tail})`;
    }
    if (rule.kind === 'color-scale') {
      return `${range} · ${t.kindColorScale} (${rule.stops.length} stop)`;
    }
    return `${range} · ${t.kindDataBar}`;
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
    } else {
      mutators.addConditionalRule(store, {
        kind: 'data-bar',
        range,
        color: barColorInput.value,
        showValue: showValueCk.checked,
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
