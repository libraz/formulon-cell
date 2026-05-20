import type { CellFormat } from '../store/store.js';

export interface ConditionalApplyFormatLabels {
  fillColor: string;
  fontColor: string;
  bold: string;
  italic: string;
  underline: string;
  strike: string;
}

export interface ConditionalApplyFormatControls {
  fillToggle: HTMLInputElement;
  fillInput: HTMLInputElement;
  fontToggle: HTMLInputElement;
  fontInput: HTMLInputElement;
  bold: HTMLInputElement;
  italic: HTMLInputElement;
  underline: HTMLInputElement;
  strike: HTMLInputElement;
}

export interface ConditionalApplyFormatOptions {
  defaultFill?: string;
  defaultFontColor?: string;
  fillChecked?: boolean;
}

const appendApplyColorRow = (
  parent: HTMLElement,
  labelText: string,
  value: string,
  checked: boolean,
): { toggle: HTMLInputElement; input: HTMLInputElement } => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row';
  const toggle = document.createElement('input');
  toggle.type = 'checkbox';
  toggle.checked = checked;
  toggle.setAttribute('aria-label', labelText);
  const label = document.createElement('span');
  label.textContent = labelText;
  const input = document.createElement('input');
  input.type = 'color';
  input.value = value;
  input.setAttribute('aria-label', labelText);
  row.append(toggle, label, input);
  parent.appendChild(row);
  return { toggle, input };
};

const appendApplyStyleCheckbox = (
  parent: HTMLElement,
  labelText: string,
  fieldName: string,
): HTMLInputElement => {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__check';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.dataset.dialogField = fieldName;
  const label = document.createElement('span');
  label.textContent = labelText;
  wrap.append(input, label);
  parent.appendChild(wrap);
  return input;
};

export const appendConditionalApplyFormatControls = (
  parent: HTMLElement,
  labels: ConditionalApplyFormatLabels,
  opts: ConditionalApplyFormatOptions = {},
): ConditionalApplyFormatControls => {
  const fill = appendApplyColorRow(
    parent,
    labels.fillColor,
    opts.defaultFill ?? '#ffeb3b',
    opts.fillChecked ?? true,
  );
  const font = appendApplyColorRow(
    parent,
    labels.fontColor,
    opts.defaultFontColor ?? '#000000',
    false,
  );
  const styleRow = document.createElement('div');
  styleRow.className = 'fc-fmtdlg__row';
  parent.appendChild(styleRow);
  return {
    fillToggle: fill.toggle,
    fillInput: fill.input,
    fontToggle: font.toggle,
    fontInput: font.input,
    bold: appendApplyStyleCheckbox(styleRow, labels.bold, 'bold'),
    italic: appendApplyStyleCheckbox(styleRow, labels.italic, 'italic'),
    underline: appendApplyStyleCheckbox(styleRow, labels.underline, 'underline'),
    strike: appendApplyStyleCheckbox(styleRow, labels.strike, 'strike'),
  };
};

export const collectConditionalApplyPatch = (
  controls: ConditionalApplyFormatControls,
): Partial<CellFormat> => {
  const apply: Partial<CellFormat> = {};
  if (controls.fillToggle.checked) apply.fill = controls.fillInput.value;
  if (controls.fontToggle.checked) apply.color = controls.fontInput.value;
  if (controls.bold.checked) apply.bold = true;
  if (controls.italic.checked) apply.italic = true;
  if (controls.underline.checked) apply.underline = true;
  if (controls.strike.checked) apply.strike = true;
  return apply;
};

export const applyPatchToConditionalApplyControls = (
  controls: ConditionalApplyFormatControls,
  patch: Partial<CellFormat> | undefined,
): void => {
  controls.fillToggle.checked = !!patch?.fill;
  if (patch?.fill) controls.fillInput.value = patch.fill;
  controls.fontToggle.checked = !!patch?.color;
  if (patch?.color) controls.fontInput.value = patch.color;
  controls.bold.checked = patch?.bold === true;
  controls.italic.checked = patch?.italic === true;
  controls.underline.checked = patch?.underline === true;
  controls.strike.checked = patch?.strike === true;
};

export const applyPresetPatchToConditionalApplyControls = (
  controls: ConditionalApplyFormatControls,
  patch: Partial<CellFormat>,
): void => {
  controls.fillToggle.checked = !!patch.fill;
  controls.fontToggle.checked = !!patch.color;
  if (patch.fill) controls.fillInput.value = patch.fill;
  if (patch.color) controls.fontInput.value = patch.color;
};
