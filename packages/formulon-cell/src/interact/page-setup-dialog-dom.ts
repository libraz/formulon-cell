// DOM builders shared by the Page Setup dialog's tabs.

export function makeRow(label: string): { row: HTMLDivElement; valueCell: HTMLSpanElement } {
  const row = document.createElement('div');
  row.className = 'fc-pgsetup__row fc-fmtdlg__row';
  const labelSpan = document.createElement('span');
  labelSpan.textContent = label;
  const valueCell = document.createElement('span');
  valueCell.className = 'fc-pgsetup__value';
  row.append(labelSpan, valueCell);
  return { row, valueCell };
}

export function makeNumberInput(value: number, step = 0.1, min = 0, max = 99): HTMLInputElement {
  const inp = document.createElement('input');
  inp.type = 'number';
  inp.step = String(step);
  inp.min = String(min);
  inp.max = String(max);
  inp.value = String(value);
  inp.className = 'fc-pgsetup__num';
  return inp;
}

export function makeTextInput(value: string, placeholder = ''): HTMLInputElement {
  const inp = document.createElement('input');
  inp.type = 'text';
  inp.value = value;
  inp.placeholder = placeholder;
  inp.autocomplete = 'off';
  inp.spellcheck = false;
  inp.className = 'fc-pgsetup__text';
  return inp;
}

export const labelize = (label: string, input: HTMLInputElement): HTMLLabelElement => {
  const lab = document.createElement('label');
  lab.className = 'fc-pgsetup__margin';
  const sp = document.createElement('span');
  sp.textContent = label;
  lab.append(sp, input);
  return lab;
};

export const makeCheck = (
  label: string,
): { labelEl: HTMLLabelElement; input: HTMLInputElement } => {
  const labelEl = document.createElement('label');
  labelEl.className = 'fc-fmtdlg__check';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.setAttribute('aria-label', label);
  const text = document.createElement('span');
  text.textContent = label;
  labelEl.append(input, text);
  return { labelEl, input };
};
