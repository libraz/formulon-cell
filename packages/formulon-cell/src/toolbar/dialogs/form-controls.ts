export interface DialogSelectOption {
  value: string;
  label: string;
}

export interface DialogSelectOptions {
  className?: string;
  fieldName?: string;
  ariaLabel?: string;
  role?: string;
}

export const appendDialogSelectOptions = (
  select: HTMLSelectElement,
  options: readonly DialogSelectOption[],
): void => {
  for (const option of options) {
    const opt = document.createElement('option');
    opt.value = option.value;
    opt.textContent = option.label;
    select.appendChild(opt);
  }
};

export const appendDialogDatalistOptions = (
  datalist: HTMLDataListElement,
  values: readonly string[],
): void => {
  for (const value of values) {
    const opt = document.createElement('option');
    opt.value = value;
    datalist.appendChild(opt);
  }
};

export const createDialogSelect = (
  options: readonly DialogSelectOption[],
  initial: string,
  opts: DialogSelectOptions = {},
): HTMLSelectElement => {
  const select = document.createElement('select');
  select.className = opts.className ?? 'fc-tb__dlg__input';
  if (opts.fieldName) {
    select.dataset.dialogField = opts.fieldName;
  }
  if (opts.ariaLabel) {
    select.setAttribute('aria-label', opts.ariaLabel);
  }
  if (opts.role) {
    select.setAttribute('role', opts.role);
  }
  appendDialogSelectOptions(select, options);
  select.value = initial;
  return select;
};

export const appendSelectRow = (
  body: HTMLElement,
  labelText: string,
  options: readonly DialogSelectOption[],
  initial: string,
  fieldName?: string,
): HTMLSelectElement => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const label = document.createElement('label');
  label.className = 'fc-tb__dlg__label';
  label.textContent = labelText;
  const select = createDialogSelect(options, initial, { fieldName });
  label.appendChild(select);
  row.appendChild(label);
  body.appendChild(row);
  return select;
};

export const appendCheckboxRow = (
  body: HTMLElement,
  labelText: string,
  initial: boolean,
  fieldName?: string,
): HTMLInputElement => {
  const label = document.createElement('label');
  label.className = 'fc-fmtdlg__checkbox';
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.checked = initial;
  if (fieldName) {
    input.dataset.dialogField = fieldName;
  }
  label.append(input, document.createTextNode(labelText));
  body.appendChild(label);
  return input;
};
