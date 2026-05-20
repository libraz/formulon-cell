export const appendSelectRow = (
  body: HTMLElement,
  labelText: string,
  options: readonly { value: string; label: string }[],
  initial: string,
  fieldName?: string,
): HTMLSelectElement => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const label = document.createElement('label');
  label.className = 'app__dlg__label';
  label.textContent = labelText;
  const select = document.createElement('select');
  select.className = 'app__dlg__input';
  if (fieldName) {
    select.dataset.dialogField = fieldName;
  }
  for (const option of options) {
    const opt = document.createElement('option');
    opt.value = option.value;
    opt.textContent = option.label;
    select.appendChild(opt);
  }
  select.value = initial;
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
