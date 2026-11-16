import { createInteractionChipButton } from '../../interact/chip-button.js';
import { projectDisabledState } from '../menu-a11y.js';
import { appendDialogMessage } from './shell.js';

export interface DialogSelectOption {
  value: string;
  label: string;
  /** A row that cannot be picked, such as a separator. */
  disabled?: boolean;
}

export interface DialogDatalistOption {
  value: string;
  label?: string;
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
    if (option.disabled) projectDisabledState(opt, true, null);
    select.appendChild(opt);
  }
};

export const appendDialogDatalistOptions = (
  datalist: HTMLDataListElement,
  values: readonly (string | DialogDatalistOption)[],
): void => {
  for (const entry of values) {
    const opt = document.createElement('option');
    if (typeof entry === 'string') {
      opt.value = entry;
    } else {
      opt.value = entry.value;
      if (entry.label !== undefined) opt.label = entry.label;
    }
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
  className?: string,
): HTMLSelectElement => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const label = document.createElement('label');
  label.className = 'fc-tb__dlg__label';
  label.textContent = labelText;
  const select = createDialogSelect(options, initial, { fieldName, className });
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

export interface RadioOptionConfig {
  name: string;
  label: string;
  checked: boolean;
  value?: string;
}

/** One `label.fc-fmtdlg__radio > input[type=radio] + span` row. */
export const createRadioOption = (
  doc: Document,
  config: RadioOptionConfig,
): { wrap: HTMLLabelElement; input: HTMLInputElement } => {
  const wrap = doc.createElement('label');
  wrap.className = 'fc-fmtdlg__radio';
  const input = doc.createElement('input');
  input.type = 'radio';
  input.name = config.name;
  if (config.value !== undefined) input.value = config.value;
  input.checked = config.checked;
  const text = doc.createElement('span');
  text.textContent = config.label;
  wrap.append(input, text);
  return { wrap, input };
};

export interface RadioGroupConfig<T extends string> {
  /** Accessible name of the group. */
  ariaLabel: string;
  /** Visible caption above the group, when given (an empty string still draws the paragraph). */
  caption?: string;
  options: readonly { value: T; label: string }[];
  /** Whether the option at `index` starts checked. */
  isChecked: (value: T, index: number) => boolean;
}

/** Appends a radio group under `body` and returns the group element and a reader for its checked value. */
export const appendRadioGroup = <T extends string>(
  body: HTMLElement,
  config: RadioGroupConfig<T>,
): { group: HTMLDivElement; read: () => T | undefined } => {
  const doc = body.ownerDocument;
  const group = doc.createElement('div');
  group.className = 'fc-fmtdlg__choice-grid fc-tb__dlg__choices';
  group.setAttribute('role', 'radiogroup');
  group.setAttribute('aria-label', config.ariaLabel);
  if (config.caption !== undefined) {
    appendDialogMessage(body, config.caption);
  }
  const name = `app-radio-${Math.random().toString(36).slice(2)}`;
  for (const [index, option] of config.options.entries()) {
    group.appendChild(
      createRadioOption(doc, {
        name,
        value: option.value,
        label: option.label,
        checked: config.isChecked(option.value, index),
      }).wrap,
    );
  }
  body.appendChild(group);
  return {
    group,
    read: () =>
      group.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.value as T | undefined,
  };
};

/** A single-select `role=listbox` of chip buttons with Arrow Up/Down navigation. */
export const createChipListbox = (
  names: readonly string[],
  initial: number,
  ariaLabel: string,
): { list: HTMLDivElement; selected: () => number; choose: (index: number) => void } => {
  const list = document.createElement('div');
  list.className = 'fc-movecopy__list';
  list.setAttribute('role', 'listbox');
  list.setAttribute('aria-label', ariaLabel);
  let selected = initial;
  const options = names.map((name, index) => {
    const option = createInteractionChipButton({
      className: 'fc-movecopy__item',
      label: name,
      role: 'option',
    });
    option.addEventListener('click', () => select(index));
    list.appendChild(option);
    return option;
  });
  const select = (index: number): void => {
    selected = Math.max(0, Math.min(names.length - 1, index));
    options.forEach((option, i) => {
      option.setAttribute('aria-selected', i === selected ? 'true' : 'false');
      option.tabIndex = i === selected ? 0 : -1;
    });
    options[selected]?.focus({ preventScroll: true });
  };
  list.addEventListener('keydown', (event) => {
    if (event.key === 'ArrowDown') select(selected + 1);
    else if (event.key === 'ArrowUp') select(selected - 1);
    else return;
    event.preventDefault();
  });
  return { list, selected: () => selected, choose: select };
};

export interface NamedRadioRowConfig {
  name: string;
  value: string;
  title: string;
  note: string;
  checked: boolean;
  /** Called with `value` when the radio becomes checked. */
  onSelect: (value: string) => void;
}

/** A list row with a radio, a bold title and a muted note. */
export const appendNamedRadioRow = (
  list: HTMLElement,
  config: NamedRadioRowConfig,
): HTMLInputElement => {
  const label = document.createElement('label');
  label.className = 'fc-fmtdlg__row fc-fmtdlg__row--block fc-tb__dlg__check';
  const radio = document.createElement('input');
  radio.type = 'radio';
  radio.name = config.name;
  radio.value = config.value;
  radio.checked = config.checked;
  radio.addEventListener('change', () => {
    if (radio.checked) config.onSelect(config.value);
  });
  const body = document.createElement('span');
  const title = document.createElement('strong');
  title.textContent = config.title;
  const note = document.createElement('span');
  note.className = 'fc-tb__dlg__note';
  note.textContent = config.note;
  body.append(title, note);
  label.append(radio, body);
  list.appendChild(label);
  return radio;
};
