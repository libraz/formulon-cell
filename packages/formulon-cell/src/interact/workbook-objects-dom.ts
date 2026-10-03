import { createDialogButton } from './dialog-shell.js';

export function createWorkbookObjectsActionButton(
  label: string,
  opts: { primary?: boolean; type?: 'button' | 'submit' } = {},
): HTMLButtonElement {
  const button = createDialogButton({
    label,
    baseClass: 'fc-objects__action',
    variant: opts.primary ? 'primary' : undefined,
    primaryClass: 'fc-objects__action--primary',
  });
  button.type = opts.type ?? 'button';
  return button;
}

export function pivotEditField(label: string, control: HTMLElement): HTMLLabelElement {
  const row = document.createElement('label');
  row.className = 'fc-objects__pivot-edit-field';
  const text = document.createElement('span');
  text.textContent = label;
  row.append(text, control);
  return row;
}

export function pivotEditCheck(label: string, control: HTMLInputElement): HTMLLabelElement {
  const row = document.createElement('label');
  row.className = 'fc-objects__pivot-edit-check';
  row.append(control, document.createTextNode(label));
  return row;
}
