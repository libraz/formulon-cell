import type { InsertCopiedCellsDirection } from '../commands/clipboard/insert-copied-cells.js';
import type { Strings } from '../i18n/strings.js';

export interface InsertCopiedCellsDialogDeps {
  strings: Strings;
  onSubmit(direction: InsertCopiedCellsDirection): void;
}

export function openInsertCopiedCellsDialog(deps: InsertCopiedCellsDialogDeps): void {
  const t = deps.strings.insertCopiedCellsDialog;
  document.querySelector('.fc-insertcopied')?.remove();

  const root = document.createElement('div');
  root.className = 'fc-insertcopied';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'true');
  root.setAttribute('aria-label', t.title);

  const panel = document.createElement('div');
  panel.className = 'fc-insertcopied__panel';

  const title = document.createElement('div');
  title.className = 'fc-insertcopied__title';
  title.textContent = t.title;

  const choices = document.createElement('div');
  choices.className = 'fc-insertcopied__choices';
  const name = `fc-insertcopied-${Math.random().toString(36).slice(2)}`;
  choices.append(radio(name, 'right', t.shiftRight, false), radio(name, 'down', t.shiftDown, true));

  const footer = document.createElement('div');
  footer.className = 'fc-insertcopied__footer';
  const cancel = document.createElement('button');
  cancel.type = 'button';
  cancel.className = 'fc-insertcopied__button fc-insertcopied__button--secondary';
  cancel.textContent = t.cancel;
  const ok = document.createElement('button');
  ok.type = 'button';
  ok.className = 'fc-insertcopied__button fc-insertcopied__button--primary';
  ok.textContent = t.ok;
  footer.append(cancel, ok);

  const close = (): void => root.remove();
  cancel.addEventListener('click', close);
  ok.addEventListener('click', () => {
    const checked = root.querySelector<HTMLInputElement>('input[type="radio"]:checked');
    const direction = checked?.value === 'right' ? 'right' : 'down';
    close();
    deps.onSubmit(direction);
  });
  root.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    }
  });
  root.addEventListener('mousedown', (e) => {
    if (e.target === root) close();
  });

  panel.append(title, choices, footer);
  root.append(panel);
  document.body.appendChild(root);
  ok.focus({ preventScroll: true });
}

function radio(
  name: string,
  value: InsertCopiedCellsDirection,
  label: string,
  checked: boolean,
): HTMLLabelElement {
  const row = document.createElement('label');
  row.className = 'fc-insertcopied__choice';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  const mark = document.createElement('span');
  mark.className = 'fc-insertcopied__radio';
  const text = document.createElement('span');
  text.textContent = label;
  row.append(input, mark, text);
  return row;
}
