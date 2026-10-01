import type { InsertCopiedCellsDirection } from '../commands/clipboard/insert-copied-cells.js';
import type { Strings } from '../i18n/strings.js';
import { appendDialogActions } from './dialog-shell.js';
import { overlayPortalFor } from './overlay-portal.js';

export interface InsertCopiedCellsDialogDeps {
  host: HTMLElement;
  strings: Strings;
  onSubmit(direction: InsertCopiedCellsDirection): void;
}

export function openInsertCopiedCellsDialog(deps: InsertCopiedCellsDialogDeps): void {
  const t = deps.strings.insertCopiedCellsDialog;
  const ownerDocument = deps.host.ownerDocument;
  const portal = overlayPortalFor(deps.host);
  portal.querySelector('.fc-insertcopied')?.remove();

  const root = ownerDocument.createElement('div');
  root.className = 'fc-insertcopied';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'true');
  root.setAttribute('aria-label', t.title);

  const panel = ownerDocument.createElement('div');
  panel.className = 'fc-insertcopied__panel';

  const title = ownerDocument.createElement('div');
  title.className = 'fc-insertcopied__title';
  title.textContent = t.title;

  const choices = ownerDocument.createElement('div');
  choices.className = 'fc-insertcopied__choices';
  const name = `fc-insertcopied-${Math.random().toString(36).slice(2)}`;
  choices.append(
    radio(ownerDocument, name, 'right', t.shiftRight, false),
    radio(ownerDocument, name, 'down', t.shiftDown, true),
  );

  const footer = ownerDocument.createElement('div');
  footer.className = 'fc-insertcopied__footer';
  const { cancelBtn: cancel, okBtn: ok } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
    buttonBaseClass: 'fc-insertcopied__button',
    buttonPrimaryClass: 'fc-insertcopied__button--primary',
    buttonSecondaryClass: 'fc-insertcopied__button--secondary',
  });

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
  portal.appendChild(root);
  ok.focus({ preventScroll: true });
}

function radio(
  ownerDocument: Document,
  name: string,
  value: InsertCopiedCellsDirection,
  label: string,
  checked: boolean,
): HTMLLabelElement {
  const row = ownerDocument.createElement('label');
  row.className = 'fc-insertcopied__choice';
  const input = ownerDocument.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  const mark = ownerDocument.createElement('span');
  mark.className = 'fc-insertcopied__radio';
  const text = ownerDocument.createElement('span');
  text.textContent = label;
  row.append(input, mark, text);
  return row;
}
