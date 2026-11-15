import type { InsertCopiedCellsDirection } from '../commands/clipboard/insert-copied-cells.js';
import type { Strings } from '../i18n/strings.js';
import { appendDialogActions, createDialogShell } from './dialog-shell.js';

export interface InsertCopiedCellsDialogDeps {
  host: HTMLElement;
  strings: Strings;
  onSubmit(direction: InsertCopiedCellsDirection): void;
}

// One open dialog per host; opening again replaces it.
const openDialogs = new WeakMap<HTMLElement, () => void>();

export function openInsertCopiedCellsDialog(deps: InsertCopiedCellsDialogDeps): void {
  const t = deps.strings.insertCopiedCellsDialog;
  const ownerDocument = deps.host.ownerDocument;
  openDialogs.get(deps.host)?.();

  const shell = createDialogShell({
    host: deps.host,
    className: 'fc-insertcopied',
    ariaLabel: t.title,
    onDismiss: () => close(),
  });
  const close = (): void => {
    if (openDialogs.get(deps.host) !== close) return;
    openDialogs.delete(deps.host);
    shell.close();
    shell.dispose();
  };
  openDialogs.set(deps.host, close);

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

  shell.on(cancel, 'click', close);
  shell.on(ok, 'click', () => {
    const checked = shell.panel.querySelector<HTMLInputElement>('input[type="radio"]:checked');
    const direction = checked?.value === 'right' ? 'right' : 'down';
    close();
    deps.onSubmit(direction);
  });

  shell.panel.append(title, choices, footer);
  shell.open();
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
