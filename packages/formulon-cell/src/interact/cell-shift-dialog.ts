import type { DeleteCellsDirection, InsertCellsDirection } from '../commands/cell-shift.js';
import type { Strings } from '../i18n/strings.js';
import { appendDialogActions, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

export interface CellShiftDialogDeps {
  host: HTMLElement;
  strings: Strings;
  kind: 'insert' | 'delete';
  onSubmit(direction: InsertCellsDirection | DeleteCellsDirection): void;
}

interface CellShiftDialogRecord {
  instanceId: string | undefined;
  disposed: boolean;
  close(): void;
  dispose(): void;
}

const dialogs = new WeakMap<HTMLElement, CellShiftDialogRecord>();

/** Open the Excel-style direction chooser used by Insert Cells and Delete Cells. */
export function openCellShiftDialog(deps: CellShiftDialogDeps): void {
  const { host, strings, kind } = deps;
  const t = strings.ribbonMenu;
  const ownerDocument = host.ownerDocument;
  dialogs.get(host)?.close();

  const title = kind === 'insert' ? t.insertCells : t.deleteCells;
  let record: CellShiftDialogRecord | null = null;
  const shell = createDialogShell({
    host,
    className: 'fc-cellshift',
    ariaLabel: title,
    onDismiss: () => record?.close(),
  });
  const dispose = (): void => {
    if (!record || record.disposed) return;
    record.disposed = true;
    shell.dispose();
    if (dialogs.get(host) === record) dialogs.delete(host);
  };
  const close = (): void => {
    if (!record || record.disposed) return;
    shell.close();
    dispose();
  };

  const panel = shell.panel;
  const heading = ownerDocument.createElement('div');
  heading.className = 'fc-cellshift__title';
  heading.textContent = title;
  const choices = ownerDocument.createElement('div');
  choices.className = 'fc-cellshift__choices';
  const name = `fc-cellshift-${Math.random().toString(36).slice(2)}`;
  const options =
    kind === 'insert'
      ? ([
          ['right', t.insertShiftRight],
          ['down', t.insertShiftDown],
        ] as const)
      : ([
          ['left', t.deleteShiftLeft],
          ['up', t.deleteShiftUp],
        ] as const);
  for (const [value, label] of options)
    choices.appendChild(radio(ownerDocument, name, value, label, value === options[1][0]));

  const footer = ownerDocument.createElement('div');
  footer.className = 'fc-cellshift__footer';
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: strings.hyperlinkDialog.cancel,
    okLabel: strings.hyperlinkDialog.ok,
    buttonBaseClass: 'fc-cellshift__button',
    buttonPrimaryClass: 'fc-cellshift__button--primary',
    buttonSecondaryClass: 'fc-cellshift__button--secondary',
  });
  panel.append(heading, choices, footer);

  record = {
    instanceId: host.dataset.fcInstId,
    disposed: false,
    close,
    dispose,
  };
  dialogs.set(host, record);

  const submit = (): void => {
    if (!record || record.disposed || !shell.isOpen()) return;
    const selected = panel.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.value;
    const direction =
      selected === 'down' || selected === 'right' || selected === 'up' || selected === 'left'
        ? selected
        : null;
    close();
    if (direction) deps.onSubmit(direction);
  };

  shell.on(cancelBtn, 'click', close);
  shell.on(okBtn, 'click', submit);
  shell.on(shell.overlay, 'keydown', (rawEvent) => {
    const event = rawEvent as KeyboardEvent;
    if (!isSubmitEnter(event)) return;
    const target = event.target;
    if (!target || typeof target !== 'object' || !('nodeType' in target)) return;
    const element = target as Element;
    if (element.nodeType !== 1 || element.tagName !== 'INPUT') return;
    const input = target as HTMLInputElement;
    if (input.type !== 'radio' || input !== ownerDocument.activeElement) return;
    event.preventDefault();
    event.stopPropagation();
    submit();
  });

  shell.open();
  okBtn.focus({ preventScroll: true });
}

/** Dispose a host's current dialog without returning focus to a removed opener. */
export function disposeCellShiftDialog(host: HTMLElement, instanceId?: string): void {
  const record = dialogs.get(host);
  if (!record || (instanceId !== undefined && record.instanceId !== instanceId)) return;
  record.dispose();
}

function radio(
  ownerDocument: Document,
  name: string,
  value: string,
  label: string,
  checked: boolean,
): HTMLLabelElement {
  const row = ownerDocument.createElement('label');
  row.className = 'fc-cellshift__choice';
  const input = ownerDocument.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  const mark = ownerDocument.createElement('span');
  mark.className = 'fc-cellshift__radio';
  const text = ownerDocument.createElement('span');
  text.textContent = label;
  row.append(input, mark, text);
  return row;
}
