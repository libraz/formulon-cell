import {
  type DeleteCellsDirection,
  deleteCells,
  type InsertCellsDirection,
  insertCells,
} from '../commands/cell-shift.js';
import type { History } from '../commands/history.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../commands/structure.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { appendDialogActions, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

/** What the dialog submits: a shift direction, or whole rows / columns. */
export type CellShiftChoice = InsertCellsDirection | DeleteCellsDirection | 'rows' | 'cols';

export interface CellShiftDialogDeps {
  host: HTMLElement;
  strings: Strings;
  kind: 'insert' | 'delete';
  onSubmit(choice: CellShiftChoice): void;
}

interface CellShiftDialogRecord {
  instanceId: string | undefined;
  disposed: boolean;
  close(): void;
  dispose(): void;
}

const dialogs = new WeakMap<HTMLElement, CellShiftDialogRecord>();

const CHOICES: readonly CellShiftChoice[] = ['right', 'down', 'left', 'up', 'rows', 'cols'];

/** Open the Excel-style chooser used by Insert Cells and Delete Cells: the two
 *  shift directions, then Entire row and Entire column. */
export function openCellShiftDialog(deps: CellShiftDialogDeps): void {
  const { host, strings, kind } = deps;
  const t = strings.cellShiftDialog;
  const ownerDocument = host.ownerDocument;
  dialogs.get(host)?.close();

  const title = kind === 'insert' ? t.insertTitle : t.deleteTitle;
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
  const shifts: readonly (readonly [CellShiftChoice, string])[] =
    kind === 'insert'
      ? [
          ['right', t.shiftRight],
          ['down', t.shiftDown],
        ]
      : [
          ['left', t.shiftLeft],
          ['up', t.shiftUp],
        ];
  const options = [...shifts, ['rows', t.entireRow], ['cols', t.entireColumn]] as const;
  // The vertical shift is the default, except that Mac's Delete dialog opens
  // on the left shift.
  const mac = isMacHost(host);
  const defaultChoice = (kind === 'delete' && mac ? shifts[0] : shifts[1])?.[0];
  for (const [value, label] of options)
    choices.appendChild(radio(ownerDocument, name, value, label, value === defaultChoice));

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
    const choice = CHOICES.find((c) => c === selected) ?? null;
    close();
    if (choice) deps.onSubmit(choice);
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

export interface CellShiftTarget {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  history: History | null;
}

/** Apply an Insert / Delete cells dialog choice to `range`: shift the cells,
 *  or insert / delete the whole rows or columns the range spans. A shift that
 *  belongs to the other dialog does nothing. Returns whether the sheet changed. */
export function runCellShiftChoice(
  target: CellShiftTarget,
  kind: 'insert' | 'delete',
  range: Range,
  choice: CellShiftChoice,
): boolean {
  const { store, wb, history } = target;
  const r0 = Math.min(range.r0, range.r1);
  const r1 = Math.max(range.r0, range.r1);
  const c0 = Math.min(range.c0, range.c1);
  const c1 = Math.max(range.c0, range.c1);
  const norm: Range = { sheet: range.sheet, r0, c0, r1, c1 };
  if (choice === 'rows') {
    return (kind === 'insert' ? insertRows : deleteRows)(store, wb, history, r0, r1 - r0 + 1);
  }
  if (choice === 'cols') {
    return (kind === 'insert' ? insertCols : deleteCols)(store, wb, history, c0, c1 - c0 + 1);
  }
  if (kind === 'insert') {
    return choice === 'down' || choice === 'right'
      ? insertCells(store, wb, history, norm, choice)
      : false;
  }
  return choice === 'up' || choice === 'left'
    ? deleteCells(store, wb, history, norm, choice)
    : false;
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
