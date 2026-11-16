import { cellShiftChoices } from '../commands/built-in-command-policy.js';
import type { DeleteCellsDirection, InsertCellsDirection } from '../commands/cell-shift.js';
import type { History } from '../commands/history.js';
import { columnShiftBand, rowShiftBand } from '../commands/interaction-effects.js';
import type { InteractionOrigin } from '../commands/interaction-policy.js';
import {
  deleteCellsAction,
  insertCellsAction,
  type RowColActionContext,
  rowColWrite,
} from '../commands/row-col-actions.js';
import { deleteCols, deleteRows, insertCols, insertRows } from '../commands/structure.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { Strings } from '../i18n/strings.js';
import { normalizeRange } from '../store/selection-geometry.js';
import type { SpreadsheetStore } from '../store/store.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogActions, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

/** What the dialog submits: a shift direction, or whole rows / columns. */
export type CellShiftChoice = InsertCellsDirection | DeleteCellsDirection | 'rows' | 'cols';

export interface CellShiftDialogDeps {
  host: HTMLElement;
  strings: Strings;
  kind: 'insert' | 'delete';
  /** The choices that may run; the others are disabled. Omitted, every choice is offered. */
  allowed?: { readonly shift: boolean; readonly rows: boolean; readonly cols: boolean };
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
  const allowed = deps.allowed ?? { shift: true, rows: true, cols: true };
  const permitted = (choice: CellShiftChoice): boolean =>
    choice === 'rows' ? allowed.rows : choice === 'cols' ? allowed.cols : allowed.shift;
  // With no choice it may run, it does not open.
  if (!allowed.shift && !allowed.rows && !allowed.cols) return;
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
  const preferred = (kind === 'delete' && mac ? shifts[0] : shifts[1])?.[0];
  const defaultChoice =
    preferred && permitted(preferred)
      ? preferred
      : options.map(([value]) => value).find((value) => permitted(value));
  for (const [value, label] of options) {
    const row = radio(ownerDocument, name, value, label, value === defaultChoice);
    const input = row.querySelector('input');
    if (input && !permitted(value)) projectDisabledState(input, true, null);
    choices.appendChild(row);
  }

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
    const input = event.target as HTMLInputElement | null;
    if (input?.tagName !== 'INPUT') return;
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
  /** The entry point that opened the dialog. */
  origin: InteractionOrigin;
  commandId: string;
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
  const ctx: RowColActionContext = { ...target, workbook: wb };
  const norm = normalizeRange(range);
  const { r0, c0, r1, c1 } = norm;
  const sheet = store.getState().data.sheetIndex;
  if (choice === 'rows') {
    const band = [{ kind: 'range', range: rowShiftBand(sheet, r0) }] as const;
    return rowColWrite(ctx, kind === 'insert' ? 'insertRows' : 'deleteRows', band, false, () =>
      (kind === 'insert' ? insertRows : deleteRows)(store, wb, history, r0, r1 - r0 + 1),
    );
  }
  if (choice === 'cols') {
    const band = [{ kind: 'range', range: columnShiftBand(sheet, c0) }] as const;
    return rowColWrite(
      ctx,
      kind === 'insert' ? 'insertColumns' : 'deleteColumns',
      band,
      false,
      () => (kind === 'insert' ? insertCols : deleteCols)(store, wb, history, c0, c1 - c0 + 1),
    );
  }
  if (kind === 'insert') {
    return choice === 'down' || choice === 'right' ? insertCellsAction(ctx, norm, choice) : false;
  }
  return choice === 'up' || choice === 'left' ? deleteCellsAction(ctx, norm, choice) : false;
}

/** Opens the Insert / Delete Cells dialog over `range` with the choices `target` may run,
 *  applies the chosen one, and calls `onApplied` when the sheet changed. */
export function runCellShiftFlow(
  host: HTMLElement,
  strings: Strings,
  kind: 'insert' | 'delete',
  target: CellShiftTarget,
  range: Range,
  onApplied: () => void,
): void {
  openCellShiftDialog({
    host,
    strings,
    kind,
    allowed: cellShiftChoices(target.store, kind, target.origin),
    onSubmit: (choice) => {
      if (runCellShiftChoice(target, kind, range, choice)) onApplied();
    },
  });
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
