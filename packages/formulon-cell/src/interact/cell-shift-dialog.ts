import type { DeleteCellsDirection, InsertCellsDirection } from '../commands/cell-shift.js';
import type { Strings } from '../i18n/strings.js';
import { appendDialogActions } from './dialog-shell.js';
import { overlayPortalFor } from './overlay-portal.js';

export interface CellShiftDialogDeps {
  strings: Strings;
  kind: 'insert' | 'delete';
  onSubmit(direction: InsertCellsDirection | DeleteCellsDirection): void;
}

/** Open the Excel-style direction chooser used by Insert Cells and Delete Cells. */
export function openCellShiftDialog(deps: CellShiftDialogDeps): void {
  const { strings, kind } = deps;
  const t = strings.ribbonMenu;
  document.querySelector('.fc-cellshift')?.remove();

  const root = document.createElement('div');
  root.className = 'fc-cellshift';
  root.setAttribute('role', 'dialog');
  root.setAttribute('aria-modal', 'true');
  root.setAttribute('aria-label', kind === 'insert' ? t.insertCells : t.deleteCells);
  const panel = document.createElement('div');
  panel.className = 'fc-cellshift__panel';
  const title = document.createElement('div');
  title.className = 'fc-cellshift__title';
  title.textContent = kind === 'insert' ? t.insertCells : t.deleteCells;
  const choices = document.createElement('div');
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
    choices.appendChild(radio(name, value, label, value === options[1][0]));

  const footer = document.createElement('div');
  footer.className = 'fc-cellshift__footer';
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: strings.hyperlinkDialog.cancel,
    okLabel: strings.hyperlinkDialog.ok,
    buttonBaseClass: 'fc-cellshift__button',
    buttonPrimaryClass: 'fc-cellshift__button--primary',
    buttonSecondaryClass: 'fc-cellshift__button--secondary',
  });
  const close = (): void => root.remove();
  cancelBtn.addEventListener('click', close);
  okBtn.addEventListener('click', () => {
    const selected = root.querySelector<HTMLInputElement>('input[type="radio"]:checked')?.value;
    if (selected === 'down' || selected === 'right' || selected === 'up' || selected === 'left') {
      deps.onSubmit(selected);
    }
    close();
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
  root.appendChild(panel);
  overlayPortalFor(document.querySelector('.fc-host')).appendChild(root);
  okBtn.focus({ preventScroll: true });
}

function radio(name: string, value: string, label: string, checked: boolean): HTMLLabelElement {
  const row = document.createElement('label');
  row.className = 'fc-cellshift__choice';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  input.checked = checked;
  const mark = document.createElement('span');
  mark.className = 'fc-cellshift__radio';
  const text = document.createElement('span');
  text.textContent = label;
  row.append(input, mark, text);
  return row;
}
