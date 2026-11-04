import { createInteractionChipButton } from '../../interact/chip-button.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface UnhideSheetDialogStrings {
  title: string;
  list: string;
  ok: string;
  cancel: string;
}

export interface UnhideSheetDialogOptions {
  strings: UnhideSheetDialogStrings;
  /** Element inside the spreadsheet host; anchors the dialog so it follows the host platform. */
  host?: HTMLElement;
  /** Names of the hidden sheets, in the order listed. */
  sheets: readonly string[];
}

/** Unhide dialog: one list of the hidden sheets. Resolves to the chosen position, or null. */
export const showUnhideSheetDialog = (opts: UnhideSheetDialogOptions): Promise<number | null> =>
  new Promise<number | null>((resolve) => {
    const { strings } = opts;
    const shell = createDialogShell({ host: opts.host, title: strings.title, bodyVariant: 'app' });
    shell.panel.classList.add('fc-movecopy', 'fc-unhidesheet');

    const label = document.createElement('div');
    label.className = 'fc-tb__dlg__label';
    label.textContent = strings.list;
    const list = document.createElement('div');
    list.className = 'fc-movecopy__list';
    list.setAttribute('role', 'listbox');
    list.setAttribute('aria-label', strings.list);

    let selected = 0;
    const options = opts.sheets.map((name, index) => {
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
      selected = Math.max(0, Math.min(options.length - 1, index));
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

    shell.body.append(label, list);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const lifecycle = installDialogLifecycle<number | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => lifecycle.finish(selected);
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, okBtn);
    select(0);
  });
