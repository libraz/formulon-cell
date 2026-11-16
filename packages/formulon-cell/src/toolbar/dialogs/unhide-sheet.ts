import { createChipListbox } from './form-controls.js';
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
    const chips = createChipListbox(opts.sheets, 0, strings.list);
    const { list } = chips;

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
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void => lifecycle.finish(chips.selected());

    mountDialog(shell, okBtn);
    chips.choose(0);
  });
