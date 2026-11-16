import { projectDisabledState } from '../menu-a11y.js';
import { createChipListbox } from './form-controls.js';
import {
  appendDialogActions,
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from './shell.js';

export interface MoveCopySheetDialogStrings {
  title: string;
  targetBook: string;
  before: string;
  toEnd: string;
  makeCopy: string;
  ok: string;
  cancel: string;
}

export interface MoveCopySheetDialogOptions {
  strings: MoveCopySheetDialogStrings;
  /** Element inside the spreadsheet host; anchors the dialog so it follows the host platform. */
  host?: HTMLElement;
  bookName: string;
  sheets: readonly string[];
  /** The sheet being moved. */
  current: number;
  /** Why "Create a copy" is unavailable; null when the copy may be made. */
  copyDisabledReason: string | null;
}

export interface MoveCopySheetDialogResult {
  /** Index of the sheet to go in front of; the sheet count means the end. */
  before: number;
  /** True when "Create a copy" was checked: copy the sheet instead of moving it. */
  copy: boolean;
}

/** Move or Copy dialog: the position to move the sheet to, or to copy it to. */
export const showMoveCopySheetDialog = (
  opts: MoveCopySheetDialogOptions,
): Promise<MoveCopySheetDialogResult | null> =>
  new Promise<MoveCopySheetDialogResult | null>((resolve) => {
    const { strings } = opts;
    const shell = createDialogShell({ host: opts.host, title: strings.title, bodyVariant: 'app' });
    shell.panel.classList.add('fc-movecopy');

    const bookLabel = document.createElement('div');
    bookLabel.className = 'fc-tb__dlg__label';
    bookLabel.textContent = strings.targetBook;
    const book = document.createElement('div');
    book.className = 'fc-movecopy__book';
    book.textContent = opts.bookName;
    book.setAttribute('role', 'combobox');
    book.setAttribute('aria-label', strings.targetBook);

    const beforeLabel = document.createElement('div');
    beforeLabel.className = 'fc-tb__dlg__label';
    beforeLabel.textContent = strings.before;
    const names = [...opts.sheets, strings.toEnd];
    const chips = createChipListbox(
      names,
      Math.min(opts.current + 1, opts.sheets.length),
      strings.before,
    );
    const { list } = chips;

    const copyLabel = document.createElement('label');
    copyLabel.className = 'fc-movecopy__copy fc-fmtdlg__checkbox';
    const copy = document.createElement('input');
    copy.type = 'checkbox';
    projectDisabledState(copy, opts.copyDisabledReason !== null, opts.copyDisabledReason);
    copyLabel.append(copy, document.createTextNode(strings.makeCopy));

    shell.body.append(bookLabel, book, beforeLabel, list, copyLabel);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: strings.cancel,
      okLabel: strings.ok,
    });

    const lifecycle = installDialogLifecycle<MoveCopySheetDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
      buttons: { okBtn, cancelBtn },
    });
    const onOk = (): void =>
      lifecycle.finish({ before: chips.selected(), copy: copy.checked && !copy.disabled });

    mountDialog(shell, okBtn);
    chips.choose(chips.selected());
  });
