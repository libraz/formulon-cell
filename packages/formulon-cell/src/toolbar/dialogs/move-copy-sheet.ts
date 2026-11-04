import { createInteractionChipButton } from '../../interact/chip-button.js';
import { projectDisabledState } from '../menu-a11y.js';
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
  copyUnavailable: string;
  ok: string;
  cancel: string;
}

export interface MoveCopySheetDialogOptions {
  strings: MoveCopySheetDialogStrings;
  bookName: string;
  sheets: readonly string[];
  /** The sheet being moved. */
  current: number;
}

export interface MoveCopySheetDialogResult {
  /** Index of the sheet to go in front of; the sheet count means the end. */
  before: number;
}

/** Move or Copy dialog. Only moving is available: the copy checkbox is shown disabled. */
export const showMoveCopySheetDialog = (
  opts: MoveCopySheetDialogOptions,
): Promise<MoveCopySheetDialogResult | null> =>
  new Promise<MoveCopySheetDialogResult | null>((resolve) => {
    const { strings } = opts;
    const shell = createDialogShell({ title: strings.title, bodyVariant: 'app' });
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
    const list = document.createElement('div');
    list.className = 'fc-movecopy__list';
    list.setAttribute('role', 'listbox');
    list.setAttribute('aria-label', strings.before);

    const names = [...opts.sheets, strings.toEnd];
    let selected = Math.min(opts.current + 1, opts.sheets.length);
    const options = names.map((name, index) => {
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
      selected = Math.max(0, Math.min(names.length - 1, index));
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

    const copyLabel = document.createElement('label');
    copyLabel.className = 'fc-movecopy__copy fc-fmtdlg__checkbox';
    const copy = document.createElement('input');
    copy.type = 'checkbox';
    projectDisabledState(copy, true, strings.copyUnavailable);
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
    });
    const onOk = (): void => lifecycle.finish({ before: selected });
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));

    mountDialog(shell, okBtn);
    select(selected);
  });
