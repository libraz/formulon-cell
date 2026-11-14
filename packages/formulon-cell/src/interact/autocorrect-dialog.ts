import { proofingLocale } from '../commands/spelling.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { appendAutoCorrectOptions } from './autocorrect-options.js';
import { appendDialogButton, appendDialogFrame, createDialogShell } from './dialog-shell.js';

const openDialogs = new WeakMap<SpreadsheetInstance, () => void>();

/** The name of the language the lists are for, in the UI language. */
const languageName = (locale: string): string => {
  try {
    return (
      new Intl.DisplayNames([locale], { type: 'language' }).of(proofingLocale(locale)) ?? locale
    );
  } catch {
    return locale;
  }
};

/** Opens the AutoCorrect options: the replacement list and its switches. Every change applies at
 *  once and reaches the host through `autoCorrect.onChange`. */
export function openAutoCorrectDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
  const t = instance.i18n.strings.autoCorrectDialog;
  const title = t.title.replace('{language}', languageName(instance.i18n.locale));
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-autocorrect',
    ariaLabel: title,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { body, footer } = appendDialogFrame(shell, {
    title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-autocorrect__panel'],
    bodyClass: 'fc-fmtdlg__body fc-autocorrect__body',
    footerClass: 'fc-fmtdlg__footer fc-autocorrect__footer',
  });

  const options = appendAutoCorrectOptions(instance, body, shell.on, title, 'fc-autocorrect');
  const closeButton = appendDialogButton(footer, { label: t.close, variant: 'primary' });

  shell.on(closeButton, 'click', () => close());

  let closed = false;
  function close(): void {
    if (closed) return;
    closed = true;
    openDialogs.delete(instance);
    shell.close();
    shell.dispose();
  }
  openDialogs.set(instance, close);

  shell.open();
  options.focus();
}

/** Closes open AutoCorrect options, as unmounting does. */
export function closeAutoCorrectDialog(instance: SpreadsheetInstance): void {
  openDialogs.get(instance)?.();
}
