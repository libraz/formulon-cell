import type { History } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { isMacHost } from '../extensions/ui-options.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { appendCalculationOptions } from './calculation-options.js';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';

export interface IterativeDialogDeps {
  host: HTMLElement;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so every setting always targets the live engine. */
  getWb: () => WorkbookHandle | null;
  /** Live input locale. Defaults to the canonical English grammar. */
  getLocale?: () => string;
  strings?: Strings;
  /** Store mirroring the reference style and date system. Without it those
   *  settings go to the engine directly. */
  store?: SpreadsheetStore;
  /** Undo stack for the constants Set precision as displayed rounds. */
  history?: History | null;
}

export interface IterativeDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

/**
 * Calculation options dialog — the desktop Formulas / Calculation options
 * pages. The controls come from `appendCalculationOptions`, shared with the
 * Mac Preferences Calculation pane.
 */
export function attachIterativeDialog(deps: IterativeDialogDeps): IterativeDialogHandle {
  const { host } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.iterativeDialog;

  const shell = createDialogShell({
    host,
    className: 'fc-iterdlg',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });

  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel'],
    headerClass: 'fc-fmtdlg__header fc-iterdlg__header',
    bodyClass: 'fc-fmtdlg__body fc-iterdlg__body',
    footerClass: 'fc-fmtdlg__footer fc-iterdlg__footer',
  });

  const options = appendCalculationOptions({ ...deps, strings }, body, shell.on, 'fc-iterdlg');

  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
    buttonBaseClass: 'fc-fmtdlg__btn fc-iterdlg__btn',
    buttonPrimaryClass: 'fc-fmtdlg__btn--primary fc-iterdlg__btn--primary',
  });

  // The Mac pane has no buttons; a close control in the header dismisses it.
  const closeBtn = appendDialogButton(header, {
    label: '',
    baseClass: 'fc-fmtdlg__close',
  });
  closeBtn.setAttribute('aria-label', strings.watchPanel.close);
  closeBtn.title = strings.watchPanel.close;
  closeBtn.addEventListener('click', () => api.close());

  const onOk = (): void => {
    if (options.commit()) api.close();
  };

  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', () => api.close());
  shell.on(shell.overlay, 'keydown', (e) => {
    const event = e as KeyboardEvent;
    if (isSubmitEnter(event)) {
      event.preventDefault();
      onOk();
    }
  });

  const api: IterativeDialogHandle = {
    open(): void {
      const mac = isMacHost(host);
      footer.hidden = mac;
      closeBtn.hidden = !mac;
      options.refresh();
      shell.open();
      deferInitialFocus(shell.overlay, () => options.focus());
    },
    close(): void {
      shell.close();
    },
    detach(): void {
      shell.dispose();
    },
  };
  return api;
}
