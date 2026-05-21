// Shared shell + focus-trap utilities used by every dialog in this folder.
// Each `showX(...)` wires up its own widgets but reuses this scaffold so
// keyboard handling, focus restoration, and aria wiring stay consistent.

const FOCUSABLE_DIALOG_SELECTOR = [
  'button',
  'input',
  'select',
  'textarea',
  'a[href]',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

const focusableDialogItems = (root: HTMLElement): HTMLElement[] =>
  Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE_DIALOG_SELECTOR)).filter((el) => {
    if (el.closest('[hidden],[aria-hidden="true"]')) return false;
    if ('disabled' in el && (el as HTMLButtonElement | HTMLInputElement).disabled) return false;
    return el.tabIndex >= 0;
  });

export const trapDialogTab = (root: HTMLElement, event: KeyboardEvent): boolean => {
  if (event.key !== 'Tab') return false;
  const items = focusableDialogItems(root);
  if (items.length === 0) {
    event.preventDefault();
    root.focus({ preventScroll: true });
    return true;
  }
  const first = items[0];
  const last = items[items.length - 1];
  if (event.shiftKey && document.activeElement === first) {
    event.preventDefault();
    last?.focus({ preventScroll: true });
    return true;
  }
  if (!event.shiftKey && document.activeElement === last) {
    event.preventDefault();
    first?.focus({ preventScroll: true });
    return true;
  }
  return false;
};

export const restoreDialogFocus = (overlay: HTMLElement, opener: HTMLElement | null): void => {
  if (!opener) return;
  if (overlay.contains(document.activeElement) || document.activeElement === document.body) {
    opener.focus({ preventScroll: true });
  }
};

export interface DialogShellOptions {
  title: string;
  /** Defaults to `dialog`; pass `alertdialog` for confirm/message dialogs. */
  role?: 'dialog' | 'alertdialog';
  /** Override aria-label when the visible title differs from the accessible name. */
  ariaLabel?: string;
  /** Body class names. Some dialogs use the extra `app__dlg__body` modifier; opt
   *  in by passing `{ bodyVariant: 'app' }`. Defaults to the bare base class so
   *  existing visual snapshots stay byte-identical. */
  bodyVariant?: 'base' | 'app';
}

export interface DialogShell {
  overlay: HTMLDivElement;
  panel: HTMLDivElement;
  header: HTMLDivElement;
  body: HTMLDivElement;
  footer: HTMLDivElement;
  opener: HTMLElement | null;
}

/** Builds the overlay → panel → header/body/footer skeleton shared by every
 *  app dialog and records the previously-focused element so focus can be
 *  restored on close. The caller is responsible for appending widgets to
 *  `body`/`footer` and calling `document.body.appendChild(overlay)`. */
export const createDialogShell = (opts: DialogShellOptions): DialogShell => {
  const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg app__dlg';
  overlay.setAttribute('role', opts.role ?? 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', opts.ariaLabel ?? opts.title);

  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel app__dlg__panel';
  overlay.appendChild(panel);

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = opts.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className =
    opts.bodyVariant === 'app' ? 'fc-fmtdlg__body app__dlg__body' : 'fc-fmtdlg__body';
  panel.appendChild(body);

  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);

  return { overlay, panel, header, body, footer, opener };
};

export interface DialogButtonOptions {
  label: string;
  variant?: 'primary' | 'secondary';
  destructive?: boolean;
}

export const appendDialogButton = (
  footer: HTMLElement,
  opts: DialogButtonOptions,
): HTMLButtonElement => {
  const button = document.createElement('button');
  button.type = 'button';
  const classes = ['fc-fmtdlg__btn'];
  if (opts.variant === 'primary') classes.push('fc-fmtdlg__btn--primary');
  if (opts.destructive) classes.push('app__dlg__btn--danger');
  button.className = classes.join(' ');
  button.textContent = opts.label;
  footer.appendChild(button);
  return button;
};

/** Convenience: cancel pair (secondary + primary), appended in that order
 *  so the primary action sits on the right like every other app dialog. */
export const appendDialogActions = (
  footer: HTMLElement,
  opts: { cancelLabel: string; okLabel: string; destructive?: boolean },
): { cancelBtn: HTMLButtonElement; okBtn: HTMLButtonElement } => {
  const cancelBtn = appendDialogButton(footer, { label: opts.cancelLabel });
  const okBtn = appendDialogButton(footer, {
    label: opts.okLabel,
    variant: 'primary',
    destructive: opts.destructive,
  });
  return { cancelBtn, okBtn };
};

export interface DialogChoiceButtonOptions {
  label: string;
  className?: string;
  title?: string;
  ariaLabel?: string;
}

export const createDialogChoiceButton = (opts: DialogChoiceButtonOptions): HTMLButtonElement => {
  const button = document.createElement('button');
  button.type = 'button';
  button.className = opts.className ?? 'app__cf-choice';
  button.textContent = opts.label;
  if (opts.title) button.title = opts.title;
  button.setAttribute('aria-label', opts.ariaLabel ?? opts.label);
  return button;
};

export interface DialogLifecycleHooks<T> {
  shell: DialogShell;
  resolve: (value: T) => void;
  onSubmit?: () => void;
  onCancel: () => T;
}

type DialogCanceller = () => void;

const activeDialogCancellers = new Set<DialogCanceller>();

export const cancelOpenAppDialogs = (): void => {
  for (const cancel of Array.from(activeDialogCancellers)) cancel();
};

/** Wires up the common close/keyboard flow. Returns a `finish` closure that
 *  cleans up listeners, restores focus, removes the overlay, and resolves
 *  the promise. */
export const installDialogLifecycle = <T>(
  hooks: DialogLifecycleHooks<T>,
): {
  finish: (value: T) => void;
  onKey: (event: KeyboardEvent) => void;
} => {
  const { shell, resolve, onSubmit, onCancel } = hooks;
  let done = false;
  const cancel = (): void => {
    finish(onCancel());
  };
  const finish = (value: T): void => {
    if (done) return;
    done = true;
    activeDialogCancellers.delete(cancel);
    shell.overlay.removeEventListener('keydown', onKey);
    restoreDialogFocus(shell.overlay, shell.opener);
    shell.overlay.remove();
    resolve(value);
  };
  const onKey = (event: KeyboardEvent): void => {
    event.stopPropagation();
    if (trapDialogTab(shell.overlay, event)) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      finish(onCancel());
    } else if (event.key === 'Enter' && onSubmit) {
      event.preventDefault();
      onSubmit();
    }
  };
  shell.overlay.addEventListener('click', (event) => {
    if (event.target === shell.overlay) finish(onCancel());
  });
  shell.overlay.addEventListener('keydown', onKey);
  activeDialogCancellers.add(cancel);
  return { finish, onKey };
};

export const mountDialog = (
  shell: DialogShell,
  focusInit: HTMLElement | (() => void) | null,
): void => {
  document.body.appendChild(shell.overlay);
  if (!focusInit) return;
  requestAnimationFrame(() => {
    if (typeof focusInit === 'function') focusInit();
    else focusInit.focus({ preventScroll: true });
  });
};

/** Adds a labelled text input row in the standard dialog body. */
export const appendInputRow = (
  body: HTMLElement,
  labelText: string,
  config: {
    type?: 'text' | 'number';
    initial?: string;
    placeholder?: string;
    min?: number;
    max?: number;
    step?: number;
  } = {},
): HTMLInputElement => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
  const label = document.createElement('label');
  label.className = 'app__dlg__label';
  label.textContent = labelText;
  const input = document.createElement('input');
  input.type = config.type ?? 'text';
  input.className = 'app__dlg__input';
  if (config.initial !== undefined) input.value = config.initial;
  if (config.placeholder) input.placeholder = config.placeholder;
  if (typeof config.min === 'number') input.min = String(config.min);
  if (typeof config.max === 'number') input.max = String(config.max);
  if (typeof config.step === 'number') input.step = String(config.step);
  label.appendChild(input);
  row.appendChild(label);
  body.appendChild(row);
  return input;
};

export const appendErrorRow = (body: HTMLElement): HTMLDivElement => {
  const errorRow = document.createElement('div');
  errorRow.className = 'app__dlg__error';
  errorRow.setAttribute('role', 'alert');
  errorRow.hidden = true;
  body.appendChild(errorRow);
  return errorRow;
};

export const focusAndSelectInput = (input: HTMLInputElement): void => {
  input.focus({ preventScroll: true });
  input.select();
};

export const showDialogError = (errorRow: HTMLElement, message: string): void => {
  errorRow.textContent = message;
  errorRow.hidden = false;
};

export const clearDialogError = (errorRow: HTMLElement): void => {
  errorRow.hidden = true;
  errorRow.textContent = '';
};

export const showInputError = (
  errorRow: HTMLElement,
  input: HTMLInputElement,
  message: string,
): void => {
  showDialogError(errorRow, message);
  focusAndSelectInput(input);
};

export interface DialogNameField {
  input: HTMLInputElement;
  focus: () => void;
  valueOrError: (errorRow: HTMLElement, message: string) => string | null;
}

export const appendDialogNameField = (
  body: HTMLElement,
  labelText: string,
  initial: string,
): DialogNameField => {
  const input = appendInputRow(body, labelText, { initial });
  input.dataset.dialogField = 'name';
  return {
    input,
    focus: () => focusAndSelectInput(input),
    valueOrError: (errorRow, message) => {
      const value = input.value.trim();
      if (!value) {
        showInputError(errorRow, input, message);
        return null;
      }
      return value;
    },
  };
};
