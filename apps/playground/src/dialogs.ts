export interface PromptOptions {
  title: string;
  label: string;
  initial?: string;
  placeholder?: string;
  okLabel?: string;
  cancelLabel?: string;
  validate?: (value: string) => string | null;
}

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

const trapDialogTab = (root: HTMLElement, event: KeyboardEvent): boolean => {
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

const restoreDialogFocus = (overlay: HTMLElement, opener: HTMLElement | null): void => {
  if (!opener) return;
  if (overlay.contains(document.activeElement) || document.activeElement === document.body) {
    opener.focus({ preventScroll: true });
  }
};

/** Excel 365-styled modal prompt. Returns the entered value, or `null`
 *  when the user cancels. Replaces native `window.prompt`. */
export const showPrompt = (opts: PromptOptions): Promise<string | null> => {
  return new Promise<string | null>((resolve) => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const overlay = document.createElement('div');
    overlay.className = 'fc-fmtdlg app__dlg';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', opts.title);

    const panel = document.createElement('div');
    panel.className = 'fc-fmtdlg__panel app__dlg__panel';
    overlay.appendChild(panel);

    const header = document.createElement('div');
    header.className = 'fc-fmtdlg__header';
    header.textContent = opts.title;
    panel.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-fmtdlg__body';
    panel.appendChild(body);

    const row = document.createElement('div');
    row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
    const label = document.createElement('label');
    label.className = 'app__dlg__label';
    label.textContent = opts.label;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'app__dlg__input';
    input.value = opts.initial ?? '';
    if (opts.placeholder) input.placeholder = opts.placeholder;
    label.appendChild(input);
    row.appendChild(label);
    body.appendChild(row);

    const errorRow = document.createElement('div');
    errorRow.className = 'app__dlg__error';
    errorRow.setAttribute('role', 'alert');
    errorRow.hidden = true;
    body.appendChild(errorRow);

    const footer = document.createElement('div');
    footer.className = 'fc-fmtdlg__footer';
    panel.appendChild(footer);

    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'fc-fmtdlg__btn';
    cancelBtn.textContent = opts.cancelLabel ?? 'Cancel';
    footer.appendChild(cancelBtn);

    const okBtn = document.createElement('button');
    okBtn.type = 'button';
    okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
    okBtn.textContent = opts.okLabel ?? 'OK';
    footer.appendChild(okBtn);

    let done = false;
    const finish = (value: string | null): void => {
      if (done) return;
      done = true;
      overlay.removeEventListener('keydown', onKey);
      restoreDialogFocus(overlay, opener);
      overlay.remove();
      resolve(value);
    };
    const onOk = (): void => {
      const v = input.value;
      const err = opts.validate?.(v) ?? null;
      if (err) {
        errorRow.textContent = err;
        errorRow.hidden = false;
        input.focus();
        input.select();
        return;
      }
      finish(v);
    };
    const onCancel = (): void => finish(null);
    const onKey = (e: KeyboardEvent): void => {
      e.stopPropagation();
      if (trapDialogTab(overlay, e)) return;
      if (e.key === 'Escape') {
        e.preventDefault();
        onCancel();
      } else if (e.key === 'Enter') {
        e.preventDefault();
        onOk();
      }
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', onCancel);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) onCancel();
    });
    overlay.addEventListener('keydown', onKey);

    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      input.focus();
      input.select();
    });
  });
};

export interface ConfirmOptions {
  title: string;
  message: string;
  okLabel?: string;
  cancelLabel?: string;
  destructive?: boolean;
}

/** Excel 365-styled modal confirm. Returns true on accept, false on
 *  cancel/dismiss. Replaces native `window.confirm`. */
export const showConfirm = (opts: ConfirmOptions): Promise<boolean> => {
  return new Promise<boolean>((resolve) => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const overlay = document.createElement('div');
    overlay.className = 'fc-fmtdlg app__dlg';
    overlay.setAttribute('role', 'alertdialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', opts.title);

    const panel = document.createElement('div');
    panel.className = 'fc-fmtdlg__panel app__dlg__panel';
    overlay.appendChild(panel);

    const header = document.createElement('div');
    header.className = 'fc-fmtdlg__header';
    header.textContent = opts.title;
    panel.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-fmtdlg__body app__dlg__body';
    const msg = document.createElement('p');
    msg.className = 'app__dlg__message';
    msg.textContent = opts.message;
    body.appendChild(msg);
    panel.appendChild(body);

    const footer = document.createElement('div');
    footer.className = 'fc-fmtdlg__footer';
    panel.appendChild(footer);

    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'fc-fmtdlg__btn';
    cancelBtn.textContent = opts.cancelLabel ?? 'Cancel';
    footer.appendChild(cancelBtn);

    const okBtn = document.createElement('button');
    okBtn.type = 'button';
    okBtn.className = `fc-fmtdlg__btn fc-fmtdlg__btn--primary${
      opts.destructive ? ' app__dlg__btn--danger' : ''
    }`;
    okBtn.textContent = opts.okLabel ?? 'OK';
    footer.appendChild(okBtn);

    let done = false;
    const finish = (value: boolean): void => {
      if (done) return;
      done = true;
      overlay.removeEventListener('keydown', onKey);
      restoreDialogFocus(overlay, opener);
      overlay.remove();
      resolve(value);
    };
    const onKey = (e: KeyboardEvent): void => {
      e.stopPropagation();
      if (trapDialogTab(overlay, e)) return;
      if (e.key === 'Escape') {
        e.preventDefault();
        finish(false);
      } else if (e.key === 'Enter') {
        e.preventDefault();
        finish(true);
      }
    };
    okBtn.addEventListener('click', () => finish(true));
    cancelBtn.addEventListener('click', () => finish(false));
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) finish(false);
    });
    overlay.addEventListener('keydown', onKey);

    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      okBtn.focus();
    });
  });
};

export interface MessageOptions {
  title: string;
  message: string;
  okLabel?: string;
}

/** Excel 365-styled message dialog. Use for one-button errors/info instead
 *  of native `window.alert`, keeping focus and stacking inside app chrome. */
export const showMessage = (opts: MessageOptions): Promise<void> => {
  return new Promise<void>((resolve) => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const overlay = document.createElement('div');
    overlay.className = 'fc-fmtdlg app__dlg';
    overlay.setAttribute('role', 'alertdialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', opts.title);

    const panel = document.createElement('div');
    panel.className = 'fc-fmtdlg__panel app__dlg__panel';
    overlay.appendChild(panel);

    const header = document.createElement('div');
    header.className = 'fc-fmtdlg__header';
    header.textContent = opts.title;
    panel.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-fmtdlg__body app__dlg__body';
    const msg = document.createElement('p');
    msg.className = 'app__dlg__message';
    msg.textContent = opts.message;
    body.appendChild(msg);
    panel.appendChild(body);

    const footer = document.createElement('div');
    footer.className = 'fc-fmtdlg__footer';
    panel.appendChild(footer);

    const okBtn = document.createElement('button');
    okBtn.type = 'button';
    okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
    okBtn.textContent = opts.okLabel ?? 'OK';
    footer.appendChild(okBtn);

    let done = false;
    const finish = (): void => {
      if (done) return;
      done = true;
      overlay.removeEventListener('keydown', onKey);
      restoreDialogFocus(overlay, opener);
      overlay.remove();
      resolve();
    };
    const onKey = (e: KeyboardEvent): void => {
      e.stopPropagation();
      if (trapDialogTab(overlay, e)) return;
      if (e.key === 'Escape' || e.key === 'Enter') {
        e.preventDefault();
        finish();
      }
    };
    okBtn.addEventListener('click', finish);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) finish();
    });
    overlay.addEventListener('keydown', onKey);

    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      okBtn.focus();
    });
  });
};

export interface ReportItem {
  severity: 'warning' | 'info';
  label: string;
  detail: string;
}

export interface ReportOptions {
  title: string;
  items: readonly ReportItem[];
  emptyLabel?: string;
  closeLabel?: string;
}

/** Excel 365-styled one-button report dialog for Review/Add-ins surfaces. */
export const showReport = (opts: ReportOptions): Promise<void> => {
  return new Promise<void>((resolve) => {
    const opener = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const overlay = document.createElement('div');
    overlay.className = 'fc-fmtdlg app__dlg';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', opts.title);

    const panel = document.createElement('div');
    panel.className = 'fc-fmtdlg__panel app__dlg__panel';
    overlay.appendChild(panel);

    const header = document.createElement('div');
    header.className = 'fc-fmtdlg__header';
    header.textContent = opts.title;
    panel.appendChild(header);

    const body = document.createElement('div');
    body.className = 'fc-fmtdlg__body app__dlg__body';
    panel.appendChild(body);

    const list = document.createElement('div');
    list.className = 'app__dlg__list';
    if (opts.items.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'app__dlg__note';
      empty.textContent = opts.emptyLabel ?? 'No issues found.';
      list.appendChild(empty);
    } else {
      for (const item of opts.items) {
        const row = document.createElement('div');
        row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
        const label = document.createElement('strong');
        label.textContent = `${item.severity === 'warning' ? 'Warning' : 'Info'} · ${item.label}`;
        const detail = document.createElement('div');
        detail.textContent = item.detail;
        row.append(label, detail);
        list.appendChild(row);
      }
    }
    body.appendChild(list);

    const footer = document.createElement('div');
    footer.className = 'fc-fmtdlg__footer';
    panel.appendChild(footer);

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
    closeBtn.textContent = opts.closeLabel ?? 'Close';
    footer.appendChild(closeBtn);

    let done = false;
    const finish = (): void => {
      if (done) return;
      done = true;
      overlay.removeEventListener('keydown', onKey);
      restoreDialogFocus(overlay, opener);
      overlay.remove();
      resolve();
    };
    const onKey = (e: KeyboardEvent): void => {
      e.stopPropagation();
      if (trapDialogTab(overlay, e)) return;
      if (e.key === 'Escape' || e.key === 'Enter') {
        e.preventDefault();
        finish();
      }
    };
    closeBtn.addEventListener('click', finish);
    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) finish();
    });
    overlay.addEventListener('keydown', onKey);

    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      closeBtn.focus();
    });
  });
};
