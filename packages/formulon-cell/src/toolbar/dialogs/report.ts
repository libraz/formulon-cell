import { createDialogShell, installDialogLifecycle, mountDialog } from './shell.js';

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
  infoLabel?: string;
  warningLabel?: string;
}

/** Excel 365-styled one-button report dialog for Review/Add-ins surfaces. */
export const showReport = (opts: ReportOptions): Promise<void> =>
  new Promise<void>((resolve) => {
    const shell = createDialogShell({ title: opts.title, bodyVariant: 'app' });

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
        const tag =
          item.severity === 'warning'
            ? (opts.warningLabel ?? 'Warning')
            : (opts.infoLabel ?? 'Info');
        label.textContent = `${tag} · ${item.label}`;
        const detail = document.createElement('div');
        detail.textContent = item.detail;
        row.append(label, detail);
        list.appendChild(row);
      }
    }
    shell.body.appendChild(list);

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
    closeBtn.textContent = opts.closeLabel ?? 'Close';
    shell.footer.appendChild(closeBtn);

    const lifecycle = installDialogLifecycle<void>({
      shell,
      resolve: () => resolve(),
      onCancel: () => undefined,
      onSubmit: () => lifecycle.finish(undefined),
    });
    closeBtn.addEventListener('click', () => lifecycle.finish(undefined));

    mountDialog(shell, closeBtn);
  });
