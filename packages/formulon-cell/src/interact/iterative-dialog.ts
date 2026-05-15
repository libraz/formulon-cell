import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { createDialogShell } from './dialog-shell.js';

export interface IterativeDialogDeps {
  host: HTMLElement;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so flipping the iterative mode always targets the live engine. */
  getWb: () => WorkbookHandle | null;
  strings?: Strings;
}

export interface IterativeDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

interface IterativeSettings {
  enabled: boolean;
  maxIterations: number;
  maxChange: number;
}

const DEFAULTS: IterativeSettings = { enabled: false, maxIterations: 100, maxChange: 0.001 };

/**
 * Spreadsheet-style "Enable iterative calculation" controls. Mirrors File →
 * Options → Formulas → Calculation options. The dialog persists settings on
 * `wb.setIterative` and wires a progress callback that surfaces residual /
 * iteration counts to a status-bar span via the `fc:iterative-progress`
 * custom event.
 */
export function attachIterativeDialog(deps: IterativeDialogDeps): IterativeDialogHandle {
  const { host, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.iterativeDialog;

  const shell = createDialogShell({
    host,
    className: 'fc-iterdlg',
    ariaLabel: t.title,
    onDismiss: () => api.close(),
  });

  const header = document.createElement('div');
  header.className = 'fc-iterdlg__header';
  header.textContent = t.title;
  shell.panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-iterdlg__body';
  shell.panel.appendChild(body);

  const note = document.createElement('p');
  note.className = 'fc-iterdlg__note';
  note.textContent = t.note;
  body.appendChild(note);

  const enableLabel = document.createElement('label');
  enableLabel.className = 'fc-iterdlg__row';
  const enableInput = document.createElement('input');
  enableInput.type = 'checkbox';
  const enableSpan = document.createElement('span');
  enableSpan.textContent = t.enable;
  enableLabel.append(enableInput, enableSpan);
  body.appendChild(enableLabel);

  const maxIterRow = document.createElement('label');
  maxIterRow.className = 'fc-iterdlg__row';
  const maxIterLabel = document.createElement('span');
  maxIterLabel.textContent = t.maxIterations;
  const maxIterInput = document.createElement('input');
  maxIterInput.type = 'number';
  maxIterInput.min = '1';
  maxIterInput.max = '32767';
  maxIterInput.step = '1';
  maxIterRow.append(maxIterLabel, maxIterInput);
  body.appendChild(maxIterRow);

  const maxChangeRow = document.createElement('label');
  maxChangeRow.className = 'fc-iterdlg__row';
  const maxChangeLabel = document.createElement('span');
  maxChangeLabel.textContent = t.maxChange;
  const maxChangeInput = document.createElement('input');
  maxChangeInput.type = 'text';
  maxChangeInput.spellcheck = false;
  maxChangeInput.autocomplete = 'off';
  maxChangeRow.append(maxChangeLabel, maxChangeInput);
  body.appendChild(maxChangeRow);

  const status = document.createElement('div');
  status.className = 'fc-iterdlg__status';
  body.appendChild(status);

  const footer = document.createElement('div');
  footer.className = 'fc-iterdlg__footer';
  shell.panel.appendChild(footer);
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-iterdlg__btn';
  cancelBtn.textContent = t.cancel;
  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-iterdlg__btn fc-iterdlg__btn--primary';
  okBtn.textContent = t.ok;
  footer.append(cancelBtn, okBtn);

  const draft: IterativeSettings = { ...DEFAULTS };

  const syncControls = (): void => {
    enableInput.checked = draft.enabled;
    maxIterInput.value = String(draft.maxIterations);
    maxChangeInput.value = String(draft.maxChange);
    maxIterInput.disabled = !draft.enabled;
    maxChangeInput.disabled = !draft.enabled;
  };

  const onEnable = (): void => {
    draft.enabled = enableInput.checked;
    syncControls();
  };
  const onIter = (): void => {
    const n = Number.parseInt(maxIterInput.value, 10);
    if (Number.isFinite(n)) draft.maxIterations = Math.max(1, Math.min(32767, n));
  };
  const onChange = (): void => {
    const n = Number.parseFloat(maxChangeInput.value);
    if (Number.isFinite(n) && n > 0) draft.maxChange = n;
  };

  const onOk = (): void => {
    const wb = getWb();
    if (!wb) {
      api.close();
      return;
    }
    const ok = wb.setIterative(draft.enabled, draft.maxIterations, draft.maxChange);
    if (!ok) {
      status.textContent = t.unsupported;
      return;
    }
    if (draft.enabled) {
      // Wire the progress callback to a host event the chrome can render.
      wb.setIterativeProgress((iteration, maxResidual, maxIterations) => {
        host.dispatchEvent(
          new CustomEvent('fc:iterative-progress', {
            detail: { iteration, maxResidual, maxIterations },
          }),
        );
      });
    } else {
      wb.setIterativeProgress(null);
    }
    api.close();
  };

  shell.on(enableInput, 'change', onEnable);
  shell.on(maxIterInput, 'input', onIter);
  shell.on(maxChangeInput, 'input', onChange);
  shell.on(okBtn, 'click', onOk);
  shell.on(cancelBtn, 'click', () => api.close());
  shell.on(shell.overlay, 'keydown', (e) => {
    const event = e as KeyboardEvent;
    event.stopPropagation();
    if (event.key === 'Enter') {
      event.preventDefault();
      onOk();
    }
  });

  const api: IterativeDialogHandle = {
    open(): void {
      status.textContent = '';
      syncControls();
      shell.open();
      requestAnimationFrame(() => enableInput.focus());
    },
    close(): void {
      shell.close();
      host.focus();
    },
    detach(): void {
      shell.dispose();
    },
  };
  return api;
}
