import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';

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
 * Excel-style "Enable iterative calculation" controls. Mirrors File →
 * Options → Formulas → Calculation options. The dialog persists settings on
 * `wb.setIterative` and wires a progress callback that surfaces residual /
 * iteration counts to a status-bar span via the `fc:iterative-progress`
 * custom event.
 */
export function attachIterativeDialog(deps: IterativeDialogDeps): IterativeDialogHandle {
  const { host, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.iterativeDialog;

  const overlay = document.createElement('div');
  overlay.className = 'fc-iterdlg';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = 'fc-iterdlg__panel';
  overlay.appendChild(panel);

  const header = document.createElement('div');
  header.className = 'fc-iterdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-iterdlg__body';
  panel.appendChild(body);

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
  panel.appendChild(footer);
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-iterdlg__btn';
  cancelBtn.textContent = t.cancel;
  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-iterdlg__btn fc-iterdlg__btn--primary';
  okBtn.textContent = t.ok;
  footer.append(cancelBtn, okBtn);

  host.appendChild(overlay);

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
  const onCancel = (): void => api.close();

  const onOverlayClick = (e: MouseEvent): void => {
    if (e.target === overlay) api.close();
  };
  const onOverlayKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    }
  };

  enableInput.addEventListener('change', onEnable);
  maxIterInput.addEventListener('input', onIter);
  maxChangeInput.addEventListener('input', onChange);
  okBtn.addEventListener('click', onOk);
  cancelBtn.addEventListener('click', onCancel);
  overlay.addEventListener('click', onOverlayClick);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: IterativeDialogHandle = {
    open(): void {
      status.textContent = '';
      syncControls();
      overlay.hidden = false;
      requestAnimationFrame(() => enableInput.focus());
    },
    close(): void {
      overlay.hidden = true;
      host.focus();
    },
    detach(): void {
      enableInput.removeEventListener('change', onEnable);
      maxIterInput.removeEventListener('input', onIter);
      maxChangeInput.removeEventListener('input', onChange);
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      overlay.removeEventListener('click', onOverlayClick);
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };
  return api;
}
