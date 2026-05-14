import {
  boundingRange,
  findMatchingCells,
  type GoToScope,
  type GoToSpecialKind,
} from '../commands/goto-special.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { inheritHostTokens } from './inherit-host-tokens.js';

export interface GoToDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  /** Lazy workbook accessor — keeps the dialog in lockstep with `setWorkbook`
   *  swaps so a Go To Special invocation always queries the live engine. */
  getWb: () => WorkbookHandle;
  strings?: Strings;
}

export interface GoToDialogHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const KINDS: readonly GoToSpecialKind[] = [
  'blanks',
  'non-blanks',
  'formulas',
  'constants',
  'numbers',
  'text',
  'errors',
  'data-validation',
  'conditional-format',
];

/**
 * Spreadsheet-style "Go To Special" dialog. Lets the user pick a category (blanks,
 * formulas, errors, validation, …) and rewrites the selection to the
 * bounding range of every matching cell on the active sheet (or just inside
 * the current selection when one is provided).
 *
 * Lifecycle mirrors the other modals: `attach…()` mounts a hidden overlay
 * on the host; `open()` shows it, `close()` hides it, `detach()` tears it
 * down. The OK button runs the predicate; on no matches an inline status
 * message is shown and the dialog stays open. On 1+ matches the selection
 * jumps to the bounding rect and the dialog closes.
 */
export function attachGoToDialog(deps: GoToDialogDeps): GoToDialogHandle {
  const { host, store, getWb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.goToDialog;

  const overlay = document.createElement('div');
  overlay.className = 'fc-fmtdlg fc-goto';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-modal', 'true');
  overlay.setAttribute('aria-label', t.title);
  overlay.hidden = true;

  const panel = document.createElement('div');
  panel.className = 'fc-fmtdlg__panel fc-goto__panel';
  overlay.appendChild(panel);

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body';
  panel.appendChild(body);

  // ── Scope (only meaningful when current selection is multi-cell) ───────
  const scopeLegend = document.createElement('div');
  scopeLegend.className = 'fc-goto__legend';
  scopeLegend.textContent = t.scopeLabel;
  body.appendChild(scopeLegend);

  const scopeGroup = document.createElement('div');
  scopeGroup.className = 'fc-goto__scope';
  scopeGroup.setAttribute('role', 'radiogroup');
  scopeGroup.setAttribute('aria-label', t.scopeLabel);
  body.appendChild(scopeGroup);

  const scopeName = `fc-goto-scope-${Math.random().toString(36).slice(2, 8)}`;
  const makeScopeRadio = (value: GoToScope, label: string): HTMLInputElement => {
    const wrap = document.createElement('label');
    wrap.className = 'fc-goto__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = scopeName;
    input.value = value;
    const span = document.createElement('span');
    span.textContent = label;
    wrap.append(input, span);
    scopeGroup.appendChild(wrap);
    return input;
  };
  const scopeSheet = makeScopeRadio('sheet', t.scopeSheet);
  const scopeSelection = makeScopeRadio('selection', t.scopeSelection);
  scopeSheet.checked = true;

  // ── Kind list ──────────────────────────────────────────────────────────
  const kindLegend = document.createElement('div');
  kindLegend.className = 'fc-goto__legend';
  kindLegend.textContent = t.kindLabel;
  body.appendChild(kindLegend);

  const kindList = document.createElement('div');
  kindList.className = 'fc-goto__kinds';
  kindList.setAttribute('role', 'radiogroup');
  kindList.setAttribute('aria-label', t.kindLabel);
  body.appendChild(kindList);

  const kindName = `fc-goto-kind-${Math.random().toString(36).slice(2, 8)}`;
  const kindLabels: Record<GoToSpecialKind, string> = {
    blanks: t.kindBlanks,
    'non-blanks': t.kindNonBlanks,
    formulas: t.kindFormulas,
    constants: t.kindConstants,
    numbers: t.kindNumbers,
    text: t.kindText,
    errors: t.kindErrors,
    'data-validation': t.kindDataValidation,
    'conditional-format': t.kindConditionalFormat,
  };
  const kindInputs = new Map<GoToSpecialKind, HTMLInputElement>();
  for (const k of KINDS) {
    const wrap = document.createElement('label');
    wrap.className = 'fc-goto__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = kindName;
    input.value = k;
    const span = document.createElement('span');
    span.textContent = kindLabels[k];
    wrap.append(input, span);
    kindList.appendChild(wrap);
    kindInputs.set(k, input);
  }
  // Default selection — spreadsheets open with `constants` highlighted; we mirror that.
  (kindInputs.get('constants') as HTMLInputElement).checked = true;

  // Inline status (shown when a search returns zero results).
  const statusLine = document.createElement('div');
  statusLine.className = 'fc-goto__status';
  statusLine.setAttribute('role', 'status');
  statusLine.setAttribute('aria-live', 'polite');
  body.appendChild(statusLine);

  // Footer: Cancel / OK
  const footer = document.createElement('div');
  footer.className = 'fc-fmtdlg__footer';
  panel.appendChild(footer);
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'fc-fmtdlg__btn';
  cancelBtn.textContent = t.cancel;
  const okBtn = document.createElement('button');
  okBtn.type = 'button';
  okBtn.className = 'fc-fmtdlg__btn fc-fmtdlg__btn--primary';
  okBtn.textContent = t.ok;
  footer.append(cancelBtn, okBtn);

  // Body-portal so the modal escapes `.fc-host`'s `contain: strict`.
  inheritHostTokens(host, overlay);
  document.body.appendChild(overlay);

  const isSelectionMulti = (): boolean => {
    const r = store.getState().selection.range;
    return r.r1 > r.r0 || r.c1 > r.c0;
  };

  const syncScopeAvailability = (): void => {
    const multi = isSelectionMulti();
    // Disable selection scope when only one cell is selected — matching the spreadsheet convention,
    // which silently widens to "active sheet" in that case.
    scopeSelection.disabled = !multi;
    if (!multi) {
      scopeSheet.checked = true;
    }
  };

  const getCheckedKind = (): GoToSpecialKind => {
    for (const [k, input] of kindInputs) if (input.checked) return k;
    return 'constants';
  };
  const getCheckedScope = (): GoToScope => (scopeSelection.checked ? 'selection' : 'sheet');

  const onOk = (): void => {
    statusLine.textContent = '';
    const kind = getCheckedKind();
    const scope = getCheckedScope();
    const wb = getWb();
    const matches = findMatchingCells(wb, store, scope, kind);
    if (matches.length === 0) {
      statusLine.textContent = t.noResults;
      return;
    }
    const range = boundingRange(matches);
    const first = matches[0];
    if (!first) return;
    store.setState((s) => ({
      ...s,
      selection: {
        active: first,
        anchor: first,
        range,
        extraRanges: [],
      },
    }));
    api.close();
  };
  const onCancel = (): void => api.close();

  const onOverlayClick = (e: MouseEvent): void => {
    if (e.target === overlay) api.close();
  };
  const onOverlayKey = (e: KeyboardEvent): void => {
    e.stopPropagation();
    if (e.key === 'Escape') {
      e.preventDefault();
      api.close();
    } else if (e.key === 'Enter') {
      // Enter on any control inside the modal commits — matches the spreadsheet convention.
      e.preventDefault();
      onOk();
    }
  };

  okBtn.addEventListener('click', onOk);
  cancelBtn.addEventListener('click', onCancel);
  overlay.addEventListener('click', onOverlayClick);
  overlay.addEventListener('keydown', onOverlayKey);

  const api: GoToDialogHandle = {
    open(): void {
      statusLine.textContent = '';
      syncScopeAvailability();
      overlay.hidden = false;
      requestAnimationFrame(() => {
        // Focus the first kind radio so keyboard users can immediately arrow
        // through the list.
        const first = kindInputs.values().next().value;
        if (first) first.focus();
      });
    },
    close(): void {
      overlay.hidden = true;
      host.focus();
    },
    detach(): void {
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      overlay.removeEventListener('click', onOverlayClick);
      overlay.removeEventListener('keydown', onOverlayKey);
      overlay.remove();
    },
  };
  return api;
}
