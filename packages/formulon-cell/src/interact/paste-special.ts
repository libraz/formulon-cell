import {
  type PasteOperation,
  type PasteSpecialOptions,
  type PasteWhat,
  pasteSpecial,
} from '../commands/clipboard/paste-special.js';
import type { ClipboardSnapshot } from '../commands/clipboard/snapshot.js';
import { type History, recordFormatChange } from '../commands/history.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SpreadsheetStore } from '../store/store.js';
import { createDialogShell } from './dialog-shell.js';

export interface PasteSpecialDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  /** Source for the structured clipboard. The clipboard adapter mutates this
   *  snapshot on copy/cut; we read it lazily on open. */
  getSnapshot: () => ClipboardSnapshot | null;
  /** UI string dictionary. */
  strings?: Strings;
  /** Shared history. When provided, the entire paste (cell writes + format
   *  changes) is bundled into a single undoable transaction. */
  history?: History | null;
  /** Refresh cached cells after the paste. */
  onAfterCommit: () => void;
}

export interface PasteSpecialHandle {
  open(): void;
  close(): void;
  detach(): void;
}

const buildWhatOptions = (s: Strings): { id: PasteWhat; label: string }[] => {
  const t = s.pasteSpecialDialog;
  return [
    { id: 'all', label: t.pasteAll },
    { id: 'formulas', label: t.pasteFormulas },
    { id: 'values', label: t.pasteValues },
    { id: 'formats', label: t.pasteFormats },
    { id: 'formulas-and-numfmt', label: t.pasteFormulasAndNumFmt },
    { id: 'values-and-numfmt', label: t.pasteValuesAndNumFmt },
  ];
};

const buildOpOptions = (s: Strings): { id: PasteOperation; label: string }[] => {
  const t = s.pasteSpecialDialog;
  return [
    { id: 'none', label: t.opNone },
    { id: 'add', label: t.opAdd },
    { id: 'subtract', label: t.opSubtract },
    { id: 'multiply', label: t.opMultiply },
    { id: 'divide', label: t.opDivide },
  ];
};

/**
 * desktop-spreadsheet-compatible "形式を選択して貼り付け" dialog. Operates on the
 * structured snapshot captured when the user copied within the workbook.
 * If no internal snapshot is available, the dialog refuses to open.
 */
export function attachPasteSpecial(deps: PasteSpecialDeps): PasteSpecialHandle {
  const { host, store, wb } = deps;
  const strings = deps.strings ?? defaultStrings;
  const t = strings.pasteSpecialDialog;
  const whatOptions = buildWhatOptions(strings);
  const opOptions = buildOpOptions(strings);

  const shell = createDialogShell({
    host,
    className: 'fc-pastesp',
    ariaLabel: t.title,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  shell.panel.classList.add('fc-fmtdlg__panel', 'fc-pastesp__panel');
  const { overlay, panel } = shell;

  const header = document.createElement('div');
  header.className = 'fc-fmtdlg__header';
  header.textContent = t.title;
  panel.appendChild(header);

  const body = document.createElement('div');
  body.className = 'fc-fmtdlg__body fc-pastesp__body';
  panel.appendChild(body);

  // Two-column layout: paste section (left), operation section (right).
  const cols = document.createElement('div');
  cols.className = 'fc-pastesp__cols';
  body.appendChild(cols);

  const whatGroup = makeFieldset(t.sectionPaste);
  cols.appendChild(whatGroup.fieldset);
  const whatRadios = new Map<PasteWhat, HTMLInputElement>();
  const whatName = `fc-pastesp-what-${Math.random().toString(36).slice(2, 8)}`;
  for (const opt of whatOptions) {
    const { input, label } = makeRadio(whatName, opt.id, opt.label);
    whatRadios.set(opt.id, input);
    whatGroup.body.appendChild(label);
  }

  const opGroup = makeFieldset(t.sectionOperation);
  cols.appendChild(opGroup.fieldset);
  const opRadios = new Map<PasteOperation, HTMLInputElement>();
  const opName = `fc-pastesp-op-${Math.random().toString(36).slice(2, 8)}`;
  for (const opt of opOptions) {
    const { input, label } = makeRadio(opName, opt.id, opt.label);
    opRadios.set(opt.id, input);
    opGroup.body.appendChild(label);
  }

  // Bottom row: skipBlanks + transpose
  const bottomRow = document.createElement('div');
  bottomRow.className = 'fc-pastesp__bottomrow';
  body.appendChild(bottomRow);

  const skipBlanks = makeCheck(t.skipBlanks);
  const transpose = makeCheck(t.transpose);
  bottomRow.append(skipBlanks.label, transpose.label);

  // Footer
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

  // Initial defaults
  const setDefaults = (): void => {
    const allWhat = whatRadios.get('all');
    const noneOp = opRadios.get('none');
    if (allWhat) allWhat.checked = true;
    if (noneOp) noneOp.checked = true;
    skipBlanks.input.checked = false;
    transpose.input.checked = false;
  };
  setDefaults();

  const close = (): void => {
    shell.close();
    host.focus();
  };

  const open = (): void => {
    if (!deps.getSnapshot()) {
      // Nothing on the internal clipboard — the dialog has nothing to paste.
      // Spreadsheets fall back to a much older dialog here; we simply no-op so the
      // standard ⌘V paste path can fill in.
      return;
    }
    setDefaults();
    shell.open();
    // Focus first radio for keyboard nav.
    whatRadios.get('all')?.focus();
  };

  const history = deps.history ?? null;
  const apply = (): void => {
    const snap = deps.getSnapshot();
    if (!snap) {
      close();
      return;
    }
    const what =
      [...whatRadios.entries()].find(([, el]) => el.checked)?.[0] ?? ('all' as PasteWhat);
    const operation =
      [...opRadios.entries()].find(([, el]) => el.checked)?.[0] ?? ('none' as PasteOperation);
    const opts: PasteSpecialOptions = {
      what,
      operation,
      skipBlanks: skipBlanks.input.checked,
      transpose: transpose.input.checked,
    };
    // Bundle every value/format mutation into one undoable step. Cell writes
    // route through the workbook (which pushes to the same history); the
    // format slice change goes through recordFormatChange to capture both.
    let result: ReturnType<typeof pasteSpecial> = null;
    if (history) {
      history.begin();
      try {
        recordFormatChange(history, store, () => {
          result = pasteSpecial(store.getState(), store, wb, snap, opts);
        });
      } finally {
        history.end();
      }
    } else {
      result = pasteSpecial(store.getState(), store, wb, snap, opts);
    }
    close();
    if (result) deps.onAfterCommit();
  };

  shell.on(cancelBtn, 'click', close);
  shell.on(okBtn, 'click', apply);
  const onKey = (e: KeyboardEvent): void => {
    if (overlay.hidden) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      close();
    } else if (e.key === 'Enter') {
      e.preventDefault();
      apply();
    }
  };
  shell.on(document, 'keydown', onKey as EventListener, true);

  return {
    open,
    close,
    detach() {
      shell.dispose();
    },
  };
}

function makeFieldset(legend: string): { fieldset: HTMLDivElement; body: HTMLDivElement } {
  const fieldset = document.createElement('div');
  fieldset.className = 'fc-pastesp__group';
  const lg = document.createElement('div');
  lg.className = 'fc-pastesp__legend';
  lg.textContent = legend;
  const body = document.createElement('div');
  body.className = 'fc-pastesp__list';
  body.setAttribute('role', 'radiogroup');
  body.setAttribute('aria-label', legend);
  fieldset.append(lg, body);
  return { fieldset, body };
}

function makeRadio(
  name: string,
  value: string,
  label: string,
): { input: HTMLInputElement; label: HTMLLabelElement } {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__radio';
  const input = document.createElement('input');
  input.type = 'radio';
  input.name = name;
  input.value = value;
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return { input, label: wrap };
}

function makeCheck(label: string): { input: HTMLInputElement; label: HTMLLabelElement } {
  const wrap = document.createElement('label');
  wrap.className = 'fc-fmtdlg__check';
  const input = document.createElement('input');
  input.type = 'checkbox';
  const span = document.createElement('span');
  span.textContent = label;
  wrap.append(input, span);
  return { input, label: wrap };
}
