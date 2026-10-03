import type { History } from '../commands/history.js';
import { setSparkline } from '../commands/sparkline.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import { colLabel } from '../render/geometry.js';
import type { SparklineKind, SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';

/** A small Insert Sparkline chooser. The source may be any
 *  rectangular A1 range; the destination is one cell because the engine's
 *  sparkline model stores one spec per host cell. */
export interface MacSparklineDialogDeps {
  host: HTMLElement;
  store: SpreadsheetStore;
  history?: History | null;
  getWb?: () => WorkbookHandle | null;
  strings?: Strings;
  /** Optional policy gate supplied by a host action provider. */
  canCommit?: (source: Range, destination: Addr) => boolean;
  onCommitted?: (source: Range, destination: Addr, kind: SparklineKind) => void;
}

export interface MacSparklineDialogHandle {
  open(): void;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  detach(): void;
}

interface ResolvedRange extends Range {
  explicitSheet: boolean;
}

const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

const formatRange = (range: Range): string => {
  const start = `${colLabel(range.c0)}${range.r0 + 1}`;
  const end = `${colLabel(range.c1)}${range.r1 + 1}`;
  return start === end ? start : `${start}:${end}`;
};

const formatSheetRange = (sheetName: string, range: Range): string => {
  const body = formatRange(range);
  if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName)) return `${sheetName}!${body}`;
  return `'${sheetName.replace(/'/g, "''")}'!${body}`;
};

const sheetIndex = (wb: WorkbookHandle, name: string): number => {
  const target = name.toLocaleLowerCase();
  for (let i = 0; i < wb.sheetCount; i += 1) {
    if (wb.sheetName(i).toLocaleLowerCase() === target) return i;
  }
  return -1;
};

const resolveRange = (
  raw: string,
  fallbackSheet: number,
  wb: WorkbookHandle | null,
): ResolvedRange | null => {
  const parsed = parseRangeRef(raw);
  if (!parsed) return null;
  const sheet =
    parsed.sheetName === null ? fallbackSheet : wb ? sheetIndex(wb, parsed.sheetName) : -1;
  if (sheet < 0) return null;
  return {
    sheet,
    r0: parsed.r0,
    c0: parsed.c0,
    r1: parsed.r1,
    c1: parsed.c1,
    explicitSheet: parsed.sheetName !== null,
  };
};

const kindLabels = (strings: Strings): Readonly<Record<SparklineKind, string>> => ({
  line: strings.quickAnalysis.actions.sparkLine,
  column: strings.quickAnalysis.actions.sparkColumn,
  'win-loss': strings.quickAnalysis.actions.sparkWinLoss,
});

/** Attach the Mac Insert > Sparkline chooser. It intentionally delegates the
 *  mutation to `setSparkline`, so protected destination cells and undo history
 *  retain the same behaviour as Quick Analysis and the instance API. */
export function attachMacSparklineDialog(deps: MacSparklineDialogDeps): MacSparklineDialogHandle {
  const { host, store } = deps;
  let strings = deps.strings ?? defaultStrings;

  const shell = createDialogShell({
    host,
    className: 'fc-macsparkdlg',
    ariaLabel: strings.macSparkline.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: strings.macSparkline.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-macsparkdlg__panel'],
    bodyClass: 'fc-fmtdlg__body fc-macsparkdlg__body',
  });

  const sourceInput = document.createElement('input');
  sourceInput.type = 'text';
  sourceInput.className = 'fc-fmtdlg__input fc-macsparkdlg__source';
  sourceInput.autocomplete = 'off';
  sourceInput.spellcheck = false;
  const destinationInput = document.createElement('input');
  destinationInput.type = 'text';
  destinationInput.className = 'fc-fmtdlg__input fc-macsparkdlg__destination';
  destinationInput.autocomplete = 'off';
  destinationInput.spellcheck = false;
  const kindSelect = document.createElement('select');
  kindSelect.className = 'fc-fmtdlg__select fc-macsparkdlg__type';
  const error = document.createElement('div');
  error.className = 'fc-fmtdlg__error fc-macsparkdlg__error';
  error.setAttribute('role', 'alert');
  error.dataset.macSparklineError = 'true';
  error.hidden = true;

  const labelRow = (input: HTMLElement): { row: HTMLLabelElement; text: HTMLSpanElement } => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row fc-macsparkdlg__row';
    const text = document.createElement('span');
    row.append(text, input);
    return { row, text };
  };
  const sourceRow = labelRow(sourceInput);
  const destinationRow = labelRow(destinationInput);
  const kindRow = labelRow(kindSelect);
  body.append(sourceRow.row, destinationRow.row, kindRow.row, error);
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: strings.macSparkline.cancel,
    okLabel: strings.macSparkline.ok,
  });

  let lastSource: ResolvedRange | null = null;
  let lastDestination: Addr | null = null;
  let lastKind: SparklineKind = 'line';

  const currentSheet = (): number => store.getState().data.sheetIndex;
  const setError = (message: string | null): void => {
    error.hidden = message === null;
    error.textContent = message ?? '';
    projectDisabledState(okBtn, message !== null, message, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.macSparkline.ok,
    });
  };

  const validate = (): boolean => {
    const wb = deps.getWb?.() ?? null;
    const source = resolveRange(sourceInput.value, currentSheet(), wb);
    if (!source) {
      lastSource = null;
      lastDestination = null;
      setError(strings.macSparkline.invalidSource);
      return false;
    }
    const destination = resolveRange(destinationInput.value, currentSheet(), wb);
    if (!destination || destination.r0 !== destination.r1 || destination.c0 !== destination.c1) {
      lastSource = null;
      lastDestination = null;
      setError(strings.macSparkline.locationOneCell);
      return false;
    }
    if (
      source.r0 < 0 ||
      source.c0 < 0 ||
      source.r1 > MAX_ROW ||
      source.c1 > MAX_COL ||
      destination.r0 < 0 ||
      destination.c0 < 0 ||
      destination.r0 > MAX_ROW ||
      destination.c0 > MAX_COL
    ) {
      lastSource = null;
      lastDestination = null;
      setError(strings.macSparkline.outOfBounds);
      return false;
    }
    const destinationAddr: Addr = {
      sheet: destination.sheet,
      row: destination.r0,
      col: destination.c0,
    };
    const kind = kindSelect.value as SparklineKind;
    if (kind !== 'line' && kind !== 'column' && kind !== 'win-loss') {
      lastSource = null;
      lastDestination = null;
      setError(strings.macSparkline.chooseType);
      return false;
    }
    if (deps.canCommit && !deps.canCommit(source, destinationAddr)) {
      lastSource = null;
      lastDestination = null;
      setError(strings.macSparkline.unavailable);
      return false;
    }
    lastSource = source;
    lastDestination = destinationAddr;
    lastKind = kind;
    setError(null);
    return true;
  };

  const sourceForSpec = (source: ResolvedRange, destination: Addr): string => {
    const raw = sourceInput.value.trim().replace(/^=/, '');
    if (source.explicitSheet || source.sheet !== destination.sheet) {
      const wb = deps.getWb?.() ?? null;
      const name = wb?.sheetName(source.sheet) ?? `Sheet${source.sheet + 1}`;
      return formatSheetRange(name, source);
    }
    return raw || formatRange(source);
  };

  const onChange = (): void => {
    validate();
  };
  const onCancel = (): void => api.close();
  const onOk = (): void => {
    if (!validate() || !lastSource || !lastDestination) return;
    const ok = setSparkline(
      store,
      lastDestination,
      {
        kind: lastKind,
        source: sourceForSpec(lastSource, lastDestination),
        showNegative: lastKind !== 'line',
      },
      deps.history ?? null,
    );
    if (!ok) {
      setError(strings.macSparkline.protectedLocation);
      return;
    }
    deps.onCommitted?.(lastSource, lastDestination, lastKind);
    api.close();
  };
  const onKey = (event: KeyboardEvent): void => {
    if (event.key === 'Enter' && !event.isComposing) {
      event.preventDefault();
      onOk();
    }
  };

  shell.on(sourceInput, 'input', onChange);
  shell.on(destinationInput, 'input', onChange);
  shell.on(kindSelect, 'change', onChange);
  shell.on(cancelBtn, 'click', onCancel);
  shell.on(okBtn, 'click', onOk);
  shell.on(sourceInput, 'keydown', onKey as EventListener);
  shell.on(destinationInput, 'keydown', onKey as EventListener);

  const refreshLabels = (): void => {
    const labels = kindLabels(strings);
    const t = strings.macSparkline;
    header.textContent = t.title;
    shell.setAriaLabel(t.title);
    sourceRow.text.textContent = t.dataRange;
    destinationRow.text.textContent = t.location;
    kindRow.text.textContent = t.type;
    sourceInput.setAttribute('aria-label', t.dataRange);
    destinationInput.setAttribute('aria-label', t.location);
    kindSelect.setAttribute('aria-label', t.type);
    cancelBtn.textContent = t.cancel;
    okBtn.textContent = t.ok;
    const selectedKind = kindSelect.value;
    const options: readonly SparklineKind[] = ['line', 'column', 'win-loss'];
    kindSelect.replaceChildren();
    appendDialogSelectOptions(
      kindSelect,
      options.map((kind) => ({ value: kind, label: labels[kind] })),
    );
    if (options.includes(selectedKind as SparklineKind)) kindSelect.value = selectedKind;
  };

  const api: MacSparklineDialogHandle = {
    open(): void {
      const range = store.getState().selection.range;
      sourceInput.value = formatRange(range);
      const destinationCol = range.c1 < MAX_COL ? range.c1 + 1 : range.c0;
      destinationInput.value = `${colLabel(destinationCol)}${range.r0 + 1}`;
      kindSelect.value = 'line';
      lastSource = null;
      lastDestination = null;
      refreshLabels();
      validate();
      shell.open();
      requestAnimationFrame(() => sourceInput.focus());
    },
    close(): void {
      shell.close();
      host.focus();
    },
    refresh(): void {
      refreshLabels();
      if (shell.isOpen()) validate();
    },
    setStrings(next: Strings): void {
      strings = next;
      refreshLabels();
      if (shell.isOpen()) validate();
    },
    detach(): void {
      shell.dispose();
    },
  };

  refreshLabels();
  return api;
}
