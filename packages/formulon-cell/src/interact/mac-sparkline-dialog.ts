import type { History } from '../commands/history.js';
import {
  defaultSparklineGroupSettings,
  type SparklineInsert,
  setSparklinesAction,
  sparklineTypeOf,
} from '../commands/sparkline.js';
import { replaceSparklineGroupAction } from '../commands/sparkline-design.js';
import { splitDataForLocation } from '../commands/sparkline-layout.js';
import {
  type ResolvedSparklineRange,
  resolveSparklineRange,
  sparklineSourceText,
} from '../commands/sparkline-ranges.js';
import { formatA1Range, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { defaultStrings, type Strings } from '../i18n/strings.js';
import type { SparklineKind, SpreadsheetStore } from '../store/store.js';
import { appendDialogSelectOptions } from '../toolbar/dialogs/form-controls.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  deferInitialFocus,
} from './dialog-shell.js';
import { createRangePickFormatter } from './range-pick-format.js';
import { attachRangePickerButton, updateRangePickerLabel } from './range-picker-control.js';

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
  /** Called after the dialog closes, whether it was confirmed or dismissed. */
  onClosed?: () => void;
}

/** An existing group the dialog edits: the data and location range texts it opens with. The
 *  group keeps its settings and gets one sparkline per location cell. */
export interface MacSparklineEditTarget {
  groupId: string;
  data: string;
  location: string;
}

export interface MacSparklineDialogHandle {
  /** Open the Insert Sparkline chooser, or Edit Sparklines for `edit`. */
  open(edit?: MacSparklineEditTarget): void;
  close(): void;
  refresh(): void;
  setStrings(next: Strings): void;
  detach(): void;
}

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
    ariaLabel: strings.mac.sparkline.title,
    onDismiss: () => api.close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: strings.mac.sparkline.title,
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
  const pickFormat = createRangePickFormatter({
    store,
    sheetName: (index) => deps.getWb?.()?.sheetName(index) ?? `Sheet${index + 1}`,
  });
  const subscribeToRangeChanges = pickFormat.subscribe;
  const pickers = [
    attachRangePickerButton(sourceInput, {
      label: strings.pivotTableDialog.rangePickerSelect,
      getValue: () => pickFormat.range(),
      subscribeToRangeChanges,
      kind: 'sparkline-source',
    }),
    attachRangePickerButton(destinationInput, {
      label: strings.pivotTableDialog.rangePickerSelect,
      getValue: () => pickFormat.cell(),
      subscribeToRangeChanges,
      kind: 'sparkline-location',
    }),
  ];
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: strings.mac.sparkline.cancel,
    okLabel: strings.mac.sparkline.ok,
  });

  let lastSource: ResolvedSparklineRange | null = null;
  let lastDestination: Addr | null = null;
  let lastKind: SparklineKind = 'line';
  let lastEntries: SparklineInsert[] = [];
  let editing: MacSparklineEditTarget | null = null;

  const currentSheet = (): number => store.getState().data.sheetIndex;
  const setError = (message: string | null): void => {
    error.hidden = message === null;
    error.textContent = message ?? '';
    projectDisabledState(okBtn, message !== null, message, {
      datasetKey: 'disabledReason',
      titlePrefix: strings.mac.sparkline.ok,
    });
  };

  const validate = (): boolean => {
    const wb = deps.getWb?.() ?? null;
    const source = resolveSparklineRange(sourceInput.value, currentSheet(), wb);
    if (!source) {
      lastSource = null;
      lastDestination = null;
      setError(strings.mac.sparkline.invalidSource);
      return false;
    }
    const destination = resolveSparklineRange(destinationInput.value, currentSheet(), wb);
    const oneCell =
      destination !== null &&
      destination.r0 === destination.r1 &&
      destination.c0 === destination.c1;
    if (!destination || (!editing && !oneCell)) {
      lastSource = null;
      lastDestination = null;
      setError(strings.mac.sparkline.locationOneCell);
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
      setError(strings.mac.sparkline.outOfBounds);
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
      setError(strings.mac.sparkline.chooseType);
      return false;
    }
    if (deps.canCommit && !deps.canCommit(source, destinationAddr)) {
      lastSource = null;
      lastDestination = null;
      setError(strings.mac.sparkline.unavailable);
      return false;
    }
    if (editing) {
      const parts = splitDataForLocation(source, destination);
      if (!parts) {
        lastSource = null;
        lastDestination = null;
        setError(strings.mac.sparkline.invalidSource);
        return false;
      }
      const width = destination.c1 - destination.c0 + 1;
      lastEntries = parts.map((part, i) => ({
        addr: {
          sheet: destination.sheet,
          row: destination.r0 + Math.floor(i / width),
          col: destination.c0 + (i % width),
        },
        source: sparklineSourceText(
          { ...part, explicitSheet: source.explicitSheet },
          destination.sheet,
          wb,
        ),
      }));
    }
    lastSource = source;
    lastDestination = destinationAddr;
    lastKind = kind;
    setError(null);
    return true;
  };

  const sourceForSpec = (source: ResolvedSparklineRange, destination: Addr): string =>
    sparklineSourceText(source, destination.sheet, deps.getWb?.() ?? null);

  const onChange = (): void => {
    validate();
  };
  const onCancel = (): void => api.close();
  const onOk = (): void => {
    if (!validate() || !lastSource || !lastDestination) return;
    if (editing) {
      const done = replaceSparklineGroupAction(
        {
          store,
          history: deps.history ?? null,
          origin: 'ribbon',
          commandId: 'sparklineEditData',
        },
        editing.groupId,
        lastEntries,
      );
      if (!done) {
        setError(strings.mac.sparkline.protectedLocation);
        return;
      }
      api.close();
      return;
    }
    const ok = setSparklinesAction(
      {
        store,
        history: deps.history ?? null,
        origin: 'ribbon',
        commandId: 'mac.insert.sparkline',
      },
      [{ addr: lastDestination, source: sourceForSpec(lastSource, lastDestination) }],
      defaultSparklineGroupSettings(sparklineTypeOf(lastKind)),
    );
    if (!ok) {
      setError(strings.mac.sparkline.protectedLocation);
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
    const t = strings.mac.sparkline;
    const title = editing ? strings.ribbon.sparklineDialog.editGroupTitle : t.title;
    header.textContent = title;
    shell.setAriaLabel(title);
    kindRow.row.hidden = editing !== null;
    sourceRow.text.textContent = t.dataRange;
    destinationRow.text.textContent = t.location;
    kindRow.text.textContent = t.type;
    sourceInput.setAttribute('aria-label', t.dataRange);
    destinationInput.setAttribute('aria-label', t.location);
    kindSelect.setAttribute('aria-label', t.type);
    cancelBtn.textContent = t.cancel;
    okBtn.textContent = t.ok;
    for (const picker of pickers) {
      updateRangePickerLabel(picker, strings.pivotTableDialog.rangePickerSelect);
    }
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
    open(edit?: MacSparklineEditTarget): void {
      const range = store.getState().selection.range;
      editing = edit ?? null;
      pickFormat.begin();
      sourceInput.value = edit ? edit.data : formatA1Range(range);
      // The desktop dialog opens with the location empty; OK stays disabled
      // without flagging an error the user has not made yet.
      destinationInput.value = edit ? edit.location : '';
      kindSelect.value = 'line';
      lastSource = null;
      lastDestination = null;
      refreshLabels();
      validate();
      if (!edit) error.hidden = true;
      shell.open();
      deferInitialFocus(shell.overlay, () => sourceInput.focus());
    },
    close(): void {
      shell.close();
      host.focus();
      deps.onClosed?.();
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
