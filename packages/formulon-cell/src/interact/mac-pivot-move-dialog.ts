import { formatMacRangeAddress, parseMacRange } from '../commands/mac-data-tools.js';
import { findPivotTableAtCell } from '../engine/passthrough-sync.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendDialogActions, appendDialogFrame, createDialogShell } from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';
import { movePivotTable } from './pivot-move.js';
import { createRangePickFormatter } from './range-pick-format.js';
import { attachRangePickerButton, updateRangePickerLabel } from './range-picker-control.js';

interface PivotMoveDialogHandle {
  open(): boolean;
  detach(): void;
}

const handles = new WeakMap<SpreadsheetInstance, PivotMoveDialogHandle>();

function attachPivotMoveDialog(instance: SpreadsheetInstance): PivotMoveDialogHandle {
  const labels = () => ({
    ...instance.i18n.strings.pivotTableDialog,
    ...instance.i18n.strings.mac.pivotTableDialog,
  });
  let t = labels();
  let target: { sheet: number; pivotIndex: number } | null = null;
  const shell = createDialogShell({
    host: instance.host,
    className: 'fc-mac-pivotmove',
    ariaLabel: t.moveTitle,
    onDismiss: () => close(),
  });
  shell.overlay.classList.add('fc-fmtdlg');
  const { header, body, footer } = appendDialogFrame(shell, {
    title: t.moveTitle,
    panelClasses: ['fc-fmtdlg__panel', 'fc-mac-pivotmove__panel'],
    bodyClass: 'fc-fmtdlg__body fc-mac-pivotmove__body',
    footerClass: 'fc-fmtdlg__footer fc-mac-pivotmove__footer',
  });
  const prompt = document.createElement('div');
  const radio = (value: 'new' | 'existing'): { row: HTMLLabelElement; input: HTMLInputElement } => {
    const row = document.createElement('label');
    row.className = 'fc-fmtdlg__row fc-mac-pivotmove__radio';
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = 'fc-mac-pivotmove-placement';
    input.value = value;
    row.append(input, document.createElement('span'));
    return { row, input };
  };
  const newSheet = radio('new');
  const existing = radio('existing');
  const locationRow = document.createElement('label');
  locationRow.className = 'fc-fmtdlg__row fc-mac-pivotmove__row';
  const locationText = document.createElement('span');
  const location = document.createElement('input');
  location.type = 'text';
  location.id = 'fc-mac-pivotmove-location';
  location.className = 'fc-fmtdlg__input';
  location.autocomplete = 'off';
  location.spellcheck = false;
  locationRow.append(locationText, location);
  const status = document.createElement('div');
  status.className = 'fc-mac-pivotmove__status';
  status.setAttribute('role', 'status');
  body.append(prompt, newSheet.row, existing.row, locationRow, status);
  const pickFormat = createRangePickFormatter({
    store: instance.store,
    sheetName: (index) => instance.workbook.sheetName(index),
  });
  const picker = attachRangePickerButton(location, {
    label: t.rangePickerSelect,
    getValue: () => {
      const r = instance.store.getState().selection.range;
      return pickFormat.range({ range: { ...r, r1: r.r0, c1: r.c0 } });
    },
    subscribeToRangeChanges: pickFormat.subscribe,
    kind: 'pivot-move-location',
  });
  const { cancelBtn, okBtn } = appendDialogActions(footer, {
    cancelLabel: t.cancel,
    okLabel: t.ok,
    buttonBaseClass: 'fc-fmtdlg__btn fc-mac-pivotmove__btn',
  });
  okBtn.dataset.fcMacAction = 'pivot-move-ok';
  cancelBtn.dataset.fcMacAction = 'pivot-move-cancel';
  newSheet.input.dataset.fcMacAction = 'pivot-move-new';
  existing.input.dataset.fcMacAction = 'pivot-move-existing';

  const syncPlacement = (): void => {
    const off = newSheet.input.checked;
    const reason = off ? t.destinationRequiresExistingWorksheet : null;
    projectDisabledState(location, off, reason, { datasetKey: 'disabledReason' });
    projectDisabledState(picker, off, reason, { datasetKey: 'disabledReason' });
  };
  const applyLabels = (): void => {
    t = labels();
    shell.setAriaLabel(t.moveTitle);
    header.textContent = t.moveTitle;
    prompt.textContent = t.destinationSection;
    (newSheet.row.lastElementChild as HTMLElement).textContent = t.newWorksheet;
    (existing.row.lastElementChild as HTMLElement).textContent = t.existingWorksheet;
    locationText.textContent = t.moveLocation;
    okBtn.textContent = t.ok;
    cancelBtn.textContent = t.cancel;
    updateRangePickerLabel(picker, t.rangePickerSelect);
  };
  applyLabels();
  const unsubscribeLocale = instance.i18n.on('change', applyLabels);

  const close = (): void => {
    target = null;
    shell.close();
    instance.host.focus();
  };

  const run = (): void => {
    if (!target) return;
    let destination: Parameters<typeof movePivotTable>[2];
    if (newSheet.input.checked) {
      destination = 'new-sheet';
    } else {
      const parsed = parseMacRange(
        instance.workbook,
        location.value,
        instance.store.getState().data.sheetIndex,
      );
      if (!parsed) {
        status.dataset.state = 'error';
        status.textContent = t.invalidDestination;
        return;
      }
      destination = { sheet: parsed.range.sheet, row: parsed.range.r0, col: parsed.range.c0 };
    }
    const moved = movePivotTable(
      {
        store: instance.store,
        workbook: instance.workbook,
        history: instance.history,
        origin: 'ribbon',
        commandId: 'pivotMove',
      },
      target,
      destination,
    );
    if (!moved) {
      status.dataset.state = 'error';
      status.textContent = t.engineFailed;
      return;
    }
    close();
  };

  shell.on(okBtn, 'click', run);
  shell.on(cancelBtn, 'click', close);
  shell.on(newSheet.input, 'change', syncPlacement);
  shell.on(existing.input, 'change', syncPlacement);
  shell.on(shell.overlay, 'keydown', (event) => {
    const e = event as KeyboardEvent;
    if (!isSubmitEnter(e)) return;
    e.preventDefault();
    run();
  });

  return {
    open(): boolean {
      const pivot = findPivotTableAtCell(
        instance.workbook,
        instance.store.getState().selection.active,
      );
      if (!pivot) return false;
      target = { sheet: pivot.sheetIndex, pivotIndex: pivot.pivotIndex };
      pickFormat.begin();
      applyLabels();
      const bounds = instance.workbook
        .pivotBounds(pivot.sheetIndex)
        .find((b) => b.pivotIndex === pivot.pivotIndex)?.range;
      // The field starts on the report's current top-left, relative form.
      location.value = bounds
        ? formatMacRangeAddress(instance.workbook, { ...bounds, r1: bounds.r0, c1: bounds.c0 })
        : '';
      existing.input.checked = true;
      newSheet.input.checked = false;
      syncPlacement();
      delete status.dataset.state;
      status.textContent = '';
      shell.open();
      queueMicrotask(() => {
        location.focus();
        location.select();
      });
      return true;
    },
    detach(): void {
      unsubscribeLocale();
      shell.dispose();
    },
  };
}

/** Open the PivotTable Analyze → Move PivotTable dialog for the report at the
 *  active cell. False when the active cell is not inside a report. */
export function openMacPivotMoveDialog(instance: SpreadsheetInstance): boolean {
  let handle = handles.get(instance);
  if (!handle) {
    handle = attachPivotMoveDialog(instance);
    handles.set(instance, handle);
  }
  return handle.open();
}

/** Dispose the Move PivotTable overlay owned by an instance, if it was opened. */
export function disposeMacPivotMoveDialog(instance: SpreadsheetInstance): void {
  const handle = handles.get(instance);
  if (!handle) return;
  handles.delete(instance);
  handle.detach();
}
