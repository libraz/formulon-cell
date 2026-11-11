import { BlanksAs } from '@libraz/formulon';
import {
  type ResolvedSparklineRange,
  resolveSparklineRange,
} from '../../commands/sparkline-ranges.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { createRangePickFormatter } from '../../interact/range-pick-format.js';
import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import type { SpreadsheetStore } from '../../store/store.js';
import { appendCheckboxRow } from './form-controls.js';
import {
  appendDialogActions,
  appendErrorRow,
  createDialogShell,
  focusAndSelectInput,
  installDialogLifecycle,
  mountDialog,
  showInputError,
} from './shell.js';

export interface SparklineRangeDialogOptions {
  host?: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle | null;
  title: string;
  label: string;
  /** The range text the field opens with. */
  initial: string;
  okLabel: string;
  cancelLabel: string;
  /** Shown when the typed text is not a range; also the picker button's label. */
  invalidMessage: string;
  pickerLabel: string;
}

/** A dialog with one range field: the Edit Single Sparkline's Data and Date Axis dialogs. Returns
 *  the range typed or picked, or null on cancel. */
export const showSparklineRangeDialog = (
  opts: SparklineRangeDialogOptions,
): Promise<ResolvedSparklineRange | null> =>
  new Promise((resolve) => {
    const shell = createDialogShell({ title: opts.title, host: opts.host });
    const row = shell.body.ownerDocument.createElement('label');
    row.className = 'fc-fmtdlg__row fc-fmtdlg__row--block';
    const text = shell.body.ownerDocument.createElement('span');
    text.className = 'fc-tb__dlg__label';
    text.textContent = opts.label;
    const input = shell.body.ownerDocument.createElement('input');
    input.type = 'text';
    input.className = 'fc-tb__dlg__input';
    input.autocomplete = 'off';
    input.spellcheck = false;
    input.value = opts.initial;
    input.setAttribute('aria-label', opts.label);
    row.append(text, input);
    shell.body.appendChild(row);
    const pick = createRangePickFormatter({
      store: opts.store,
      sheetName: (index) => opts.getWb()?.sheetName(index) ?? `Sheet${index + 1}`,
    });
    pick.begin();
    attachRangePickerButton(input, {
      label: opts.pickerLabel,
      getValue: () => pick.range(),
      subscribeToRangeChanges: pick.subscribe,
      kind: 'sparkline-range',
    });
    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });
    const lifecycle = installDialogLifecycle<ResolvedSparklineRange | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void => {
      const range = resolveSparklineRange(
        input.value,
        opts.store.getState().data.sheetIndex,
        opts.getWb(),
      );
      if (!range) {
        showInputError(errorRow, input, opts.invalidMessage);
        return;
      }
      lifecycle.finish(range);
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));
    mountDialog(shell, () => focusAndSelectInput(input));
  });

export interface SparklineHiddenEmptyDialogOptions {
  host?: HTMLElement;
  title: string;
  emptyCellsLabel: string;
  gaps: string;
  zero: string;
  connect: string;
  showHidden: string;
  okLabel: string;
  cancelLabel: string;
  initial: { emptyCellsAs: BlanksAs; displayHidden: boolean };
}

export interface SparklineHiddenEmptyResult {
  emptyCellsAs: BlanksAs;
  displayHidden: boolean;
}

const EMPTY_CELLS_RADIO = 'sparkline-empty-cells';

/** The Hidden and Empty Cell Settings dialog. Returns the chosen settings, or null on cancel. */
export const showSparklineHiddenEmptyDialog = (
  opts: SparklineHiddenEmptyDialogOptions,
): Promise<SparklineHiddenEmptyResult | null> =>
  new Promise((resolve) => {
    const shell = createDialogShell({ title: opts.title, host: opts.host });
    const caption = shell.body.ownerDocument.createElement('div');
    caption.className = 'fc-tb__dlg__label';
    caption.textContent = opts.emptyCellsLabel;
    shell.body.appendChild(caption);
    const options: readonly [BlanksAs, string][] = [
      [BlanksAs.Gap, opts.gaps],
      [BlanksAs.Zero, opts.zero],
      [BlanksAs.Span, opts.connect],
    ];
    for (const [value, label] of options) {
      const wrap = shell.body.ownerDocument.createElement('label');
      wrap.className = 'fc-fmtdlg__radio';
      const radio = shell.body.ownerDocument.createElement('input');
      radio.type = 'radio';
      radio.name = EMPTY_CELLS_RADIO;
      radio.value = String(value);
      radio.checked = value === opts.initial.emptyCellsAs;
      const span = shell.body.ownerDocument.createElement('span');
      span.textContent = label;
      wrap.append(radio, span);
      shell.body.appendChild(wrap);
    }
    const showHidden = appendCheckboxRow(
      shell.body,
      opts.showHidden,
      opts.initial.displayHidden,
      'displayHidden',
    );
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: opts.cancelLabel,
      okLabel: opts.okLabel,
    });
    const lifecycle = installDialogLifecycle<SparklineHiddenEmptyResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const onOk = (): void =>
      lifecycle.finish({
        emptyCellsAs: Number(
          shell.body.querySelector<HTMLInputElement>(`input[name="${EMPTY_CELLS_RADIO}"]:checked`)
            ?.value ?? opts.initial.emptyCellsAs,
        ) as BlanksAs,
        displayHidden: showHidden.checked,
      });
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));
    mountDialog(shell, okBtn);
  });
