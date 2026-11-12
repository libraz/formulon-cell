import { BlanksAs, SparklineAxisType } from '@libraz/formulon';
import {
  type ResolvedSparklineRange,
  resolveSparklineRange,
} from '../../commands/sparkline-ranges.js';
import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { appendDialogTabPair } from '../../interact/dialog-shell.js';
import { createRangePickFormatter } from '../../interact/range-pick-format.js';
import { attachRangePickerButton } from '../../interact/range-picker-control.js';
import type { SpreadsheetStore } from '../../store/store.js';
import { projectDisabledState } from '../menu-a11y.js';
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

/** The text of the Axis dialog. */
export interface SparklineAxisDialogText {
  title: string;
  horizontal: string;
  vertical: string;
  general: string;
  date: string;
  dateRange: string;
  showAxis: string;
  plotRtl: string;
  rtlUnsupported: string;
  minimum: string;
  maximum: string;
  autoEach: string;
  sameAll: string;
  customValue: string;
  invalidRange: string;
  pickerLabel: string;
  ok: string;
  cancel: string;
}

/** One end of the vertical axis: how it is set, and the custom value when it has one. */
export interface SparklineAxisBound {
  type: SparklineAxisType;
  value: number | null;
}

/** The axis settings the dialog opens with. */
export interface SparklineAxisDialogInitial {
  dateAxis: boolean;
  /** The date range text, empty when the group has none. */
  dateRange: string;
  displayXAxis: boolean;
  min: SparklineAxisBound;
  max: SparklineAxisBound;
}

export interface SparklineAxisDialogResult {
  dateAxis: boolean;
  /** The date range, set whenever `dateAxis` is. */
  dateRange: ResolvedSparklineRange | null;
  displayXAxis: boolean;
  min: SparklineAxisBound;
  max: SparklineAxisBound;
}

export interface SparklineAxisDialogOptions {
  host?: HTMLElement;
  store: SpreadsheetStore;
  getWb: () => WorkbookHandle | null;
  text: SparklineAxisDialogText;
  initial: SparklineAxisDialogInitial;
}

const AXIS_BOUND_TYPES = [
  SparklineAxisType.Individual,
  SparklineAxisType.Group,
  SparklineAxisType.Custom,
] as const;

/** The Axis dialog of the Mac app: the horizontal axis on one tab and the two vertical bounds on
 *  the other, applied together on OK. Plot Data Right-to-Left stays greyed. */
export const showSparklineAxisDialog = (
  opts: SparklineAxisDialogOptions,
): Promise<SparklineAxisDialogResult | null> =>
  new Promise((resolve) => {
    const t = opts.text;
    const shell = createDialogShell({ title: t.title, host: opts.host });
    shell.overlay.classList.add('fc-macsparkaxis');
    const doc = shell.body.ownerDocument;
    const el = <K extends keyof HTMLElementTagNameMap>(tag: K, className?: string) => {
      const node = doc.createElement(tag);
      if (className) node.className = className;
      return node;
    };
    const radio = (name: string, label: string, checked: boolean) => {
      const wrap = el('label', 'fc-fmtdlg__radio');
      const input = el('input');
      input.type = 'radio';
      input.name = name;
      input.checked = checked;
      const span = el('span');
      span.textContent = label;
      wrap.append(input, span);
      return { wrap, input };
    };
    const textField = (label: string, value: string) => {
      const input = el('input', 'fc-tb__dlg__input');
      input.type = 'text';
      input.autocomplete = 'off';
      input.spellcheck = false;
      input.value = value;
      input.setAttribute('aria-label', label);
      return input;
    };

    const tabs = el('div', 'fc-macsparkaxis__tabs');
    tabs.setAttribute('role', 'tablist');
    const panels = el('div', 'fc-macsparkaxis__panels');
    shell.body.append(tabs, panels);
    const tab = (id: 'horizontal' | 'vertical') => ({
      id,
      ...appendDialogTabPair(tabs, panels, {
        id,
        label: t[id],
        tabId: `fc-macsparkaxis-tab-${id}`,
        panelId: `fc-macsparkaxis-panel-${id}`,
        tabClass: 'fc-fmtdlg__tab',
        panelClass: 'fc-macsparkaxis__panel',
      }),
    });
    const pairs = [tab('horizontal'), tab('vertical')] as const;
    const activate = (id: 'horizontal' | 'vertical'): void => {
      for (const pair of pairs) {
        const on = pair.id === id;
        pair.button.setAttribute('aria-selected', on ? 'true' : 'false');
        pair.button.tabIndex = on ? 0 : -1;
        pair.panel.hidden = !on;
      }
    };
    for (const pair of pairs) pair.button.addEventListener('click', () => activate(pair.id));
    activate('horizontal');
    const [{ panel: horizontal }, { panel: vertical }] = pairs;

    // Horizontal: the axis type, the date range under Date Axis, then the two display options.
    const general = radio('fc-macsparkaxis-type', t.general, !opts.initial.dateAxis);
    const date = radio('fc-macsparkaxis-type', t.date, opts.initial.dateAxis);
    const dateCaption = el('div', 'fc-tb__dlg__label fc-macsparkaxis__indent');
    dateCaption.textContent = t.dateRange;
    const dateInput = textField(t.dateRange, opts.initial.dateRange);
    const dateRow = el('div', 'fc-macsparkaxis__indent');
    dateRow.appendChild(dateInput);
    const pick = createRangePickFormatter({
      store: opts.store,
      sheetName: (index) => opts.getWb()?.sheetName(index) ?? `Sheet${index + 1}`,
    });
    pick.begin();
    const picker = attachRangePickerButton(dateInput, {
      label: t.pickerLabel,
      getValue: () => pick.range(),
      subscribeToRangeChanges: pick.subscribe,
      kind: 'sparkline-range',
    });
    const rule = el('hr', 'fc-macsparkaxis__rule');
    horizontal.append(general.wrap, date.wrap, dateCaption, dateRow, rule);
    const showAxis = appendCheckboxRow(horizontal, t.showAxis, opts.initial.displayXAxis);
    const rtl = appendCheckboxRow(horizontal, t.plotRtl, false);

    // Vertical: each bound is a caption over a box of three choices, the last with its value.
    const bound = (name: string, caption: string, initial: SparklineAxisBound) => {
      const head = el('div', 'fc-tb__dlg__label');
      head.textContent = caption;
      const box = el('div', 'fc-macsparkaxis__box');
      box.setAttribute('role', 'radiogroup');
      box.setAttribute('aria-label', caption);
      const choices = AXIS_BOUND_TYPES.map((type, index) => {
        const label = [t.autoEach, t.sameAll, t.customValue][index] as string;
        return { type, ...radio(name, label, initial.type === type) };
      });
      const value = textField(`${caption} ${t.customValue}`, initial.value?.toString() ?? '');
      value.classList.add('fc-macsparkaxis__value');
      choices[2]?.wrap.appendChild(value);
      box.append(...choices.map((choice) => choice.wrap));
      vertical.append(head, box);
      return { choices, value };
    };
    const min = bound('fc-macsparkaxis-min', t.minimum, opts.initial.min);
    const max = bound('fc-macsparkaxis-max', t.maximum, opts.initial.max);

    const errorRow = appendErrorRow(shell.body);
    const { cancelBtn, okBtn } = appendDialogActions(shell.footer, {
      cancelLabel: t.cancel,
      okLabel: t.ok,
    });
    cancelBtn.dataset.dialogAction = 'cancel';

    const sync = (): void => {
      const grey = (el: HTMLInputElement | HTMLButtonElement, disabled: boolean): void =>
        projectDisabledState(el, disabled, null);
      grey(dateInput, !date.input.checked);
      for (const end of [min, max]) grey(end.value, !end.choices[2]?.input.checked);
      grey(picker, dateInput.disabled);
    };
    projectDisabledState(rtl, true, t.rtlUnsupported, { titlePrefix: t.plotRtl });
    shell.body.addEventListener('change', sync);
    sync();

    const lifecycle = installDialogLifecycle<SparklineAxisDialogResult | null>({
      shell,
      resolve,
      onCancel: () => null,
      onSubmit: () => onOk(),
    });
    const boundOf = (end: typeof min): SparklineAxisBound | null => {
      const type =
        end.choices.find((choice) => choice.input.checked)?.type ?? SparklineAxisType.Individual;
      if (type !== SparklineAxisType.Custom) return { type, value: null };
      const value = Number(end.value.value.trim());
      if (end.value.value.trim() === '' || !Number.isFinite(value)) {
        activate('vertical');
        showInputError(errorRow, end.value, t.customValue);
        return null;
      }
      return { type, value };
    };
    const onOk = (): void => {
      let dateRange: ResolvedSparklineRange | null = null;
      if (date.input.checked) {
        dateRange = resolveSparklineRange(
          dateInput.value,
          opts.store.getState().data.sheetIndex,
          opts.getWb(),
        );
        if (!dateRange) {
          activate('horizontal');
          showInputError(errorRow, dateInput, t.invalidRange);
          return;
        }
      }
      const minBound = boundOf(min);
      if (!minBound) return;
      const maxBound = boundOf(max);
      if (!maxBound) return;
      lifecycle.finish({
        dateAxis: date.input.checked,
        dateRange,
        displayXAxis: showAxis.checked,
        min: minBound,
        max: maxBound,
      });
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', () => lifecycle.finish(null));
    mountDialog(shell, () => general.input.focus());
  });
