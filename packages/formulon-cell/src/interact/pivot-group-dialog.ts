import { PivotDatePeriod } from '@libraz/formulon';
import { parseLocalizedNumericInput } from '../commands/coerce-input.js';
import { parseDateTimeInput } from '../commands/input-grammar.js';
import { formatNumber } from '../format/number-format.js';
import type { Strings } from '../i18n/strings.js';
import {
  appendDialogActions,
  appendDialogFrame,
  createDialogShell,
  type DialogShell,
  deferInitialFocus,
  focusAndSelectInput,
} from './dialog-shell.js';
import { isSubmitEnter } from './mac-dialog-keys.js';
import { pivotFilterInputFor } from './pivot-field-settings.js';
import type { PivotCommitContext } from './pivot-field-settings-dialog-apply.js';
import {
  groupPivotField,
  PIVOT_GROUP_PERIODS,
  type PivotGroupBounds,
  type PivotGroupSettings,
  type PivotGroupTarget,
  pivotGroupSettings,
} from './pivot-group.js';

const PERIOD_KEYS = {
  [PivotDatePeriod.Seconds]: 'seconds',
  [PivotDatePeriod.Minutes]: 'minutes',
  [PivotDatePeriod.Hours]: 'hours',
  [PivotDatePeriod.Days]: 'days',
  [PivotDatePeriod.Months]: 'months',
  [PivotDatePeriod.Quarters]: 'quarters',
  [PivotDatePeriod.Years]: 'years',
} as const;

type GroupDialogStrings = Strings['pivotGroupDialog'];

const textInput = (value: string, label: string): HTMLInputElement => {
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'fc-fmtdlg__input';
  input.autocomplete = 'off';
  input.spellcheck = false;
  input.value = value;
  input.setAttribute('aria-label', label);
  return input;
};

/** A bound row: its checkbox follows the data; typing a bound replaces the automatic one. */
const boundRow = (shell: DialogShell, label: string, auto: boolean, value: string) => {
  const row = document.createElement('div');
  row.className = 'fc-fmtdlg__row fc-pivotgroup__row';
  const check = document.createElement('label');
  const box = document.createElement('input');
  box.type = 'checkbox';
  box.checked = auto;
  const text = document.createElement('span');
  text.textContent = label;
  check.append(box, text);
  const input = textInput(value, label);
  row.append(check, input);
  shell.on(input, 'input', () => {
    box.checked = false;
  });
  return { row, box, input };
};

const labelledRow = (label: string, control: HTMLElement, className: string): HTMLElement => {
  const row = document.createElement('label');
  row.className = `fc-fmtdlg__row ${className}`;
  const text = document.createElement('span');
  text.textContent = label;
  row.append(text, control);
  return row;
};

/** Reads `input` with `read`; on a miss it takes focus and the result is null. */
const readOrFocus = <T>(input: HTMLInputElement, read: (text: string) => T | null): T | null => {
  const value = read(input.value);
  if (value === null) focusAndSelectInput(input);
  return value;
};

/** How a bound is shown and read: a date for a date field, a number for a number field. */
interface BoundCodec {
  show(value: number): string;
  read(text: string): number | null;
}

function boundCodec(ctx: PivotCommitContext, kind: PivotGroupSettings['kind']): BoundCodec {
  const { locale, date1904 } = pivotFilterInputFor(ctx.wb);
  if (kind === 'number') {
    return {
      show: (value) => formatNumber(value, { kind: 'general' }, locale),
      read: (text) => {
        const value = parseLocalizedNumericInput(text, locale);
        return value !== null && Number.isFinite(value) ? value : null;
      },
    };
  }
  return {
    show: (serial) =>
      formatNumber(
        serial,
        Number.isInteger(serial)
          ? { kind: 'date', pattern: 'm/d/yy' }
          : { kind: 'datetime', pattern: 'm/d/yy h:mm' },
        locale,
        date1904,
      ),
    read: (text) => parseDateTimeInput(text, locale, date1904 === true)?.value ?? null,
  };
}

/** The By controls of a mode: the period list and Number of days for a date field, the interval
 *  for a number field. `read` completes the settings from the bounds, null when a control does not
 *  read (it then holds focus). */
interface ByControls {
  rows: HTMLElement[];
  sync(okBtn: HTMLButtonElement): void;
  read(bounds: PivotGroupBounds): PivotGroupSettings | null;
}

function byControls(
  settings: PivotGroupSettings,
  t: GroupDialogStrings,
  codec: BoundCodec,
): ByControls {
  if (settings.kind === 'number') {
    const by = textInput(codec.show(settings.interval), t.numberBy);
    return {
      rows: [labelledRow(t.numberBy, by, 'fc-pivotgroup__row')],
      sync() {},
      read(bounds) {
        const interval = readOrFocus(by, (text) => {
          const value = codec.read(text);
          return value !== null && value > 0 ? value : null;
        });
        return interval === null ? null : { kind: 'number', ...bounds, interval };
      },
    };
  }
  const list = document.createElement('select');
  list.multiple = true;
  list.size = PIVOT_GROUP_PERIODS.length;
  list.className = 'fc-pivotgroup__periods';
  for (const period of PIVOT_GROUP_PERIODS) {
    const option = document.createElement('option');
    option.value = String(period);
    option.textContent = t[PERIOD_KEYS[period]];
    option.selected = (settings.periods & period) !== 0;
    list.append(option);
  }
  const days = textInput(String(settings.intervalDays), t.numberOfDays);
  days.inputMode = 'numeric';
  const periods = (): number =>
    Array.from(list.selectedOptions).reduce((mask, option) => mask | Number(option.value), 0);
  return {
    rows: [
      labelledRow(t.by, list, 'fc-pivotgroup__by'),
      labelledRow(t.numberOfDays, days, 'fc-pivotgroup__row'),
    ],
    sync(okBtn) {
      // Number of days applies to Days alone; OK needs a period.
      const mask = periods();
      days.disabled = mask !== PivotDatePeriod.Days;
      okBtn.disabled = mask === 0;
    },
    read(bounds) {
      const mask = periods();
      if (mask === 0) return null;
      const intervalDays = days.disabled
        ? 1
        : readOrFocus(days, (text) => {
            const n = Number(text.trim());
            return Number.isInteger(n) && n >= 1 ? n : null;
          });
      return intervalDays === null
        ? null
        : { kind: 'date', ...bounds, periods: mask, intervalDays };
    },
  };
}

/** Open the Grouping dialog of a date or number field; OK groups it as one undo step. False when
 *  the target holds no date or number to group. */
export function openPivotGroupDialog(
  ctx: PivotCommitContext & { host: HTMLElement; strings: Strings },
  target: PivotGroupTarget,
): boolean {
  const settings = pivotGroupSettings(ctx.wb, target);
  if (!settings) return false;
  const t = ctx.strings.pivotGroupDialog;
  const shell = createDialogShell({
    host: ctx.host,
    className: 'fc-pivotgroup',
    ariaLabel: t.title,
    onDismiss: () => shell.dispose(),
  });
  const codec = boundCodec(ctx, settings.kind);
  const start = boundRow(shell, t.startingAt, settings.autoStart, codec.show(settings.start));
  const end = boundRow(shell, t.endingAt, settings.autoEnd, codec.show(settings.end));
  const by = byControls(settings, t, codec);
  shell.overlay.classList.add('fc-fmtdlg');
  const frame = appendDialogFrame(shell, {
    title: t.title,
    panelClasses: ['fc-fmtdlg__panel', 'fc-pivotgroup__panel'],
    bodyClass: 'fc-fmtdlg__body fc-pivotgroup__body',
    footerClass: 'fc-fmtdlg__footer fc-pivotgroup__footer',
  });
  const status = document.createElement('div');
  status.className = 'fc-pivotgroup__status';
  status.setAttribute('role', 'status');
  frame.body.append(start.row, end.row, ...by.rows, status);
  const { cancelBtn, okBtn } = appendDialogActions(frame.footer, {
    cancelLabel: ctx.strings.pivotTableDialog.cancel,
    okLabel: ctx.strings.pivotTableDialog.ok,
    buttonBaseClass: 'fc-fmtdlg__btn fc-pivotgroup__btn',
  });
  by.sync(okBtn);

  // A ticked bound follows the data; a typed one must read.
  const bound = (row: typeof start, value: number): number | null =>
    row.box.checked ? value : readOrFocus(row.input, codec.read);
  const run = (): void => {
    const startValue = bound(start, settings.start);
    if (startValue === null) return;
    const endValue = bound(end, settings.end);
    if (endValue === null) return;
    const next = by.read({
      autoStart: start.box.checked,
      autoEnd: end.box.checked,
      start: startValue,
      end: endValue,
    });
    if (!next) return;
    const grouped = groupPivotField(ctx, target, next);
    if (!grouped) {
      status.dataset.state = 'error';
      status.textContent = ctx.strings.commandRefusal.engineFailed;
      return;
    }
    shell.dispose();
  };

  shell.on(okBtn, 'click', run);
  shell.on(cancelBtn, 'click', () => shell.dispose());
  shell.on(frame.body, 'change', () => by.sync(okBtn));
  shell.on(shell.overlay, 'keydown', (event) => {
    const e = event as KeyboardEvent;
    if (!isSubmitEnter(e)) return;
    e.preventDefault();
    run();
  });
  shell.open();
  focusAndSelectInput(start.input);
  deferInitialFocus(shell.overlay, () => focusAndSelectInput(start.input));
  return true;
}
