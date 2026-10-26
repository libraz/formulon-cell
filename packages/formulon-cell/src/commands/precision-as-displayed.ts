import { effectiveCellFormatWithPending } from '../engine/cell-format-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { NumFmt, SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';

/** Round half away from zero at `decimals`. The scaled value is first cut to
 *  15 significant digits so a binary 1.005 rounds up to 1.01 the way the
 *  display does. Values whose scaled form leaves the exact-integer range
 *  already show no hidden digits and are kept. */
const roundDecimal = (value: number, decimals: number): number => {
  const factor = 10 ** decimals;
  const scaled = Math.abs(value) * factor;
  if (!Number.isFinite(scaled) || scaled >= Number.MAX_SAFE_INTEGER) return value;
  const rounded = Math.round(Number(scaled.toPrecision(15))) / factor;
  if (!Number.isFinite(rounded)) return value;
  return rounded === 0 ? 0 : Math.sign(value) * rounded;
};

/** The value a cell keeps under "Set precision as displayed": rounded to the
 *  digits its number format shows. General, text, date/time, special and
 *  custom codes keep full precision. */
export function displayedPrecisionValue(value: number, numFmt: NumFmt | undefined): number {
  if (!Number.isFinite(value) || !numFmt) return value;
  switch (numFmt.kind) {
    case 'fixed':
    case 'currency':
    case 'accounting':
      return roundDecimal(value, numFmt.decimals);
    case 'percent':
      return roundDecimal(value, numFmt.decimals + 2);
    case 'scientific':
      return value === 0 ? 0 : Number(value.toPrecision(Math.min(100, numFmt.decimals + 1)));
    default:
      return value;
  }
}

/** Permanently round every number constant in the workbook to its displayed
 *  precision, as turning on "Set precision as displayed" does. Formulas keep
 *  their full-precision results. Returns how many cells changed; the writes
 *  are one undo step. */
export function applyPrecisionAsDisplayed(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  history: History | null,
): number {
  const state = store.getState();
  const writes: { addr: Parameters<WorkbookHandle['setNumber']>[0]; value: number }[] = [];
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    for (const cell of wb.physicalCells(sheet)) {
      if (cell.formula !== null || cell.value.kind !== 'number') continue;
      const numFmt = effectiveCellFormatWithPending(wb, state, cell.addr)?.numFmt;
      const rounded = displayedPrecisionValue(cell.value.value, numFmt);
      if (Number.isFinite(rounded) && rounded !== cell.value.value) {
        writes.push({ addr: cell.addr, value: rounded });
      }
    }
  }
  if (writes.length === 0) return 0;
  history?.begin();
  try {
    wb.withBatchedRecalc(() => {
      for (const { addr, value } of writes) wb.setNumber(addr, value);
    });
  } finally {
    history?.end();
  }
  return writes.length;
}

const writePrecisionFlag = (wb: WorkbookHandle, store: SpreadsheetStore, on: boolean): void => {
  wb.setCalcSetting('fullPrecision', !on);
  store.setState((s) =>
    s.ui.precisionAsDisplayed === on ? s : { ...s, ui: { ...s.ui, precisionAsDisplayed: on } },
  );
};

/** Turn "Set precision as displayed" on or off. Turning it on rounds the
 *  stored constants, and that sweep and the option undo together; turning it
 *  off is an option change only, as in the desktop app. */
export function setPrecisionAsDisplayed(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  history: History | null,
  on: boolean,
): void {
  if (!on) {
    writePrecisionFlag(wb, store, false);
    return;
  }
  const record = history !== null && !history.isReplaying();
  if (record) history.begin();
  try {
    applyPrecisionAsDisplayed(wb, store, record ? history : null);
    writePrecisionFlag(wb, store, true);
    if (record) {
      history.push({
        undo: () => writePrecisionFlag(wb, store, false),
        redo: () => writePrecisionFlag(wb, store, true),
      });
    }
  } finally {
    if (record) history.end();
  }
}
