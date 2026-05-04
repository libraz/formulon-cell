import { makeRangeResolver } from '../engine/range-resolver.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellValidation, SpreadsheetStore } from '../store/store.js';
import { isCellWritable, warnProtected } from './protection.js';
import { type ValidationOutcome, validateAgainst } from './validate.js';

export type CoercedInput =
  | { kind: 'blank' }
  | { kind: 'formula'; text: string }
  | { kind: 'number'; value: number }
  | { kind: 'bool'; value: boolean }
  | { kind: 'text'; value: string };

const NUMERIC = /^-?\d+(\.\d+)?(e[+-]?\d+)?$/i;

/**
 * Map a user-typed string to the right primitive setter. Pure — depends on
 * neither the engine nor the store. Shared between the keyboard path,
 * formula bar, and clipboard paste so the rules stay in one place.
 */
export function coerceInput(raw: string): CoercedInput {
  const trimmed = raw.trim();
  if (trimmed === '') return { kind: 'blank' };
  if (trimmed.startsWith('=')) return { kind: 'formula', text: trimmed };
  if (trimmed === 'TRUE' || trimmed === 'FALSE') {
    return { kind: 'bool', value: trimmed === 'TRUE' };
  }
  if (NUMERIC.test(trimmed)) {
    const n = Number(trimmed);
    if (!Number.isNaN(n)) return { kind: 'number', value: n };
  }
  return { kind: 'text', value: raw };
}

/**
 * Dispatch a coerced input to the workbook adapter. Callers can either use
 * this helper or call `coerceInput` and switch themselves (e.g. when paste
 * needs to skip writes for unchanged cells).
 */
export function writeCoerced(wb: WorkbookHandle, a: Addr, c: CoercedInput): void {
  switch (c.kind) {
    case 'blank':
      wb.setBlank(a);
      return;
    case 'formula':
      wb.setFormula(a, c.text);
      return;
    case 'number':
      wb.setNumber(a, c.value);
      return;
    case 'bool':
      wb.setBool(a, c.value);
      return;
    case 'text':
      wb.setText(a, c.value);
  }
}

/** Convenience: coerce + write in one call for the common keyboard path.
 *  When `store` is supplied, sheet-protection is checked first — locked cells
 *  on protected sheets emit a console warning and the write is skipped. */
export function writeInput(
  wb: WorkbookHandle,
  a: Addr,
  raw: string,
  store?: SpreadsheetStore,
): void {
  if (store && !isCellWritable(store.getState(), a)) {
    warnProtected(a);
    return;
  }
  writeCoerced(wb, a, coerceInput(raw));
}

/** Coerce + validate + write. When validation rejects with severity `stop`,
 *  the write is skipped and the outcome is returned so the caller can surface
 *  the error. `warning` and `information` outcomes still write through but
 *  the message is returned for an inline toast. Range-backed list sources
 *  resolve against `wb` rooted at `a.sheet`. When `store` is supplied,
 *  sheet-protection is gated first; gated cells return `{ ok: true }` and
 *  emit a console warning rather than writing through. */
export function writeInputValidated(
  wb: WorkbookHandle,
  a: Addr,
  raw: string,
  validation: CellValidation | undefined,
  store?: SpreadsheetStore,
): ValidationOutcome {
  if (store && !isCellWritable(store.getState(), a)) {
    warnProtected(a);
    return { ok: true };
  }
  const coerced = coerceInput(raw);
  if (!validation) {
    writeCoerced(wb, a, coerced);
    return { ok: true };
  }
  const outcome = validateAgainst(validation, coerced, makeRangeResolver(wb, a.sheet));
  if (outcome.ok || outcome.severity !== 'stop') {
    writeCoerced(wb, a, coerced);
  }
  return outcome;
}
