import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';

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

/** Convenience: coerce + write in one call for the common keyboard path. */
export function writeInput(wb: WorkbookHandle, a: Addr, raw: string): void {
  writeCoerced(wb, a, coerceInput(raw));
}
