import type { RangeResolver } from '../engine/range-resolver.js';
import type { CellValue, Range } from '../engine/types.js';
import { syncValidationsToEngine } from '../engine/validation-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellValidation, SpreadsheetStore, ValidationOp } from '../store/store.js';
import type { CoercedInput } from './coerce-input.js';
import { applyFormatSnapshot, captureFormatSnapshot, type History } from './history.js';
import { isCellWritable } from './protection.js';

export type ValidationOutcome =
  | { ok: true }
  | { ok: false; severity: 'stop' | 'warning' | 'information'; message: string };

type BoundedKind = 'whole' | 'decimal' | 'date' | 'time' | 'textLength';
type BoundedValidation = Extract<CellValidation, { kind: BoundedKind }>;
type ListValidation = Extract<CellValidation, { kind: 'list' }>;

const rangeContains = (range: Range, row: number, col: number): boolean =>
  row >= range.r0 && row <= range.r1 && col >= range.c0 && col <= range.c1;

export function clearValidationInRange(store: SpreadsheetStore, range: Range): number {
  let cleared = 0;
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const [key, current] of s.format.formats) {
      if (!current.validation) continue;
      const [sheet, row, col] = key.split(':').map(Number);
      if (sheet !== range.sheet || row === undefined || col === undefined) continue;
      if (!rangeContains(range, row, col)) continue;
      const addr = { sheet, row, col };
      if (!isCellWritable(s, addr)) continue;
      const { validation: _validation, ...next } = current;
      if (Object.keys(next).length === 0) formats.delete(key);
      else formats.set(key, next);
      cleared += 1;
    }
    return { ...s, format: { ...s.format, formats } };
  });
  return cleared;
}

export function clearValidationInRangeWithEngine(
  store: SpreadsheetStore,
  history: History | null,
  wb: WorkbookHandle | null,
  range: Range,
): number {
  let cleared = 0;
  const sync = (): void => {
    if (wb) syncValidationsToEngine(wb, store, range.sheet);
  };
  if (!history || history.isReplaying()) {
    cleared = clearValidationInRange(store, range);
    sync();
    return cleared;
  }
  const before = captureFormatSnapshot(store.getState());
  cleared = clearValidationInRange(store, range);
  if (cleared === 0) return cleared;
  const after = captureFormatSnapshot(store.getState());
  sync();
  history.push({
    undo: () => {
      applyFormatSnapshot(store, before);
      sync();
    },
    redo: () => {
      applyFormatSnapshot(store, after);
      sync();
    },
  });
  return cleared;
}

/** Evaluate an existing cell value against a data-validation rule. This mirrors
 *  the typed-input path closely enough for renderers and ribbon commands:
 *  numbers stay numeric, booleans stay boolean, and text that looks like a
 *  number/bool is interpreted the same way a user-typed value is. */
export function cellValueViolatesValidation(
  value: CellValue,
  validation: CellValidation | undefined,
  resolveRange?: RangeResolver,
): boolean {
  if (!validation) return false;
  if (value.kind === 'error') return false;
  const outcome = validateAgainst(validation, coerceCellValue(value), resolveRange);
  return !outcome.ok;
}

/** Materialize a list-validation's source to a flat string array. Inline
 *  literals pass through; range refs route through `resolveRange`. Returns
 *  `[]` when the resolver isn't supplied for a range-backed list — the
 *  validator then accepts any input (spreadsheet parity: an unresolved DV list
 *  doesn't reject; the dropdown is just empty). */
export function resolveListValues(
  validation: ListValidation,
  resolveRange?: RangeResolver,
): string[] {
  if (Array.isArray(validation.source)) return validation.source;
  if (!resolveRange) return [];
  return resolveRange(validation.source.ref);
}

/**
 * Test a coerced cell input against a `CellValidation` rule. Returns
 * `{ ok: true }` when the input is acceptable; otherwise an outcome carrying
 * the configured severity (stop/warning/information) plus a human-readable
 * default message. Callers decide how to surface the outcome (the keyboard
 * path rejects on `stop`, accepts but logs on `warning`/`information`).
 *
 * Blank input is always accepted when the rule has `allowBlank: true` (the
 * desktop default). Formula input bypasses validation entirely — spreadsheets do
 * the same; the constraint is on the literal user-typed value, not on
 * downstream calculation results.
 */
/**
 * Evaluates a custom-validation formula for the cell under test and returns
 * whether it is satisfied. Returns `null` when evaluation is unavailable (e.g.
 * the stub engine), in which case the validator accepts the input for parity.
 */
export type CustomValidationEvaluator = (formula: string) => boolean | null;

export function validateAgainst(
  validation: CellValidation,
  input: CoercedInput,
  resolveRange?: RangeResolver,
  evalCustom?: CustomValidationEvaluator,
): ValidationOutcome {
  if (input.kind === 'formula') return { ok: true };
  if (input.kind === 'blank') {
    return validation.allowBlank !== false ? { ok: true } : reject(validation, 'blank not allowed');
  }
  switch (validation.kind) {
    case 'list': {
      const text = inputAsText(input);
      const values = resolveListValues(validation, resolveRange);
      // Range-backed list with no resolver / empty resolution: accept anything
      // (spreadsheet parity — it just disables the constraint silently).
      if (!Array.isArray(validation.source) && values.length === 0) return { ok: true };
      return listValueMatches(values, text)
        ? { ok: true }
        : reject(validation, listMessage(values));
    }
    case 'whole':
    case 'decimal': {
      if (input.kind !== 'number') return reject(validation, 'expected number');
      if (validation.kind === 'whole' && !Number.isInteger(input.value)) {
        return reject(validation, 'expected whole number');
      }
      return checkBounded(input.value, validation, 'number');
    }
    case 'date':
    case 'time': {
      if (input.kind !== 'number') return reject(validation, `expected ${validation.kind}`);
      return checkBounded(input.value, validation, validation.kind);
    }
    case 'textLength': {
      const len = inputAsText(input).length;
      return checkBounded(len, validation, 'length');
    }
    case 'custom': {
      // Excel evaluates the custom formula and rejects when it is not TRUE.
      // Without an evaluator (renderers) or when the engine can't evaluate
      // (stub → null), accept for parity rather than block silently.
      if (!validation.formula || !evalCustom) return { ok: true };
      const satisfied = evalCustom(validation.formula);
      if (satisfied === null) return { ok: true };
      return satisfied
        ? { ok: true }
        : reject(validation, 'value does not satisfy the custom rule');
    }
    default:
      return { ok: true };
  }
}

function inputAsText(input: Exclude<CoercedInput, { kind: 'formula' }>): string {
  switch (input.kind) {
    case 'blank':
      return '';
    case 'number':
      return String(input.value);
    case 'bool':
      return input.value ? 'TRUE' : 'FALSE';
    case 'text':
      return input.value;
  }
}

function normalizeListText(value: string): string {
  return value.toLocaleLowerCase();
}

function listValueMatches(values: readonly string[], text: string): boolean {
  const normalized = normalizeListText(text);
  return values.some((value) => normalizeListText(value) === normalized);
}

const NUMERIC = /^[+-]?(?:(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;

function coerceCellValue(value: Exclude<CellValue, { kind: 'error' }>): CoercedInput {
  switch (value.kind) {
    case 'blank':
      return { kind: 'blank' };
    case 'number':
      return { kind: 'number', value: value.value };
    case 'bool':
      return { kind: 'bool', value: value.value };
    case 'text': {
      const text = value.value.trim();
      if (text === '') return { kind: 'blank' };
      if (/^(true|false)$/i.test(text)) return { kind: 'bool', value: /^true$/i.test(text) };
      if (NUMERIC.test(text)) return { kind: 'number', value: Number(text.replaceAll(',', '')) };
      return { kind: 'text', value: value.value };
    }
  }
}

function checkBounded(
  value: number,
  validation: BoundedValidation,
  noun: string,
): ValidationOutcome {
  const a = validation.a;
  const b = validation.b ?? a;
  let ok = true;
  switch (validation.op) {
    case 'between':
      ok = value >= Math.min(a, b) && value <= Math.max(a, b);
      break;
    case 'notBetween':
      ok = value < Math.min(a, b) || value > Math.max(a, b);
      break;
    case '=':
      ok = value === a;
      break;
    case '<>':
      ok = value !== a;
      break;
    case '<':
      ok = value < a;
      break;
    case '<=':
      ok = value <= a;
      break;
    case '>':
      ok = value > a;
      break;
    case '>=':
      ok = value >= a;
      break;
  }
  return ok ? { ok: true } : reject(validation, boundedMessage(validation.op, a, b, noun));
}

function reject(meta: CellValidation, defaultMessage: string): ValidationOutcome {
  return {
    ok: false,
    severity: meta.errorStyle ?? 'stop',
    message: meta.errorMessage || defaultMessage,
  };
}

function listMessage(values: string[]): string {
  const sample = values.slice(0, 3).join(', ');
  return values.length <= 3
    ? `value must be one of: ${sample}`
    : `value must be one of: ${sample}, …`;
}

function boundedMessage(op: ValidationOp, a: number, b: number, noun: string): string {
  switch (op) {
    case 'between':
      return `${noun} must be between ${a} and ${b}`;
    case 'notBetween':
      return `${noun} must be outside ${a}..${b}`;
    case '=':
      return `${noun} must equal ${a}`;
    case '<>':
      return `${noun} must not equal ${a}`;
    case '<':
      return `${noun} must be less than ${a}`;
    case '<=':
      return `${noun} must be ≤ ${a}`;
    case '>':
      return `${noun} must be greater than ${a}`;
    case '>=':
      return `${noun} must be ≥ ${a}`;
  }
}
