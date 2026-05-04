import type { RangeResolver } from '../engine/range-resolver.js';
import type { CellValidation, ValidationOp } from '../store/store.js';
import type { CoercedInput } from './coerce-input.js';

export type ValidationOutcome =
  | { ok: true }
  | { ok: false; severity: 'stop' | 'warning' | 'information'; message: string };

type BoundedKind = 'whole' | 'decimal' | 'date' | 'time' | 'textLength';
type BoundedValidation = Extract<CellValidation, { kind: BoundedKind }>;
type ListValidation = Extract<CellValidation, { kind: 'list' }>;

/** Materialize a list-validation's source to a flat string array. Inline
 *  literals pass through; range refs route through `resolveRange`. Returns
 *  `[]` when the resolver isn't supplied for a range-backed list — the
 *  validator then accepts any input (Excel parity: an unresolved DV list
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
 * Excel default). Formula input bypasses validation entirely — Excel does
 * the same; the constraint is on the literal user-typed value, not on
 * downstream calculation results.
 */
export function validateAgainst(
  validation: CellValidation,
  input: CoercedInput,
  resolveRange?: RangeResolver,
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
      // (Excel parity — it just disables the constraint silently).
      if (!Array.isArray(validation.source) && values.length === 0) return { ok: true };
      return values.includes(text) ? { ok: true } : reject(validation, listMessage(values));
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
    case 'custom':
      return { ok: true };
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
