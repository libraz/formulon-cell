import { addrKey, parseAddrKey } from '../engine/address.js';
import { makeRangeResolver, type RangeResolver } from '../engine/range-resolver.js';
import type { Addr, CellValue, Range } from '../engine/types.js';
import { syncValidationsToEngine } from '../engine/validation-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { subtractFromRangeValidations } from '../store/range-validations.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import {
  type CellValidation,
  mutators,
  type SpreadsheetStore,
  type ValidationOp,
} from '../store/store.js';
import type { CoercedInput } from './coerce-input.js';
import type { History } from './history.js';
import { isCellWritable } from './protection.js';
import { applyFormatSnapshot, captureFormatSnapshot } from './slice-history.js';

export type ValidationOutcome =
  | { ok: true }
  | { ok: false; severity: 'stop' | 'warning' | 'information'; message: string };

type BoundedKind = 'whole' | 'decimal' | 'date' | 'time' | 'textLength';
type BoundedValidation = Extract<CellValidation, { kind: BoundedKind }>;
type ListValidation = Extract<CellValidation, { kind: 'list' }>;

/** Whether formulon's 0.13 proposed-value validator can evaluate this rule
 * without falling back to the cell layer. Custom rules stay local because the
 * upstream evaluator resolves their formulas against the saved workbook value.
 * A list source is safe once the synchronized engine rule is checked below. */
export function nativeValidationSupports(validation: CellValidation, input: CoercedInput): boolean {
  if (input.kind === 'formula') return false;
  if (validation.kind === 'custom') return false;
  if (validation.kind === 'list') return true;
  return (
    validation.kind === 'whole' ||
    validation.kind === 'decimal' ||
    validation.kind === 'date' ||
    validation.kind === 'time' ||
    validation.kind === 'textLength'
  );
}

/** Evaluate a proposed value with formulon's typed validation API. `null`
 * means the engine cannot authoritatively answer this rule, so callers should
 * run the existing JavaScript evaluator. */
export function nativeValidationOutcome(
  wb: WorkbookHandle,
  addr: { sheet: number; row: number; col: number },
  input: CoercedInput,
  validation: CellValidation,
): ValidationOutcome | null {
  if (!nativeValidationSupports(validation, input)) return null;
  if (typeof wb.validateValue !== 'function') return null;
  const value = coercedInputAsCellValue(input);
  const result = wb.validateValue(addr.sheet, addr.row, addr.col, value);
  if (!result?.hasRule) return null;
  if (!validationMatchesEngineRule(wb, addr, validation, result.ruleIndex)) return null;
  if (result.valid) return { ok: true };
  const local = validateAgainst(validation, input);
  const severity =
    result.errorStyle === 1 ? 'warning' : result.errorStyle === 2 ? 'information' : 'stop';
  return {
    ok: false,
    severity,
    message: local.ok ? validation.errorMessage || 'value is invalid' : local.message,
  };
}

function validationMatchesEngineRule(
  wb: WorkbookHandle,
  addr: { sheet: number; row: number; col: number },
  validation: CellValidation,
  ruleIndex: number,
): boolean {
  if (typeof wb.getValidationsForSheet !== 'function') return false;
  const entries = wb.getValidationsForSheet(addr.sheet);
  const entry = entries[ruleIndex];
  if (!entry?.ranges.some((range) => rangeContainsAddr(range, addr))) return false;
  const type =
    validation.kind === 'whole'
      ? 1
      : validation.kind === 'decimal'
        ? 2
        : validation.kind === 'list'
          ? 3
          : validation.kind === 'date'
            ? 4
            : validation.kind === 'time'
              ? 5
              : validation.kind === 'textLength'
                ? 6
                : 7;
  if (entry.type !== type) return false;
  const expectedErrorStyle =
    validation.errorStyle === 'warning' ? 1 : validation.errorStyle === 'information' ? 2 : 0;
  if (entry.errorStyle !== expectedErrorStyle) return false;
  if (entry.allowBlank !== (validation.allowBlank !== false)) return false;
  if (entry.showErrorMessage !== (validation.showErrorMessage !== false)) return false;
  if (validation.kind === 'custom') {
    return normalizeFormula(entry.formula1) === normalizeFormula(validation.formula);
  }
  if (validation.kind === 'list') {
    if (Array.isArray(validation.source)) {
      return normalizeListFormula(entry.formula1) === `"${validation.source.join(',')}"`;
    }
    return normalizeListFormula(entry.formula1) === normalizeListFormula(validation.source.ref);
  }
  const op =
    validation.op === 'between'
      ? 0
      : validation.op === 'notBetween'
        ? 1
        : validation.op === '='
          ? 2
          : validation.op === '<>'
            ? 3
            : validation.op === '>'
              ? 4
              : validation.op === '<'
                ? 5
                : validation.op === '>='
                  ? 6
                  : 7;
  if (entry.op !== op || !boundMatches(entry.formula1, validation.aFormula, validation.a)) {
    return false;
  }
  const needsFormula2 = validation.op === 'between' || validation.op === 'notBetween';
  return (
    !needsFormula2 ||
    boundMatches(entry.formula2, validation.bFormula, validation.b ?? validation.a)
  );
}

function boundMatches(
  engineFormula: string,
  formula: string | undefined,
  constant: number,
): boolean {
  return formula === undefined
    ? Number(engineFormula.replace(/^=/, '')) === constant
    : normalizeFormula(engineFormula) === normalizeFormula(formula);
}

function normalizeFormula(formula: string): string {
  return formula.trim().replace(/^=/, '').replace(/\s+/g, '').toLowerCase();
}

function normalizeListFormula(formula: string): string {
  return formula.trim().replace(/^=/, '').trim();
}

function coercedInputAsCellValue(input: CoercedInput): CellValue {
  switch (input.kind) {
    case 'number':
      return { kind: 'number', value: input.value };
    case 'bool':
      return { kind: 'bool', value: input.value };
    case 'error':
      return { kind: 'error', code: input.code, text: input.text };
    case 'text':
      return { kind: 'text', value: input.value };
    case 'blank':
      return { kind: 'blank' };
    case 'formula':
      return { kind: 'blank' };
  }
}

/** Enumerate native invalid cells page by page. `null` signals an older
 * engine, allowing the existing cell-format scan to remain the fallback. */
export function listInvalidValidationCells(wb: WorkbookHandle, sheet: number): Addr[] | null {
  if (!wb.capabilities?.validationEvaluation || typeof wb.listInvalidCells !== 'function')
    return null;
  const out: Addr[] = [];
  let cursor: number | null = null;
  const seenCursors = new Set<number>();
  do {
    if (cursor !== null) {
      if (seenCursors.has(cursor)) return null;
      seenCursors.add(cursor);
    }
    const page = wb.listInvalidCells(sheet, cursor, 512);
    if (!page) return null;
    out.push(...page.cells);
    if (page.nextCursor !== null && page.cells.length === 0) return null;
    if (page.nextCursor !== null && page.nextCursor === cursor) return null;
    cursor = page.nextCursor;
  } while (cursor !== null);
  return out;
}

/** Replace the validation-circle set with formulon's paged invalid-cell list,
 * optionally restricted to a selection. Returns `null` when the native list
 * surface is unavailable. */
export function circleInvalidValidationDataWithEngine(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  range?: Range,
): number | null {
  const cells = listInvalidValidationCells(wb, sheet);
  if (!cells) return null;
  const filtered = range ? cells.filter((addr) => rangeContainsAddr(range, addr)) : cells;
  const keys = new Set(store.getState().errorIndicators.validationCircles);
  const found = new Set<string>();
  for (const addr of filtered) {
    const key = addrKey(addr);
    found.add(key);
    keys.add(key);
  }

  // formulon's pager is intentionally limited to populated cells. Keep the
  // cell-layer scan for validation-formatted blanks, especially rules with
  // `allowBlank: false`, so Circle Invalid Data retains its previous coverage.
  const state = store.getState();
  const resolveRange = makeRangeResolver(wb, sheet);
  for (const [key, format] of state.format.formats) {
    if (!format.validation) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet || (range && !rangeContainsAddr(range, addr))) continue;
    const value = state.data.cells.get(key)?.value ?? { kind: 'blank' as const };
    if (!cellValueViolatesValidation(value, format.validation, resolveRange)) continue;
    found.add(key);
    keys.add(key);
  }
  mutators.setValidationCircles(store, keys);
  return found.size;
}

export function clearValidationInRange(store: SpreadsheetStore, range: Range): number {
  let cleared = 0;
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const [key, current] of s.format.formats) {
      if (!current.validation) continue;
      const addr = parseAddrKey(key);
      if (!addr || addr.sheet !== range.sheet) continue;
      if (!rangeContainsAddr(range, addr)) continue;
      if (!isCellWritable(s, addr)) continue;
      const { validation: _validation, ...next } = current;
      if (Object.keys(next).length === 0) formats.delete(key);
      else formats.set(key, next);
      cleared += 1;
    }
    const rangeValidations = subtractFromRangeValidations(s.format.rangeValidations, range);
    if (rangeValidations !== s.format.rangeValidations) cleared += 1;
    return { ...s, format: { ...s.format, formats, rangeValidations } };
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
      return checkBounded(input.value, validation, 'number', resolveRange);
    }
    case 'date':
    case 'time': {
      if (input.kind !== 'number') return reject(validation, `expected ${validation.kind}`);
      return checkBounded(input.value, validation, validation.kind, resolveRange);
    }
    case 'textLength': {
      const len = inputAsText(input).length;
      return checkBounded(len, validation, 'length', resolveRange);
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
    case 'error':
      return input.text;
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

/** Coerce a stored cell value into the shape typed input would produce. */
export function coerceCellValue(value: Exclude<CellValue, { kind: 'error' }>): CoercedInput {
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
  resolveRange?: RangeResolver,
): ValidationOutcome {
  const a = resolveBound(validation.aFormula, validation.a, resolveRange);
  const b = resolveBound(validation.bFormula, validation.b ?? validation.a, resolveRange);
  // A formula limit this layer cannot evaluate defers to the engine's verdict.
  if (a === null || b === null) return { ok: true };
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

/** Numeric value of a limit: the constant, or a single-cell formula limit read
 *  through the resolver. `null` when the formula cannot be evaluated here. */
function resolveBound(
  formula: string | undefined,
  constant: number,
  resolveRange?: RangeResolver,
): number | null {
  if (formula === undefined) return constant;
  if (!resolveRange) return null;
  const values = resolveRange(formula);
  if (values.length !== 1 || values[0]?.trim() === '') return null;
  const n = Number(values[0]);
  return Number.isFinite(n) ? n : null;
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
