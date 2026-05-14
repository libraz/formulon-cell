import type {
  CellFormat,
  CellValidation,
  SpreadsheetStore,
  ValidationErrorStyle,
  ValidationMeta,
  ValidationOp,
} from '../store/store.js';
import { addrKey } from './address.js';
import { parseRangeRef } from './range-resolver.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** OOXML validation `type` ordinal — 0 none, 1 whole, 2 decimal, 3 list,
 *  4 date, 5 time, 6 textLength, 7 custom. */
const DV_TYPE: Record<Exclude<CellValidation['kind'], 'list' | 'custom'>, number> = {
  whole: 1,
  decimal: 2,
  date: 4,
  time: 5,
  textLength: 6,
};
const DV_TYPE_LIST = 3;
const DV_TYPE_CUSTOM = 7;

/** OOXML op ordinal — 0 between, 1 notBetween, 2 equal, 3 notEqual,
 *  4 lessThan, 5 lessThanOrEqual, 6 greaterThan, 7 greaterThanOrEqual. */
const OP_ORDINAL: Record<ValidationOp, number> = {
  between: 0,
  notBetween: 1,
  '=': 2,
  '<>': 3,
  '<': 4,
  '<=': 5,
  '>': 6,
  '>=': 7,
};
const OP_FROM_ORDINAL: Record<number, ValidationOp> = {
  0: 'between',
  1: 'notBetween',
  2: '=',
  3: '<>',
  4: '<',
  5: '<=',
  6: '>',
  7: '>=',
};

const ERROR_STYLE_ORDINAL: Record<ValidationErrorStyle, number> = {
  stop: 0,
  warning: 1,
  information: 2,
};
const ERROR_STYLE_FROM_ORDINAL: Record<number, ValidationErrorStyle> = {
  0: 'stop',
  1: 'warning',
  2: 'information',
};

/**
 * Hydrate FormatSlice `validation` fields from engine entries on `sheet`.
 * Every supported `type` ordinal is surfaced — `list` parses inline literals,
 * `whole`/`decimal`/`date`/`time`/`textLength` decode formula1/formula2 to
 * numbers, `custom` carries the formula1 string verbatim. Range-ref list
 * sources still drop (no formula expansion at hydrate time).
 */
export function hydrateValidationsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.dataValidation) return;
  const entries = wb.getValidationsForSheet(sheet);
  if (entries.length === 0) return;

  store.setState((s) => {
    const formats = new Map(s.format.formats);
    for (const v of entries) {
      const decoded = decodeValidation(v);
      if (!decoded) continue;
      for (const r of v.ranges) {
        if (r.sheet !== sheet) continue;
        for (let row = r.r0; row <= r.r1; row += 1) {
          for (let col = r.c0; col <= r.c1; col += 1) {
            const k = addrKey({ sheet, row, col });
            const cur = formats.get(k);
            const next: CellFormat = { ...(cur ?? {}), validation: decoded };
            formats.set(k, next);
          }
        }
      }
    }
    return { ...s, format: { formats } };
  });
}

/**
 * Replace the engine's data-validation rules on `sheet` with whatever
 * FormatSlice currently asserts. Cells sharing identical validation config
 * (same kind/op/values/messages) coalesce into one rule with multiple
 * single-cell ranges to keep rule count low. No-op when the capability flag
 * is off.
 */
export function syncValidationsToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.dataValidation) return;
  const buckets = new Map<string, { validation: CellValidation; ranges: Range[] }>();
  const formats = store.getState().format.formats;
  for (const [key, fmt] of formats) {
    if (!fmt.validation) continue;
    const [sStr, rStr, cStr] = key.split(':');
    if (sStr === undefined || rStr === undefined || cStr === undefined) continue;
    const sIdx = Number.parseInt(sStr, 10);
    if (sIdx !== sheet) continue;
    const row = Number.parseInt(rStr, 10);
    const col = Number.parseInt(cStr, 10);
    const sig = JSON.stringify(fmt.validation);
    let bucket = buckets.get(sig);
    if (!bucket) {
      bucket = { validation: fmt.validation, ranges: [] };
      buckets.set(sig, bucket);
    }
    bucket.ranges.push({ sheet, r0: row, c0: col, r1: row, c1: col });
  }
  wb.clearValidations(sheet);
  for (const { validation, ranges } of buckets.values()) {
    if (ranges.length === 0) continue;
    const encoded = encodeValidation(validation, ranges);
    if (!encoded) continue;
    wb.addValidationEntry(sheet, encoded);
  }
}

interface EngineValidationEntry {
  ranges: Range[];
  type: number;
  op: number;
  errorStyle: number;
  allowBlank: boolean;
  showInputMessage: boolean;
  showErrorMessage: boolean;
  formula1: string;
  formula2: string;
  errorTitle: string;
  errorMessage: string;
  promptTitle: string;
  promptMessage: string;
}

function decodeValidation(v: EngineValidationEntry): CellValidation | null {
  // Only surface fields that depart from desktop defaults so round-tripping a
  // simple `{ kind: 'list', source: [...] }` stays minimal in the store.
  const errorStyle = ERROR_STYLE_FROM_ORDINAL[v.errorStyle] ?? 'stop';
  const meta: ValidationMeta = {
    ...(v.allowBlank === false ? { allowBlank: false } : {}),
    ...(errorStyle !== 'stop' ? { errorStyle } : {}),
    ...(v.showInputMessage === true ? { showInputMessage: true } : {}),
    ...(v.showErrorMessage === true ? { showErrorMessage: true } : {}),
    ...(v.errorTitle ? { errorTitle: v.errorTitle } : {}),
    ...(v.errorMessage ? { errorMessage: v.errorMessage } : {}),
    ...(v.promptTitle ? { promptTitle: v.promptTitle } : {}),
    ...(v.promptMessage ? { promptMessage: v.promptMessage } : {}),
  };
  switch (v.type) {
    case DV_TYPE_LIST: {
      const literal = parseInlineList(v.formula1);
      if (literal) return { kind: 'list', source: literal, ...meta };
      const ref = parseRangeRefSource(v.formula1);
      if (ref) return { kind: 'list', source: { ref }, ...meta };
      return null;
    }
    case DV_TYPE_CUSTOM: {
      const formula = v.formula1.replace(/^=/, '').trim();
      if (!formula) return null;
      return { kind: 'custom', formula, ...meta };
    }
    default: {
      const kind = ordinalToBoundedKind(v.type);
      if (!kind) return null;
      const op = OP_FROM_ORDINAL[v.op] ?? '=';
      const a = parseNumberFormula(v.formula1);
      if (a === null) return null;
      const b = v.formula2 ? parseNumberFormula(v.formula2) : null;
      if (op === 'between' || op === 'notBetween') {
        if (b === null) return null;
        return { kind, op, a, b, ...meta };
      }
      return { kind, op, a, ...meta };
    }
  }
}

function ordinalToBoundedKind(
  ordinal: number,
): 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | null {
  for (const [k, v] of Object.entries(DV_TYPE) as [keyof typeof DV_TYPE, number][]) {
    if (v === ordinal) return k;
  }
  return null;
}

function encodeValidation(
  validation: CellValidation,
  ranges: Range[],
): {
  ranges: Range[];
  type: number;
  op?: number;
  errorStyle?: number;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  formula1?: string;
  formula2?: string;
  errorTitle?: string;
  errorMessage?: string;
  promptTitle?: string;
  promptMessage?: string;
} | null {
  const metaPayload = encodeMeta(validation);
  switch (validation.kind) {
    case 'list':
      return {
        ranges,
        type: DV_TYPE_LIST,
        formula1: encodeListSource(validation.source),
        ...metaPayload,
      };
    case 'custom':
      if (!validation.formula.trim()) return null;
      return {
        ranges,
        type: DV_TYPE_CUSTOM,
        formula1: ensureFormulaPrefix(validation.formula),
        ...metaPayload,
      };
    default: {
      const type = DV_TYPE[validation.kind];
      const op = OP_ORDINAL[validation.op];
      const formula1 = String(validation.a);
      const needsB = validation.op === 'between' || validation.op === 'notBetween';
      const formula2 = needsB ? String(validation.b ?? validation.a) : '';
      return {
        ranges,
        type,
        op,
        formula1,
        ...(formula2 ? { formula2 } : {}),
        ...metaPayload,
      };
    }
  }
}

function encodeMeta(validation: ValidationMeta): {
  errorStyle?: number;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorTitle?: string;
  errorMessage?: string;
  promptTitle?: string;
  promptMessage?: string;
} {
  return {
    ...(validation.errorStyle !== undefined
      ? { errorStyle: ERROR_STYLE_ORDINAL[validation.errorStyle] }
      : {}),
    ...(validation.allowBlank !== undefined
      ? { allowBlank: validation.allowBlank }
      : { allowBlank: true }),
    ...(validation.showInputMessage !== undefined
      ? { showInputMessage: validation.showInputMessage }
      : { showInputMessage: true }),
    ...(validation.showErrorMessage !== undefined
      ? { showErrorMessage: validation.showErrorMessage }
      : { showErrorMessage: true }),
    ...(validation.errorTitle ? { errorTitle: validation.errorTitle } : {}),
    ...(validation.errorMessage ? { errorMessage: validation.errorMessage } : {}),
    ...(validation.promptTitle ? { promptTitle: validation.promptTitle } : {}),
    ...(validation.promptMessage ? { promptMessage: validation.promptMessage } : {}),
  };
}

function ensureFormulaPrefix(formula: string): string {
  const t = formula.trim();
  return t.startsWith('=') ? t : `=${t}`;
}

/** Parse spreadsheet-style inline list literals. Returns null when the source
 *  string is empty, a range reference, or otherwise unparseable. */
function parseInlineList(formula: string): string[] | null {
  const trimmed = formula.trim().replace(/^=/, '');
  if (!trimmed) return null;
  if (/[!$:]/.test(trimmed)) return null;
  const inner = trimmed.startsWith('"') && trimmed.endsWith('"') ? trimmed.slice(1, -1) : trimmed;
  const parts = inner
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return parts.length > 0 ? parts : null;
}

/** Pull out a range reference (Sheet1!$A$1:$A$10 or $A$1:$A$10) from formula1
 *  for list-kind DV. Returns the cleaned ref (no leading `=`) when it parses
 *  as an A1-style range, otherwise null. */
function parseRangeRefSource(formula: string): string | null {
  const trimmed = formula.trim().replace(/^=/, '');
  if (!trimmed) return null;
  // Reject inline-string literals so this never claims a list of values.
  if (trimmed.startsWith('"')) return null;
  return parseRangeRef(trimmed) ? trimmed : null;
}

function parseNumberFormula(formula: string): number | null {
  const t = formula.trim().replace(/^=/, '');
  if (!t) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

/** Encode either a literal string-array or a range-ref source to formula1.
 *  Inline literals always wrap in double quotes so spreadsheets parse them back as
 *  a list (rather than a function call); range refs pass through with a
 *  leading `=` to disambiguate from a single-token literal. */
function encodeListSource(source: string[] | { ref: string }): string {
  if (Array.isArray(source)) return `"${source.join(',')}"`;
  const ref = source.ref.trim().replace(/^=/, '');
  return `=${ref}`;
}
