import { isWideValidation } from '../store/range-validations.js';
import type {
  CellFormat,
  CellValidation,
  SpreadsheetStore,
  ValidationBoundFormulas,
  ValidationErrorStyle,
  ValidationMeta,
  ValidationOp,
} from '../store/store.js';
import { addrKey, parseAddrKey } from './address.js';
import type { EngineSyncOptions } from './cell-format-sync.js';
import type { Addr, Range } from './types.js';
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
 *  4 greaterThan, 5 lessThan, 6 greaterThanOrEqual, 7 lessThanOrEqual. */
const OP_ORDINAL: Record<ValidationOp, number> = {
  between: 0,
  notBetween: 1,
  '=': 2,
  '<>': 3,
  '>': 4,
  '<': 5,
  '>=': 6,
  '<=': 7,
};
const OP_FROM_ORDINAL: Record<number, ValidationOp> = {
  0: 'between',
  1: 'notBetween',
  2: '=',
  3: '<>',
  4: '>',
  5: '<',
  6: '>=',
  7: '<=',
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

const strictSyncError = (operation: string, address: string): Error =>
  new Error(`Strict engine sync failed: validations ${operation} at ${address}`);

/**
 * Hydrate FormatSlice `validation` fields from engine entries on `sheet`.
 * Every supported `type` ordinal is surfaced — `list` parses inline literals,
 * `whole`/`decimal`/`date`/`time`/`textLength` decode formula1/formula2 to
 * numbers, `custom` carries the formula1 string verbatim. Range-ref list
 * sources retain their range expression for the native and local validators.
 */
export function hydrateValidationsFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.dataValidation) return;
  const entries = wb.getValidationsForSheet(sheet);

  // The engine is authoritative: rules it no longer holds for this sheet go.
  store.setState((s) => {
    const formats = new Map<string, CellFormat>();
    let stale = false;
    for (const [key, fmt] of s.format.formats) {
      if (!fmt.validation || parseAddrKey(key)?.sheet !== sheet) {
        formats.set(key, fmt);
        continue;
      }
      stale = true;
      const { validation: _stale, ...rest } = fmt;
      if (Object.keys(rest).length > 0) formats.set(key, rest);
    }
    const wide = (s.format.rangeValidations ?? []).flatMap((rule) => {
      const ranges = rule.ranges.filter((range) => range.sheet !== sheet);
      return ranges.length > 0 ? [{ ...rule, ranges }] : [];
    });
    const staleWide = (s.format.rangeValidations ?? []).some((rule) =>
      rule.ranges.some((range) => range.sheet === sheet),
    );
    if (entries.length === 0 && !stale && !staleWide) return s;
    for (const v of entries) {
      const decoded = decodeValidation(v);
      if (!decoded) continue;
      const ranges = v.ranges.filter((r) => r.sheet === sheet);
      if (isWideValidation(ranges)) {
        wide.push({ ranges, validation: decoded });
        continue;
      }
      for (const r of ranges) {
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
    return { ...s, format: { ...s.format, formats, rangeValidations: wide } };
  });
}

/**
 * Replace the engine's data-validation rules on `sheet` with whatever
 * FormatSlice currently asserts. Cells sharing identical validation config
 * (same kind/op/values/messages) coalesce into one rule whose ranges are the
 * merged rectangles of those cells; range-stored rules go out as they are.
 * No-op when the capability flag is off.
 */
export function syncValidationsToEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
  options?: EngineSyncOptions,
): void {
  if (!wb.capabilities.dataValidation) return;
  const strict = options?.strict === true;
  const buckets = new Map<string, { validation: CellValidation; cells: Addr[] }>();
  const formats = store.getState().format.formats;
  for (const [key, fmt] of formats) {
    if (!fmt.validation) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const sig = JSON.stringify(fmt.validation);
    let bucket = buckets.get(sig);
    if (!bucket) {
      bucket = { validation: fmt.validation, cells: [] };
      buckets.set(sig, bucket);
    }
    bucket.cells.push(addr);
  }
  const rules: Array<{ validation: CellValidation; ranges: Range[] }> = [];
  for (const { validation, cells } of buckets.values()) {
    rules.push({ validation, ranges: coalesceCells(sheet, cells) });
  }
  for (const rule of store.getState().format.rangeValidations ?? []) {
    const ranges = rule.ranges.filter((range) => range.sheet === sheet);
    if (ranges.length > 0) rules.push({ validation: rule.validation, ranges });
  }
  const clearSucceeded = wb.clearValidations(sheet);
  if (strict && !clearSucceeded) {
    throw strictSyncError('clearValidations', `sheet:${sheet}`);
  }
  for (const { validation, ranges } of rules) {
    if (ranges.length === 0) continue;
    const encoded = encodeValidation(validation, ranges);
    const address = ranges[0]
      ? `${ranges[0].sheet}:${ranges[0].r0}:${ranges[0].c0}`
      : `sheet:${sheet}`;
    if (!encoded) {
      if (strict) throw strictSyncError('encodeValidation', address);
      continue;
    }
    const addSucceeded = wb.addValidationEntry(sheet, encoded);
    if (strict && !addSucceeded) {
      throw strictSyncError('addValidationEntry', address);
    }
  }
}

/** Merge cells into rectangles: horizontal runs per row, then identical runs
 *  on consecutive rows stack into one block. */
function coalesceCells(sheet: number, cells: readonly Addr[]): Range[] {
  const byRow = new Map<number, number[]>();
  for (const { row, col } of cells) {
    const cols = byRow.get(row);
    if (cols) cols.push(col);
    else byRow.set(row, [col]);
  }
  const open = new Map<string, Range>();
  const out: Range[] = [];
  let prevRow = -2;
  for (const row of [...byRow.keys()].sort((a, b) => a - b)) {
    const cols = (byRow.get(row) as number[]).sort((a, b) => a - b);
    const next = new Map<string, Range>();
    let c0 = cols[0] as number;
    let c1 = c0;
    const flush = (): void => {
      const runKey = `${c0}:${c1}`;
      const stacked = row === prevRow + 1 ? open.get(runKey) : undefined;
      if (stacked) {
        const grown = { ...stacked, r1: row };
        next.set(runKey, grown);
        out[out.indexOf(stacked)] = grown;
      } else {
        const fresh: Range = { sheet, r0: row, c0, r1: row, c1 };
        next.set(runKey, fresh);
        out.push(fresh);
      }
    };
    for (const col of cols.slice(1)) {
      if (col === c1 + 1) {
        c1 = col;
        continue;
      }
      flush();
      c0 = col;
      c1 = col;
    }
    flush();
    open.clear();
    for (const [k, v] of next) open.set(k, v);
    prevRow = row;
  }
  return out;
}

interface EngineValidationEntry {
  ranges: Range[];
  type: number;
  op: number;
  errorStyle: number;
  allowBlank: boolean;
  showInputMessage: boolean;
  showErrorMessage: boolean;
  showDropDown?: boolean;
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
    // Preserve an explicit `false` (error alert disabled). The store convention
    // omits the field when enabled (the default) and only records the disabled
    // case, so decoding `true` must NOT emit the field.
    ...(v.showErrorMessage === false ? { showErrorMessage: false } : {}),
    // The native API exposes the user-facing meaning (arrow shown/hidden),
    // while the engine handles OOXML's inverted attribute internally.
    ...(v.showDropDown === false ? { showDropdown: false } : {}),
    ...(v.errorTitle ? { errorTitle: v.errorTitle } : {}),
    ...(v.errorMessage ? { errorMessage: v.errorMessage } : {}),
    ...(v.promptTitle ? { promptTitle: v.promptTitle } : {}),
    ...(v.promptMessage ? { promptMessage: v.promptMessage } : {}),
  };
  switch (v.type) {
    case DV_TYPE_LIST: {
      const literal = parseInlineList(v.formula1);
      if (literal) return { kind: 'list', source: literal, ...meta };
      const ref = parseListFormulaSource(v.formula1);
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
      const f1 = stripEquals(v.formula1);
      if (!f1) return null;
      const a = parseNumberFormula(f1);
      const bounds: ValidationBoundFormulas = a === null ? { aFormula: f1 } : {};
      if (op === 'between' || op === 'notBetween') {
        const f2 = stripEquals(v.formula2);
        if (!f2) return null;
        const b = parseNumberFormula(f2);
        if (b === null) bounds.bFormula = f2;
        return { kind, op, a: a ?? 0, b: b ?? 0, ...bounds, ...meta };
      }
      return { kind, op, a: a ?? 0, ...bounds, ...meta };
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
  showDropDown?: boolean;
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
        // Formulon's public validation API consumes the raw OOXML formula
        // without the optional leading `=`.
        formula1: validation.formula.trim().replace(/^=/, '').trim(),
        ...metaPayload,
      };
    default: {
      const type = DV_TYPE[validation.kind];
      const op = OP_ORDINAL[validation.op];
      const formula1 = validation.aFormula ?? String(validation.a);
      const needsB = validation.op === 'between' || validation.op === 'notBetween';
      const formula2 = needsB ? (validation.bFormula ?? String(validation.b ?? validation.a)) : '';
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
  showDropDown?: boolean;
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
    ...(validation.showDropdown === false ? { showDropDown: false } : {}),
    ...(validation.errorTitle ? { errorTitle: validation.errorTitle } : {}),
    ...(validation.errorMessage ? { errorMessage: validation.errorMessage } : {}),
    ...(validation.promptTitle ? { promptTitle: validation.promptTitle } : {}),
    ...(validation.promptMessage ? { promptMessage: validation.promptMessage } : {}),
  };
}

/** Parse an inline list literal (`"a,b,c"`). Returns null for anything else:
 *  a range, a defined name or a formula all stay formulas. */
function parseInlineList(formula: string): string[] | null {
  const trimmed = formula.trim().replace(/^=/, '');
  if (!trimmed) return null;
  const quoted = trimmed.length >= 2 && trimmed.startsWith('"') && trimmed.endsWith('"');
  // A bare `a,b` is a literal only when it cannot be a function call.
  if (!quoted && (!trimmed.includes(',') || /[!$:()]/.test(trimmed))) return null;
  const inner = quoted ? trimmed.slice(1, -1) : trimmed;
  const parts = inner
    .split(',')
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  return parts.length > 0 ? parts : null;
}

/** List source that is not a literal: an A1 range, a defined name or a formula
 *  such as `INDIRECT(...)`, returned without the leading `=`. */
function parseListFormulaSource(formula: string): string | null {
  const trimmed = stripEquals(formula);
  return trimmed || null;
}

function stripEquals(formula: string): string {
  return formula.trim().replace(/^=/, '').trim();
}

function parseNumberFormula(formula: string): number | null {
  const t = stripEquals(formula);
  if (!t) return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

/** Encode either a literal string-array or a range-ref source to formula1.
 *  Inline literals always wrap in double quotes so spreadsheets parse them back as
 *  a list (rather than a function call). Formulon's validation API follows the
 *  OOXML contract and expects range refs without the optional leading `=`. */
function encodeListSource(source: string[] | { ref: string }): string {
  if (Array.isArray(source)) return `"${source.join(',')}"`;
  const ref = source.ref.trim().replace(/^=/, '');
  return ref;
}
