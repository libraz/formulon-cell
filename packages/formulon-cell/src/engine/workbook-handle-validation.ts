import type {
  ComparisonOperator,
  ErrorCode,
  Sheet,
  ValidationErrorStyle,
  ValidationType,
} from '@libraz/formulon';
import { isFormulonError, orFallback } from './result.js';
import type { DataValidationInput, EngineCapabilities, Range, Workbook } from './types.js';
import { type CellValue, type Value, ValueKind } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleValidationMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

function ws(handle: unknown, sheet: number): Sheet {
  return wb(handle).sheet(sheet);
}

/** Run an engine call; an engine refusal is a soft `false`, anything else rethrows. */
function attempt(fn: () => void): boolean {
  try {
    fn();
    return true;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return false;
  }
}

function toEngineValue(value: CellValue): Value {
  switch (value.kind) {
    case 'number':
      return { kind: ValueKind.Number, number: value.value };
    case 'bool':
      return { kind: ValueKind.Bool, boolean: value.value };
    case 'text':
      return { kind: ValueKind.Text, text: value.value };
    case 'error':
      return { kind: ValueKind.Error, errorCode: value.code as ErrorCode };
    default:
      return { kind: ValueKind.Blank };
  }
}

export abstract class WorkbookHandleValidationMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Check a proposed cell value against the engine's first covering rule.
   * Returns `null` when the 0.13 validation-evaluation surface is absent or
   * the call fails, allowing the cell layer's evaluator to remain the fallback. */
  validateValue(
    sheet: number,
    row: number,
    col: number,
    value: CellValue,
  ): {
    hasRule: boolean;
    valid: boolean;
    ruleIndex: number;
    errorStyle: number;
  } | null {
    assertAlive(this);
    if (!this.capabilities.validationEvaluation) return null;
    const sh = ws(this, sheet);
    const result = orFallback(null, () => sh.validateValue(row, col, toEngineValue(value)));
    if (!result) return null;
    // The engine names the applied rule by id; callers address it by position.
    const ruleIndex =
      result.ruleId === null
        ? -1
        : orFallback([], () => sh.listValidations()).findIndex((v) => v.id === result.ruleId);
    return {
      hasRule: result.ruleId !== null,
      valid: result.valid,
      ruleIndex,
      errorStyle: result.errorStyle ?? 0,
    };
  }

  /** Enumerate populated cells that fail a stored validation rule. Results
   * are paged by the native engine; callers can resume with `nextCursor`. */
  listInvalidCells(
    sheet: number,
    cursor: number | null = null,
    limit: number | null = null,
  ): { cells: { sheet: number; row: number; col: number }[]; nextCursor: number | null } | null {
    assertAlive(this);
    if (!this.capabilities.validationEvaluation) return null;
    const result = orFallback(null, () => ws(this, sheet).getInvalidCells({ cursor, limit }));
    if (!result) return null;
    return {
      cells: result.cells.map((cell) => ({ sheet, row: cell.row, col: cell.col })),
      nextCursor: result.nextCursor,
    };
  }

  /** Snapshot of every validation entry on `sheet`. Each entry can apply to
   *  multiple ranges (`ranges`) and carries an spreadsheet-style descriptor: numeric
   *  `type` ordinal (0 none, 1 whole, 2 decimal, 3 list, 4 date, 5 time,
   *  6 textLength, 7 custom), `op` as the engine's `ComparisonOperator`,
   *  formula1/2 strings (a stated formula keeps its leading `=`; none reads
   *  empty), and the surrounding error/prompt metadata. `showDropDown` is
   *  whether the in-cell arrow is shown. Empty when `capabilities.dataValidation`
   *  is off or when the engine returns no rules. */
  getValidationsForSheet(sheet: number): {
    ranges: Range[];
    type: number;
    op: ComparisonOperator;
    errorStyle: ValidationErrorStyle;
    allowBlank: boolean;
    showInputMessage: boolean;
    showErrorMessage: boolean;
    showDropDown: boolean;
    formula1: string;
    formula2: string;
    errorTitle: string;
    errorMessage: string;
    promptTitle: string;
    promptMessage: string;
  }[] {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return [];
    return orFallback([], () => ws(this, sheet).listValidations()).map((v) => ({
      ranges: v.ranges.map((m) => ({
        sheet,
        r0: m.firstRow,
        c0: m.firstCol,
        r1: m.lastRow,
        c1: m.lastCol,
      })),
      type: v.type,
      op: v.operator,
      errorStyle: v.errorStyle,
      allowBlank: v.allowBlank,
      showInputMessage: v.showInputMessage,
      showErrorMessage: v.showErrorMessage,
      showDropDown: v.showInCellDropdown,
      formula1: v.formula1 ?? '',
      formula2: v.formula2 ?? '',
      errorTitle: v.errorTitle ?? '',
      errorMessage: v.errorMessage ?? '',
      promptTitle: v.promptTitle ?? '',
      promptMessage: v.promptMessage ?? '',
    }));
  }

  /** Append a data-validation rule to `sheet`. Ranges are inclusive `Range`
   *  records; everything else mirrors the upstream `DataValidationInput` shape.
   *  Returns false on engine failure or when capability is off. */
  addValidationEntry(
    sheet: number,
    input: {
      ranges: Range[];
      type: ValidationType;
      op?: ComparisonOperator;
      errorStyle?: ValidationErrorStyle;
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
    },
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return false;
    const dv: DataValidationInput = {
      type: input.type,
      ranges: input.ranges.map((r) => ({
        firstRow: r.r0,
        firstCol: r.c0,
        lastRow: r.r1,
        lastCol: r.c1,
      })),
      ...(input.op !== undefined ? { operator: input.op } : {}),
      ...(input.errorStyle !== undefined ? { errorStyle: input.errorStyle } : {}),
      ...(input.allowBlank !== undefined ? { allowBlank: input.allowBlank } : {}),
      ...(input.showInputMessage !== undefined ? { showInputMessage: input.showInputMessage } : {}),
      ...(input.showErrorMessage !== undefined ? { showErrorMessage: input.showErrorMessage } : {}),
      ...(input.showDropDown !== undefined ? { showInCellDropdown: input.showDropDown } : {}),
      ...(input.formula1 !== undefined ? { formula1: input.formula1 } : {}),
      ...(input.formula2 !== undefined ? { formula2: input.formula2 } : {}),
      ...(input.errorTitle !== undefined ? { errorTitle: input.errorTitle } : {}),
      ...(input.errorMessage !== undefined ? { errorMessage: input.errorMessage } : {}),
      ...(input.promptTitle !== undefined ? { promptTitle: input.promptTitle } : {}),
      ...(input.promptMessage !== undefined ? { promptMessage: input.promptMessage } : {}),
    };
    return attempt(() => {
      ws(this, sheet).addValidation(dv);
    });
  }

  /** Remove the validation rule at `index` on `sheet`. */
  removeValidationAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return false;
    return attempt(() => ws(this, sheet).removeValidation(index));
  }

  /** Drop every validation rule on `sheet`. */
  clearValidations(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return false;
    return attempt(() => ws(this, sheet).clearValidations());
  }
}

export function installValidationMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleValidationMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleValidationMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
