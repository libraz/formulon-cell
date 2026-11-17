import type {
  ComparisonOperator,
  ErrorCode,
  ValidationErrorStyle,
  ValidationType,
} from '@libraz/formulon';
import { rangeOfRect, rectOfRange } from './address.js';
import { attemptWrite, orFallback } from './result.js';
import type { DataValidationInput, EngineCapabilities, Range } from './types.js';
import { type CellValue, type Value, ValueKind } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { assertAlive, ws } from './workbook-handle-internals.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleValidationMethods {}
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
      ranges: v.ranges.map((m) => rangeOfRect(sheet, m)),
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
      ranges: input.ranges.map(rectOfRange),
      ...(input.op !== undefined ? { operator: input.op } : {}),
      ...(input.errorStyle !== undefined ? { errorStyle: input.errorStyle } : {}),
      ...(input.allowBlank !== undefined ? { allowBlank: input.allowBlank } : {}),
      ...(input.showInputMessage !== undefined ? { showInputMessage: input.showInputMessage } : {}),
      ...(input.showErrorMessage !== undefined ? { showErrorMessage: input.showErrorMessage } : {}),
      ...(input.showDropDown !== undefined ? { showInCellDropdown: input.showDropDown } : {}),
      // The engine refuses an empty string; an empty text is the absent one.
      ...(input.formula1 ? { formula1: input.formula1 } : {}),
      ...(input.formula2 ? { formula2: input.formula2 } : {}),
      ...(input.errorTitle ? { errorTitle: input.errorTitle } : {}),
      ...(input.errorMessage ? { errorMessage: input.errorMessage } : {}),
      ...(input.promptTitle ? { promptTitle: input.promptTitle } : {}),
      ...(input.promptMessage ? { promptMessage: input.promptMessage } : {}),
    };
    return attemptWrite(() => {
      ws(this, sheet).addValidation(dv);
    });
  }

  /** Remove the validation rule at `index` on `sheet`. */
  removeValidationAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return false;
    return attemptWrite(() => ws(this, sheet).removeValidation(index));
  }

  /** Drop every validation rule on `sheet`. */
  clearValidations(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return false;
    return attemptWrite(() => ws(this, sheet).clearValidations());
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
