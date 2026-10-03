import type { DataValidationInput, EngineCapabilities, Range, Workbook } from './types.js';
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

export abstract class WorkbookHandleValidationMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Snapshot of every validation entry on `sheet`. Each entry can apply to
   *  multiple ranges (`ranges`) and carries an spreadsheet-style descriptor: numeric
   *  `type` ordinal (0 none, 1 whole, 2 decimal, 3 list, 4 date, 5 time,
   *  6 textLength, 7 custom), numeric `op` ordinal (0 between … 7 lessThanOrEqual),
   *  formula1/2 strings, and the surrounding error/prompt metadata. Empty
   *  when `capabilities.dataValidation` is off or when the engine returns
   *  no rules. */
  getValidationsForSheet(sheet: number): {
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
  }[] {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return [];
    const arr = wb(this).getValidations(sheet);
    if (!arr.status.ok) return [];
    return arr.map((v) => ({
      ranges: v.ranges.map((m) => ({
        sheet,
        r0: m.firstRow,
        c0: m.firstCol,
        r1: m.lastRow,
        c1: m.lastCol,
      })),
      type: v.type,
      op: v.op,
      errorStyle: v.errorStyle,
      allowBlank: v.allowBlank,
      showInputMessage: v.showInputMessage,
      showErrorMessage: v.showErrorMessage,
      showDropDown:
        'showDropDown' in v && typeof v.showDropDown === 'boolean' ? v.showDropDown : undefined,
      formula1: v.formula1,
      formula2: v.formula2,
      errorTitle: v.errorTitle,
      errorMessage: v.errorMessage,
      promptTitle: v.promptTitle,
      promptMessage: v.promptMessage,
    }));
  }

  /** Append a data-validation rule to `sheet`. Ranges are inclusive `Range`
   *  records; everything else mirrors the upstream `DataValidationInput` shape.
   *  Returns false on engine failure or when capability is off. */
  addValidationEntry(
    sheet: number,
    input: {
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
      ...(input.op !== undefined ? { op: input.op } : {}),
      ...(input.errorStyle !== undefined ? { errorStyle: input.errorStyle } : {}),
      ...(input.allowBlank !== undefined ? { allowBlank: input.allowBlank } : {}),
      ...(input.showInputMessage !== undefined ? { showInputMessage: input.showInputMessage } : {}),
      ...(input.showErrorMessage !== undefined ? { showErrorMessage: input.showErrorMessage } : {}),
      ...(input.showDropDown !== undefined ? { showDropDown: input.showDropDown } : {}),
      ...(input.formula1 !== undefined ? { formula1: input.formula1 } : {}),
      ...(input.formula2 !== undefined ? { formula2: input.formula2 } : {}),
      ...(input.errorTitle !== undefined ? { errorTitle: input.errorTitle } : {}),
      ...(input.errorMessage !== undefined ? { errorMessage: input.errorMessage } : {}),
      ...(input.promptTitle !== undefined ? { promptTitle: input.promptTitle } : {}),
      ...(input.promptMessage !== undefined ? { promptMessage: input.promptMessage } : {}),
    };
    const s = wb(this).addValidation(sheet, dv);
    return s.ok;
  }

  /** Remove the validation rule at `index` on `sheet`. */
  removeValidationAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return false;
    const s = wb(this).removeValidationAt(sheet, index);
    return s.ok;
  }

  /** Drop every validation rule on `sheet`. */
  clearValidations(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.dataValidation) return false;
    const s = wb(this).clearValidations(sheet);
    return s.ok;
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
