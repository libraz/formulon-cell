import { inputLocaleForWorkbook } from '../commands/coerce-input.js';
import { formatA1FormulaAsR1C1 } from '../commands/refs.js';
import type { State } from '../store/store.js';
import { addrKey } from './address.js';
import { localeOrdinalForSpreadsheetProfile, localizeFormula } from './function-locale.js';
import type { Addr, CellValue } from './types.js';
import { formatBoolean, formatEditableNumber } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';

export interface FormatCellForEditOptions {
  formulaHidden?: boolean;
  sheetProtected?: boolean;
  formulaOverride?: string | null;
  formatFormula?: (formula: string) => string;
  locale?: string;
}

/** Resolve a stored A1 formula to the text shown by an editor. Native R1C1
 * output is selected before function aliases are localized so references keep
 * the engine's exact relative-address semantics. */
export function formulaForEdit(
  wb: WorkbookHandle,
  addr: Addr,
  formula: string | null,
  r1c1 = false,
): string | null {
  if (!formula) return null;
  if (!r1c1) {
    return typeof wb.getFormulaForDisplay === 'function'
      ? wb.getFormulaForDisplay(formula)
      : formula;
  }
  // The engine spells the references in the profile's R1C1 letters along with its separators.
  const localized =
    typeof wb.localizeFormulaR1C1Text === 'function'
      ? wb.localizeFormulaR1C1Text(formula, addr)
      : null;
  if (localized !== null) return localized;
  let native: string | null = null;
  if (wb.capabilities.formulaR1C1 && addr.sheet >= 0 && addr.sheet < wb.sheetCount) {
    try {
      native = wb.cellFormulaR1C1(addr);
    } catch {
      // A transient workbook/sheet transition can invalidate the native read;
      // the A1 formula below remains a safe source for conversion.
    }
  }
  const resolved = native ?? formatA1FormulaAsR1C1(formula, addr);
  // Native R1C1 output already resolved the exact relative references. Apply
  // full profile conversion only after that resolution so separators and
  // localized literals are displayed consistently with A1 editing.
  if (typeof wb.getFormulaForDisplay === 'function') {
    return wb.getFormulaForDisplay(resolved);
  }
  const profile = typeof wb.spreadsheetProfileId === 'function' ? wb.spreadsheetProfileId() : null;
  return localizeFormula(wb, resolved, localeOrdinalForSpreadsheetProfile(profile));
}

/**
 * Resolve the seed text the formula bar / inline editor should display
 * for `cell`. Honours four cases in priority order:
 *
 *   1. Cells with an explicit formula text — show it verbatim.
 *   2. Cells whose cached value is a lambda (no formula stored): pull the
 *      LAMBDA body via `getLambdaText` and prefix `=` so the editor sees
 *      something parseable. Desktop spreadsheets do the same — formula bar shows the
 *      lambda definition, not `#CALC!`.
 *   3. Plain values — render them as the editor would type them
 *      (numbers without thousand separators, booleans as TRUE/FALSE,
 *      errors as their canonical sentinel).
 *   4. Anything else (e.g. blanks) — empty string.
 *
 * `wb` and `addr` are optional so callers without engine context (e.g.
 * tests, isolated formatting paths) can still resolve cases 1, 3, 4.
 */
export function formatCellForEdit(
  cell: { value: CellValue; formula: string | null } | undefined,
  wb?: WorkbookHandle,
  addr?: Addr,
  options?: FormatCellForEditOptions,
): string {
  const hideFormula = options?.formulaHidden === true && options.sheetProtected === true;
  const formula =
    options && 'formulaOverride' in options ? options.formulaOverride : (cell?.formula ?? null);
  const displayFormula = (text: string): string => options?.formatFormula?.(text) ?? text;
  if (!cell) return formula ? (hideFormula ? '' : displayFormula(formula)) : '';
  if (formula) return hideFormula ? '' : displayFormula(formula);
  if (wb && addr) {
    const lambda = wb.getLambdaText(addr);
    if (lambda) {
      const localized =
        typeof wb.getFormulaForDisplay === 'function'
          ? wb.getFormulaForDisplay(`=${lambda}`)
          : `=${lambda}`;
      return hideFormula ? '' : displayFormula(localized);
    }
  }
  const v = cell.value;
  switch (v.kind) {
    case 'number':
      return formatEditableNumber(v.value, options?.locale);
    case 'bool':
      return formatBoolean(v.value, options?.locale);
    case 'text':
      return v.value;
    case 'error':
      return v.text;
    default:
      return '';
  }
}

/** Edit-seed text for the cell at `addr`, honouring the stored formula, R1C1
 *  mode, hidden formulas on protected sheets and the workbook input locale. */
export function editSeedForAddr(state: State, wb: WorkbookHandle, addr: Addr): string {
  const key = addrKey(addr);
  const cell = state.data.cells.get(key);
  const storedFormula = wb.cellFormula(addr) ?? cell?.formula ?? null;
  return formatCellForEdit(cell, wb, addr, {
    formulaOverride: formulaForEdit(wb, addr, storedFormula, state.ui.r1c1),
    formulaHidden: state.format.formats.get(key)?.formulaHidden === true,
    sheetProtected: state.protection.protectedSheets.has(addr.sheet),
    locale: inputLocaleForWorkbook(wb),
  });
}
