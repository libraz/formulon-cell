import type { Sheet as EngineSheet } from '@libraz/formulon';
import type { MemberSpec, SheetContext } from './facade-context.js';

/** Cell values, formulas and evaluation: engine `Sheet` members of the same name. */
export const SHEET_CELL_MEMBERS = {
  read: ['getTypedValue', 'getFormula', 'evaluateTypedFormula', 'evaluateTypedFormulaArray'],
  cellWrite: ['setNumber', 'setText', 'setBool', 'setError', 'setBlank', 'setFormula'],
} as const satisfies MemberSpec<SheetContext>;

type CellMemberName =
  | (typeof SHEET_CELL_MEMBERS.read)[number]
  | (typeof SHEET_CELL_MEMBERS.cellWrite)[number];

declare module './workbook.js' {
  interface Sheet extends Pick<EngineSheet, CellMemberName> {}
}
