import type { State } from '../store/types.js';
import { SheetKind } from './types.js';

/** The type of `sheet` as the store holds it; a sheet the store has no kind for is a worksheet. */
export const sheetKindOf = (state: State, sheet: number): SheetKind =>
  state.layout.sheetKinds?.get(sheet) ?? SheetKind.Worksheet;

/** Whether `sheet` is a chartsheet. */
export const isChartsheet = (state: State, sheet: number): boolean =>
  sheetKindOf(state, sheet) === SheetKind.Chartsheet;

/** Whether `sheet` holds cells; only a worksheet does. */
export const sheetHasCells = (state: State, sheet: number): boolean =>
  sheetKindOf(state, sheet) === SheetKind.Worksheet;
