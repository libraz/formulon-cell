import type { State } from '../store/types.js';
import { SheetKind } from './types.js';

/** Where a kind is read from: the store, or a workbook handle when no store is at hand. */
export type SheetKindSource = State | { sheetKind(sheet: number): SheetKind };

/** The type of `sheet` as `source` holds it; a sheet the store has no kind for is a worksheet. */
export const sheetKindOf = (source: SheetKindSource, sheet: number): SheetKind =>
  'sheetKind' in source
    ? source.sheetKind(sheet)
    : (source.layout.sheetKinds?.get(sheet) ?? SheetKind.Worksheet);

/** Whether `sheet` is a chartsheet. */
export const isChartsheet = (source: SheetKindSource, sheet: number): boolean =>
  sheetKindOf(source, sheet) === SheetKind.Chartsheet;

/** Whether `sheet` holds cells; only a worksheet does. */
export const sheetHasCells = (source: SheetKindSource, sheet: number): boolean =>
  sheetKindOf(source, sheet) === SheetKind.Worksheet;

/** Whether `sheet` carries a view; only a worksheet and a chartsheet do. */
export const sheetHasView = (source: SheetKindSource, sheet: number): boolean => {
  const kind = sheetKindOf(source, sheet);
  return kind === SheetKind.Worksheet || kind === SheetKind.Chartsheet;
};

/** Whether `sheet` can be zoomed: whether it carries a view. */
export const sheetZoomable = sheetHasView;
