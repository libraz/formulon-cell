import { formatA1Cell } from '../engine/address.js';

/** Capability surfaced by the inline editor used by the pointer layer to
 *  detect a live formula edit and inject clicked cell references. */
export interface RangeInsertTarget {
  isFormulaEdit: () => boolean;
  insertRefAtCaret: (ref: string) => void;
}

const r1c1Axis = (prefix: 'R' | 'C', target: number, base: number): string => {
  const delta = target - base;
  return delta === 0 ? prefix : `${prefix}[${delta}]`;
};
const r1c1RefOf = (row: number, col: number, base: { row: number; col: number }): string =>
  `${r1c1Axis('R', row, base.row)}${r1c1Axis('C', col, base.col)}`;
export const refOf = (
  row: number,
  col: number,
  mode: { r1c1: boolean; base: { row: number; col: number } },
): string => (mode.r1c1 ? r1c1RefOf(row, col, mode.base) : formatA1Cell(row, col));
export const rangeRefOf = (
  a: { row: number; col: number },
  b: { row: number; col: number },
  mode: { r1c1: boolean; base: { row: number; col: number } },
): string => {
  if (a.row === b.row && a.col === b.col) return refOf(a.row, a.col, mode);
  const r0 = Math.min(a.row, b.row);
  const r1 = Math.max(a.row, b.row);
  const c0 = Math.min(a.col, b.col);
  const c1 = Math.max(a.col, b.col);
  return `${refOf(r0, c0, mode)}:${refOf(r1, c1, mode)}`;
};
