/** Border preset names and their per-cell expansion into format entries. */

import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { CellBorderSide, CellBorders } from '../store/store.js';
import type { SelectionFormatEntry } from './selection-format-plan.js';

export type BorderPreset =
  | 'none'
  | 'outline'
  | 'thickOutline'
  | 'all'
  | 'inside'
  | 'insideHorizontal'
  | 'insideVertical'
  | 'top'
  | 'bottom'
  | 'left'
  | 'right'
  | 'doubleBottom'
  | 'thickBottom'
  | 'topAndBottom'
  | 'topAndThickBottom'
  | 'topAndDoubleBottom'
  | 'diagonalDown'
  | 'diagonalUp';

type BorderSideName = keyof CellBorders;

export const borderEntriesForPlan = (
  plan: { ranges: readonly Range[]; cells: readonly Addr[] },
  preset: BorderPreset,
  side: CellBorderSide,
  thickSide: CellBorderSide,
  doubleSide: CellBorderSide,
): SelectionFormatEntry[] => {
  const entries = new Map<string, SelectionFormatEntry>();
  const add = (addr: Addr, name: BorderSideName, value: CellBorderSide): void => {
    const key = addrKey(addr);
    const current = entries.get(key);
    if (!current) {
      entries.set(key, { addr, patch: { borders: { [name]: value } } });
      return;
    }
    if (current.patch?.borders) current.patch.borders[name] = value;
  };

  for (const range of plan.ranges) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        if (preset === 'all') {
          add(addr, 'top', side);
          add(addr, 'right', side);
          add(addr, 'bottom', side);
          add(addr, 'left', side);
        } else if (preset === 'none') {
          add(addr, 'top', false);
          add(addr, 'right', false);
          add(addr, 'bottom', false);
          add(addr, 'left', false);
          add(addr, 'diagonalDown', false);
          add(addr, 'diagonalUp', false);
        } else if (preset === 'inside') {
          if (row > range.r0) add(addr, 'top', side);
          if (col > range.c0) add(addr, 'left', side);
        } else if (preset === 'insideHorizontal') {
          if (row > range.r0) add(addr, 'top', side);
        } else if (preset === 'insideVertical') {
          if (col > range.c0) add(addr, 'left', side);
        } else if (preset === 'diagonalDown') {
          add(addr, 'diagonalDown', side);
        } else if (preset === 'diagonalUp') {
          add(addr, 'diagonalUp', side);
        } else if (preset === 'outline') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', side);
          if (col === range.c0) add(addr, 'left', side);
          if (col === range.c1) add(addr, 'right', side);
        } else if (preset === 'thickOutline') {
          if (row === range.r0) add(addr, 'top', thickSide);
          if (row === range.r1) add(addr, 'bottom', thickSide);
          if (col === range.c0) add(addr, 'left', thickSide);
          if (col === range.c1) add(addr, 'right', thickSide);
        } else if (preset === 'top' && row === range.r0) {
          add(addr, 'top', side);
        } else if (preset === 'bottom' && row === range.r1) {
          add(addr, 'bottom', side);
        } else if (preset === 'left' && col === range.c0) {
          add(addr, 'left', side);
        } else if (preset === 'right' && col === range.c1) {
          add(addr, 'right', side);
        } else if (preset === 'doubleBottom' && row === range.r1) {
          add(addr, 'bottom', doubleSide);
        } else if (preset === 'thickBottom' && row === range.r1) {
          add(addr, 'bottom', thickSide);
        } else if (preset === 'topAndBottom') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', side);
        } else if (preset === 'topAndThickBottom') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', thickSide);
        } else if (preset === 'topAndDoubleBottom') {
          if (row === range.r0) add(addr, 'top', side);
          if (row === range.r1) add(addr, 'bottom', doubleSide);
        }
      }
    }
  }
  return [...entries.values()];
};
