import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import { sheetFontPatchFor } from '../engine/whole-sheet-font-sync.js';
import type { CellFormat, State } from './types.js';

export const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

const formatTargetKey = (state: State, addr: Addr): string =>
  state.merges.byCell.get(addrKey(addr)) ?? addrKey(addr);

export const sameFormatTarget = (state: State, a: Addr, b: Addr): boolean =>
  formatTargetKey(state, a) === formatTargetKey(state, b);

/** The cell's own format — a merged cell reads its merge anchor — over the
 *  worksheet-wide font default, or undefined when neither is set. */
export function directCellFormat(state: State, addr: Addr): CellFormat | undefined {
  const stored = state.format.formats.get(formatTargetKey(state, addr));
  const sheetFont = sheetFontPatchFor(state, addr.sheet);
  if (!stored && !sheetFont) return undefined;
  const next = { ...(sheetFont ?? {}), ...(stored ?? {}) };
  // An axis stored as `undefined` was cleared; drop the key so it does not
  // read as an explicit axis.
  if (next.fontFamily === undefined) delete next.fontFamily;
  if (next.fontSize === undefined) delete next.fontSize;
  return next;
}
