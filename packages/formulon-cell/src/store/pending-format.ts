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

export function formatWithPending(state: State, addr: Addr): CellFormat | undefined {
  const targetKey = formatTargetKey(state, addr);
  const stored = state.format.formats.get(targetKey);
  const sheetFont = sheetFontPatchFor(state, addr.sheet);
  const pending = state.ui.pendingFormat;
  const pendingFormat =
    pending && formatTargetKey(state, pending.addr) === targetKey ? pending.format : undefined;
  if (!stored && !sheetFont && !pendingFormat) return undefined;
  // Worksheet defaults are the implicit base. A direct cell format and a
  // pending editor format take precedence over that base when they actually
  // carry the same font axis.
  const next = { ...(sheetFont ?? {}), ...(stored ?? {}), ...(pendingFormat ?? {}) };
  // `undefined` means that an axis was cleared, so it must not mask a
  // worksheet-wide implicit font default (or make the effective projection
  // look like it carries an explicit axis).
  if (next.fontFamily === undefined) delete next.fontFamily;
  if (next.fontSize === undefined) delete next.fontSize;
  return next;
}
