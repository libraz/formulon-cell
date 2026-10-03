import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
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
  const pending = state.ui.pendingFormat;
  if (!pending || formatTargetKey(state, pending.addr) !== targetKey) return stored;
  return { ...(stored ?? {}), ...pending.format };
}
