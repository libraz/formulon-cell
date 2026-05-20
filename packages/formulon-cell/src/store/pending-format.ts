import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { CellFormat, State } from './types.js';

export const sameAddr = (a: Addr, b: Addr): boolean =>
  a.sheet === b.sheet && a.row === b.row && a.col === b.col;

export function formatWithPending(state: State, addr: Addr): CellFormat | undefined {
  const stored = state.format.formats.get(addrKey(addr));
  const pending = state.ui.pendingFormat;
  if (!pending || !sameAddr(pending.addr, addr)) return stored;
  return { ...(stored ?? {}), ...pending.format };
}
