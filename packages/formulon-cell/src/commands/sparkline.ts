import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import { mutators, type Sparkline, type SpreadsheetStore } from '../store/store.js';
import type { History } from './history.js';
import { isCellWritable, warnProtected } from './protection.js';
import { recordSparklineChange } from './slice-history.js';

export interface SparklineEntry {
  addr: Addr;
  spec: Sparkline;
}

export function listSparklines(state: {
  sparkline: { sparklines: ReadonlyMap<string, Sparkline> };
}): readonly SparklineEntry[] {
  const out: SparklineEntry[] = [];
  for (const [key, spec] of state.sparkline.sparklines) {
    const addr = parseAddrKey(key);
    if (addr) out.push({ addr, spec: { ...spec } });
  }
  return out.sort(
    (a, b) => a.addr.sheet - b.addr.sheet || a.addr.row - b.addr.row || a.addr.col - b.addr.col,
  );
}

export function sparklineAt(
  state: { sparkline: { sparklines: ReadonlyMap<string, Sparkline> } },
  addr: Addr,
): Sparkline | null {
  const spec = state.sparkline.sparklines.get(addrKey(addr));
  return spec ? { ...spec } : null;
}

export function setSparkline(
  store: SpreadsheetStore,
  addr: Addr,
  spec: Sparkline,
  history: History | null = null,
): boolean {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return false;
  }
  recordSparklineChange(history, store, () => {
    mutators.setSparkline(store, addr, spec);
  });
  return true;
}

export function clearSparkline(
  store: SpreadsheetStore,
  addr: Addr,
  history: History | null = null,
): boolean {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return false;
  }
  recordSparklineChange(history, store, () => {
    mutators.clearSparkline(store, addr);
  });
  return true;
}

export function clearSparklinesInRange(
  store: SpreadsheetStore,
  range: Range,
  history: History | null = null,
): number {
  const state = store.getState();
  const targets = listSparklines(state)
    .filter(({ addr }) => rangeContainsAddr(range, addr))
    .map(({ addr }) => addr);
  if (targets.length === 0) return 0;
  const writable = targets.filter((addr) => {
    if (isCellWritable(state, addr)) return true;
    warnProtected(addr);
    return false;
  });
  if (writable.length === 0) return 0;
  recordSparklineChange(history, store, () => {
    for (const addr of writable) mutators.clearSparkline(store, addr);
  });
  return writable.length;
}
