import type { Addr, Range } from '../engine/types.js';
import { addrKey } from '../engine/workbook-handle.js';
import { mutators, type Sparkline, type SpreadsheetStore } from '../store/store.js';

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

export function setSparkline(store: SpreadsheetStore, addr: Addr, spec: Sparkline): void {
  mutators.setSparkline(store, addr, spec);
}

export function clearSparkline(store: SpreadsheetStore, addr: Addr): void {
  mutators.clearSparkline(store, addr);
}

export function clearSparklinesInRange(store: SpreadsheetStore, range: Range): void {
  mutators.clearSparklinesInRange(store, range);
}

const parseAddrKey = (key: string): Addr | null => {
  const parts = key.split(':');
  if (parts.length !== 3) return null;
  const sheetPart = parts[0];
  const rowPart = parts[1];
  const colPart = parts[2];
  if (sheetPart === undefined || rowPart === undefined || colPart === undefined) return null;
  const sheet = Number(sheetPart);
  const row = Number(rowPart);
  const col = Number(colPart);
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
};
