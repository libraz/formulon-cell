import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

export interface HyperlinkEntry {
  addr: Addr;
  target: string;
  display?: string;
  tooltip?: string;
}

export function hyperlinkAt(state: State, addr: Addr): string | null {
  const target = state.format.formats.get(addrKey(addr))?.hyperlink;
  return target && target.length > 0 ? target : null;
}

export function listHyperlinks(state: State, sheet = state.data.sheetIndex): HyperlinkEntry[] {
  const out: HyperlinkEntry[] = [];
  for (const [key, fmt] of state.format.formats) {
    if (!fmt.hyperlink) continue;
    const parts = key.split(':').map((n) => Number(n));
    const s = parts[0] ?? -1;
    const row = parts[1] ?? -1;
    const col = parts[2] ?? -1;
    if (s !== sheet) continue;
    out.push({ addr: { sheet: s, row, col }, target: fmt.hyperlink });
  }
  return out.sort((a, b) => a.addr.row - b.addr.row || a.addr.col - b.addr.col);
}

export function listEngineHyperlinks(workbook: WorkbookHandle, sheet: number): HyperlinkEntry[] {
  return workbook.getHyperlinks(sheet).map((h) => ({
    addr: { sheet, row: h.row, col: h.col },
    target: h.target,
    display: h.display,
    tooltip: h.tooltip,
  }));
}

export function setHyperlink(
  store: SpreadsheetStore,
  addr: Addr,
  target: string,
  workbook?: WorkbookHandle,
): void {
  const range = { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col };
  const next = target.trim();
  mutators.setRangeFormat(store, range, { hyperlink: next.length > 0 ? next : undefined });
  if (workbook) flushFormatToEngine(workbook, store, addr.sheet);
}

export function clearHyperlink(
  store: SpreadsheetStore,
  addr: Addr,
  workbook?: WorkbookHandle,
): void {
  const range = { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col };
  mutators.setRangeFormat(store, range, { hyperlink: undefined });
  if (workbook) flushFormatToEngine(workbook, store, addr.sheet);
}
