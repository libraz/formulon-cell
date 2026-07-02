import { addrKey } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';
import { isCellWritable, warnProtected } from './protection.js';

export interface HyperlinkEntry {
  addr: Addr;
  target: string;
  display?: string;
  tooltip?: string;
}

export interface SetHyperlinkOptions {
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
    const entry: HyperlinkEntry = { addr: { sheet: s, row, col }, target: fmt.hyperlink };
    if (fmt.hyperlinkDisplay) entry.display = fmt.hyperlinkDisplay;
    if (fmt.hyperlinkTooltip) entry.tooltip = fmt.hyperlinkTooltip;
    out.push(entry);
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
  options: SetHyperlinkOptions = {},
): void {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return;
  }
  const range = { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col };
  const next = target.trim();
  const prev = store.getState().format.formats.get(addrKey(addr));
  const display = options.display ?? prev?.hyperlinkDisplay;
  const tooltip = options.tooltip ?? prev?.hyperlinkTooltip;
  mutators.setRangeFormat(store, range, {
    hyperlink: next.length > 0 ? next : undefined,
    hyperlinkDisplay: next.length > 0 ? display : undefined,
    hyperlinkTooltip: next.length > 0 ? tooltip : undefined,
  });
  if (workbook) flushFormatToEngine(workbook, store, addr.sheet);
}

export function clearHyperlink(
  store: SpreadsheetStore,
  addr: Addr,
  workbook?: WorkbookHandle,
): void {
  if (!isCellWritable(store.getState(), addr)) {
    warnProtected(addr);
    return;
  }
  if (hyperlinkAt(store.getState(), addr) === null) return;
  const range = { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col };
  mutators.setRangeFormat(store, range, {
    hyperlink: undefined,
    hyperlinkDisplay: undefined,
    hyperlinkTooltip: undefined,
  });
  if (workbook) flushFormatToEngine(workbook, store, addr.sheet);
}
