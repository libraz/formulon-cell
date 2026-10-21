import { addrKey, parseAddrKey } from '../engine/address.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';
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

interface SpanAnchor {
  sheet: number;
  row: number;
  col: number;
  fmt: CellFormat;
}

// Span anchors per formats map; the store replaces the map on every change.
const spanAnchorCache = new WeakMap<ReadonlyMap<string, CellFormat>, SpanAnchor[]>();

function spanAnchors(formats: ReadonlyMap<string, CellFormat>): SpanAnchor[] {
  let anchors = spanAnchorCache.get(formats);
  if (anchors) return anchors;
  anchors = [];
  for (const [key, fmt] of formats) {
    if (!fmt.hyperlinkSpan || !(fmt.hyperlink || fmt.hyperlinkLocation)) continue;
    const addr = parseAddrKey(key);
    if (addr) anchors.push({ ...addr, fmt });
  }
  spanAnchorCache.set(formats, anchors);
  return anchors;
}

/** The cell and format carrying the hyperlink that covers `addr`: the cell
 *  itself, or the anchor of a multi-cell link whose span contains it. */
export function hyperlinkCoverAt(state: State, addr: Addr): { addr: Addr; fmt: CellFormat } | null {
  const formats = state.format.formats;
  const own = formats.get(addrKey(addr));
  if (own && (own.hyperlink || own.hyperlinkLocation)) return { addr, fmt: own };
  for (const a of spanAnchors(formats)) {
    const span = a.fmt.hyperlinkSpan;
    if (
      span &&
      a.sheet === addr.sheet &&
      addr.row >= a.row &&
      addr.row <= span.lastRow &&
      addr.col >= a.col &&
      addr.col <= span.lastCol
    ) {
      return { addr: { sheet: a.sheet, row: a.row, col: a.col }, fmt: a.fmt };
    }
  }
  return null;
}

/** The link target covering `addr`. A location-only link reads as `#<location>`. */
export function hyperlinkAt(state: State, addr: Addr): string | null {
  const fmt = hyperlinkCoverAt(state, addr)?.fmt;
  if (!fmt) return null;
  if (fmt.hyperlink && fmt.hyperlink.length > 0) return fmt.hyperlink;
  return fmt.hyperlinkLocation ? `#${fmt.hyperlinkLocation}` : null;
}

/** Whether `url` may be followed from the UI: only http(s), mailto and tel,
 *  so a hostile cell value cannot smuggle a `javascript:` URL. */
export function isSafeHyperlinkTarget(url: string): boolean {
  const lower = url.trim().toLowerCase();
  return (
    lower.startsWith('http://') ||
    lower.startsWith('https://') ||
    lower.startsWith('mailto:') ||
    lower.startsWith('tel:')
  );
}

/** Whether `target` addresses a place in this document (`#Sheet2!A1`, `#Name`). */
export function isInDocumentHyperlinkTarget(target: string): boolean {
  const t = target.trim();
  return t.startsWith('#') && t.length > 1;
}

/** Whether the UI can follow `target`: a safe external URL or an in-document location. */
export function isFollowableHyperlinkTarget(target: string): boolean {
  return isSafeHyperlinkTarget(target) || isInDocumentHyperlinkTarget(target);
}

export function listHyperlinks(state: State, sheet = state.data.sheetIndex): HyperlinkEntry[] {
  const out: HyperlinkEntry[] = [];
  for (const [key, fmt] of state.format.formats) {
    if (!fmt.hyperlink) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    const entry: HyperlinkEntry = { addr, target: fmt.hyperlink };
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
  const keep = next.length > 0;
  // `#Sheet2!A1` addresses a place in the document; an unchanged target keeps
  // the location and span it was loaded with.
  const location = next === prev?.hyperlink ? prev.hyperlinkLocation : undefined;
  mutators.setRangeFormat(store, range, {
    hyperlink: keep ? next : undefined,
    hyperlinkLocation: keep ? (next.startsWith('#') ? next.slice(1) : location) : undefined,
    hyperlinkSpan: keep ? prev?.hyperlinkSpan : undefined,
    hyperlinkDisplay: keep ? display : undefined,
    hyperlinkTooltip: keep ? tooltip : undefined,
  });
  if (workbook) flushFormatToEngine(workbook, store, addr.sheet);
}

export function clearHyperlink(
  store: SpreadsheetStore,
  addr: Addr,
  workbook?: WorkbookHandle,
): void {
  const cover = hyperlinkCoverAt(store.getState(), addr);
  if (!cover) return;
  const anchor = cover.addr;
  if (!isCellWritable(store.getState(), anchor)) {
    warnProtected(anchor);
    return;
  }
  const range = {
    sheet: anchor.sheet,
    r0: anchor.row,
    c0: anchor.col,
    r1: anchor.row,
    c1: anchor.col,
  };
  mutators.setRangeFormat(store, range, {
    hyperlink: undefined,
    hyperlinkLocation: undefined,
    hyperlinkSpan: undefined,
    hyperlinkDisplay: undefined,
    hyperlinkTooltip: undefined,
  });
  if (workbook) flushFormatToEngine(workbook, store, anchor.sheet);
}
