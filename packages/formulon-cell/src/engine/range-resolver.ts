import { formatCell } from './value.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** Resolve a literal string-array list source, or evaluate a range reference
 *  against the workbook to produce one string per non-blank cell. The returned
 *  values are de-duplicated in source order — Excel's list dropdown does the
 *  same for range-backed lists. Returns `[]` when the ref can't be parsed or
 *  the engine reports no values. */
export type RangeResolver = (ref: string) => string[];

/** Build a resolver bound to a workbook handle and a fallback sheet index
 *  (used when the ref omits a sheet prefix). */
export function makeRangeResolver(wb: WorkbookHandle, fallbackSheet: number): RangeResolver {
  return (ref: string): string[] => resolveRangeRef(wb, ref, fallbackSheet);
}

export function resolveRangeRef(wb: WorkbookHandle, ref: string, fallbackSheet: number): string[] {
  const parsed = parseRangeRef(ref);
  if (!parsed) return [];
  let sheet = fallbackSheet;
  if (parsed.sheetName !== null) {
    const idx = sheetIndexByName(wb, parsed.sheetName);
    if (idx < 0) return [];
    sheet = idx;
  }
  const seen = new Set<string>();
  const out: string[] = [];
  for (let row = parsed.r0; row <= parsed.r1; row += 1) {
    for (let col = parsed.c0; col <= parsed.c1; col += 1) {
      const value = wb.getValue({ sheet, row, col });
      if (value.kind === 'blank') continue;
      const text = formatCell(value);
      if (!text || seen.has(text)) continue;
      seen.add(text);
      out.push(text);
    }
  }
  return out;
}

interface ParsedRangeRef {
  sheetName: string | null;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

const ATOM_RE = /^\$?([A-Za-z]+)\$?(\d+)$/;

/** Parse `Sheet1!$A$1:$B$5`, `'Sheet 1'!A1:B5`, `A1:B5`, `A1`, `$A$1`. */
export function parseRangeRef(raw: string): ParsedRangeRef | null {
  const trimmed = raw.trim().replace(/^=/, '');
  if (!trimmed) return null;
  let body = trimmed;
  let sheetName: string | null = null;
  const bang = trimmed.indexOf('!');
  if (bang !== -1) {
    let prefix = trimmed.slice(0, bang);
    body = trimmed.slice(bang + 1);
    if (prefix.startsWith("'") && prefix.endsWith("'")) {
      prefix = prefix.slice(1, -1).replace(/''/g, "'");
    }
    if (!prefix) return null;
    sheetName = prefix;
  }
  const parts = body.split(':');
  if (parts.length === 0 || parts.length > 2) return null;
  const head = parseAtom(parts[0] ?? '');
  if (!head) return null;
  const tail = parts.length === 2 ? parseAtom(parts[1] ?? '') : head;
  if (!tail) return null;
  return {
    sheetName,
    r0: Math.min(head.row, tail.row),
    c0: Math.min(head.col, tail.col),
    r1: Math.max(head.row, tail.row),
    c1: Math.max(head.col, tail.col),
  };
}

function parseAtom(raw: string): { row: number; col: number } | null {
  const m = ATOM_RE.exec(raw.trim());
  if (!m) return null;
  const letters = m[1] ?? '';
  const digits = m[2] ?? '';
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  }
  col -= 1;
  const row = Number.parseInt(digits, 10) - 1;
  if (col < 0 || row < 0 || col > 16383 || row > 1048575) return null;
  return { row, col };
}

function sheetIndexByName(wb: WorkbookHandle, name: string): number {
  const n = wb.sheetCount;
  const target = name.toLowerCase();
  for (let i = 0; i < n; i += 1) {
    if (wb.sheetName(i).toLowerCase() === target) return i;
  }
  return -1;
}

/** True when a list source carries a range reference rather than literal values. */
export function isRangeSource(source: string[] | { ref: string }): source is { ref: string } {
  return !Array.isArray(source);
}
