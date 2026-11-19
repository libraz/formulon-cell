import type { Addr, CellValue } from '../engine/types.js';

/** Plain-text values above `beforeRow` in column `col` for the "pick from
 *  list" popover: text-only (formulas, numbers, blanks all skip), deduped,
 *  nearest-first, at most ten. Takes the sheet's populated cells in one pass
 *  rather than probing each row — per-row `cellFormula` lookups are O(n) and
 *  would run on every keystroke. */
export function pickListValues(
  cells: Iterable<{ addr: Addr; value: CellValue; formula: string | null }>,
  col: number,
  beforeRow: number,
): string[] {
  const hits: { row: number; text: string }[] = [];
  for (const e of cells) {
    if (e.addr.col !== col) continue;
    if (e.addr.row >= beforeRow) continue;
    // Formulas don't contribute — the pick-list is verbatim text only.
    if (e.formula !== null) continue;
    if (e.value.kind !== 'text') continue;
    const text = e.value.value;
    if (text.length === 0) continue;
    hits.push({ row: e.addr.row, text });
  }
  // Nearest-first: highest row index wins.
  hits.sort((a, b) => b.row - a.row);
  const out: string[] = [];
  const seen = new Set<string>();
  for (const h of hits) {
    if (seen.has(h.text)) continue;
    seen.add(h.text);
    out.push(h.text);
    if (out.length >= 10) break;
  }
  return out;
}

/** Pick From Drop-down List: the text entries of the column's block of filled cells around
 *  `row` (stopping at the first blank, the active cell left out), folded case-insensitively with
 *  the topmost spelling kept, sorted. Numbers and formulas are left out. */
export function pickFromListValues(
  cells: Iterable<{ addr: Addr; value: CellValue; formula: string | null }>,
  col: number,
  row: number,
): string[] {
  const byRow = new Map<number, { value: CellValue; formula: string | null }>();
  for (const e of cells) {
    if (e.addr.col === col && e.value.kind !== 'blank') byRow.set(e.addr.row, e);
  }
  let top = row;
  while (top > 0 && byRow.has(top - 1)) top -= 1;
  let bottom = row;
  while (byRow.has(bottom + 1)) bottom += 1;
  const byFold = new Map<string, string>();
  for (let r = top; r <= bottom; r += 1) {
    const entry = r === row ? undefined : byRow.get(r);
    if (!entry || entry.formula !== null || entry.value.kind !== 'text') continue;
    const text = entry.value.value;
    const fold = text.toLocaleLowerCase();
    if (text.length > 0 && !byFold.has(fold)) byFold.set(fold, text);
  }
  return [...byFold.values()].sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: 'base' }),
  );
}
