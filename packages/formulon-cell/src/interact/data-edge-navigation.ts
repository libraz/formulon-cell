import { formatHasContent, isPopulated } from '../commands/current-region.js';
import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { State } from '../store/store.js';

const MIN_JUMP_EDGE_GAP_PROBES = 32;

const clamp = (a: Addr, row: number, col: number): Addr => ({
  sheet: a.sheet,
  row: Math.max(0, Math.min(MAX_ROW, row)),
  col: Math.max(0, Math.min(MAX_COL, col)),
});

function sparsePopulatedOnLine(
  s: State,
  a: Addr,
  row: number,
  col: number,
  dRow: number,
  dCol: number,
): Addr | null {
  const start = dRow !== 0 ? row : col;
  const forward = dRow !== 0 ? dRow > 0 : dCol > 0;
  let nearest: number | null = null;
  const consider = (key: string): void => {
    const addr = parseAddrKey(key);
    if (
      !addr ||
      addrKey(addr) !== key ||
      addr.sheet !== a.sheet ||
      addr.row < 0 ||
      addr.row > MAX_ROW ||
      addr.col < 0 ||
      addr.col > MAX_COL
    ) {
      return;
    }
    if (dRow !== 0 ? addr.col !== a.col : addr.row !== a.row) return;

    const position = dRow !== 0 ? addr.row : addr.col;
    if (forward ? position < start : position > start) return;
    if (nearest === null || (forward ? position < nearest : position > nearest)) {
      nearest = position;
    }
  };

  for (const key of s.data.cells.keys()) consider(key);
  for (const [key, format] of s.format.formats.entries()) {
    if (formatHasContent(format)) consider(key);
  }

  if (nearest === null) return null;
  return clamp(a, dRow !== 0 ? nearest : a.row, dCol !== 0 ? nearest : a.col);
}

/**
 * Spreadsheet-style Ctrl+Arrow jump. Three behaviors based on neighbor state:
 *  - empty origin            → jump to next populated cell (or sheet edge)
 *  - origin populated, adjacent populated → jump to last consecutive populated cell
 *  - origin populated, adjacent empty     → jump past the gap to next populated cell
 */
export function jumpToDataEdge(s: State, a: Addr, dRow: number, dCol: number): Addr {
  const sheet = a.sheet;
  const limitRow = dRow > 0 ? MAX_ROW : 0;
  const limitCol = dCol > 0 ? MAX_COL : 0;
  // Step 1: probe one step ahead.
  let r = a.row + dRow;
  let c = a.col + dCol;
  // Already at the edge — clamp.
  if (r < 0 || r > MAX_ROW || c < 0 || c > MAX_COL) return a;

  const originPop = isPopulated(s, sheet, a.row, a.col);
  const nextPop = isPopulated(s, sheet, r, c);

  if (originPop && nextPop) {
    // Walk until next cell is empty.
    while (
      r + dRow >= 0 &&
      r + dRow <= MAX_ROW &&
      c + dCol >= 0 &&
      c + dCol <= MAX_COL &&
      isPopulated(s, sheet, r + dRow, c + dCol)
    ) {
      r += dRow;
      c += dCol;
    }
    return clamp(a, r, c);
  }
  // Either origin is blank, or origin is populated but adjacent is blank.
  // Let moderate gaps in a larger map stay on direct probes before scanning
  // sparse model content. Reaching the sheet edge still stops the loop early.
  const sparseScanAfter = Math.max(
    MIN_JUMP_EDGE_GAP_PROBES,
    s.data.cells.size + s.format.formats.size,
  );
  let probes = 0;
  while (r >= 0 && r <= MAX_ROW && c >= 0 && c <= MAX_COL) {
    if (isPopulated(s, sheet, r, c)) return clamp(a, r, c);
    if ((dRow !== 0 && r === limitRow) || (dCol !== 0 && c === limitCol)) break;
    r += dRow;
    c += dCol;
    probes += 1;
    if (probes === sparseScanAfter) {
      const sparseTarget = sparsePopulatedOnLine(s, a, r, c, dRow, dCol);
      if (sparseTarget) return sparseTarget;
      break;
    }
  }
  // No populated cell found — go to sheet edge in that direction.
  return clamp(a, dRow !== 0 ? limitRow : a.row, dCol !== 0 ? limitCol : a.col);
}
