import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { addrKey } from '../engine/workbook-handle.js';
import type { State } from '../store/store.js';
import { shiftFormulaRefs } from './refs.js';

/**
 * Fill direction inferred from the relationship between source and dest.
 * Diagonal extensions (drag bottom-right corner past both axes) split into
 * two passes: row direction first, then column.
 */
type FillDir = 'down' | 'up' | 'right' | 'left' | 'copy';

interface SourceCell {
  row: number;
  col: number;
  formula: string | null;
  numeric: number | null;
  text: string | null;
  bool: boolean | null;
  blank: boolean;
}

const numericFromText = (s: string): { prefix: string; n: number } | null => {
  const m = s.match(/^(.*?)(-?\d+)$/);
  if (!m) return null;
  const prefix = m[1] ?? '';
  const n = Number(m[2]);
  if (!Number.isFinite(n)) return null;
  return { prefix, n };
};

/** Custom-list series — Excel ships these by default. Each list represents
 *  a closed cycle that auto-fill steps through. Casing of the source is
 *  preserved by matching against the lowercased list. */
const CUSTOM_LISTS: readonly string[][] = [
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ],
  ['日', '月', '火', '水', '木', '金', '土'],
  ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'],
  ['Q1', 'Q2', 'Q3', 'Q4'],
  ['第1四半期', '第2四半期', '第3四半期', '第4四半期'],
];

/** Find the custom list (if any) that contains every source value. Returns
 *  the list and the list-indices of each source cell, or null when no single
 *  list matches all of them. */
function matchCustomList(values: string[]): { list: readonly string[]; indices: number[] } | null {
  for (const list of CUSTOM_LISTS) {
    const indices: number[] = [];
    let ok = true;
    for (const v of values) {
      const idx = list.findIndex((item) => item.toLowerCase() === v.toLowerCase());
      if (idx < 0) {
        ok = false;
        break;
      }
      indices.push(idx);
    }
    if (ok) return { list, indices };
  }
  return null;
}

function readSource(state: State, src: Range): SourceCell[][] {
  const sheet = src.sheet;
  const out: SourceCell[][] = [];
  for (let r = src.r0; r <= src.r1; r += 1) {
    const row: SourceCell[] = [];
    for (let c = src.c0; c <= src.c1; c += 1) {
      const cell = state.data.cells.get(addrKey({ sheet, row: r, col: c }));
      if (!cell) {
        row.push({
          row: r,
          col: c,
          formula: null,
          numeric: null,
          text: null,
          bool: null,
          blank: true,
        });
        continue;
      }
      const v = cell.value;
      row.push({
        row: r,
        col: c,
        formula: cell.formula,
        numeric: v.kind === 'number' ? v.value : null,
        text: v.kind === 'text' ? v.value : null,
        bool: v.kind === 'bool' ? v.value : null,
        blank: v.kind === 'blank' && !cell.formula,
      });
    }
    out.push(row);
  }
  return out;
}

function detectDirection(src: Range, dest: Range): FillDir {
  if (dest.r1 > src.r1 && dest.r0 === src.r0 && dest.c0 === src.c0 && dest.c1 === src.c1) {
    return 'down';
  }
  if (dest.r0 < src.r0 && dest.r1 === src.r1 && dest.c0 === src.c0 && dest.c1 === src.c1) {
    return 'up';
  }
  if (dest.c1 > src.c1 && dest.c0 === src.c0 && dest.r0 === src.r0 && dest.r1 === src.r1) {
    return 'right';
  }
  if (dest.c0 < src.c0 && dest.c1 === src.c1 && dest.r0 === src.r0 && dest.r1 === src.r1) {
    return 'left';
  }
  // Diagonal or arbitrary — treat as 2D copy/cycle.
  return 'copy';
}

interface SeriesProjection {
  /** Project the value at extension index `i` (1-based; 1 = first cell beyond
   *  source in the fill direction). Returns null when no value should be written. */
  at(
    i: number,
  ): { kind: 'number'; value: number } | { kind: 'text'; value: string } | { kind: 'blank' } | null;
}

/**
 * Inspect a 1D source line and produce a series projector. Excel's heuristic:
 *  - all-numeric, length >= 2 → linear extrapolation (step = avg consecutive diff)
 *  - all-numeric, length 1   → copy
 *  - "Item 1", "Item 2"      → increment trailing integer
 *  - "Item 1"                → copy
 *  - mixed                   → cycle
 */
function buildProjection(line: SourceCell[]): SeriesProjection {
  if (line.length === 0) {
    return { at: () => null };
  }
  // All numeric?
  if (line.every((c) => c.numeric !== null && c.formula === null)) {
    if (line.length === 1) {
      const v = line[0]?.numeric ?? 0;
      return { at: () => ({ kind: 'number', value: v }) };
    }
    let stepSum = 0;
    for (let i = 1; i < line.length; i += 1) {
      stepSum += (line[i]?.numeric ?? 0) - (line[i - 1]?.numeric ?? 0);
    }
    const step = stepSum / (line.length - 1);
    const last = line[line.length - 1]?.numeric ?? 0;
    return { at: (i) => ({ kind: 'number', value: last + step * i }) };
  }

  // Custom list match (e.g. Mon/Tue/Wed, Jan/Feb...) — preserves the casing
  // of the *list*, not the source.
  if (line.every((c) => c.text !== null && c.formula === null)) {
    const values = line.map((c) => c.text ?? '');
    const lm = matchCustomList(values);
    if (lm) {
      const lastIdx = lm.indices[lm.indices.length - 1] ?? 0;
      const step =
        lm.indices.length >= 2
          ? (lastIdx - (lm.indices[0] ?? 0)) / Math.max(1, lm.indices.length - 1)
          : 1;
      const stepInt = Math.round(step) || 1;
      return {
        at: (i) => {
          const target = lastIdx + stepInt * i;
          const len = lm.list.length;
          const idx = ((target % len) + len) % len;
          return { kind: 'text', value: lm.list[idx] ?? '' };
        },
      };
    }
  }

  // Trailing-integer text? Need every cell to have the same prefix and a numeric tail.
  if (line.every((c) => c.text !== null && c.formula === null)) {
    const parsed = line.map((c) => numericFromText(c.text ?? ''));
    if (parsed.every((p) => p !== null)) {
      const first = parsed[0];
      if (first && parsed.every((p) => p?.prefix === first.prefix)) {
        if (line.length === 1) {
          // Single cell: increment by 1 each step.
          return {
            at: (i) => ({ kind: 'text', value: `${first.prefix}${first.n + i}` }),
          };
        }
        let stepSum = 0;
        for (let i = 1; i < parsed.length; i += 1) {
          stepSum += (parsed[i]?.n ?? 0) - (parsed[i - 1]?.n ?? 0);
        }
        const step = stepSum / (parsed.length - 1);
        const last = parsed[parsed.length - 1]?.n ?? 0;
        return {
          at: (i) => ({ kind: 'text', value: `${first.prefix}${Math.round(last + step * i)}` }),
        };
      }
    }
  }

  // Cycle/copy.
  return {
    at: (i) => {
      const idx = (((i - 1) % line.length) + line.length) % line.length;
      const c = line[idx];
      if (!c) return null;
      if (c.numeric !== null) return { kind: 'number', value: c.numeric };
      if (c.text !== null) return { kind: 'text', value: c.text };
      if (c.bool !== null) return { kind: 'text', value: c.bool ? 'TRUE' : 'FALSE' };
      return { kind: 'blank' };
    },
  };
}

const writeProjected = (
  wb: WorkbookHandle,
  sheet: number,
  row: number,
  col: number,
  projected: ReturnType<SeriesProjection['at']>,
): void => {
  if (!projected) return;
  if (projected.kind === 'number') wb.setNumber({ sheet, row, col }, projected.value);
  else if (projected.kind === 'text') wb.setText({ sheet, row, col }, projected.value);
  else wb.setBlank({ sheet, row, col });
};

export interface FillOptions {
  /** Bypass series detection and tile the source verbatim (Excel: Ctrl-drag). */
  copyOnly?: boolean;
}

/**
 * Fill the cells in `dest` (which contains `src` as a sub-rect) by projecting
 * the source range outward. Returns true if the fill produced any writes.
 *
 * Formula handling: when the source contains formulas we currently copy them
 * verbatim (no relative-reference translation yet). For pure values, the
 * series detector picks linear / increment / copy as Excel would.
 */
export function fillRange(
  state: State,
  wb: WorkbookHandle,
  src: Range,
  dest: Range,
  opts?: FillOptions,
): boolean {
  if (dest.r0 === src.r0 && dest.r1 === src.r1 && dest.c0 === src.c0 && dest.c1 === src.c1) {
    return false;
  }
  const dir = opts?.copyOnly ? 'copy' : detectDirection(src, dest);
  const sheet = src.sheet;
  const source = readSource(state, src);

  // Per-cell formula tile with relative-ref shifting. Returns true when used.
  const tileFormula = (destRow: number, destCol: number, srcCell: SourceCell): boolean => {
    if (!srcCell.formula) return false;
    const shifted = shiftFormulaRefs(srcCell.formula, destRow - srcCell.row, destCol - srcCell.col);
    wb.setFormula({ sheet, row: destRow, col: destCol }, shifted);
    return true;
  };

  if (dir === 'down' || dir === 'up') {
    // Fill column-by-column. Each column gets its own projection.
    const cols = src.c1 - src.c0 + 1;
    for (let c = 0; c < cols; c += 1) {
      const line: SourceCell[] = source.map((row) => row[c] as SourceCell);
      const allFormula = line.length > 0 && line.every((cc) => cc.formula !== null);
      const proj = allFormula ? null : buildProjection(line);
      if (dir === 'down') {
        const ext = dest.r1 - src.r1;
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r1 + i;
          const destCol = src.c0 + c;
          if (allFormula) {
            const srcCell = line[(i - 1) % line.length] as SourceCell;
            tileFormula(destRow, destCol, srcCell);
          } else {
            writeProjected(wb, sheet, destRow, destCol, proj?.at(i) ?? null);
          }
        }
      } else {
        // Up: extension index 1 is the row just above src.r0.
        const ext = src.r0 - dest.r0;
        const reversed = [...line].reverse();
        const projUp = allFormula ? null : buildProjection(reversed);
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r0 - i;
          const destCol = src.c0 + c;
          if (allFormula) {
            const srcCell = reversed[(i - 1) % reversed.length] as SourceCell;
            tileFormula(destRow, destCol, srcCell);
          } else {
            writeProjected(wb, sheet, destRow, destCol, projUp?.at(i) ?? null);
          }
        }
      }
    }
    return true;
  }

  if (dir === 'right' || dir === 'left') {
    const rows = src.r1 - src.r0 + 1;
    for (let r = 0; r < rows; r += 1) {
      const line: SourceCell[] = source[r] ?? [];
      const allFormula = line.length > 0 && line.every((cc) => cc.formula !== null);
      const proj = allFormula ? null : buildProjection(line);
      if (dir === 'right') {
        const ext = dest.c1 - src.c1;
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r0 + r;
          const destCol = src.c1 + i;
          if (allFormula) {
            const srcCell = line[(i - 1) % line.length] as SourceCell;
            tileFormula(destRow, destCol, srcCell);
          } else {
            writeProjected(wb, sheet, destRow, destCol, proj?.at(i) ?? null);
          }
        }
      } else {
        const ext = src.c0 - dest.c0;
        const reversed = [...line].reverse();
        const projLeft = allFormula ? null : buildProjection(reversed);
        for (let i = 1; i <= ext; i += 1) {
          const destRow = src.r0 + r;
          const destCol = src.c0 - i;
          if (allFormula) {
            const srcCell = reversed[(i - 1) % reversed.length] as SourceCell;
            tileFormula(destRow, destCol, srcCell);
          } else {
            writeProjected(wb, sheet, destRow, destCol, projLeft?.at(i) ?? null);
          }
        }
      }
    }
    return true;
  }

  // 2D / arbitrary: tile source over dest.
  const sR = src.r1 - src.r0 + 1;
  const sC = src.c1 - src.c0 + 1;
  for (let r = dest.r0; r <= dest.r1; r += 1) {
    for (let c = dest.c0; c <= dest.c1; c += 1) {
      if (r >= src.r0 && r <= src.r1 && c >= src.c0 && c <= src.c1) continue; // skip source
      const sr = (((r - src.r0) % sR) + sR) % sR;
      const sc = (((c - src.c0) % sC) + sC) % sC;
      const cell = source[sr]?.[sc];
      if (!cell || cell.blank) {
        wb.setBlank({ sheet, row: r, col: c });
        continue;
      }
      if (cell.formula) {
        const shifted = shiftFormulaRefs(cell.formula, r - cell.row, c - cell.col);
        wb.setFormula({ sheet, row: r, col: c }, shifted);
      } else if (cell.numeric !== null) {
        wb.setNumber({ sheet, row: r, col: c }, cell.numeric);
      } else if (cell.text !== null) {
        wb.setText({ sheet, row: r, col: c }, cell.text);
      } else if (cell.bool !== null) {
        wb.setBool({ sheet, row: r, col: c }, cell.bool);
      }
    }
  }
  return true;
}

/** Compute the dest range for a drag from the source's bottom-right corner
 *  to a target cell (the cursor position). Excel locks the extension to
 *  whichever axis is most extended — diagonal drags don't extend both axes
 *  (unless the target is fully inside both). */
export function fillDestFor(src: Range, target: { row: number; col: number }): Range {
  // Decide whether to extend rows or cols based on which delta is larger.
  const dRow =
    target.row > src.r1 ? target.row - src.r1 : target.row < src.r0 ? src.r0 - target.row : 0;
  const dCol =
    target.col > src.c1 ? target.col - src.c1 : target.col < src.c0 ? src.c0 - target.col : 0;
  if (dRow === 0 && dCol === 0) return src;
  if (dRow >= dCol) {
    return target.row > src.r1
      ? { sheet: src.sheet, r0: src.r0, c0: src.c0, r1: target.row, c1: src.c1 }
      : { sheet: src.sheet, r0: target.row, c0: src.c0, r1: src.r1, c1: src.c1 };
  }
  return target.col > src.c1
    ? { sheet: src.sheet, r0: src.r0, c0: src.c0, r1: src.r1, c1: target.col }
    : { sheet: src.sheet, r0: src.r0, c0: target.col, r1: src.r1, c1: src.c1 };
}
