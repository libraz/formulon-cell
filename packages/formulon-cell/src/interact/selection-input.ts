import type { CellBatchOperation } from '../commands/interaction-policy.js';
import { mergeAt } from '../commands/merge.js';
import { shiftFormulaRefs } from '../commands/refs.js';
import { addrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { formatWithPending } from '../store/pending-format.js';
import { subtractRange } from '../store/selection-geometry.js';
import type { State } from '../store/store.js';

export interface SelectionInputChange {
  readonly addr: Addr;
  readonly input: string;
}

export interface SelectionInputBatch {
  readonly changes: readonly SelectionInputChange[];
  readonly operation: Extract<CellBatchOperation, 'valueEdit' | 'formulaEdit'>;
}

const MAX_SELECTION_INPUT_CELLS = 100_000;

/** Shown when {@link buildSelectionInputBatch} refuses a selection. */
export const SELECTION_INPUT_LIMIT_MESSAGE =
  'The selection cannot be filled together. A fill can include at most 100,000 unique cells.';
const MAX_ROW = 1_048_575;
const MAX_COL = 16_383;

const rangeArea = (range: Range): number => (range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1);

const validRange = (range: Range): boolean =>
  Number.isSafeInteger(range.sheet) &&
  range.sheet >= 0 &&
  Number.isSafeInteger(range.r0) &&
  Number.isSafeInteger(range.c0) &&
  Number.isSafeInteger(range.r1) &&
  Number.isSafeInteger(range.c1) &&
  range.r0 >= 0 &&
  range.c0 >= 0 &&
  range.r1 <= MAX_ROW &&
  range.c1 <= MAX_COL &&
  range.r0 <= range.r1 &&
  range.c0 <= range.c1 &&
  Number.isSafeInteger(rangeArea(range));

/**
 * Plan a bounded set of edits for a multi-cell selection. Overlapping ranges
 * are unioned before cells are materialized, merge bodies are omitted, and
 * relative formula references are shifted from the anchor cell.
 */
export function buildSelectionInputBatch(
  state: State,
  raw: string,
  anchor: Addr,
): SelectionInputBatch | null {
  const inputRanges = [state.selection.range, ...(state.selection.extraRanges ?? [])];
  const uniqueRanges: Range[] = [];
  let totalCells = 0;

  for (const range of inputRanges) {
    if (!validRange(range)) return null;
    let uncovered: Range[] = [{ ...range }];
    for (const existing of uniqueRanges) {
      uncovered = uncovered.flatMap((piece) => subtractRange(piece, existing));
      if (uncovered.length === 0) break;
    }
    const addedCells = uncovered.reduce((sum, piece) => sum + rangeArea(piece), 0);
    if (!Number.isSafeInteger(addedCells) || totalCells + addedCells > MAX_SELECTION_INPUT_CELLS) {
      return null;
    }
    totalCells += addedCells;
    uniqueRanges.push(...uncovered);
  }

  const changes: SelectionInputChange[] = [];
  const seen = new Set<string>();
  let hasFormula = false;
  const formulaRaw = raw.trim();
  const formula = formulaRaw.startsWith('=');

  for (const range of uniqueRanges) {
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        const key = addrKey(addr);
        if (seen.has(key)) continue;
        const merge = mergeAt(state, addr);
        if (merge && (merge.r0 !== row || merge.c0 !== col)) continue;

        seen.add(key);
        const format = formatWithPending(state, addr);
        const forceText = format?.numFmt?.kind === 'text';
        const isFormula = formula && !forceText;
        hasFormula ||= isFormula;
        changes.push({
          addr,
          input: isFormula ? shiftFormulaRefs(formulaRaw, row - anchor.row, col - anchor.col) : raw,
        });
      }
    }
  }

  if (changes.length === 0) return null;
  return {
    changes,
    operation: hasFormula ? 'formulaEdit' : 'valueEdit',
  };
}
