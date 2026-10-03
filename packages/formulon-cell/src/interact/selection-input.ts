import { writeCoerced, writeInputValidated } from '../commands/coerce-input.js';
import type { CellBatchOperation } from '../commands/interaction-policy.js';
import { mergeAt } from '../commands/merge.js';
import { shiftFormulaRefs } from '../commands/refs.js';
import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { formatWithPending, sameAddr } from '../store/pending-format.js';
import { rangeArea, subtractRange } from '../store/selection-geometry.js';
import { type CellFormat, mutators, type SpreadsheetStore, type State } from '../store/store.js';

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

interface SelectionInputRejection {
  severity: 'stop';
  title?: string;
  message: string;
}

/** Outcome of {@link writeSelectionInput}; `rejected` carries the blocking validation alert. */
export type SelectionInputWriteResult =
  | { status: 'applied' }
  | { status: 'limitExceeded' }
  | { status: 'rejected'; outcome: SelectionInputRejection };

/**
 * Write `raw` to every selected cell without an interaction controller,
 * shifting relative formula refs from `anchor`. A `stop` validation on any
 * value cell aborts the remaining writes; on success the anchor's pending
 * format is applied.
 */
export function writeSelectionInput(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  state: State,
  raw: string,
  anchor: Addr,
): SelectionInputWriteResult {
  const ranges = [state.selection.range, ...(state.selection.extraRanges ?? [])];
  const totalCells = ranges.reduce((sum, r) => sum + rangeArea(r), 0);
  if (totalCells > MAX_SELECTION_INPUT_CELLS) return { status: 'limitExceeded' };
  const sheet = state.data.sheetIndex;
  const isFormula = raw.startsWith('=');
  // Validated write that mirrors the anchor's stop-rejection handling. Returns
  // the alert when a `stop` rule blocked the entry (the whole fill aborts) so DV
  // bites on every filled cell, not just the anchor.
  const writeValidatedOrAbort = (
    target: Addr,
    text: string,
    fmt: CellFormat | undefined,
  ): SelectionInputRejection | null => {
    const outcome = writeInputValidated(wb, target, text, fmt?.validation, store);
    if (!outcome.ok && outcome.severity === 'stop') {
      return {
        severity: outcome.severity,
        title: fmt?.validation?.errorTitle,
        message: outcome.message,
      };
    }
    return null;
  };
  // One recalc for the whole fill instead of one per written cell.
  const rejection = wb.withBatchedRecalc((): SelectionInputRejection | null => {
    for (const r of ranges) {
      for (let row = r.r0; row <= r.r1; row += 1) {
        for (let col = r.c0; col <= r.c1; col += 1) {
          const target = { sheet, row, col };
          const fmt =
            target.sheet === anchor.sheet && target.row === anchor.row && target.col === anchor.col
              ? formatWithPending(store.getState(), target)
              : state.format.formats.get(addrKey(target));
          const forceText = fmt?.numFmt?.kind === 'text';
          if (isFormula && !forceText) {
            // Formula fill: relative refs shift by the paste offset. Formula
            // results aren't run through DV here (Excel validates the typed
            // entry, not the recomputed result).
            const shifted = shiftFormulaRefs(raw, row - anchor.row, col - anchor.col);
            try {
              writeCoerced(wb, target, { kind: 'formula', text: shifted });
            } catch (err) {
              console.warn('formulon-cell: writeCoerced failed', err);
            }
          } else {
            // Value fill (including text-formatted cells): every target validates
            // against its own rule via the store-aware coercion path.
            const rejected = writeValidatedOrAbort(target, raw, fmt);
            if (rejected) return rejected;
          }
        }
      }
    }
    return null;
  });
  if (rejection) return { status: 'rejected', outcome: rejection };
  const pending = store.getState().ui.pendingFormat;
  if (pending && sameAddr(pending.addr, anchor)) {
    mutators.setCellFormat(store, anchor, pending.format);
    mutators.setPendingFormat(store, null);
  }
  return { status: 'applied' };
}
