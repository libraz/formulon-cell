import {
  coerceInputForCell,
  validateCoercedInput,
  writeCoerced,
  writeInputValidated,
} from '../commands/coerce-input.js';
import type { CellBatchOperation } from '../commands/interaction-policy.js';
import { mergeAt } from '../commands/merge.js';
import { isCellWritable } from '../commands/protection.js';
import { shiftFormulaRefs } from '../commands/refs.js';
import { addrKey, MAX_COL, MAX_ROW } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { formatWithPending, sameAddr } from '../store/pending-format.js';
import { rangeArea, subtractRange } from '../store/selection-geometry.js';
import { mutators, type SpreadsheetStore, type State } from '../store/store.js';

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
 * shifting relative formula refs from `anchor`. Atomic: a `stop` validation on
 * any writable value cell rejects the fill with nothing written. Protected
 * cells are skipped. On success the anchor's pending format is applied.
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

  // Preflight: plan every write and validate value targets before touching the workbook.
  const plan: { target: Addr; formula?: string }[] = [];
  for (const r of ranges) {
    for (let row = r.r0; row <= r.r1; row += 1) {
      for (let col = r.c0; col <= r.c1; col += 1) {
        const target = { sheet, row, col };
        const fmt = sameAddr(target, anchor)
          ? formatWithPending(store.getState(), target)
          : state.format.formats.get(addrKey(target));
        if (isFormula && fmt?.numFmt?.kind !== 'text') {
          // Formula results aren't run through DV (Excel validates the typed entry).
          plan.push({ target, formula: shiftFormulaRefs(raw, row - anchor.row, col - anchor.col) });
          continue;
        }
        plan.push({ target });
        if (!fmt?.validation || !isCellWritable(state, target)) continue;
        const outcome = validateCoercedInput(
          wb,
          target,
          coerceInputForCell(state, target, raw),
          fmt.validation,
        );
        if (!outcome.ok && outcome.severity === 'stop') {
          return {
            status: 'rejected',
            outcome: {
              severity: 'stop',
              title: fmt.validation.errorTitle,
              message: outcome.message,
            },
          };
        }
      }
    }
  }

  // One recalc for the whole fill instead of one per written cell.
  wb.withBatchedRecalc(() => {
    for (const { target, formula } of plan) {
      if (formula === undefined) {
        // Already validated; this applies protection, coercion and implicit formats.
        writeInputValidated(wb, target, raw, undefined, store);
        continue;
      }
      try {
        writeCoerced(wb, target, { kind: 'formula', text: formula });
      } catch (err) {
        console.warn('formulon-cell: writeCoerced failed', err);
      }
    }
  });
  const pending = store.getState().ui.pendingFormat;
  if (pending && sameAddr(pending.addr, anchor)) {
    mutators.setCellFormat(store, anchor, pending.format);
    mutators.setPendingFormat(store, null);
  }
  return { status: 'applied' };
}
