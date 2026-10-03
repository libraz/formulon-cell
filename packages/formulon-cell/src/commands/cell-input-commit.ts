// Single-cell input commit shared by the inline editor and the formula bar.
// A registered interaction controller owns the write whenever one exists, so
// both surfaces record the same undo entry and change batch; the direct
// workbook path serves only controller-less (standalone) hosts.

import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { formatWithPending, sameFormatTarget } from '../store/pending-format.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { coerceInputForCell, validateCoercedInput, writeInputValidated } from './coerce-input.js';
import { interactionControllerFor } from './interaction-controller.js';
import { mergeAnchorFor, mergeCellsFor } from './interaction-effects.js';
import { isCellWritable, warnProtected } from './protection.js';

export type CellInputOperation = 'valueEdit' | 'formulaEdit';

/** Validation or policy alert for the committing surface to present. */
export interface CellInputAlert {
  severity: 'stop' | 'warning' | 'information';
  title?: string;
  message: string;
}

/**
 * - `applied`: written; `notice` is a non-blocking validation alert to surface.
 * - `rejected`: nothing written; `alert` is null when the controller gave no reason.
 * - `failed`: the write threw.
 */
export type CellInputCommitResult =
  | { status: 'applied'; notice: CellInputAlert | null }
  | { status: 'rejected'; operation: CellInputOperation; alert: CellInputAlert | null }
  | { status: 'failed'; operation: CellInputOperation; error: unknown };

export interface CellInputCommit {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  addr: Addr;
  raw: string;
  origin: 'editor' | 'formulaBar';
}

/** Batch operation a typed entry performs; Text-formatted cells keep `=` literal. */
export const inputOperation = (raw: string, forceText = false): CellInputOperation =>
  !forceText && raw.trimStart().startsWith('=') ? 'formulaEdit' : 'valueEdit';

/** Write one typed entry and settle the cell's pending format. */
export function commitCellInput({
  store,
  wb,
  addr,
  raw,
  origin,
}: CellInputCommit): CellInputCommitResult {
  const state = store.getState();
  const anchor = mergeAnchorFor(state, addr);
  const fmt = formatWithPending(state, anchor);
  const operation = inputOperation(raw, fmt?.numFmt?.kind === 'text');
  const controller = interactionControllerFor(store);
  let notice: CellInputAlert | null = null;
  try {
    const mergeCells = mergeCellsFor(state, addr);
    if (mergeCells === null) {
      return {
        status: 'rejected',
        operation,
        alert: { severity: 'stop', message: 'merged cell range exceeds the limit' },
      };
    }
    const protectedCell = mergeCells.find((cell) => !isCellWritable(state, cell));
    if (protectedCell) {
      warnProtected(protectedCell);
      return {
        status: 'rejected',
        operation,
        alert: { severity: 'stop', message: 'cell is protected' },
      };
    }
    if (controller) {
      // The controller blocks `stop` rules itself but reports neither the
      // rule's title nor non-blocking outcomes, so validate up front.
      const validation = fmt?.validation;
      const outcome = validation
        ? validateCoercedInput(wb, anchor, coerceInputForCell(state, anchor, raw), validation)
        : { ok: true as const };
      if (!outcome.ok) {
        const alert = {
          severity: outcome.severity,
          title: validation?.errorTitle,
          message: outcome.message,
        };
        if (outcome.severity === 'stop') return { status: 'rejected', operation, alert };
        notice = alert;
      }
      const result = controller.execute({
        type: 'cellBatch',
        operation,
        origin,
        changes: [{ addr, input: raw }],
        denied: 'reject',
      });
      if (result.status === 'rejected') {
        const reason = result.rejected[0]?.reason;
        return {
          status: 'rejected',
          operation,
          alert: reason ? { severity: 'stop', message: reason } : null,
        };
      }
      // The controller already applied the pending format to the written cell.
      const currentState = store.getState();
      const currentPending = currentState.ui.pendingFormat;
      if (currentPending && sameFormatTarget(currentState, currentPending.addr, anchor))
        mutators.setPendingFormat(store, null);
      return { status: 'applied', notice };
    }
    const outcome = writeInputValidated(wb, anchor, raw, fmt?.validation, store);
    if (!outcome.ok) {
      const alert = {
        severity: outcome.severity,
        title: fmt?.validation?.errorTitle,
        message: outcome.message,
      };
      if (outcome.severity === 'stop') return { status: 'rejected', operation, alert };
      notice = alert;
    }
  } catch (error) {
    return { status: 'failed', operation, error };
  }
  const currentState = store.getState();
  const currentPending = currentState.ui.pendingFormat;
  if (currentPending && sameFormatTarget(currentState, currentPending.addr, anchor)) {
    mutators.setCellFormat(store, anchor, currentPending.format);
    mutators.setPendingFormat(store, null);
  }
  return { status: 'applied', notice };
}
