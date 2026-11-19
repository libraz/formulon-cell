// The formula bar's two write paths: a selection fill and a single-cell entry.
// They decide and write; the bar maps the result onto its UI.
import { type CellInputCommitResult, commitCellInput } from '../commands/cell-input-commit.js';
import { interactionControllerFor } from '../commands/interaction-controller.js';
import { localizedDecisionReason } from '../commands/permission-reasons.js';
import type { Addr } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';
import {
  buildSelectionInputBatch,
  SELECTION_INPUT_LIMIT_REFUSAL,
  writeSelectionInput,
} from '../interact/selection-input.js';
import type { SpreadsheetStore, State } from '../store/store.js';

/** Outcome of a selection fill: refused with the text to show, or applied against the live workbook. */
export type FillSelectionResult =
  | { status: 'refused'; message: string; title?: string }
  | { status: 'applied'; wb: WorkbookHandle };

/** Fills the selection with the typed entry, through the interaction controller when one is registered. */
export function fillSelectionFromFormulaBar(
  store: SpreadsheetStore,
  wb: () => WorkbookHandle,
  state: State,
  raw: string,
  anchor: Addr,
  getStrings: () => Strings,
): FillSelectionResult {
  const batch = buildSelectionInputBatch(state, raw, anchor);
  if (!batch) {
    return {
      status: 'refused',
      message: localizedDecisionReason(SELECTION_INPUT_LIMIT_REFUSAL, getStrings()),
    };
  }
  const controller = interactionControllerFor(store);
  if (!controller) {
    // The direct path serves only hosts without a registered controller.
    const written = writeSelectionInput(wb(), store, state, raw, anchor);
    if (written.status === 'limitExceeded') {
      return {
        status: 'refused',
        message: localizedDecisionReason(SELECTION_INPUT_LIMIT_REFUSAL, getStrings()),
      };
    }
    if (written.status === 'rejected') {
      return {
        status: 'refused',
        message: written.outcome.message,
        title: written.outcome.title,
      };
    }
    return { status: 'applied', wb: wb() };
  }

  let result: ReturnType<typeof controller.execute>;
  try {
    result = controller.execute({
      type: 'cellBatch',
      operation: batch.operation,
      origin: 'formulaBar',
      changes: batch.changes,
    });
  } catch (err) {
    console.warn('formulon-cell: formula-bar selection fill failed', err);
    return {
      status: 'refused',
      message: localizedDecisionReason(
        { code: 'invalid', reasonKey: 'cellsNotFilled' },
        getStrings(),
      ),
    };
  }
  if (result.status === 'refused') {
    const [rejection] = result.rejections;
    return {
      status: 'refused',
      message: rejection
        ? localizedDecisionReason(rejection, getStrings())
        : localizedDecisionReason({ code: 'operationDenied' }, getStrings()),
    };
  }
  return { status: 'applied', wb: wb() };
}

/** Writes one typed entry at `addr`; a thrown write is logged here and returned as `failed`. */
export function writeFormulaBarEntry(
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  addr: Addr,
  raw: string,
): CellInputCommitResult {
  const result = commitCellInput({ store, wb, addr, raw, origin: 'formulaBar', typed: true });
  if (result.status === 'failed') {
    console.warn('formulon-cell: formula-bar write failed', result.error);
  }
  return result;
}
