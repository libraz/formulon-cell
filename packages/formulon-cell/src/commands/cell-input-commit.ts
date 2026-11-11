// Single-cell input commit shared by the inline editor and the formula bar.
// A registered interaction controller owns the write whenever one exists, so
// both surfaces record the same undo entry and change batch; the direct
// workbook path serves only controller-less (standalone) hosts.

import { addrKey } from '../engine/address.js';
import { effectiveCellFormat } from '../engine/cell-format-projection.js';
import { flushFormatToEngine } from '../engine/cell-format-sync.js';
import type { Addr } from '../engine/types.js';
import { sheetFontPatchFor } from '../engine/whole-sheet-font-sync.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { formulaCallsTable } from '../engine/workbook-handle-data-tables.js';
import { directCellFormat } from '../store/direct-format.js';
import { validationAt } from '../store/range-validations.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { correctTypedEntry } from './autocorrect.js';
import { extendTablesForEntry } from './clipboard/table-extension.js';
import {
  coerceInputForCell,
  inputLocaleForWorkbook,
  validateCoercedInput,
  writeInputValidated,
} from './coerce-input.js';
import { dataTableGuardAlert, dataTableOccupies, tableFunctionAlert } from './data-table-guard.js';
import { interactionControllerFor, nativeAxisFontFormatFor } from './interaction-controller.js';
import { mergeAnchorFor, mergeCellsFor } from './interaction-effects.js';
import { pivotGuardAlert, pivotOccupies } from './pivot-guard.js';
import { isCellWritable, warnProtected } from './protection.js';

export type CellInputOperation = 'valueEdit' | 'formulaEdit';

/** Validation or policy alert for the committing surface to present. */
export interface CellInputAlert {
  severity: 'stop' | 'warning' | 'information';
  title?: string;
  message: string;
  /** Set when the presenting surface words the alert itself: `protectedCell`
   *  is a change to a locked cell of a protected sheet. */
  code?: 'protectedCell';
}

/** The alert raised by a change to a locked cell of a protected sheet. The
 *  message is a fallback; the alert dialog shows the localized wording. */
export const protectedCellAlert = (): CellInputAlert => ({
  severity: 'stop',
  code: 'protectedCell',
  message: 'cell is protected',
});

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
  /** The user typed `raw`, so AutoCorrect applies; a list pick or a conversion is entered as is. */
  typed?: boolean;
}

/** Batch operation a typed entry performs; Text-formatted cells keep `=` literal. */
export const inputOperation = (raw: string, forceText = false): CellInputOperation =>
  !forceText && raw.trimStart().startsWith('=') ? 'formulaEdit' : 'valueEdit';

/** Write one typed entry to the active cell. */
export function commitCellInput({
  store,
  wb,
  addr,
  raw: entered,
  origin,
  typed = false,
}: CellInputCommit): CellInputCommitResult {
  const raw = typed ? correctTypedEntry(store, entered) : entered;
  const state = store.getState();
  const anchor = mergeAnchorFor(state, addr);
  const fmt = directCellFormat(state, anchor);
  const ruleValidation = fmt?.validation ?? validationAt(state.format, anchor);
  let operation = inputOperation(raw, fmt?.numFmt?.kind === 'text');
  const controller = interactionControllerFor(store);
  let inputLocale: ReturnType<typeof inputLocaleForWorkbook>;
  let directInputAxis: ReturnType<typeof nativeAxisFontFormatFor>;
  let directInputFont: { fontFamily?: string; fontSize?: number } | undefined;
  let notice: CellInputAlert | null = null;
  try {
    // These reads cross the native adapter boundary and can throw when a
    // caller races a commit with WorkbookHandle.dispose(). Keep them inside
    // the same failure boundary as the write so the editor retains its raw
    // caret/input state instead of escaping through the preflight.
    const effectiveFormat = effectiveCellFormat(wb, state, anchor) ?? fmt;
    operation = inputOperation(raw, effectiveFormat?.numFmt?.kind === 'text');
    inputLocale = inputLocaleForWorkbook(wb);
    directInputAxis = controller ? undefined : nativeAxisFontFormatFor(wb, anchor);
    directInputFont = controller
      ? undefined
      : (() => {
          const patch = sheetFontPatchFor(state, anchor.sheet);
          // A saved row/column XF is pinned back after the value write. Keep
          // that raw axis implicit; materialising its paint projection as a
          // store format would turn a future inherited cell into a direct XF.
          if (!patch || directInputAxis) return undefined;
          const candidate = patch;
          const direct = state.format.formats.get(addrKey(anchor));
          const inherited: { fontFamily?: string; fontSize?: number } = {};
          if (candidate.fontFamily !== undefined && direct?.fontFamily === undefined) {
            inherited.fontFamily = candidate.fontFamily;
          }
          if (candidate.fontSize !== undefined && direct?.fontSize === undefined) {
            inherited.fontSize = candidate.fontSize;
          }
          return Object.keys(inherited).length > 0 ? inherited : undefined;
        })();
    if (
      pivotOccupies(wb, {
        sheet: anchor.sheet,
        r0: anchor.row,
        c0: anchor.col,
        r1: anchor.row,
        c1: anchor.col,
      })
    ) {
      return { status: 'rejected', operation, alert: pivotGuardAlert(wb) };
    }
    if (
      dataTableOccupies(wb, {
        sheet: anchor.sheet,
        r0: anchor.row,
        c0: anchor.col,
        r1: anchor.row,
        c1: anchor.col,
      })
    ) {
      return { status: 'rejected', operation, alert: dataTableGuardAlert(wb) };
    }
    const typed = coerceInputForCell(state, anchor, raw, inputLocale, wb);
    if (typed.kind === 'formula' && formulaCallsTable(typed.text)) {
      return { status: 'rejected', operation, alert: tableFunctionAlert(wb) };
    }
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
        alert: protectedCellAlert(),
      };
    }
    if (controller) {
      // The controller blocks `stop` rules itself but reports neither the
      // rule's title nor non-blocking outcomes, so validate up front.
      const validation = ruleValidation;
      const outcome = validation
        ? validateCoercedInput(
            wb,
            anchor,
            coerceInputForCell(state, anchor, raw, inputLocale, wb),
            validation,
          )
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
      });
      if (result.status === 'refused') {
        const reason = result.rejections[0]?.reason;
        return {
          status: 'rejected',
          operation,
          alert: reason ? { severity: 'stop', message: reason } : null,
        };
      }
      if (raw !== '') extendTablesForEntry(store, wb, controller.history, anchor);
      return { status: 'applied', notice };
    }
    const outcome = writeInputValidated(wb, anchor, raw, ruleValidation, store, inputLocale);
    if (!outcome.ok) {
      const alert = {
        severity: outcome.severity,
        title: ruleValidation?.errorTitle,
        message: outcome.message,
      };
      if (outcome.severity === 'stop') return { status: 'rejected', operation, alert };
      notice = alert;
    }
  } catch (error) {
    return { status: 'failed', operation, error };
  }
  if (directInputAxis) {
    try {
      if (!wb.setCellXfIndex(anchor.sheet, anchor.row, anchor.col, directInputAxis.xfIndex)) {
        return {
          status: 'failed',
          operation,
          error: new Error(`could not preserve inherited font XF at ${addrKey(anchor)}`),
        };
      }
    } catch (error) {
      return { status: 'failed', operation, error };
    }
  }
  if (directInputFont) {
    const currentFormat = store.getState().format.formats.get(addrKey(anchor));
    mutators.setCellFormat(store, anchor, { ...directInputFont, ...(currentFormat ?? {}) });
  }
  if (directInputAxis) flushFormatToEngine(wb, store, anchor.sheet);
  if (raw !== '') extendTablesForEntry(store, wb, null, anchor);
  return { status: 'applied', notice };
}
