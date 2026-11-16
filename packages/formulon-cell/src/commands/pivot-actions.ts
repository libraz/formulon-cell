import { PivotAxis, type Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { SpreadsheetStore } from '../store/store.js';
import { authorizedWriteOr } from './authorized-write.js';
import type { History } from './history.js';
import type { InteractionOrigin, OperationEffect } from './interaction-policy.js';

/** Who runs a pivot change: the store whose controller decides it, and the entry point. */
export interface PivotAuthorization {
  readonly store: SpreadsheetStore;
  readonly origin: InteractionOrigin;
  readonly commandId: string;
}

/** The cells a report covers, or its anchor cell when it has not laid out yet. */
export function pivotReportRange(
  wb: WorkbookHandle,
  sheet: number,
  pivotIndex: number,
): Range | null {
  const bounds = wb.pivotBounds(sheet).find((entry) => entry.pivotIndex === pivotIndex)?.range;
  return bounds ? { ...bounds } : null;
}

/** Rows a report's filter block takes above it: each report filter plus a blank row, or none. */
export function pivotFilterBlockRows(definition: {
  readonly fields: readonly { readonly axis: number | null }[];
}): number {
  const filters = definition.fields.filter((field) => field.axis === PivotAxis.Page).length;
  return filters > 0 ? filters + 1 : 0;
}

/**
 * Run a pivot change as one authorized step. `effects` are the report's range
 * and the destination of a new or moved report. A change that adds a sheet
 * (a report or move on a new sheet) also declares a workbook-wide `sheet`
 * intent, so workbook-structure protection refuses it. `write` records its
 * own undo entry, as before.
 */
export function pivotWrite<T>(
  auth: PivotAuthorization,
  history: History | null | undefined,
  effects: readonly OperationEffect[],
  refused: T,
  write: () => T,
  options: { readonly addsSheet?: boolean } = {},
): T {
  return authorizedWriteOr(
    { store: auth.store, history, origin: auth.origin },
    (origin) => [
      { operation: 'pivot', origin, commandId: auth.commandId, effects },
      ...(options.addsSheet
        ? [
            {
              operation: 'sheet' as const,
              origin,
              commandId: auth.commandId,
              effects: [{ kind: 'workbook' as const }],
            },
          ]
        : []),
    ],
    refused,
    write,
  );
}

/** The effects of a change to the report `pivotIndex` on `sheet`: its range, or the sheet. */
export function pivotReportEffects(
  wb: WorkbookHandle,
  sheet: number,
  pivotIndex: number,
): readonly OperationEffect[] {
  const range = pivotReportRange(wb, sheet, pivotIndex);
  return [range ? { kind: 'range', range } : { kind: 'sheet', sheet }];
}
