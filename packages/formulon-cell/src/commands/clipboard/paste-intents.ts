import { MAX_COL, MAX_ROW } from '../../engine/address.js';
import type { Addr, Range } from '../../engine/types.js';
import { isWholeColumnRange, isWholeRowRange } from '../../store/selection-geometry.js';
import type { State } from '../../store/store.js';
import { cellShiftBand, columnShiftBand, rowShiftBand } from '../interaction-effects.js';
import type { InteractionOrigin, OperationEffect, OperationIntent } from '../interaction-policy.js';
import { isSheetProtected } from '../protection.js';
import { tablesGrownByPaste } from './table-extension.js';

/**
 * The intents of a paste: one `paste` over the whole destination and, for a cut, its source,
 * plus a `table` change for each table the paste grows. A protected sheet's tables never grow.
 */
export function pasteIntents(
  state: State,
  origin: InteractionOrigin,
  commandId: string,
  destination: Range | null,
  cutSource: Range | null = null,
): OperationIntent[] {
  const effects = [destination, cutSource]
    .filter((range): range is Range => range !== null)
    .map((range): OperationEffect => ({ kind: 'range', range }));
  const grown =
    destination && !isSheetProtected(state, destination.sheet)
      ? tablesGrownByPaste(state, destination)
      : [];
  return [
    ...(effects.length > 0 ? [{ operation: 'paste' as const, origin, commandId, effects }] : []),
    ...grown.map(
      ({ range }): OperationIntent => ({
        operation: 'table',
        origin,
        commandId,
        effects: [{ kind: 'range', range }],
      }),
    ),
  ];
}

type ShiftOperation = 'insertRows' | 'insertColumns' | 'deleteRows' | 'deleteColumns';

const shiftIntent = (
  operation: ShiftOperation,
  origin: InteractionOrigin,
  commandId: string,
  range: Range,
): OperationIntent => ({ operation, origin, commandId, effects: [{ kind: 'range', range }] });

/**
 * The intents of Insert Copied (Cut) Cells of a whole-row (column) copy `source` at `at`: as many
 * lines inserted, a cut's source lines deleted, and the copy pasted into the new lines. Null when
 * `source` is not a band of whole rows or whole columns.
 */
export function insertCopiedBandIntents(
  state: State,
  origin: InteractionOrigin,
  commandId: string,
  source: Range,
  cut: boolean,
  at: Addr,
): OperationIntent[] | null {
  const rows = isWholeRowRange(source);
  if (rows === isWholeColumnRange(source)) return null;
  const count = rows ? source.r1 - source.r0 + 1 : source.c1 - source.c0 + 1;
  const first = rows ? at.row : at.col;
  const destination: Range = rows
    ? { sheet: at.sheet, r0: first, c0: 0, r1: first + count - 1, c1: MAX_COL }
    : { sheet: at.sheet, r0: 0, c0: first, r1: MAX_ROW, c1: first + count - 1 };
  const shift = (sheet: number, from: number): Range =>
    rows ? rowShiftBand(sheet, from) : columnShiftBand(sheet, from);
  return [
    shiftIntent(rows ? 'insertRows' : 'insertColumns', origin, commandId, shift(at.sheet, first)),
    ...(cut
      ? [
          shiftIntent(
            rows ? 'deleteRows' : 'deleteColumns',
            origin,
            commandId,
            shift(source.sheet, rows ? source.r0 : source.c0),
          ),
        ]
      : []),
    ...pasteIntents(state, origin, commandId, destination, cut ? source : null),
  ];
}

/**
 * The intents of Insert Copied (Cut) Cells of a `rows` x `cols` block at `at`: the cells shifted
 * `direction`, a cut's source closed up the same way, and the block pasted.
 */
export function insertCopiedBlockIntents(
  state: State,
  origin: InteractionOrigin,
  commandId: string,
  size: { readonly rows: number; readonly cols: number },
  cutSource: Range | null,
  at: Addr,
  direction: 'down' | 'right',
): OperationIntent[] {
  const vertical = direction === 'down';
  const destination: Range = {
    sheet: at.sheet,
    r0: at.row,
    c0: at.col,
    r1: Math.min(MAX_ROW, at.row + size.rows - 1),
    c1: Math.min(MAX_COL, at.col + size.cols - 1),
  };
  return [
    shiftIntent(
      vertical ? 'insertRows' : 'insertColumns',
      origin,
      commandId,
      cellShiftBand(destination, vertical),
    ),
    ...(cutSource
      ? [
          shiftIntent(
            vertical ? 'deleteRows' : 'deleteColumns',
            origin,
            commandId,
            cellShiftBand(cutSource, vertical),
          ),
        ]
      : []),
    ...pasteIntents(state, origin, commandId, destination, cutSource),
  ];
}
