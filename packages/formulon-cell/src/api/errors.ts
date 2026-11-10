import { FormulonError, StatusCode } from '@libraz/formulon';

/**
 * Why formulon-cell refused a call that the engine itself would have accepted.
 *
 * - `disposed`: the object was disposed.
 * - `feature-disabled`: the feature the call needs is switched off in `ui.features`.
 * - `policy-denied`: the interaction policy denies the operation.
 * - `protected`: sheet or workbook protection blocks the operation.
 * - `limit-exceeded`: the input is above a formulon-cell size limit.
 * - `workbook-in-use`: the workbook is already mounted by another live instance.
 */
export type SpreadsheetFailure =
  | 'disposed'
  | 'feature-disabled'
  | 'policy-denied'
  | 'protected'
  | 'limit-exceeded'
  | 'workbook-in-use';

const FAILURES: ReadonlySet<string> = new Set<SpreadsheetFailure>([
  'disposed',
  'feature-disabled',
  'policy-denied',
  'protected',
  'limit-exceeded',
  'workbook-in-use',
]);

const FAILURE_PAIR = /(?:^|\s)failure=(\S+)/;

/**
 * Returns the formulon-cell failure kind a thrown error carries, or `null` for anything else (an engine
 * failure, a `TypeError`, a non-error value).
 *
 * @example
 * try { sheet.dialogs.open('pivotTable'); }
 * catch (e) { if (spreadsheetFailureOf(e) === 'feature-disabled') hideButton(); }
 */
export function spreadsheetFailureOf(error: unknown): SpreadsheetFailure | null {
  if (!(error instanceof FormulonError)) return null;
  const kind = FAILURE_PAIR.exec(error.context)?.[1];
  return kind !== undefined && FAILURES.has(kind) ? (kind as SpreadsheetFailure) : null;
}

/** Extra `key=value` pairs and the wrapped cause of a formulon-cell failure. */
export interface SpreadsheetFailureDetail {
  readonly context?: Readonly<Record<string, string | number>>;
  readonly cause?: unknown;
}

/**
 * Builds the `FormulonError` for a formulon-cell precondition failure: `BindingInvalidHandle` for
 * `disposed`, `PreconditionFailed` otherwise, with `failure=<kind>` first in `context`.
 */
export function spreadsheetFailure(
  failure: SpreadsheetFailure,
  message: string,
  detail: SpreadsheetFailureDetail = {},
): FormulonError {
  const pairs = [`failure=${failure}`];
  for (const [key, value] of Object.entries(detail.context ?? {})) pairs.push(`${key}=${value}`);
  const code =
    failure === 'disposed' ? StatusCode.BindingInvalidHandle : StatusCode.PreconditionFailed;
  const error = new FormulonError(code, message, pairs.join(' '));
  if ('cause' in detail) {
    Object.defineProperty(error, 'cause', {
      value: detail.cause,
      writable: true,
      configurable: true,
      enumerable: false,
    });
  }
  return error;
}
