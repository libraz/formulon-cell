import { FormulonError, type StatusCode } from '@libraz/formulon';
import { type SpreadsheetFailure, spreadsheetFailureOf } from '@libraz/formulon-cell';
import { expect } from 'vitest';

/** Runs `fn` and returns what it threw; fails the test when it returns normally. */
export const thrownBy = (fn: () => unknown): unknown => {
  try {
    fn();
  } catch (e) {
    return e;
  }
  expect.fail('expected the call to throw');
};

/** Asserts that `fn` throws a `FormulonError` with `code`. */
export const expectFormulonError = (fn: () => unknown, code: StatusCode): FormulonError => {
  const error = thrownBy(fn);
  expect(error).toBeInstanceOf(FormulonError);
  expect((error as FormulonError).code).toBe(code);
  return error as FormulonError;
};

/** Asserts that `fn` throws a formulon-cell precondition failure of the given kind. */
export const expectFailure = (fn: () => unknown, failure: SpreadsheetFailure): FormulonError => {
  const error = thrownBy(fn);
  expect(error).toBeInstanceOf(FormulonError);
  expect(spreadsheetFailureOf(error)).toBe(failure);
  return error as FormulonError;
};

/** Asserts that `promise` rejects with a formulon-cell precondition failure of the given kind. */
export const expectFailureAsync = async (
  promise: Promise<unknown>,
  failure: SpreadsheetFailure,
): Promise<void> => {
  const error = await promise.then(
    () => expect.fail('expected the promise to reject'),
    (e: unknown) => e,
  );
  expect(error).toBeInstanceOf(FormulonError);
  expect(spreadsheetFailureOf(error)).toBe(failure);
};

/** A detached host element of a fixed size, appended to the document. */
export const makeHost = (): HTMLElement => {
  const host = document.createElement('div');
  host.style.width = '800px';
  host.style.height = '600px';
  document.body.appendChild(host);
  return host;
};
