import type { NumberResult, Status } from '@libraz/formulon';

/** Turn a failed engine call into the same error shape used by mutations. */
export function checkStatus(status: Status, operation: string): void {
  if (!status.ok) throw new Error(`${operation}: ${status.message}`);
}

/** Unwrap a required numeric result without treating a failed zero as data. */
export function numberValue<T extends number>(result: NumberResult<T>, operation: string): T {
  checkStatus(result.status, operation);
  return result.value;
}
