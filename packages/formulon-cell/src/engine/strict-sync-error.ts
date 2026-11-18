/** Thrown by a strict engine sync when the engine refuses a write. */
export class StrictSyncError extends Error {}

/** The error a strict engine sync throws when `operation` fails at `address` in `area`. */
export const strictSyncError = (area: string, operation: string, address: string): Error =>
  new StrictSyncError(`Strict engine sync failed: ${area} ${operation} at ${address}`);
