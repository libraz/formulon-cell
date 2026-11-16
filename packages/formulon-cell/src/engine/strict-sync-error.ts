/** The error a strict engine sync throws when `operation` fails at `address` in `area`. */
export const strictSyncError = (area: string, operation: string, address: string): Error =>
  new Error(`Strict engine sync failed: ${area} ${operation} at ${address}`);
