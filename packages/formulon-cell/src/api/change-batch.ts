import type { CellAddress } from '@libraz/formulon';
import type { PermissionCode } from '../commands/interaction-policy.js';

/** One cell a batch of changes left unwritten, and why. */
export interface BatchRejection {
  readonly sheet: number;
  readonly at: CellAddress;
  readonly code: PermissionCode;
  readonly reason: string;
}

/**
 * The outcome of a call that writes many cells and may skip some: `applied` when every cell was
 * written (an empty batch included), `partial` when some were, `refused` when none were.
 */
export interface ChangeBatchResult {
  readonly status: 'applied' | 'partial' | 'refused';
  readonly applied: readonly { readonly sheet: number; readonly at: CellAddress }[];
  readonly rejections: readonly BatchRejection[];
}

/** Builds the result from what was written and what was skipped. */
export const changeBatchResult = (
  applied: ChangeBatchResult['applied'],
  rejections: readonly BatchRejection[],
): ChangeBatchResult => ({
  status: rejections.length === 0 ? 'applied' : applied.length === 0 ? 'refused' : 'partial',
  applied,
  rejections,
});
