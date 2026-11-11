/** A drawing whose Locked flag sheet protection reads. */
export type LockableDrawing = 'slicer' | 'note';

/**
 * Whether a drawing of `kind` is locked. The engine does not expose a drawing's Locked flag yet
 * (slicer, note shape), so every one counts as locked: on a protected sheet Edit objects alone
 * lets it be used or changed. The flag replaces this once the engine reports it.
 */
export const drawingLocked = (_kind: LockableDrawing): boolean => true;
