import type { CellBorderSide } from '../store/store.js';

/** A comparable string for one border side: equal signatures mean visually identical sides. */
export const borderSideSignature = (side: CellBorderSide | undefined): string => {
  if (side === undefined) return 'absent';
  if (typeof side === 'boolean') return `boolean:${side}`;
  return `style:${side.style};color:${side.color ?? ''}`;
};
