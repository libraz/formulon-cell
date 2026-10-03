import { type RefObject, useEffect } from 'react';
import { activateDemoModal } from '../../demo-shared/index.js';

// Modal focus trap + Esc-to-close. `activateDemoModal` lives in demo-shared
// and is shared with the Vue demo; this hook adapts it to React's effect
// model by attaching on mount/open and detaching on unmount/close.
export const useDemoModalFocus = (
  rootRef: RefObject<HTMLElement | null>,
  open: boolean,
  onClose: () => void,
): void => {
  useEffect(() => {
    if (!open) return;
    const root = rootRef.current;
    if (!root) return;
    return activateDemoModal(root, onClose);
  }, [rootRef, open, onClose]);
};
