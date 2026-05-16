// Shared open/close lifecycle for ribbon dropdown menus. Every ribbon menu has
// the same three needs: track an open flag, anchor a wrapper element, and close
// on outside mousedown / Escape. Centralizing the effect here keeps the per-menu
// components focused on layout and content.

import { type RefObject, useEffect, useRef, useState } from 'react';

export interface UseMenuOpenResult {
  open: boolean;
  setOpen: (next: boolean | ((prev: boolean) => boolean)) => void;
  wrapRef: RefObject<HTMLDivElement | null>;
}

export const useMenuOpen = (): UseMenuOpenResult => {
  const [open, setOpen] = useState(false);
  const wrapRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDocDown = (event: MouseEvent): void => {
      if (event.target instanceof Node && wrapRef.current?.contains(event.target)) return;
      setOpen(false);
    };
    const onKey = (event: globalThis.KeyboardEvent): void => {
      if (event.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDocDown, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDocDown, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open]);

  return { open, setOpen, wrapRef };
};
