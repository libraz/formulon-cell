/** Pointer travel (px) before a press on a tab becomes a reorder drag. */
const DRAG_THRESHOLD_PX = 4;
/** Height of the drop marker triangle, which sits just above the strip. */
const DROP_MARK_HEIGHT_PX = 5;

export interface SheetTabDragHandle {
  detach(): void;
}

/**
 * Drag-to-reorder for the sheet tab strip. A black triangle marks the slot the
 * tab will drop into (the left edge of the tab under the pointer's left half,
 * or the end of the strip). `onDrop(from, to)` gets sheet indexes where `to` is
 * the position the moved sheet ends at.
 */
export function attachSheetTabDrag(input: {
  sheetTabs: HTMLElement;
  onDrop(from: number, to: number): void;
}): SheetTabDragHandle {
  const { sheetTabs } = input;
  const tabsOf = (): HTMLElement[] =>
    Array.from(sheetTabs.querySelectorAll<HTMLElement>('.fc-host__sheetbar-tab'));
  const indexOf = (tab: HTMLElement): number => Number(tab.dataset.fcSheetIndex);

  let from = -1;
  let startX = 0;
  let dragging = false;
  let slot = -1;
  let indicator: HTMLElement | null = null;
  let swallowClick = false;

  const slotAt = (x: number): { slot: number; edge: number } | null => {
    const tabs = tabsOf();
    const last = tabs.at(-1);
    if (!last) return null;
    for (const tab of tabs) {
      const r = tab.getBoundingClientRect();
      if (x < r.left + r.width / 2) return { slot: indexOf(tab), edge: r.left };
    }
    return { slot: indexOf(last) + 1, edge: last.getBoundingClientRect().right };
  };

  const clear = (): void => {
    indicator?.remove();
    indicator = null;
    from = -1;
    slot = -1;
    dragging = false;
    document.removeEventListener('pointermove', onMove);
    document.removeEventListener('pointerup', onUp);
    document.removeEventListener('pointercancel', onCancel);
  };

  function onMove(e: PointerEvent): void {
    if (!dragging) {
      if (Math.abs(e.clientX - startX) < DRAG_THRESHOLD_PX) return;
      dragging = true;
      indicator = document.createElement('div');
      indicator.className = 'fc-host__sheetbar-drop';
      indicator.setAttribute('aria-hidden', 'true');
      sheetTabs.parentElement?.appendChild(indicator);
    }
    const hit = slotAt(e.clientX);
    if (!hit || !indicator) return;
    slot = hit.slot;
    const top = sheetTabs.getBoundingClientRect().top;
    indicator.style.left = `${hit.edge}px`;
    indicator.style.top = `${top - DROP_MARK_HEIGHT_PX}px`;
  }

  function onUp(): void {
    const moved = dragging;
    const target = slot;
    const source = from;
    clear();
    if (!moved) return;
    swallowClick = true;
    setTimeout(() => {
      swallowClick = false;
    }, 0);
    if (target < 0) return;
    input.onDrop(source, target > source ? target - 1 : target);
  }

  function onCancel(): void {
    clear();
  }

  const onDown = (e: PointerEvent): void => {
    if (e.button !== 0) return;
    const tab =
      e.target instanceof Element ? e.target.closest<HTMLElement>('.fc-host__sheetbar-tab') : null;
    if (!tab) return;
    clear();
    from = indexOf(tab);
    startX = e.clientX;
    document.addEventListener('pointermove', onMove);
    document.addEventListener('pointerup', onUp);
    document.addEventListener('pointercancel', onCancel);
  };
  const onClickCapture = (e: MouseEvent): void => {
    if (!swallowClick) return;
    e.stopPropagation();
    e.preventDefault();
  };

  sheetTabs.addEventListener('pointerdown', onDown);
  sheetTabs.addEventListener('click', onClickCapture, true);
  return {
    detach(): void {
      clear();
      sheetTabs.removeEventListener('pointerdown', onDown);
      sheetTabs.removeEventListener('click', onClickCapture, true);
    },
  };
}
