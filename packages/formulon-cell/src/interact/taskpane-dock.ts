/** Width limits and edge hit width of the desktop app's docked task pane, in px. */
const DOCK_MIN_WIDTH = 271;
const DOCK_MAX_WIDTH = 1001;
const EDGE_GRAB = 4;

const resizableDocks = new WeakSet<HTMLElement>();

const onMac = (dock: HTMLElement): boolean =>
  dock.closest<HTMLElement>('[data-fc-platform]')?.dataset.fcPlatform === 'mac';

const onEdge = (dock: HTMLElement, clientX: number): boolean => {
  const rect = dock.getBoundingClientRect();
  return clientX >= rect.left - EDGE_GRAB && clientX <= rect.left + EDGE_GRAB;
};

/** The dock's left edge drags its width: it follows the pointer 1:1 between the limits. */
function installDockResize(dock: HTMLElement): void {
  if (resizableDocks.has(dock)) return;
  resizableDocks.add(dock);
  dock.addEventListener('pointermove', (event) => {
    dock.style.cursor = onMac(dock) && onEdge(dock, event.clientX) ? 'ew-resize' : '';
  });
  dock.addEventListener('pointerdown', (event) => {
    if (event.button !== 0 || !onMac(dock) || !onEdge(dock, event.clientX)) return;
    const doc = dock.ownerDocument;
    const startX = event.clientX;
    const startWidth = dock.getBoundingClientRect().width;
    const move = (e: MouseEvent): void => {
      const width = Math.min(
        Math.max(startWidth + startX - e.clientX, DOCK_MIN_WIDTH),
        DOCK_MAX_WIDTH,
      );
      dock.style.flexBasis = `${width}px`;
    };
    const stop = (): void => {
      doc.removeEventListener('pointermove', move);
      doc.removeEventListener('pointerup', stop);
      doc.removeEventListener('pointercancel', stop);
    };
    doc.addEventListener('pointermove', move);
    doc.addEventListener('pointerup', stop);
    doc.addEventListener('pointercancel', stop);
    event.preventDefault();
  });
}

/** Shows the host's task-pane dock while any pane in it is visible and hides it
 *  once none is. Each pane sets its own `hidden` first and then calls this, so
 *  one pane closing never hides another. */
export function syncTaskpaneDock(dock: HTMLElement): void {
  installDockResize(dock);
  dock.hidden = !Array.from(dock.children).some(
    (child) => child instanceof HTMLElement && !child.hidden,
  );
}
