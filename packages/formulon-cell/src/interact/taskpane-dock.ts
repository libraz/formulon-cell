/** Shows the host's task-pane dock while any pane in it is visible and hides it
 *  once none is. Each pane sets its own `hidden` first and then calls this, so
 *  one pane closing never hides another. */
export function syncTaskpaneDock(dock: HTMLElement): void {
  dock.hidden = !Array.from(dock.children).some(
    (child) => child instanceof HTMLElement && !child.hidden,
  );
}
