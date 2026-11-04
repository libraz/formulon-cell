/**
 * Maps pointer events to coordinates local to the host's grid surface
 * (`.fc-host__grid`, falling back to the host itself). The grid element is
 * looked up lazily and kept while it stays inside the host.
 */
export const createGridLocalPoint = (
  host: HTMLElement,
): ((e: { clientX: number; clientY: number }) => { x: number; y: number }) => {
  let grid: HTMLElement | null = null;
  return (e) => {
    if (!grid || !host.contains(grid)) {
      grid = host.querySelector<HTMLElement>('.fc-host__grid');
    }
    const rect = (grid ?? host).getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  };
};
