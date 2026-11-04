/**
 * Pointer shapes the desktop Mac app shows over the grid, which CSS has no
 * keyword for: black arrows over the headers and split triangles on the resize
 * boundaries, each with a white outline. Cells use the platform's cell cross.
 */

const svgCursor = (body: string, fallback: string): string => {
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 20 20">` +
    `<g fill="#000" stroke="#fff" stroke-width="1" stroke-linejoin="round">${body}</g></svg>`;
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") 10 10, ${fallback}`;
};

const DOWN_ARROW = 'M7 3h6v7h3.5L10 17 3.5 10H7z';
const RIGHT_ARROW = 'M3 7h7V3.5L17 10l-7 6.5V13H3z';
const SPLIT_HORIZONTAL = 'M1.5 10 7 5v10zM18.5 10 13 5v10z';
const SPLIT_VERTICAL = 'M10 1.5 15 7H5zM10 18.5 5 13h10z';

/** Mac cursors keyed by the grid region they belong to. */
export const MAC_GRID_CURSORS = {
  cell: 'cell',
  colHeader: svgCursor(DOWN_ARROW, 's-resize'),
  rowHeader: svgCursor(RIGHT_ARROW, 'e-resize'),
  colResize: svgCursor(SPLIT_HORIZONTAL, 'col-resize'),
  rowResize: svgCursor(SPLIT_VERTICAL, 'row-resize'),
} as const;

/** Whether the grid host renders the Mac profile. */
export const isMacHost = (host: HTMLElement): boolean =>
  (host.closest<HTMLElement>('.fc-host') ?? host).dataset.fcPlatform === 'mac';

/** The cursor for a resize boundary: the Mac split shape, else the CSS keyword. */
export const resizeCursor = (host: HTMLElement, axis: 'col' | 'row'): string => {
  if (!isMacHost(host)) return axis === 'col' ? 'col-resize' : 'row-resize';
  return axis === 'col' ? MAC_GRID_CURSORS.colResize : MAC_GRID_CURSORS.rowResize;
};
