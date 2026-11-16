/** The element that holds focus when the sheet has it: the grid of a mounted sheet. */
export const sheetFocusTarget = (host: HTMLElement): HTMLElement =>
  host.querySelector<HTMLElement>('.fc-host__grid') ?? host;

/** Gives a bare host the grid child a mounted sheet has, so keys pressed on it reach the sheet. */
export const addSheetGrid = (host: HTMLElement): HTMLElement => {
  const grid = document.createElement('div');
  grid.className = 'fc-host__grid';
  grid.tabIndex = 0;
  host.appendChild(grid);
  return grid;
};
