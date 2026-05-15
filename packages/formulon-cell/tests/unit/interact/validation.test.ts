import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachValidationList } from '../../../src/interact/validation.js';
import { setValidationChevron } from '../../../src/render/grid/hit-state.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const newWb = (): Promise<WorkbookHandle> => WorkbookHandle.createDefault({ preferStub: true });

const popover = (): HTMLElement | null =>
  document.querySelector<HTMLElement>('.fc-validation-list');
const items = (): HTMLElement[] =>
  Array.from(document.querySelectorAll<HTMLElement>('.fc-validation-list__item'));

const stubGridRect = (grid: HTMLElement): void => {
  grid.getBoundingClientRect = (): DOMRect =>
    ({ left: 0, top: 0, right: 200, bottom: 200, width: 200, height: 200, x: 0, y: 0 }) as DOMRect;
};

const firePointerInChevron = (
  grid: HTMLElement,
  chevronX: number,
  chevronY: number,
): PointerEvent => {
  const e = new PointerEvent('pointerdown', {
    clientX: chevronX + 2,
    clientY: chevronY + 2,
    button: 0,
    bubbles: true,
    cancelable: true,
    pointerId: 1,
  });
  grid.dispatchEvent(e);
  return e;
};

describe('attachValidationList', () => {
  let grid: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let onAfterCommit: () => void;

  beforeEach(async () => {
    grid = document.createElement('div');
    document.body.appendChild(grid);
    stubGridRect(grid);
    store = createSpreadsheetStore();
    wb = await newWb();
    onAfterCommit = vi.fn<() => void>();
  });

  afterEach(() => {
    setValidationChevron(null);
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  it('chevron pointerdown on a list-validated cell opens a dropdown of allowed values', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 1 },
      { validation: { kind: 'list', source: ['alpha', 'beta', 'gamma'] } },
    );
    setValidationChevron({ rect: { x: 50, y: 30, w: 12, h: 12 }, row: 1, col: 1 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    expect(popover()).toBeNull();
    firePointerInChevron(grid, 50, 30);
    expect(popover()).not.toBeNull();
    expect(items().map((i) => i.textContent)).toEqual(['alpha', 'beta', 'gamma']);
    handle.detach();
  });

  it('clicking an item writes the chosen value to the workbook and notifies onAfterCommit', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['x', 'y'] } },
    );
    setValidationChevron({ rect: { x: 70, y: 50, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 70, 50);
    const item = items().find((i) => i.textContent === 'y');
    item?.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    expect(onAfterCommit).toHaveBeenCalledTimes(1);
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'text', value: 'y' });
    // Popover closes after a pick.
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('Escape on the document closes an open popover', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['only'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 10, 10);
    expect(popover()).not.toBeNull();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('mousedown outside the popover closes it', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['only'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    firePointerInChevron(grid, 10, 10);
    expect(popover()).not.toBeNull();

    const outside = document.createElement('div');
    document.body.appendChild(outside);
    outside.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
    expect(popover()).toBeNull();
    handle.detach();
  });

  it('detach removes the pointerdown listener so subsequent chevron taps do nothing', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['only'] } },
    );
    setValidationChevron({ rect: { x: 10, y: 10, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    handle.detach();
    firePointerInChevron(grid, 10, 10);
    expect(popover()).toBeNull();
  });

  it('pointerdown outside the chevron rect does not open the popover', () => {
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'list', source: ['x'] } },
    );
    setValidationChevron({ rect: { x: 100, y: 100, w: 12, h: 12 }, row: 0, col: 0 });

    const handle = attachValidationList({ grid, store, wb, onAfterCommit });
    // Far-away click.
    firePointerInChevron(grid, 0, 0);
    expect(popover()).toBeNull();
    handle.detach();
  });
});
