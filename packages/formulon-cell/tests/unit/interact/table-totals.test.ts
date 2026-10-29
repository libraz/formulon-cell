import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { formatAsTable } from '../../../src/commands/format-as-table.js';
import { History } from '../../../src/commands/history.js';
import { toggleTableOption } from '../../../src/commands/table-design.js';
import { hydrateTableOverlaysFromEngine } from '../../../src/engine/table-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { dictionaries } from '../../../src/i18n/strings.js';
import { attachTableTotals, type TableTotalsHandle } from '../../../src/interact/table-totals.js';
import { gridHitState } from '../../../src/render/grid/hit-state.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('table totals function picker', () => {
  let wb: WorkbookHandle;
  let grid: HTMLElement;
  let handle: TableTotalsHandle;
  let store = createSpreadsheetStore();
  let history = new History();
  const onMoreFunctions = vi.fn();

  beforeEach(async () => {
    wb = await WorkbookHandle.createDefault();
    store = createSpreadsheetStore();
    history = new History();
    [
      ['Item', 'Qty'],
      ['A', 1],
      ['B', 2],
      ['C', 3],
    ].forEach((cols, row) => {
      cols.forEach((value, col) => {
        if (typeof value === 'number') wb.setNumber({ sheet: 0, row, col }, value);
        else wb.setText({ sheet: 0, row, col }, value);
      });
    });
    mutators.replaceCells(store, wb.cells(0));
    hydrateTableOverlaysFromEngine(wb, store);
    wb.attachStore(store);
    formatAsTable(store, { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 1 }, { workbook: wb });
    toggleTableOption(store, history, wb, 'showTotal');
    mutators.setActive(store, { sheet: 0, row: 4, col: 1 });
    grid = document.createElement('div');
    document.body.appendChild(grid);
    handle = attachTableTotals({
      grid,
      store,
      history,
      wb,
      strings: dictionaries.en,
      onAfterCommit: () => undefined,
      onMoreFunctions,
    });
    gridHitState(grid).totalsChevron = {
      rect: { x: 60, y: 10, w: 18, h: 20 },
      row: 4,
      col: 1,
      tableId: store.getState().tables.tables[0]?.id ?? '',
      column: 1,
    };
  });

  afterEach(() => {
    handle.detach();
    onMoreFunctions.mockClear();
    grid.remove();
    wb.dispose();
  });

  const items = (): HTMLElement[] =>
    Array.from(document.querySelectorAll<HTMLElement>('.fc-table-totals [role="option"]'));

  it('lists the totals functions and marks the current one', () => {
    expect(handle.openForActiveCell()).toBe(true);
    expect(document.querySelector('.fc-table-totals')?.getAttribute('aria-label')).toBe(
      'Totals Function',
    );
    expect(items().map((item) => item.textContent)).toEqual([
      'None',
      'Average',
      'Count',
      'Count Numbers',
      'Max',
      'Min',
      'Sum',
      'StdDev',
      'Var',
      'More Functions...',
    ]);
    expect(
      items()
        .filter((item) => item.getAttribute('aria-selected') === 'true')
        .map((item) => item.dataset.totalsFunction),
    ).toEqual(['sum']);
  });

  it('opens from a click on the chevron', () => {
    const rect = grid.getBoundingClientRect();
    grid.dispatchEvent(
      new PointerEvent('pointerdown', {
        bubbles: true,
        button: 0,
        clientX: rect.left + 65,
        clientY: rect.top + 15,
      }),
    );
    expect(items()).toHaveLength(10);
  });

  it('sets the column function, writes the engine and undoes', () => {
    handle.openForActiveCell();
    items()
      .find((item) => item.dataset.totalsFunction === 'max')
      ?.click();
    expect(items()).toHaveLength(0);
    expect(store.getState().tables.tables[0]?.totals?.[1]?.function).toBe('max');
    expect(wb.getTableStyleDetails(0)?.columns[1]?.totalsRowFunction).toBe('max');
    expect(wb.getValue({ sheet: 0, row: 4, col: 1 })).toMatchObject({ kind: 'number', value: 3 });
    history.undo();
    expect(wb.getTableStyleDetails(0)?.columns[1]?.totalsRowFunction).toBe('sum');
  });

  it('leaves the cell for typing on More Functions', () => {
    handle.openForActiveCell();
    items()
      .find((item) => item.dataset.totalsFunction === 'custom')
      ?.click();
    expect(onMoreFunctions).toHaveBeenCalledOnce();
    expect(wb.getTableStyleDetails(0)?.columns[1]?.totalsRowFunction).toBe('sum');
  });
});
