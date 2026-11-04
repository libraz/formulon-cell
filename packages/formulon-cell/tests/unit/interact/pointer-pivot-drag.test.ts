import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// Header column 52 px, header row 30 px, columns 104 px, rows 28 px: column A
// spans x 52..156, row 3 (index 2) y 86..114, and each row below adds 28.
const rowTop = (row: number): number => 30 + 28 * row;

const fire = (host: HTMLElement, type: string, x: number, y: number): void => {
  host.dispatchEvent(
    new PointerEvent(type, {
      clientX: x,
      clientY: y,
      button: 0,
      pointerId: 1,
      bubbles: true,
      cancelable: true,
    }),
  );
};

describe('dragging a report or one of its items by the selection border', () => {
  let fcHost: HTMLElement;
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let detach: () => void;

  beforeEach(async () => {
    fcHost = document.createElement('div');
    fcHost.className = 'fc-host';
    fcHost.dataset.fcPlatform = 'mac';
    host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    fcHost.appendChild(host);
    document.body.appendChild(fcHost);
    store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        defaultColWidth: 104,
        defaultRowHeight: 28,
        headerColWidth: 52,
        headerRowHeight: 30,
      },
    }));
    wb = await WorkbookHandle.createDefault();
    const rows: (string | number)[][] = [
      ['Region', 'Amt', 'Qty'],
      ['East', 5, 5],
      ['North', 3, 3],
      ['South', 7, 7],
    ];
    for (const [r, row] of rows.entries()) {
      for (const [c, v] of row.entries()) {
        const addr = { sheet: 0, row: r, col: c + 10 };
        if (typeof v === 'number') wb.setNumber(addr, v);
        else wb.setText(addr, v);
      }
    }
    const created = createPivotTableFromRange(wb, {
      source: { sheet: 0, r0: 0, c0: 10, r1: 3, c1: 12 },
      destination: { sheet: 0, row: 2, col: 0 },
      name: 'Report',
      rowField: 'Region',
      valueField: 'Amt',
    });
    expect(created.ok).toBe(true);
    mutators.replaceCells(store, wb.cells(0));
    detach = attachPointer(host, store, wb);
  });

  afterEach(() => {
    detach();
    fcHost.remove();
    wb.dispose();
  });

  const label = (row: number): unknown => {
    const v = store.getState().data.cells.get(addrKey({ sheet: 0, row, col: 0 }))?.value;
    return v?.kind === 'text' ? v.value : null;
  };

  it('reorders an item dragged by its bottom border and shows the drop row', () => {
    mutators.setActive(store, { sheet: 0, row: 5, col: 0 });
    // South sits on row 6; one pixel inside its bottom edge.
    fire(host, 'pointerdown', 100, rowTop(6) - 1);
    // Released at the header/first item boundary.
    fire(host, 'pointermove', 100, rowTop(3));
    expect(store.getState().ui.movePreview).toEqual({ sheet: 0, r0: 2, c0: 0, r1: 2, c1: 1 });
    expect(document.querySelector('.fc-fill-tooltip')?.textContent).toBe('A3:B3');
    fire(host, 'pointerup', 100, rowTop(3));
    expect(store.getState().ui.movePreview).toBeNull();
    expect([label(3), label(4), label(5)]).toEqual(['South', 'East', 'North']);
    expect(store.getState().selection.range).toMatchObject({ r0: 3, r1: 5, c0: 0, c1: 0 });
  });

  it('keeps the order when released on the row above the item', () => {
    mutators.setActive(store, { sheet: 0, row: 5, col: 0 });
    fire(host, 'pointerdown', 100, rowTop(6) - 1);
    fire(host, 'pointermove', 100, rowTop(5));
    expect(store.getState().ui.movePreview).toMatchObject({ r0: 4, r1: 4 });
    fire(host, 'pointerup', 100, rowTop(5));
    expect([label(3), label(4), label(5)]).toEqual(['East', 'North', 'South']);
  });

  it('moves a whole selected report as a report, not as cells', () => {
    mutators.setActive(store, { sheet: 0, row: 2, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 2, c0: 0, r1: 6, c1: 1 });
    fire(host, 'pointerdown', 100, rowTop(7) - 1);
    fire(host, 'pointermove', 100, rowTop(7) - 1 + 28 * 10);
    expect(store.getState().ui.movePreview).toEqual({ sheet: 0, r0: 12, c0: 0, r1: 16, c1: 1 });
    fire(host, 'pointerup', 100, rowTop(7) - 1 + 28 * 10);
    expect(wb.pivotBounds(0).map((b) => b.range)).toEqual([
      { sheet: 0, r0: 12, c0: 0, r1: 16, c1: 1 },
    ]);
    expect(store.getState().selection.range).toMatchObject({ r0: 12, r1: 16 });
  });
});
