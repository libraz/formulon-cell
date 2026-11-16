import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { createPivotTableFromRange } from '../../../src/commands/pivot-table.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import { pointPx } from '../../../src/render/geometry.js';
import { pivotCollapseButtonRect } from '../../../src/render/painters/pivot-outline-button.js';
import {
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

// Header column 52 px, header row 30 px, columns 104 px, rows 28 px.
const cellBox = (row: number) => ({ x: 52, y: 30 + 28 * row, w: 104, h: 28 });

const click = (host: HTMLElement, x: number, y: number): void => {
  for (const type of ['pointerdown', 'pointerup']) {
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
  }
};

describe('the expand/collapse button of a report row label', () => {
  let host: HTMLElement;
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let history: History;
  let detach: () => void;

  beforeEach(async () => {
    host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    document.body.appendChild(host);
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
      ['Region', 'Product', 'Amt'],
      ['East', 'A', 5],
      ['West', 'B', 3],
      ['East', 'B', 7],
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
      rowField: 'Region',
      valueField: 'Amt',
    });
    expect(created.ok).toBe(true);
    expect(wb.setPivotFieldAxis(0, 0, 1, 0)).toBe(true);
    mutators.replaceCells(store, wb.cells(0));
    mutators.setActive(store, { sheet: 0, row: 0, col: 3 });
    history = new History();
    detach = attachPointer(host, store, wb, undefined, history);
  });

  afterEach(() => {
    detach();
    host.remove();
    wb.dispose();
  });

  const eastDetail = () => wb.getPivotField(0, 0, 0)?.items[0]?.showDetail;

  it('collapses the item on a click, keeping the active cell, one undo step', () => {
    // East heads rows 4-5 on row 3 (index), the button one outline step in.
    const button = pivotCollapseButtonRect(cellBox(3), 1, pointPx(store.getState().layout));
    click(host, button.x + button.w / 2, button.y + button.h / 2);
    expect(eastDetail()).toBe(false);
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 3 });
    expect(history.undo()).toBe(true);
    expect(eastDetail()).toBe(true);
  });

  it('selects the label when the click misses the button', () => {
    click(host, 140, cellBox(3).y + 14);
    expect(eastDetail()).toBe(true);
    expect(store.getState().selection.active).toMatchObject({ row: 3, col: 0 });
  });
});
