import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { Range } from '../../../../src/engine/types.js';
import { addrKey, WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import type { SpreadsheetStore } from '../../../../src/store/store.js';
import type { CellFormat } from '../../../../src/store/types.js';

export const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../../..');

export const newWb = (): Promise<WorkbookHandle> =>
  WorkbookHandle.createDefault({ preferStub: true });

export const seed = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  cells: Array<{ row: number; col: number; value: number | string }>,
): void => {
  store.setState((s) => {
    const map = new Map(s.data.cells);
    for (const c of cells) {
      const addr = { sheet: 0, row: c.row, col: c.col };
      if (typeof c.value === 'number') {
        wb.setNumber(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'number', value: c.value }, formula: null });
      } else {
        wb.setText(addr, c.value);
        map.set(addrKey(addr), { value: { kind: 'text', value: c.value }, formula: null });
      }
    }
    return { ...s, data: { ...s.data, cells: map } };
  });
  wb.recalc();
};

export const setRange = (
  store: SpreadsheetStore,
  r0: number,
  c0: number,
  r1: number,
  c1: number,
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: 0, row: r0, col: c0 },
      anchor: { sheet: 0, row: r0, col: c0 },
      range: { sheet: 0, r0, c0, r1, c1 },
    },
  }));
};

export const setSelectionRanges = (
  store: SpreadsheetStore,
  range: Range,
  extraRanges: Range[],
): void => {
  store.setState((s) => ({
    ...s,
    selection: {
      ...s.selection,
      active: { sheet: range.sheet, row: range.r0, col: range.c0 },
      anchor: { sheet: range.sheet, row: range.r0, col: range.c0 },
      range,
      extraRanges,
    },
  }));
};

export const setFormat = (
  store: SpreadsheetStore,
  row: number,
  col: number,
  format: CellFormat,
): void => {
  store.setState((s) => {
    const formats = new Map(s.format.formats);
    formats.set(addrKey({ sheet: 0, row, col }), format);
    return { ...s, format: { ...s.format, formats } };
  });
};

export const fireContextMenu = (
  host: HTMLElement,
  x: number,
  y: number,
  init: MouseEventInit = {},
): MouseEvent => {
  const e = new MouseEvent('contextmenu', {
    clientX: x,
    clientY: y,
    bubbles: true,
    cancelable: true,
    ...init,
  });
  host.dispatchEvent(e);
  return e;
};

export const item = (id: string): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>(`.fc-ctxmenu__item[data-fc-action="${id}"]`);

export const miniItem = (id: string): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>(`.fc-ctxmenu__mini-btn[data-fc-action="${id}"]`);

export const visibleMenu = (): HTMLElement | null => {
  const root = document.querySelector<HTMLElement>('.fc-ctxmenu');
  if (!root) return null;
  return root.style.display === 'none' ? null : root;
};
