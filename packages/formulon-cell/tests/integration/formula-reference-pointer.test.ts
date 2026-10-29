import { afterEach, describe, expect, it } from 'vitest';
import { applyMerge } from '../../src/commands/merge.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { cellRectUnclamped, layoutForView } from '../../src/render/geometry.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

const pointer = (
  target: HTMLElement,
  type: 'pointerdown' | 'pointermove' | 'pointerup',
  x: number,
  y: number,
): PointerEvent => {
  const event = new PointerEvent(type, {
    bubbles: true,
    cancelable: true,
    pointerId: 1,
    button: 0,
    clientX: x,
    clientY: y,
  });
  target.dispatchEvent(event);
  return event;
};

const prepareViewport = (sheet: MountedStubSheet): void => {
  sheet.instance.store.setState((state) => ({
    ...state,
    layout: {
      ...state.layout,
      defaultColWidth: 64,
      defaultRowHeight: 20,
      headerColWidth: 32,
      headerRowHeight: 20,
    },
    viewport: { ...state.viewport, rowCount: 20, colCount: 20, widthPx: 900 },
  }));
  const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
  if (grid) {
    Object.assign(grid, {
      setPointerCapture: () => {},
      releasePointerCapture: () => {},
      hasPointerCapture: () => false,
    });
  }
};

const center = (sheet: MountedStubSheet, row: number, col: number): { x: number; y: number } => {
  const state = sheet.instance.store.getState();
  const rect = cellRectUnclamped(layoutForView(state), state.viewport, row, col);
  return { x: rect.x + rect.w / 2, y: rect.y + rect.h / 2 };
};

const mountRealSheet = async (): Promise<MountedStubSheet> => {
  const workbook = await WorkbookHandle.createDefault();
  expect(workbook.isStub).toBe(false);
  return mountStubSheet({ workbook });
};

const seedFormulaReferenceCells = (
  sheet: MountedStubSheet,
  extra: Array<{ row: number; col: number; value: number }> = [],
): void => {
  const cells = [
    { row: 0, col: 1, value: 2 },
    { row: 1, col: 1, value: 3 },
    { row: 0, col: 3, value: 7 },
    { row: 3, col: 2, value: 11 },
    { row: 4, col: 3, value: 13 },
    ...extra,
  ];
  for (const cell of cells) {
    sheet.workbook.setNumber({ sheet: 0, row: cell.row, col: cell.col }, cell.value);
  }
};

describe('integration: formula reference pointer routing', () => {
  let sheet: MountedStubSheet;

  afterEach(() => {
    if (!sheet) return;
    const workbook = sheet.workbook;
    sheet.dispose();
    if (!workbook.isStub) workbook.dispose();
  });

  it.each(['default', 'mac'] as const)(
    '%s platform inserts references from formula bar draft and commits the edited formula',
    async (platform) => {
      sheet = await mountStubSheet({ ui: { platform } });
      prepareViewport(sheet);
      const addr = { sheet: 0, row: 4, col: 4 };
      sheet.workbook.setFormula(addr, '=A1+A2');
      mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
      mutators.setActive(sheet.instance.store, addr);
      const formula = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
      const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
      expect(formula).not.toBeNull();
      expect(grid).not.toBeNull();
      formula?.focus();
      formula?.dispatchEvent(new FocusEvent('focus'));
      expect(formula?.value).toBe('=A1+A2');
      formula?.setSelectionRange(1, 3);
      const b2 = center(sheet, 1, 1);
      if (grid) {
        pointer(grid, 'pointerdown', b2.x, b2.y);
        pointer(grid, 'pointerup', b2.x, b2.y);
      }
      expect(formula?.value).toBe('=B2+A2');
      expect(sheet.workbook.cellFormula(addr)).toBe('=A1+A2');
      formula?.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
      );
      expect(sheet.workbook.cellFormula(addr)).toBe('=B2+A2');
    },
  );

  it('preserves the untouched A1 endpoint while dragging a formula-bar reference forward', async () => {
    sheet = await mountRealSheet();
    prepareViewport(sheet);
    const formulaAddr = { sheet: 0, row: 4, col: 4 };
    seedFormulaReferenceCells(sheet);
    sheet.workbook.setFormula(formulaAddr, '=SUM(B1:B2)+D1');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, formulaAddr);
    const formula = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    expect(formula).not.toBeNull();
    expect(grid).not.toBeNull();
    formula?.focus();
    formula?.dispatchEvent(new FocusEvent('focus'));
    const endpoint = formula?.value.indexOf('B2') ?? -1;
    expect(endpoint).toBeGreaterThanOrEqual(0);
    formula?.setSelectionRange(endpoint, endpoint + 2);

    const c4 = center(sheet, 3, 2);
    const d5 = center(sheet, 4, 3);
    if (grid) {
      pointer(grid, 'pointerdown', c4.x, c4.y);
      expect(formula?.value).toBe('=SUM(B1:C4)+D1');
      pointer(grid, 'pointermove', d5.x, d5.y);
      expect(formula?.value).toBe('=SUM(B1:C4:D5)+D1');
      pointer(grid, 'pointerup', d5.x, d5.y);
    }

    formula?.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    );
    expect(sheet.workbook.cellFormula(formulaAddr)).toBe('=SUM(B1:C4:D5)+D1');
    expect(sheet.workbook.getValue(formulaAddr)).toEqual({ kind: 'number', value: 43 });
  });

  it('preserves the untouched endpoint while dragging an inline formula reference forward', async () => {
    sheet = await mountRealSheet();
    prepareViewport(sheet);
    const formulaAddr = { sheet: 0, row: 4, col: 4 };
    seedFormulaReferenceCells(sheet);
    sheet.workbook.setFormula(formulaAddr, '=SUM(B1:B2)+D1');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, formulaAddr);
    sheet.host.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true }));
    const formula = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__editor');
    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    expect(formula).not.toBeNull();
    expect(grid).not.toBeNull();
    expect(formula?.value).toBe('=SUM(B1:B2)+D1');
    const endpoint = formula?.value.indexOf('B2') ?? -1;
    expect(endpoint).toBeGreaterThanOrEqual(0);
    formula?.setSelectionRange(endpoint, endpoint + 2);

    const c4 = center(sheet, 3, 2);
    const d5 = center(sheet, 4, 3);
    if (grid) {
      pointer(grid, 'pointerdown', c4.x, c4.y);
      expect(formula?.value).toBe('=SUM(B1:C4)+D1');
      pointer(grid, 'pointermove', d5.x, d5.y);
      expect(formula?.value).toBe('=SUM(B1:C4:D5)+D1');
      pointer(grid, 'pointerup', d5.x, d5.y);
    }

    formula?.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    );
    expect(sheet.workbook.cellFormula(formulaAddr)).toBe('=SUM(B1:C4:D5)+D1');
    expect(sheet.workbook.getValue(formulaAddr)).toEqual({ kind: 'number', value: 43 });
  });

  it('preserves the original endpoint when dragging a formula-bar reference in reverse R1C1 style', async () => {
    sheet = await mountRealSheet();
    prepareViewport(sheet);
    const formulaAddr = { sheet: 0, row: 4, col: 4 };
    seedFormulaReferenceCells(sheet, [
      { row: 3, col: 0, value: 17 },
      { row: 4, col: 2, value: 19 },
    ]);
    sheet.workbook.setFormula(formulaAddr, '=SUM(B1:B2)+D1');
    sheet.workbook.recalc();
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setR1C1(sheet.instance.store, true);
    mutators.setActive(sheet.instance.store, formulaAddr);
    const formula = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    expect(formula).not.toBeNull();
    expect(grid).not.toBeNull();
    formula?.focus();
    formula?.dispatchEvent(new FocusEvent('focus'));
    const endpoint = formula?.value.indexOf('R[-3]C[-3]') ?? -1;
    expect(endpoint).toBeGreaterThanOrEqual(0);
    formula?.setSelectionRange(endpoint, endpoint + 'R[-3]C[-3]'.length);

    const c5 = center(sheet, 4, 2);
    const a4 = center(sheet, 3, 0);
    if (grid) {
      pointer(grid, 'pointerdown', c5.x, c5.y);
      expect(formula?.value).toBe('=SUM(R[-4]C[-3]:RC[-2])+R[-4]C[-1]');
      pointer(grid, 'pointermove', a4.x, a4.y);
      expect(formula?.value).toBe('=SUM(R[-4]C[-3]:R[-1]C[-4]:RC[-2])+R[-4]C[-1]');
      pointer(grid, 'pointerup', a4.x, a4.y);
    }

    formula?.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
    );
    expect(sheet.workbook.cellFormula(formulaAddr)).toBe('=SUM(B1:A4:C5)+D1');
    expect(sheet.workbook.getValue(formulaAddr).kind).toBe('number');
  });

  it('uses the merged cell anchor for a clicked reference and range drag', async () => {
    sheet = await mountStubSheet({ ui: { platform: 'mac' } });
    prepareViewport(sheet);
    const formulaAddr = { sheet: 0, row: 4, col: 4 };
    sheet.workbook.setFormula(formulaAddr, '=D4');
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    applyMerge(sheet.instance.store, sheet.workbook, sheet.instance.history, {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 1,
      c1: 1,
    });
    mutators.setActive(sheet.instance.store, formulaAddr);
    const formula = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    expect(formula).not.toBeNull();
    expect(grid).not.toBeNull();
    formula?.focus();
    formula?.dispatchEvent(new FocusEvent('focus'));
    formula?.setSelectionRange(1, formula.value.length);
    const b2 = center(sheet, 1, 1);
    if (grid) {
      pointer(grid, 'pointerdown', b2.x, b2.y);
      pointer(grid, 'pointermove', center(sheet, 0, 1).x, center(sheet, 0, 1).y);
      expect(formula?.value).toBe('=A1');
      pointer(grid, 'pointermove', center(sheet, 2, 2).x, center(sheet, 2, 2).y);
      expect(formula?.value).toBe('=A1:C3');
      pointer(grid, 'pointerup', center(sheet, 2, 2).x, center(sheet, 2, 2).y);
    }
  });

  it('does not open a stale inline edit on double click while formula-bar draft is active', async () => {
    sheet = await mountStubSheet({ ui: { platform: 'mac' } });
    prepareViewport(sheet);
    const addr = { sheet: 0, row: 4, col: 4 };
    sheet.workbook.setFormula(addr, '=A1');
    mutators.replaceCells(sheet.instance.store, sheet.workbook.cells(0));
    mutators.setActive(sheet.instance.store, addr);
    const formula = sheet.host.querySelector<HTMLTextAreaElement>('.fc-host__formulabar-input');
    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    expect(formula).not.toBeNull();
    expect(grid).not.toBeNull();
    formula?.focus();
    formula?.dispatchEvent(new FocusEvent('focus'));
    formula?.setSelectionRange(1, formula.value.length);
    const b2 = center(sheet, 1, 1);
    if (grid) {
      pointer(grid, 'pointerdown', b2.x, b2.y);
      pointer(grid, 'pointerup', b2.x, b2.y);
      grid.dispatchEvent(
        new MouseEvent('dblclick', {
          bubbles: true,
          cancelable: true,
          button: 0,
          clientX: b2.x,
          clientY: b2.y,
        }),
      );
    }
    expect(formula?.value).toBe('=B2');
    expect(sheet.host.querySelector('.fc-host__editor')).toBeNull();
    expect(document.activeElement).toBe(formula);
  });

  it('still opens inline editing for an ordinary cell double click', async () => {
    sheet = await mountStubSheet();
    prepareViewport(sheet);
    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    expect(grid).not.toBeNull();
    const b2 = center(sheet, 1, 1);
    const event = new MouseEvent('dblclick', {
      bubbles: true,
      cancelable: true,
      button: 0,
      clientX: b2.x,
      clientY: b2.y,
    });
    grid?.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    expect(sheet.host.querySelector('.fc-host__editor')).not.toBeNull();
  });

  it('keeps a consumed autofit double click from also opening cell editing', async () => {
    sheet = await mountStubSheet();
    prepareViewport(sheet);
    const grid = sheet.host.querySelector<HTMLElement>('.fc-host__grid');
    expect(grid).not.toBeNull();
    const event = new MouseEvent('dblclick', {
      bubbles: true,
      cancelable: true,
      button: 0,
      clientX: 94,
      clientY: 10,
    });
    grid?.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(true);
    expect(sheet.host.querySelector('.fc-host__editor')).toBeNull();
  });
});
