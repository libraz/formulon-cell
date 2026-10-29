import { afterEach, describe, expect, it } from 'vitest';
import { applyMerge } from '../../src/commands/merge.js';
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

describe('integration: formula reference pointer routing', () => {
  let sheet: MountedStubSheet;

  afterEach(() => sheet?.dispose());

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
