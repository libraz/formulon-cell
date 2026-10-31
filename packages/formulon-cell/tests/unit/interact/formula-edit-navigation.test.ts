import { describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import {
  navigateFormulaReference,
  pickedReferenceNavigation,
} from '../../../src/interact/formula-edit-navigation.js';
import { createFormulaReferenceGesture } from '../../../src/interact/formula-reference-gesture.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('formula edit point navigation', () => {
  it('uses data-edge jumps for Ctrl arrows while preserving and collapsing the expression span', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      for (const col of [1, 2, 3, 4, 7]) {
        cells.set(addrKey({ sheet: 0, row: 1, col }), {
          value: { kind: 'number', value: col },
          formula: null,
        });
      }
      return { ...state, data: { ...state.data, cells } };
    });
    const editedCell = { sheet: 0, row: 1, col: 1 };
    const text = '=SUM()+7';
    const first = navigateFormulaReference({
      state: store.getState(),
      text,
      selectionStart: 5,
      selectionEnd: 5,
      editedCell,
      direction: 'right',
      extend: false,
      jump: true,
    });
    expect(first).toMatchObject({ value: '=SUM(E2)+7', anchor: { row: 1, col: 4 } });

    const extended = navigateFormulaReference({
      state: store.getState(),
      text: first?.value ?? '',
      selectionStart: first?.caret ?? 0,
      selectionEnd: first?.caret ?? 0,
      editedCell,
      direction: 'right',
      extend: true,
      current: first,
      jump: true,
    });
    expect(extended).toMatchObject({
      value: '=SUM(E2:H2)+7',
      anchor: { row: 1, col: 4 },
      tip: { row: 1, col: 7 },
    });

    const collapsed = navigateFormulaReference({
      state: store.getState(),
      text: extended?.value ?? '',
      selectionStart: extended?.caret ?? 0,
      selectionEnd: extended?.caret ?? 0,
      editedCell,
      direction: 'left',
      extend: false,
      current: extended,
      jump: true,
    });
    expect(collapsed).toMatchObject({
      value: '=SUM(E2)+7',
      anchor: { row: 1, col: 4 },
      tip: { row: 1, col: 4 },
    });

    const reverseExtended = navigateFormulaReference({
      state: store.getState(),
      text: first?.value ?? '',
      selectionStart: first?.caret ?? 0,
      selectionEnd: first?.caret ?? 0,
      editedCell,
      direction: 'left',
      extend: true,
      current: first,
      jump: true,
    });
    expect(reverseExtended).toMatchObject({
      value: '=SUM(B2:E2)+7',
      anchor: { row: 1, col: 4 },
      tip: { row: 1, col: 1 },
    });
    const reverseCollapsed = navigateFormulaReference({
      state: store.getState(),
      text: reverseExtended?.value ?? '',
      selectionStart: reverseExtended?.caret ?? 0,
      selectionEnd: reverseExtended?.caret ?? 0,
      editedCell,
      direction: 'right',
      extend: false,
      current: reverseExtended,
      jump: true,
    });
    expect(reverseCollapsed).toMatchObject({
      value: '=SUM(E2)+7',
      anchor: { row: 1, col: 4 },
      tip: { row: 1, col: 4 },
    });
  });

  it('applies visual RTL direction before the data jump and snaps a jumped merge body to its anchor', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      for (const col of [5, 6]) {
        cells.set(addrKey({ sheet: 0, row: 1, col }), {
          value: { kind: 'number', value: col },
          formula: null,
        });
      }
      return {
        ...state,
        data: { ...state.data, cells },
        ui: { ...state.ui, rightToLeft: true },
      };
    });
    mutators.mergeRange(store, { sheet: 0, r0: 1, c0: 4, r1: 1, c1: 5 });
    const result = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM()',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 1, col: 6 },
      direction: 'right',
      extend: false,
      jump: true,
    });

    expect(result).toMatchObject({ value: '=SUM(E2)', tip: { row: 1, col: 4 } });
  });

  it('keeps idle data-edge occupancy across hidden-column boundaries', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      for (const col of [1, 2, 3, 4, 7]) {
        cells.set(addrKey({ sheet: 0, row: 1, col }), {
          value: { kind: 'number', value: col },
          formula: null,
        });
      }
      return {
        ...state,
        data: { ...state.data, cells },
        layout: { ...state.layout, hiddenCols: new Set([4]) },
      };
    });
    const result = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM()',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 1, col: 1 },
      direction: 'right',
      extend: false,
      jump: true,
    });

    expect(result?.tip).toEqual({ sheet: 0, row: 1, col: 4 });
  });

  it('formats a Ctrl range jump in R1C1 relative to the edited cell', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => {
      const cells = new Map(state.data.cells);
      for (const col of [3, 4, 7]) {
        cells.set(addrKey({ sheet: 0, row: 3, col }), {
          value: { kind: 'number', value: col },
          formula: null,
        });
      }
      return {
        ...state,
        data: { ...state.data, cells },
        ui: { ...state.ui, r1c1: true },
      };
    });
    const editedCell = { sheet: 0, row: 3, col: 3 };
    const first = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM()',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell,
      direction: 'right',
      extend: false,
      jump: true,
      r1c1: true,
    });
    const range = navigateFormulaReference({
      state: store.getState(),
      text: first?.value ?? '',
      selectionStart: first?.caret ?? 0,
      selectionEnd: first?.caret ?? 0,
      editedCell,
      direction: 'right',
      extend: true,
      current: first,
      jump: true,
      r1c1: true,
    });

    expect(first?.tip).toEqual({ sheet: 0, row: 3, col: 4 });
    expect(range?.tip).toEqual({ sheet: 0, row: 3, col: 7 });
    expect(range?.value).toMatch(/^=SUM\(R.*:R.*\)$/u);
  });

  it('starts from the edited cell and extends or collapses a reference with Shift', () => {
    const store = createSpreadsheetStore();
    const editedCell = { sheet: 0, row: 3, col: 3 };
    const first = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM(',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell,
      direction: 'left',
      extend: false,
    });
    expect(first).toMatchObject({ value: '=SUM(C4', anchor: { sheet: 0, row: 3, col: 2 } });

    const extended = navigateFormulaReference({
      state: store.getState(),
      text: first?.value ?? '',
      selectionStart: first?.caret ?? 0,
      selectionEnd: first?.caret ?? 0,
      editedCell,
      direction: 'up',
      extend: true,
      current: first,
    });
    expect(extended).toMatchObject({
      value: '=SUM(C3:C4',
      anchor: { sheet: 0, row: 3, col: 2 },
      tip: { sheet: 0, row: 2, col: 2 },
    });

    const collapsed = navigateFormulaReference({
      state: store.getState(),
      text: extended?.value ?? '',
      selectionStart: extended?.caret ?? 0,
      selectionEnd: extended?.caret ?? 0,
      editedCell,
      direction: 'down',
      extend: false,
      current: extended,
    });
    expect(collapsed).toMatchObject({
      value: '=SUM(C4',
      anchor: { sheet: 0, row: 3, col: 2 },
      tip: { sheet: 0, row: 3, col: 2 },
    });
  });

  it('starts at the edited cell even when the caret is inside an existing range', () => {
    const store = createSpreadsheetStore();
    const text = '=SUM(C3:C4)+D1';
    const result = navigateFormulaReference({
      state: store.getState(),
      text,
      selectionStart: 5,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 3, col: 3 },
      direction: 'left',
      extend: false,
    });

    expect(result?.value).toBe('=SUM(C4)+D1');
    expect(result?.tip).toEqual({ sheet: 0, row: 3, col: 2 });
  });

  it.each([
    { ref: 'B5', r1c1: false, anchor: [4, 1], tip: [4, 1] },
    { ref: '$B$5:C6', r1c1: false, anchor: [4, 1], tip: [5, 2] },
    { ref: 'R[1]C[-2]', r1c1: true, anchor: [4, 1], tip: [4, 1] },
    { ref: 'R5C2:R6C3', r1c1: true, anchor: [4, 1], tip: [5, 2] },
  ])('builds navigation state for a picked $ref', ({ ref, r1c1, anchor, tip }) => {
    const gesture = createFormulaReferenceGesture('=SUM(', 5, 5, { r1c1 });
    const state = pickedReferenceNavigation({
      gesture,
      inserted: gesture.insert(ref),
      ref,
      editedCell: { sheet: 0, row: 3, col: 3 },
      r1c1,
    });

    expect(state?.anchor).toEqual({ sheet: 0, row: anchor[0], col: anchor[1] });
    expect(state?.tip).toEqual({ sheet: 0, row: tip[0], col: tip[1] });
  });

  it('continues from a picked range on the next step', () => {
    const store = createSpreadsheetStore();
    const gesture = createFormulaReferenceGesture('=SUM(', 5, 5);
    const inserted = gesture.insert('B5:C6');
    const current = pickedReferenceNavigation({
      gesture,
      inserted,
      ref: 'B5:C6',
      editedCell: { sheet: 0, row: 3, col: 3 },
      r1c1: false,
    });
    const result = navigateFormulaReference({
      state: store.getState(),
      text: inserted.text,
      selectionStart: inserted.caret,
      selectionEnd: inserted.caret,
      editedCell: { sheet: 0, row: 3, col: 3 },
      direction: 'down',
      extend: true,
      current,
    });

    expect(result?.value).toBe('=SUM(B5:C7');
  });

  it('records no pick state for a sheet-qualified reference', () => {
    const gesture = createFormulaReferenceGesture('=', 1, 1);
    expect(
      pickedReferenceNavigation({
        gesture,
        inserted: gesture.insert('Sheet2!B5'),
        ref: 'Sheet2!B5',
        editedCell: { sheet: 0, row: 3, col: 3 },
        r1c1: false,
      }),
    ).toBeNull();
  });

  it('does not navigate when a native selection covers string text', () => {
    const store = createSpreadsheetStore();
    const result = navigateFormulaReference({
      state: store.getState(),
      text: '="abc',
      selectionStart: 2,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 3, col: 3 },
      direction: 'left',
      extend: false,
    });

    expect(result).toBeNull();
  });

  it('replaces an explicit formula selection while preserving both surrounding sides', () => {
    const store = createSpreadsheetStore();
    const text = '=1+SUM(A1)+2';
    const result = navigateFormulaReference({
      state: store.getState(),
      text,
      selectionStart: 3,
      selectionEnd: 10,
      editedCell: { sheet: 0, row: 3, col: 3 },
      direction: 'left',
      extend: false,
    });

    expect(result?.value).toBe('=1+C4+2');
    expect(result?.tip).toEqual({ sheet: 0, row: 3, col: 2 });
  });

  it('collapses an existing range to the clamped endpoint on a plain arrow at the sheet edge', () => {
    const store = createSpreadsheetStore();
    const editedCell = { sheet: 0, row: 3, col: 2 };
    const first = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM(',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell,
      direction: 'left',
      extend: true,
    });
    expect(first?.value).toBe('=SUM(B4:C4');

    const extended = navigateFormulaReference({
      state: store.getState(),
      text: first?.value ?? '',
      selectionStart: first?.caret ?? 0,
      selectionEnd: first?.caret ?? 0,
      editedCell,
      direction: 'left',
      extend: true,
      current: first,
    });
    expect(extended?.value).toBe('=SUM(A4:C4');

    const collapsed = navigateFormulaReference({
      state: store.getState(),
      text: extended?.value ?? '',
      selectionStart: extended?.caret ?? 0,
      selectionEnd: extended?.caret ?? 0,
      editedCell,
      direction: 'left',
      extend: false,
      current: extended,
    });
    expect(collapsed?.value).toBe('=SUM(A4');
    expect(collapsed?.anchor).toEqual({ sheet: 0, row: 3, col: 0 });
    expect(collapsed?.tip).toEqual({ sheet: 0, row: 3, col: 0 });
  });

  it('steps over hidden rows and merged cells without mutating selection', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 2, r1: 3, c1: 2 });
    store.setState((state) => ({
      ...state,
      layout: { ...state.layout, hiddenRows: new Set([4]) },
    }));
    const selectionBefore = store.getState().selection;
    const result = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM(',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 2, col: 2 },
      direction: 'down',
      extend: false,
    });

    expect(result).toMatchObject({ value: '=SUM(C6', tip: { row: 5, col: 2 } });
    expect(store.getState().selection).toBe(selectionBefore);
    expect(store.getState().ui.editor.kind).toBe('idle');
  });

  it('uses visual left and right in a right-to-left sheet and consumes a clamped edge step', () => {
    const store = createSpreadsheetStore();
    store.setState((state) => ({
      ...state,
      ui: { ...state.ui, rightToLeft: true },
    }));
    const rightToLeft = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM(',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 0, col: 3 },
      direction: 'left',
      extend: false,
    });
    expect(rightToLeft).toMatchObject({ value: '=SUM(E1', tip: { row: 0, col: 4 } });

    const edge = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM(',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 0, col: 0 },
      direction: 'right',
      extend: false,
    });
    expect(edge).toMatchObject({ value: '=SUM(A1', tip: { row: 0, col: 0 } });
  });

  it('starts point navigation at the edited cell when the caret is inside an R1C1 reference', () => {
    const store = createSpreadsheetStore();
    const text = '=SUM(R[-1]C[-1])';
    const caret = text.indexOf('R[-1]C[-1]') + 5;
    const result = navigateFormulaReference({
      state: store.getState(),
      text,
      selectionStart: caret,
      selectionEnd: caret,
      editedCell: { sheet: 0, row: 3, col: 3 },
      direction: 'right',
      extend: false,
      r1c1: true,
    });

    expect(result?.value).toBe('=SUM(RC[1])');
  });

  it('leaves string, function-name, and stale-caret input to normal text navigation', () => {
    const store = createSpreadsheetStore();
    const literal = '="SUM"';
    const literalResult = navigateFormulaReference({
      state: store.getState(),
      text: literal,
      selectionStart: 4,
      selectionEnd: 4,
      editedCell: { sheet: 0, row: 0, col: 0 },
      direction: 'left',
      extend: false,
    });
    expect(literalResult).toBeNull();

    const functionName = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM(',
      selectionStart: 3,
      selectionEnd: 3,
      editedCell: { sheet: 0, row: 0, col: 0 },
      direction: 'left',
      extend: false,
    });
    expect(functionName).toBeNull();

    const first = navigateFormulaReference({
      state: store.getState(),
      text: '=SUM(',
      selectionStart: 5,
      selectionEnd: 5,
      editedCell: { sheet: 0, row: 0, col: 1 },
      direction: 'left',
      extend: false,
    });
    const stale = navigateFormulaReference({
      state: store.getState(),
      text: `${first?.value}x`,
      selectionStart: first?.caret ?? 0,
      selectionEnd: first?.caret ?? 0,
      editedCell: { sheet: 0, row: 0, col: 1 },
      direction: 'left',
      extend: false,
      current: first,
    });
    expect(stale).toBeNull();
  });
});
