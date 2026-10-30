import { describe, expect, it } from 'vitest';
import { navigateFormulaReference } from '../../../src/interact/formula-edit-navigation.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// coverwise generate, strength=2, seed=42.
// Expected references are literal fixtures, independent of the navigation code.
const cases = [
  {
    style: 'a1',
    flow: 'ltr',
    selection: 'plain',
    visibility: 'shown',
    direction: 'left',
    ref: 'C4',
  },
  {
    style: 'r1c1',
    flow: 'rtl',
    selection: 'shift',
    visibility: 'hidden',
    direction: 'left',
    ref: 'RC:RC[2]',
  },
  {
    style: 'r1c1',
    flow: 'rtl',
    selection: 'plain',
    visibility: 'shown',
    direction: 'up',
    ref: 'R[-1]C',
  },
  {
    style: 'a1',
    flow: 'ltr',
    selection: 'shift',
    visibility: 'hidden',
    direction: 'up',
    ref: 'D2:D4',
  },
  {
    style: 'a1',
    flow: 'rtl',
    selection: 'shift',
    visibility: 'shown',
    direction: 'down',
    ref: 'D4:D5',
  },
  {
    style: 'r1c1',
    flow: 'ltr',
    selection: 'plain',
    visibility: 'shown',
    direction: 'right',
    ref: 'RC[1]',
  },
  {
    style: 'a1',
    flow: 'ltr',
    selection: 'shift',
    visibility: 'shown',
    direction: 'right',
    ref: 'D4:E4',
  },
  {
    style: 'r1c1',
    flow: 'rtl',
    selection: 'shift',
    visibility: 'hidden',
    direction: 'right',
    ref: 'RC[-2]:RC',
  },
  {
    style: 'a1',
    flow: 'ltr',
    selection: 'plain',
    visibility: 'shown',
    direction: 'down',
    ref: 'D5',
  },
  {
    style: 'r1c1',
    flow: 'rtl',
    selection: 'shift',
    visibility: 'shown',
    direction: 'down',
    ref: 'RC:R[1]C',
  },
  {
    style: 'a1',
    flow: 'ltr',
    selection: 'shift',
    visibility: 'hidden',
    direction: 'down',
    ref: 'D4:D6',
  },
  {
    style: 'a1',
    flow: 'ltr',
    selection: 'plain',
    visibility: 'hidden',
    direction: 'right',
    ref: 'F4',
  },
] as const;

describe('formula point navigation combinations', () => {
  it.each(cases)(
    '$style $flow $selection $visibility $direction keeps the surrounding formula and selection',
    ({ style, flow, selection, visibility, direction, ref }) => {
      const store = createSpreadsheetStore();
      store.setState((state) => ({
        ...state,
        ui: { ...state.ui, rightToLeft: flow === 'rtl' },
        layout: {
          ...state.layout,
          hiddenRows: new Set(visibility === 'hidden' ? [2, 4] : []),
          hiddenCols: new Set(visibility === 'hidden' ? [2, 4] : []),
        },
      }));
      const state = store.getState();
      const result = navigateFormulaReference({
        state,
        text: '=SUM()+7',
        selectionStart: 5,
        selectionEnd: 5,
        editedCell: { sheet: 0, row: 3, col: 3 },
        direction,
        extend: selection === 'shift',
        r1c1: style === 'r1c1',
      });

      expect(result?.value).toBe(`=SUM(${ref})+7`);
      expect(result?.caret).toBe(5 + ref.length);
      expect(store.getState()).toBe(state);
    },
  );
});
