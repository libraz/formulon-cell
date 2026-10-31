import { describe, expect, it } from 'vitest';
import { addrKey } from '../../../src/engine/address.js';
import { navigateFormulaReference } from '../../../src/interact/formula-edit-navigation.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

// coverwise generate, strength=2, seed=42: all 36 parameter pairs covered.
const cases = [
  { flow: 'ltr', style: 'a1', selection: 'plain', direction: 'up', ref: 'D2' },
  { flow: 'rtl', style: 'r1c1', selection: 'plain', direction: 'left', ref: 'RC[2]' },
  { flow: 'ltr', style: 'r1c1', selection: 'plain', direction: 'down', ref: 'R[2]C' },
  { flow: 'rtl', style: 'a1', selection: 'shift', direction: 'right', ref: 'B4:D4' },
  { flow: 'ltr', style: 'a1', selection: 'plain', direction: 'left', ref: 'B4' },
  { flow: 'rtl', style: 'a1', selection: 'shift', direction: 'down', ref: 'D4:D6' },
  { flow: 'rtl', style: 'r1c1', selection: 'plain', direction: 'up', ref: 'R[-2]C' },
  { flow: 'rtl', style: 'r1c1', selection: 'plain', direction: 'right', ref: 'RC[-2]' },
  { flow: 'ltr', style: 'r1c1', selection: 'shift', direction: 'up', ref: 'R[-2]C:RC' },
  { flow: 'rtl', style: 'r1c1', selection: 'shift', direction: 'left', ref: 'RC:RC[2]' },
  { flow: 'ltr', style: 'r1c1', selection: 'shift', direction: 'right', ref: 'RC:RC[2]' },
] as const;

describe('formula data-edge navigation combinations', () => {
  it.each(cases)(
    '$flow $style $selection $direction retains the expression and selection',
    (scenario) => {
      const store = createSpreadsheetStore();
      store.setState((state) => {
        const cells = new Map(state.data.cells);
        for (const [row, col] of [
          [3, 3],
          [1, 3],
          [2, 3],
          [4, 3],
          [5, 3],
          [3, 1],
          [3, 2],
          [3, 4],
          [3, 5],
          [3, 7],
          [7, 3],
        ] as const) {
          cells.set(addrKey({ sheet: 0, row, col }), {
            value: { kind: 'number', value: 1 },
            formula: null,
          });
        }
        return {
          ...state,
          data: { ...state.data, cells },
          ui: { ...state.ui, rightToLeft: scenario.flow === 'rtl' },
        };
      });
      const state = store.getState();
      const result = navigateFormulaReference({
        state,
        text: '=SUM()+7',
        selectionStart: 5,
        selectionEnd: 5,
        editedCell: { sheet: 0, row: 3, col: 3 },
        direction: scenario.direction,
        extend: scenario.selection === 'shift',
        r1c1: scenario.style === 'r1c1',
        jump: true,
      });
      expect(result?.value).toBe(`=SUM(${scenario.ref})+7`);
      expect(result?.caret).toBe(5 + scenario.ref.length);
      expect(store.getState()).toBe(state);
    },
  );
});
