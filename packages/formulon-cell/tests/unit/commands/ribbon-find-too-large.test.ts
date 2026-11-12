import { describe, expect, it } from 'vitest';
import {
  executeRibbonFindAction,
  MAX_GOTO_SPECIAL_CELLS,
} from '../../../src/commands/goto-special.js';
import { addrKey } from '../../../src/engine/address.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('Find & Select over too many cells', () => {
  it('refuses as too large instead of reporting', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    try {
      const store = createSpreadsheetStore();
      // More validated cells than a selection can hold, filled straight into the store.
      store.setState((state) => {
        const cells = new Map(state.data.cells);
        for (let row = 0; row <= MAX_GOTO_SPECIAL_CELLS; row += 1) {
          cells.set(addrKey({ sheet: 0, row, col: 0 }), {
            value: { kind: 'number', value: row },
            formula: null,
          });
        }
        return {
          ...state,
          data: { ...state.data, cells },
          format: {
            ...state.format,
            rangeValidations: [
              {
                ranges: [{ sheet: 0, r0: 0, c0: 0, r1: MAX_GOTO_SPECIAL_CELLS, c1: 0 }],
                validation: { kind: 'whole', op: 'between', a: 0, b: 1_000_000 },
              },
            ],
          },
        };
      });
      const result = executeRibbonFindAction({
        store,
        workbook: wb,
        action: 'data-validation',
        strings: { findSelect: 'Find & Select', findNoMatches: 'No matches', commentNone: 'None' },
      });
      expect(result).toEqual({ kind: 'refused', reason: 'too-large' });
    } finally {
      wb.dispose();
    }
  });
});
