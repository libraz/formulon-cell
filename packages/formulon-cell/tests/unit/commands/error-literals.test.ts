import { describe, expect, it, vi } from 'vitest';
import { cellValueIsFormulaError } from '../../../src/commands/error-indicators.js';
import { findMatchingRanges } from '../../../src/commands/goto-special.js';
import { ERROR_LITERALS } from '../../../src/engine/value.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const LITERALS = ['#CIRCULAR!', '#SPILL!', '#CALC!', '#GETTING_DATA', '#DIV/0!', '#N/A'];

describe('error literals', () => {
  it('has one canonical set covering the engine names and #CIRCULAR!', () => {
    for (const literal of LITERALS) expect(ERROR_LITERALS.has(literal)).toBe(true);
    expect(ERROR_LITERALS.has('#TODO!')).toBe(false);
  });

  it('error indicators recognize every literal', () => {
    for (const literal of LITERALS) {
      expect(cellValueIsFormulaError({ kind: 'text', value: literal })).toBe(true);
    }
  });

  it('go-to-special constants-with-errors recognizes every literal', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const store = createSpreadsheetStore();
    const cells = new Map(store.getState().data.cells);
    LITERALS.forEach((literal, row) => {
      wb.setText({ sheet: 0, row, col: 0 }, literal);
      cells.set(addrKey({ sheet: 0, row, col: 0 }), {
        value: { kind: 'text', value: literal },
        formula: null,
      });
    });
    store.setState((s) => ({
      ...s,
      data: { ...s.data, cells },
      selection: { ...s.selection, range: { sheet: 0, r0: 0, c0: 0, r1: 9, c1: 0 } },
    }));
    const result = findMatchingRanges(wb, store, 'constants', {
      filters: { errors: true, text: false, numbers: false, logical: false },
    });
    expect(result).toEqual({
      status: 'found',
      ranges: [{ sheet: 0, r0: 0, c0: 0, r1: LITERALS.length - 1, c1: 0 }],
    });
  });

  it('go-to-special reads the sheet cells only for kinds that need them', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    const cellsSpy = vi.spyOn(wb, 'cells');
    const store = createSpreadsheetStore();
    findMatchingRanges(wb, store, 'current-region');
    expect(cellsSpy).not.toHaveBeenCalled();
  });
});
