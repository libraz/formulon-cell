import { describe, expect, it } from 'vitest';
import {
  clearIgnoredCellErrors,
  ignoreCellError,
  isCellErrorIgnored,
  restoreCellErrorIndicator,
  toggleCellErrorIgnored,
} from '../../../src/commands/error-indicators.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('error indicator commands', () => {
  it('ignores and restores one cell error indicator', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 2 };

    ignoreCellError(store, addr);
    ignoreCellError(store, addr);

    expect(isCellErrorIgnored(store, addr)).toBe(true);
    expect([...store.getState().errorIndicators.ignoredErrors]).toEqual(['0:1:2']);

    restoreCellErrorIndicator(store, addr);

    expect(isCellErrorIgnored(store, addr)).toBe(false);
    expect(store.getState().errorIndicators.ignoredErrors.size).toBe(0);
  });

  it('toggles the ignored state and clears all ignored cells', () => {
    const store = createSpreadsheetStore();
    const a1 = { sheet: 0, row: 0, col: 0 };
    const b2 = { sheet: 0, row: 1, col: 1 };

    expect(toggleCellErrorIgnored(store, a1)).toBe(true);
    expect(toggleCellErrorIgnored(store, a1)).toBe(false);
    expect(isCellErrorIgnored(store, a1)).toBe(false);

    ignoreCellError(store, a1);
    ignoreCellError(store, b2);
    clearIgnoredCellErrors(store);

    expect(store.getState().errorIndicators.ignoredErrors.size).toBe(0);
  });
});
