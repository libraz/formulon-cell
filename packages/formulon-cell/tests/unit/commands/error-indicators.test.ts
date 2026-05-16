import { describe, expect, it } from 'vitest';
import {
  circleInvalidValidationData,
  circleInvalidValidationDataInSheet,
  clearIgnoredCellErrors,
  clearValidationCircles,
  formulaErrorCellsInRange,
  ignoreCellError,
  isCellErrorIgnored,
  recordIgnoredErrorsChange,
  recordValidationCirclesChange,
  restoreCellErrorIndicator,
  selectNextFormulaError,
  toggleCellErrorIgnored,
} from '../../../src/commands/error-indicators.js';
import { History } from '../../../src/commands/history.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

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

  it('records ignored formula error changes as undoable visual actions', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const addr = { sheet: 0, row: 1, col: 2 };

    recordIgnoredErrorsChange(history, store, () => {
      ignoreCellError(store, addr);
    });

    expect([...store.getState().errorIndicators.ignoredErrors]).toEqual(['0:1:2']);
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(store.getState().errorIndicators.ignoredErrors.size).toBe(0);

    history.redo();
    expect([...store.getState().errorIndicators.ignoredErrors]).toEqual(['0:1:2']);

    recordIgnoredErrorsChange(history, store, () => {
      clearIgnoredCellErrors(store);
    });
    expect(store.getState().errorIndicators.ignoredErrors.size).toBe(0);

    history.undo();
    expect([...store.getState().errorIndicators.ignoredErrors]).toEqual(['0:1:2']);
  });

  it('marks and clears invalid data-validation cells in a selected range', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 99 });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'number', value: 5 });

    const count = circleInvalidValidationData(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });

    expect(count).toBe(1);
    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:0:0']);

    clearValidationCircles(store);

    expect(store.getState().errorIndicators.validationCircles.size).toBe(0);
  });

  it('records validation circle changes as undoable visual actions', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 99 });

    recordValidationCirclesChange(history, store, () => {
      circleInvalidValidationData(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    });

    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:0:0']);
    expect(history.canUndo()).toBe(true);

    history.undo();
    expect(store.getState().errorIndicators.validationCircles.size).toBe(0);

    history.redo();
    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:0:0']);

    recordValidationCirclesChange(history, store, () => clearValidationCircles(store));
    expect(store.getState().errorIndicators.validationCircles.size).toBe(0);

    history.undo();
    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:0:0']);
  });

  it('marks invalid data-validation cells across the active sheet', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 8, col: 3 },
      { validation: { kind: 'list', source: ['Open', 'Closed'] } },
    );
    mutators.setCellFormat(
      store,
      { sheet: 1, row: 0, col: 0 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 5 });
    mutators.setCell(store, { sheet: 0, row: 8, col: 3 }, { kind: 'text', value: 'Hold' });
    mutators.setCell(store, { sheet: 1, row: 0, col: 0 }, { kind: 'number', value: 99 });

    const count = circleInvalidValidationDataInSheet(store, 0);

    expect(count).toBe(1);
    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:8:3']);
  });

  it('finds formula errors in range and selects the next error cell', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 1 }, '=1');
    mutators.setCell(
      store,
      { sheet: 0, row: 0, col: 1 },
      { kind: 'error', code: 7, text: '#DIV/0!' },
      '=1/0',
    );
    mutators.setCell(
      store,
      { sheet: 0, row: 1, col: 0 },
      { kind: 'text', value: '#REF!' },
      '=A999',
    );
    mutators.setCell(store, { sheet: 0, row: 1, col: 1 }, { kind: 'text', value: '#N/A' });
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });

    expect(formulaErrorCellsInRange(store).map((addr) => `${addr.row}:${addr.col}`)).toEqual([
      '0:1',
      '1:0',
    ]);

    expect(selectNextFormulaError(store)).toEqual({ sheet: 0, row: 0, col: 1 });
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 1 });
    expect(selectNextFormulaError(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 })).toEqual({
      sheet: 0,
      row: 1,
      col: 0,
    });
  });

  it('skips ignored formula errors when selecting the next error cell', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(
      store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'error', code: 7, text: '#DIV/0!' },
      '=1/0',
    );
    mutators.setCell(
      store,
      { sheet: 0, row: 0, col: 1 },
      { kind: 'error', code: 4, text: '#REF!' },
      '=Z99',
    );
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
    ignoreCellError(store, { sheet: 0, row: 0, col: 0 });

    expect(formulaErrorCellsInRange(store)).toEqual([{ sheet: 0, row: 0, col: 1 }]);
    expect(selectNextFormulaError(store)).toEqual({ sheet: 0, row: 0, col: 1 });
  });
});
