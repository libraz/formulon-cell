import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { executeRibbonFormulaAuditingAction } from '../../../src/commands/ribbon-formula-auditing.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const strings = { errorChecking: 'Error Checking' };

const workbook = (): WorkbookHandle => ({}) as WorkbookHandle;

describe('executeRibbonFormulaAuditingAction', () => {
  it('delegates traceError to the host so visual tracing can use viewport state', () => {
    const store = createSpreadsheetStore();

    expect(
      executeRibbonFormulaAuditingAction({
        store,
        workbook: workbook(),
        history: new History(),
        action: 'traceError',
        strings,
      }),
    ).toEqual({ kind: 'trace-precedents' });
  });

  it('selects the next formula error for error checking and reports when all clear', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 1 }, '=1');
    mutators.setCell(
      store,
      { sheet: 0, row: 1, col: 0 },
      { kind: 'error', code: 7, text: '#DIV/0!' },
      '=1/0',
    );
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 0 });

    expect(
      executeRibbonFormulaAuditingAction({
        store,
        workbook: workbook(),
        history: new History(),
        action: 'errorChecking',
        strings,
      }),
    ).toEqual({ kind: 'mutated' });
    expect(store.getState().selection.active).toEqual({ sheet: 0, row: 1, col: 0 });

    mutators.ignoreError(store, { sheet: 0, row: 1, col: 0 });

    expect(
      executeRibbonFormulaAuditingAction({
        store,
        workbook: workbook(),
        history: new History(),
        action: 'errorChecking',
        strings,
      }),
    ).toEqual({ kind: 'report', report: { title: 'Error Checking', items: [] } });
  });

  it('ignores an active formula error as a single undoable visual change', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setCell(
      store,
      { sheet: 0, row: 0, col: 0 },
      { kind: 'text', value: '#REF!' },
      '=MissingName',
    );

    expect(
      executeRibbonFormulaAuditingAction({
        store,
        workbook: workbook(),
        history,
        action: 'ignoreError',
        strings,
      }),
    ).toEqual({ kind: 'mutated' });
    expect([...store.getState().errorIndicators.ignoredErrors]).toEqual(['0:0:0']);

    expect(history.undo()).toBe(true);
    expect(store.getState().errorIndicators.ignoredErrors.size).toBe(0);
  });

  it('circles invalid validation cells only in the current selection and can undo it', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 });
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
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 1, col: 0 },
      { validation: { kind: 'whole', op: 'between', a: 1, b: 10 } },
    );
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 99 });
    mutators.setCell(store, { sheet: 0, row: 0, col: 1 }, { kind: 'number', value: 5 });
    mutators.setCell(store, { sheet: 0, row: 1, col: 0 }, { kind: 'number', value: 99 });

    expect(
      executeRibbonFormulaAuditingAction({
        store,
        workbook: workbook(),
        history,
        action: 'circleInvalid',
        strings,
      }),
    ).toEqual({ kind: 'mutated' });
    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:0:0']);

    expect(history.undo()).toBe(true);
    expect(store.getState().errorIndicators.validationCircles.size).toBe(0);
  });

  it('clears validation circles as an undoable action', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setValidationCircles(store, new Set(['0:0:0', '0:1:0']));

    expect(
      executeRibbonFormulaAuditingAction({
        store,
        workbook: workbook(),
        history,
        action: 'clearCircles',
        strings,
      }),
    ).toEqual({ kind: 'mutated' });
    expect(store.getState().errorIndicators.validationCircles.size).toBe(0);

    expect(history.undo()).toBe(true);
    expect([...store.getState().errorIndicators.validationCircles]).toEqual(['0:0:0', '0:1:0']);
  });
});
