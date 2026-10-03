import { describe, expect, it } from 'vitest';
import { ignoreCellError } from '../../../src/commands/error-indicators.js';
import { History } from '../../../src/commands/history.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachNavigationPolicy } from '../../../src/interact/navigation-policy.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import type { ApplyRibbonCommandDeps } from '../../../src/toolbar/ribbon/apply-ribbon-command.js';
import {
  dispatchMacRibbonCommand,
  isMacRibbonCommandSupported,
} from '../../../src/toolbar/ribbon/mac/dispatch.js';

describe('Mac ribbon command dispatch', () => {
  it('checks formula errors across the active sheet and skips navigation-restricted cells', () => {
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
    mutators.setCell(
      store,
      { sheet: 0, row: 0, col: 2 },
      { kind: 'error', code: 6, text: '#N/A' },
      '=MissingName',
    );
    mutators.setCell(
      store,
      { sheet: 1, row: 0, col: 0 },
      { kind: 'error', code: 1, text: '#DIV/0!' },
      '=2/0',
    );
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    ignoreCellError(store, { sheet: 0, row: 0, col: 0 });

    const workbook = { sheetCount: 2 } as WorkbookHandle;
    const navigation = attachNavigationPolicy(store, () => workbook, {
      selectable: (addr) => addr.col !== 1,
    });
    const host = document.createElement('div');
    const instance = { store, host } as unknown as SpreadsheetInstance;
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;
    try {
      expect(dispatchMacRibbonCommand('mac.formulas.errorCheck.run', deps, () => false)).toBe(true);
      expect(store.getState().selection.active).toEqual({ sheet: 0, row: 0, col: 2 });
    } finally {
      navigation.dispose();
    }
  });

  it('routes all native shape and AutoSum gallery leaves to their backed operations', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const formulaWrites: string[] = [];
    const workbook = {
      setFormula: (_addr: unknown, formula: string) => formulaWrites.push(formula),
    } as unknown as WorkbookHandle;
    const instance = {
      store,
      history,
      host: document.createElement('div'),
      workbook,
    } as unknown as SpreadsheetInstance;
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;
    const shapes = {
      'mac.insert.shapeLine': 'line',
      'mac.insert.shapeArrow': 'arrow',
      'mac.insert.shapeRectangle': 'rectangle',
      'mac.insert.shapeRoundedRectangle': 'rounded-rectangle',
      'mac.insert.shapeOval': 'oval',
      'mac.insert.shapeTriangle': 'triangle',
      'mac.insert.shapeDiamond': 'diamond',
    } as const;

    for (const [id, shape] of Object.entries(shapes)) {
      expect(isMacRibbonCommandSupported(id), id).toBe(true);
      expect(
        dispatchMacRibbonCommand(id, deps, () => false),
        id,
      ).toBe(true);
      expect(store.getState().illustrations.illustrations.at(-1)?.shape, id).toBe(shape);
    }

    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'number', value: 10 });
    mutators.setActive(store, { sheet: 0, row: 1, col: 0 });
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    const functions = {
      'mac.autosum.SUM': 'SUM',
      'mac.autosum.AVERAGE': 'AVERAGE',
      'mac.autosum.COUNT': 'COUNT',
      'mac.autosum.MAX': 'MAX',
      'mac.autosum.MIN': 'MIN',
    } as const;
    for (const [id, fn] of Object.entries(functions)) {
      expect(isMacRibbonCommandSupported(id), id).toBe(true);
      expect(
        dispatchMacRibbonCommand(id, deps, () => false),
        id,
      ).toBe(true);
      expect(formulaWrites.at(-1), id).toBe(`=${fn}(A1:A1)`);
    }
  });

  it('supports and dispatches dynamic function leaves only from the current live catalog', () => {
    const store = createSpreadsheetStore();
    const opened: string[] = [];
    let liveNames: readonly string[] | null = ['ACOS'];
    const workbook = {
      functionNames: () => liveNames,
    } as unknown as WorkbookHandle;
    const instance = {
      store,
      host: document.createElement('div'),
      workbook,
      openFunctionArguments: (name: string) => opened.push(name),
    } as unknown as SpreadsheetInstance;
    const deps = {
      inst: instance,
      runtime: { projectFormatToolbar: () => undefined },
    } as unknown as ApplyRibbonCommandDeps;

    expect(isMacRibbonCommandSupported('mac.function.ACOS', new Set(['ACOS']))).toBe(true);
    expect(isMacRibbonCommandSupported('mac.function.ACOS', new Set(['SUM']))).toBe(false);
    expect(dispatchMacRibbonCommand('mac.function.ACOS', deps, () => false)).toBe(true);
    expect(opened).toEqual(['ACOS']);
    liveNames = ['SUM'];
    expect(dispatchMacRibbonCommand('mac.function.ACOS', deps, () => false)).toBe(false);
    expect(opened).toEqual(['ACOS']);
  });
});
