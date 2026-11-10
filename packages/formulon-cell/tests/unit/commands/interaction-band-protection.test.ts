import { afterEach, describe, expect, it } from 'vitest';
import { InteractionAuthorizer } from '../../../src/commands/interaction-authorizer.js';
import { rowShiftBand } from '../../../src/commands/interaction-effects.js';
import type { InteractionPolicy } from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('Protect Sheet options on a protected sheet', () => {
  let workbook: WorkbookHandle | undefined;

  afterEach(() => {
    workbook?.dispose();
    workbook = undefined;
  });

  const setup = async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    workbook = wb;
    const store = createSpreadsheetStore();
    const authorizer = new InteractionAuthorizer({
      store,
      getWb: () => wb,
      isDisposed: () => false,
    });
    const insertRows = () =>
      authorizer.canExecute({
        operation: 'insertRows',
        origin: 'ribbon',
        effects: [{ kind: 'range', range: rowShiftBand(0, 4) }],
      });
    return { store, authorizer, insertRows };
  };

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', { defaultOperation: 'allow' }],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`follows the Protect Sheet insert-rows option ${label}`, async () => {
      const { store, authorizer, insertRows } = await setup();
      authorizer.policy = policy;
      expect(insertRows()).toEqual({ allowed: true });

      setProtectedSheet(store, 0, true);
      expect(insertRows()).toMatchObject({ allowed: false, code: 'protected' });

      setProtectedSheet(store, 0, true, { permissions: { insertRows: true } });
      expect(insertRows()).toEqual({ allowed: true });

      setProtectedSheet(store, 0, true, { permissions: { deleteRows: true } });
      expect(insertRows()).toMatchObject({ allowed: false, code: 'protected' });
    });
  }

  it('lets pivot report changes through a protected sheet only with the PivotTable option', async () => {
    const { store, authorizer } = await setup();
    authorizer.policy = { defaultOperation: 'allow' };
    const pivot = () =>
      authorizer.canExecute({
        operation: 'pivot',
        origin: 'ribbon',
        effects: [{ kind: 'range', range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 2 } }],
      });
    expect(pivot()).toEqual({ allowed: true });

    setProtectedSheet(store, 0, true);
    expect(pivot()).toMatchObject({ allowed: false, code: 'protected' });

    setProtectedSheet(store, 0, true, { permissions: { pivotTables: true } });
    expect(pivot()).toEqual({ allowed: true });

    authorizer.policy = { defaultOperation: 'allow', operations: { pivot: false } };
    expect(pivot()).toMatchObject({ allowed: false, code: 'operationDenied' });
  });
});
