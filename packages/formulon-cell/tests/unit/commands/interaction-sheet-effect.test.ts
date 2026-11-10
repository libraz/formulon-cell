import { afterEach, describe, expect, it } from 'vitest';
import { InteractionAuthorizer } from '../../../src/commands/interaction-authorizer.js';
import type {
  InteractionOperation,
  InteractionPolicy,
} from '../../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import type { Range } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

describe('sheet effects', () => {
  let workbook: WorkbookHandle | undefined;

  afterEach(() => {
    workbook?.dispose();
    workbook = undefined;
  });

  const setup = async (bounds?: Range) => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    workbook = wb;
    wb.addSheet('Second');
    const store = createSpreadsheetStore();
    const authorizer = new InteractionAuthorizer({
      store,
      getWb: () => wb,
      getBounds: bounds ? () => bounds : undefined,
      isDisposed: () => false,
    });
    const decide = (operation: InteractionOperation, sheet: number) =>
      authorizer.canExecute({ operation, origin: 'ribbon', effects: [{ kind: 'sheet', sheet }] });
    return { wb, store, authorizer, decide };
  };

  for (const [label, policy] of [
    ['without a policy', undefined],
    ['under an allow-all policy', { defaultOperation: 'allow' }],
  ] as const satisfies readonly (readonly [string, InteractionPolicy | undefined])[]) {
    it(`refuses a sheet the workbook does not have ${label}`, async () => {
      const { authorizer, decide } = await setup();
      authorizer.policy = policy;
      expect(decide('sheet', 1)).toEqual({ allowed: true });
      expect(decide('sheet', 2)).toMatchObject({ allowed: false, code: 'outOfBounds' });
      expect(decide('sheet', -1)).toMatchObject({ allowed: false, code: 'outOfBounds' });
      expect(decide('sheet', 0.5)).toMatchObject({ allowed: false, code: 'outOfBounds' });
    });

    it(`exempts unprotect and page setup from sheet protection ${label}`, async () => {
      const { store, authorizer, decide } = await setup();
      authorizer.policy = policy;
      setProtectedSheet(store, 0, true);
      expect(decide('sheet', 0)).toMatchObject({ allowed: false, code: 'protected' });
      expect(decide('protection', 0)).toEqual({ allowed: true });
      expect(decide('pageSetup', 0)).toEqual({ allowed: true });
      expect(decide('sheet', 1)).toEqual({ allowed: true });
    });
  }

  it('follows a Protect Sheet option the operation has', async () => {
    const { store, authorizer, decide } = await setup();
    authorizer.policy = { defaultOperation: 'allow' };
    setProtectedSheet(store, 0, true);
    expect(decide('insertRows', 0)).toMatchObject({ allowed: false, code: 'protected' });
    setProtectedSheet(store, 0, true, { permissions: { insertRows: true } });
    expect(decide('insertRows', 0)).toEqual({ allowed: true });
  });

  it('keeps a bounded controller on its bounds sheet', async () => {
    const { authorizer, decide } = await setup({ sheet: 0, r0: 0, c0: 0, r1: 9, c1: 9 });
    authorizer.policy = { defaultOperation: 'allow' };
    expect(decide('sheet', 0)).toEqual({ allowed: true });
    expect(decide('sheet', 1)).toMatchObject({ allowed: false, code: 'outOfBounds' });
  });

  it('lets the policy decide the operation without editable-cell eligibility', async () => {
    const { authorizer, decide } = await setup();
    authorizer.policy = {
      defaultOperation: 'deny',
      operations: { sheet: true },
      editable: [{ sheet: 0, r0: 5, c0: 5, r1: 5, c1: 5 }],
    };
    expect(decide('sheet', 0)).toEqual({ allowed: true });
    expect(decide('pageSetup', 0)).toMatchObject({ allowed: false, code: 'operationDenied' });

    authorizer.policy = { readOnly: true, defaultOperation: 'allow' };
    expect(decide('protection', 0)).toMatchObject({ allowed: false, code: 'readOnly' });
  });

  it('preflights a sheet effect alongside cell effects', async () => {
    const { store, authorizer } = await setup();
    authorizer.policy = { defaultOperation: 'allow' };
    setProtectedSheet(store, 1, true);
    const canExecute = authorizer.canExecute.bind(authorizer);
    const intent = (sheet: number) => ({
      operation: 'sheet' as const,
      origin: 'undo' as const,
      effects: [
        { kind: 'cells' as const, cells: [{ sheet: 0, row: 0, col: 0 }] },
        { kind: 'sheet' as const, sheet },
      ],
    });
    expect(authorizer.preflightIntent(intent(0), canExecute)).toEqual({ allowed: true });
    expect(authorizer.preflightIntent(intent(1), canExecute)).toMatchObject({
      allowed: false,
      code: 'protected',
    });
  });
});
