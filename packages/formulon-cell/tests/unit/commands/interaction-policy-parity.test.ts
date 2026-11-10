import { afterEach, describe, expect, it } from 'vitest';

import { InteractionAuthorizer } from '../../../src/commands/interaction-authorizer.js';
import {
  INTERACTION_OPERATIONS,
  type InteractionOperation,
  type InteractionPolicy,
  policyPermits,
  viewerPolicy,
} from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const policies = (): readonly (readonly [string, InteractionPolicy])[] => [
  ['allow-all', { defaultOperation: 'allow' }],
  ['deny-all', { defaultOperation: 'deny' }],
  ['readOnly', { readOnly: true, defaultOperation: 'allow' }],
  ['viewer', viewerPolicy()],
  ...INTERACTION_OPERATIONS.map((op): readonly [string, InteractionPolicy] => [
    `deny ${op}`,
    { defaultOperation: 'allow', operations: { [op]: false } },
  ]),
];

describe('policyPermits and the authorizer agree', () => {
  let workbook: WorkbookHandle | undefined;

  afterEach(() => {
    workbook?.dispose();
    workbook = undefined;
  });

  it('decides every operation the same way under every policy shape', async () => {
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    workbook = wb;
    const authorizer = new InteractionAuthorizer({
      store: createSpreadsheetStore(),
      getWb: () => wb,
      isDisposed: () => false,
    });
    const mismatches: string[] = [];
    for (const [label, policy] of policies()) {
      authorizer.policy = policy;
      for (const operation of INTERACTION_OPERATIONS) {
        const decided = authorizer.canExecute({
          operation,
          origin: 'instanceApi',
          effects: [{ kind: 'workbook' }],
        }).allowed;
        if (policyPermits(policy, operation) !== decided)
          mismatches.push(`${label}: ${operation} permits=${!decided} authorizer=${decided}`);
      }
    }
    expect(mismatches).toEqual([]);
  });

  it('lists every operation exactly once', () => {
    const seen = new Set<InteractionOperation>(INTERACTION_OPERATIONS);
    expect(seen.size).toBe(INTERACTION_OPERATIONS.length);
    expect(seen.has('table')).toBe(true);
  });
});
