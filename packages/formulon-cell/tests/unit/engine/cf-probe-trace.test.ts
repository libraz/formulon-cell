import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  conditionalRuleToEngineInput,
  engineTakesConditionalRule,
  inputWithDxf,
} from '../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import type { ConditionalRule } from '../../../src/store/store.js';

const range = { sheet: 0, r0: 0, c0: 0, r1: 3, c1: 0 };
const accepted: ConditionalRule = {
  kind: 'formula',
  range,
  formula: '=A1>1',
  apply: { fill: '#ffc7ce' },
};
const refused: ConditionalRule = { ...accepted, formula: '=SUM()' };
const real: ConditionalRule = {
  kind: 'cell-value',
  range,
  op: '>',
  a: 10,
  apply: { fill: '#c6efce' },
};

const addReal = (wb: WorkbookHandle): void => {
  const input = conditionalRuleToEngineInput(real);
  if (!input) throw new Error('untranslatable rule');
  wb.addConditionalFormat(0, inputWithDxf(wb, input, real));
};

describe('engineTakesConditionalRule leaves the engine as it found it', () => {
  const opened: WorkbookHandle[] = [];
  const fresh = async (): Promise<WorkbookHandle> => {
    const wb = await WorkbookHandle.createDefault();
    opened.push(wb);
    return wb;
  };
  let wb: WorkbookHandle;

  beforeEach(async () => {
    wb = await fresh();
  });
  afterEach(() => {
    for (const w of opened.splice(0)) w.dispose();
  });

  it.each([
    ['an accepted rule', accepted, true],
    ['a refused rule', refused, false],
  ])('announces nothing and recalculates nothing for %s', (_name, rule, expected) => {
    addReal(wb);
    const events: string[] = [];
    const offChange = wb.subscribe((e) => events.push(e.kind));
    const offFacade = wb.subscribeFacadeChanges((c) => events.push(`facade:${c.kind}`));
    expect(engineTakesConditionalRule(wb, rule)).toBe(expected);
    offChange();
    offFacade();
    expect(events).toEqual([]);
  });

  it.each([
    ['an accepted rule', accepted],
    ['a refused rule', refused],
  ])('leaves the save diagnostics unchanged after %s', (_name, rule) => {
    addReal(wb);
    const before = wb.saveWithDiagnostics().diagnostics;
    engineTakesConditionalRule(wb, rule);
    expect(wb.saveWithDiagnostics().diagnostics).toEqual(before);
  });

  it('gives the next real rule the same id and priority, and saves the same bytes', async () => {
    const probed = wb;
    const plain = await fresh();
    engineTakesConditionalRule(probed, accepted);
    engineTakesConditionalRule(probed, refused);
    addReal(probed);
    addReal(plain);

    expect(probed.getConditionalFormats(0)).toEqual(plain.getConditionalFormats(0));
    expect(probed.dxfCount()).toBe(plain.dxfCount());
    expect(probed.save()).toEqual(plain.save());
  });
});
