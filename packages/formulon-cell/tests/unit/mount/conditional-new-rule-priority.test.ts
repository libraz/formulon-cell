import { describe, expect, it } from 'vitest';
import { conditionalRuleToEngineInput, inputWithDxf } from '../../../src/engine/cf-writeback.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mountStubSheet } from '../../test-utils/mount.js';

/** Mounts over a real engine holding one loaded rule, then adds a second via Manage Rules → New Rule. */
async function addRuleThroughNewRuleDialog() {
  const wb = await WorkbookHandle.createDefault();
  expect(wb.isStub).toBe(false);
  let sheet: Awaited<ReturnType<typeof mountStubSheet>> | undefined;
  try {
    wb.setNumber({ sheet: 0, row: 0, col: 0 }, 50);
    const loaded = {
      kind: 'cell-value',
      range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
      op: '>',
      a: 10,
      apply: { fill: '#ffc7ce' },
    } as const;
    const input = conditionalRuleToEngineInput(loaded);
    if (!input) throw new Error('untranslatable rule');
    wb.addConditionalFormat(0, inputWithDxf(wb, input, loaded));
    const before = wb.getConditionalFormats(0);
    expect(before).toHaveLength(1);

    sheet = await mountStubSheet({ workbook: wb });
    sheet.instance.openCfRulesDialog();
    const newBtn = document.querySelector<HTMLButtonElement>('.fc-cfrulesdlg__new');
    if (!newBtn) throw new Error('missing New Rule button');
    newBtn.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    const valueInput = document.querySelector<HTMLInputElement>(
      '.fc-conddlg__sub input[type="text"]',
    );
    if (!valueInput) throw new Error('missing value input');
    valueInput.value = '25';
    document
      .querySelector<HTMLButtonElement>('.fc-conddlg__addrow .fc-fmtdlg__btn--primary')
      ?.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    return { before: before[0], after: wb.getConditionalFormats(0) };
  } finally {
    sheet?.dispose();
    wb.dispose();
  }
}

describe('conditional New Rule from Manage Rules', () => {
  it('adds the new rule next to the loaded one, keeping the loaded id', async () => {
    const { before, after } = await addRuleThroughNewRuleDialog();
    expect(after).toHaveLength(2);
    expect(after.some((r) => r.id === before?.id)).toBe(true);
    expect(after.filter((r) => r.id !== before?.id)).toHaveLength(1);
  });

  // Waits for a new rule to take priority 1 and shift the existing rules, as Manage Rules does.
  it.fails('gives the new rule priority 1 and shifts the loaded rule by one', async () => {
    const { before, after } = await addRuleThroughNewRuleDialog();
    const loaded = after.find((r) => r.id === before?.id);
    const added = after.find((r) => r.id !== before?.id);
    expect(added?.priority).toBe(1);
    expect(loaded?.priority).toBe((before?.priority ?? 0) + 1);
  });
});
