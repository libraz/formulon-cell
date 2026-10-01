import type { Page } from '@playwright/test';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { SpreadsheetPage } from '../../pages/SpreadsheetPage.js';
import { runFormulaScenario } from '../../scenarios/edit.js';

vi.mock('@playwright/test', () => ({ expect }));

describe('formula E2E scenario instructions', () => {
  beforeEach(() => {
    vi.spyOn(SpreadsheetPage.prototype, 'mount').mockResolvedValue();
    vi.spyOn(SpreadsheetPage.prototype, 'expectNoStub').mockResolvedValue();
    vi.spyOn(SpreadsheetPage.prototype, 'focusHost').mockResolvedValue();
    vi.spyOn(SpreadsheetPage.prototype, 'formulaBarValue').mockResolvedValue('=SUM(A1:A3)');
  });

  afterEach(() => vi.restoreAllMocks());

  const browserInstructions = (total: number) => {
    const type = vi.fn().mockResolvedValue(undefined);
    const press = vi.fn().mockResolvedValue(undefined);
    const evaluate = vi.fn().mockResolvedValue({ kind: 'number', value: total });
    const page = { keyboard: { type, press }, evaluate } as unknown as Page;
    return { page, type, press };
  };

  it('enters consecutive rows without resetting the selection to A1', async () => {
    const { page, type, press } = browserInstructions(6);
    await runFormulaScenario(page);

    expect(SpreadsheetPage.prototype.focusHost).toHaveBeenCalledTimes(1);
    expect(type.mock.calls).toEqual([['1'], ['2'], ['3'], ['=SUM(A1:A3)']]);
    expect(press.mock.calls).toEqual([['Enter'], ['Enter'], ['Enter'], ['Enter'], ['ArrowUp']]);
  });

  it('rejects an incorrect result even when the formula text is correct', async () => {
    const { page } = browserInstructions(0);
    await expect(runFormulaScenario(page)).rejects.toThrow();
  });
});
