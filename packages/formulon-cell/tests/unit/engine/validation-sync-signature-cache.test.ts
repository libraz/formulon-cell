import { afterEach, describe, expect, it, vi } from 'vitest';
import { syncValidationsToEngine } from '../../../src/engine/validation-sync.js';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type CellValidation, createSpreadsheetStore, mutators } from '../../../src/store/store.js';

afterEach(() => vi.restoreAllMocks());

describe('validation sync signatures', () => {
  it('serializes a validation shared by many cells once', () => {
    const store = createSpreadsheetStore();
    const validation: CellValidation = { kind: 'list', source: ['a', 'b'] };
    for (let row = 0; row < 20; row += 1) {
      mutators.setCellFormat(store, { sheet: 0, row, col: 0 }, { validation });
    }
    const wb = {
      capabilities: { dataValidation: true },
      clearValidations: () => true,
      addValidationEntry: vi.fn(() => true),
    } as unknown as WorkbookHandle;
    const stringify = vi.spyOn(JSON, 'stringify');

    syncValidationsToEngine(wb, store, 0);

    const forValidation = stringify.mock.calls.filter(([arg]) => arg === validation);
    expect(forValidation).toHaveLength(1);
    expect(wb.addValidationEntry).toHaveBeenCalledTimes(1);
  });
});
