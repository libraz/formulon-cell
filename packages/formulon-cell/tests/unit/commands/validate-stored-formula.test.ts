import { describe, expect, it } from 'vitest';
import { nativeValidationOutcome } from '../../../src/commands/validate.js';
import { syncValidationsToEngine } from '../../../src/engine/validation-sync.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type CellValidation, createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('a data-validation rule written by the session', () => {
  it('takes the engine verdict when the engine respells its bound formula', async () => {
    const wb = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    wb.setNumber({ sheet: 0, row: 0, col: 1 }, 2);
    const addr = { sheet: 0, row: 0, col: 0 };
    // At least B1 * 1.50 (= 3); the engine stores the literal as 1.5.
    const validation: CellValidation = { kind: 'decimal', op: '>=', a: 0, aFormula: 'B1*1.50' };
    mutators.setCellFormat(store, addr, { validation });
    syncValidationsToEngine(wb, store, 0);
    const stored = wb.getValidationsForSheet(0)[0]?.formula1;
    expect(stored).toBeDefined();
    const outcome = nativeValidationOutcome(wb, addr, { kind: 'number', value: 1 }, validation);
    expect(outcome).toMatchObject({ ok: false });
    expect(nativeValidationOutcome(wb, addr, { kind: 'number', value: 4 }, validation)).toEqual({
      ok: true,
    });
    wb.dispose();
  });
});
