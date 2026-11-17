import { describe, expect, it } from 'vitest';
import { updatePivotDataField } from '../../../src/interact/pivot-field-settings-dialog-apply.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';
import { loadFixture } from '../engine/pivot-excel-parity/fixture.js';

describe('a value field with an empty caption', () => {
  it('keeps the empty caption through a number-format change', async () => {
    const wb = await loadFixture('pivot_blank_item.xlsx');
    try {
      const current = wb.getPivotDataField(0, 0, 0);
      if (!current) throw new Error('no data field');
      // The spreadsheet saves an emptied caption as name="".
      expect(wb.setPivotDataField(0, 0, 0, { ...current, name: '' })).toBe(true);
      const ctx = {
        wb,
        store: createSpreadsheetStore(),
        origin: 'contextMenu' as const,
        commandId: 'pivotValueFieldSettings',
        onChanged: () => {},
      };
      expect(
        updatePivotDataField(
          ctx,
          { sheetIndex: 0, pivotIndex: 0 },
          0,
          { numberFormat: '#,##0' },
          false,
        ),
      ).toBe(true);
      expect(wb.getPivotDataField(0, 0, 0)?.name).toBe('');
    } finally {
      wb.dispose();
    }
  });
});
