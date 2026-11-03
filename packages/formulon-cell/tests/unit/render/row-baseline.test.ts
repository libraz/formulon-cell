import { describe, expect, it } from 'vitest';
import { rowBaselineFonts } from '../../../src/render/row-baseline.js';
import type { CellFormat } from '../../../src/store/store.js';

describe('rowBaselineFonts', () => {
  it('flags a row whose bottom-aligned cell carries an accounting underline', () => {
    const formats = new Map<string, CellFormat>([
      ['0:5:1', { underline: 'singleAccounting' }],
      ['0:5:2', { fontSize: 20 }],
      ['0:7:1', { underline: 'doubleAccounting', vAlign: 'top' }],
      ['0:8:1', { underline: 'double' }],
    ]);
    const rows = rowBaselineFonts(formats, 0);
    expect(rows.get(5)).toMatchObject({ fontSize: 20, accountingUnderline: true });
    expect(rows.get(7)).toBeUndefined();
    expect(rows.get(8)).toBeUndefined();
  });
});
