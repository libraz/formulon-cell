import { ComparisonOperator, type Workbook } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { patchZipText } from './xlsx-zip-patch.js';

const rawOf = (wb: WorkbookHandle): Workbook => (wb as unknown as { wb: Workbook }).wb;

/** Saves a workbook whose one rule covers A1:A2, then rewrites the rule's sqref to `sqref`. */
async function loadWithSqref(sqref: string): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  try {
    for (let r = 0; r < 10; r++) wb.setNumber({ sheet: 0, row: r, col: 0 }, r);
    wb.addConditionalFormat(0, {
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 0 }],
      type: 1,
      operator: ComparisonOperator.GreaterThan,
      formula1: '1',
    });
    const saved = wb.save();
    const patched = patchZipText(saved, 'xl/worksheets/sheet1.xml', (xml) => {
      expect(xml).toContain('sqref="A1:A2"');
      return xml.replace('sqref="A1:A2"', `sqref="${sqref}"`);
    });
    return await WorkbookHandle.loadBytes(patched);
  } finally {
    wb.dispose();
  }
}

const rangesOf = async (sqref: string) => {
  const wb = await loadWithSqref(sqref);
  try {
    const rules = rawOf(wb).sheet(0).listCfRules();
    expect(rules).toHaveLength(1);
    return rules[0]?.ranges ?? [];
  } finally {
    wb.dispose();
  }
};

describe('conditional-format sqref on load', () => {
  it('keeps a single-range sqref as written', async () => {
    expect(await rangesOf('A1:A10')).toEqual([
      { firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 },
    ]);
  });

  it('reads tiling pieces "A1:A5 A6:A10" back as one joined range', async () => {
    expect(await rangesOf('A1:A5 A6:A10')).toEqual([
      { firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 },
    ]);
  });

  it('reads pieces "D2:D3 B4" back ordered by top row', async () => {
    expect(await rangesOf('D2:D3 B4')).toEqual([
      { firstRow: 1, firstCol: 3, lastRow: 2, lastCol: 3 },
      { firstRow: 3, firstCol: 1, lastRow: 3, lastCol: 1 },
    ]);
  });
});
