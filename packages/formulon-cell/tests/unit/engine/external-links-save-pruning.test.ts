import type { Workbook } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const rawOf = (wb: WorkbookHandle): Workbook => (wb as unknown as { wb: Workbook }).wb;

/** Two formulas on two books; the first is cleared, then the workbook is saved and reloaded. */
async function reloadAfterClearingFirstLink(): Promise<WorkbookHandle> {
  const wb = await WorkbookHandle.createDefault();
  try {
    const sheet = rawOf(wb).sheet(0);
    sheet.setFormula('A1', '=[A.xlsx]Sheet1!A1');
    sheet.setFormula('A2', '=[B.xlsx]Sheet1!A1');
    sheet.setBlank('A1');
    return await WorkbookHandle.loadBytes(wb.save());
  } finally {
    wb.dispose();
  }
}

describe('external links across save', () => {
  it('keeps the surviving formula pointing at its book after reload', async () => {
    const re = await reloadAfterClearingFirstLink();
    try {
      const sheet = rawOf(re).sheet(0);
      expect(sheet.getFormula('A1')).toBeNull();
      expect(sheet.getFormula('A2')).toBe('=[B.xlsx]Sheet1!A1');
    } finally {
      re.dispose();
    }
  });

  // Waits for save to drop a link no formula uses and renumber the rest.
  it.fails('lists only the book still in use, renumbered to 1', async () => {
    const re = await reloadAfterClearingFirstLink();
    try {
      const links = re.getExternalLinks();
      expect(links.map((l) => l.target)).toEqual(['B.xlsx']);
      expect(links[0]?.index).toBe(1);
    } finally {
      re.dispose();
    }
  });
});
