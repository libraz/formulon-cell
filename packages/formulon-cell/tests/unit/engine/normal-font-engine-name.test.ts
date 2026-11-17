// @vitest-environment node

import { describe, expect, it, vi } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

describe('Mac layout metrics take the Normal font name from the engine', () => {
  it('does not re-resolve a theme-linked Normal font against the theme', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      const expected = wb.getSheetDefaultLayoutPx(0);
      expect(expected?.headerDigitAdvance).toBeDefined();
      const record = wb.getFontRecord(0);
      if (!record) throw new Error('font 0 missing');
      vi.spyOn(wb, 'getResolvedFontRecord').mockReturnValue({ ...record, name: 'Unlisted Face' });
      expect(wb.getSheetDefaultLayoutPx(0)).toEqual(expected);
    } finally {
      wb.dispose();
    }
  });
});
