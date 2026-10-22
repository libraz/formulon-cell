// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

describe('getResolvedFontRecord theme script faces', () => {
  it.each([
    ['ja', '游ゴシック'],
    ['zh', '等线'],
    ['zh-TW', '新細明體'],
    ['ko', '맑은 고딕'],
    ['th', 'Tahoma'],
    ['en', 'Aptos Narrow'],
  ])('resolves the %s Mac Normal font to %s', async (locale, name) => {
    const wb = await WorkbookHandle.createDefault({ locale, platform: 'mac' });
    try {
      expect(wb.getResolvedFontRecord(0)?.name).toBe(name);
    } finally {
      wb.dispose();
    }
  });

  it('takes the major script face for a major-scheme font', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'zh', platform: 'mac' });
    try {
      const base = wb.getFontRecord(0);
      if (!base) throw new Error('font 0 missing');
      expect(base.charset).toBe(134);
      const index = wb.addFontRecord({ ...base, scheme: 1 });
      expect(wb.getResolvedFontRecord(index)?.name).toBe('等线 Light');
    } finally {
      wb.dispose();
    }
  });

  it('keeps the latin face for a charset without a theme script', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'zh', platform: 'mac' });
    try {
      const base = wb.getFontRecord(0);
      if (!base) throw new Error('font 0 missing');
      const index = wb.addFontRecord({ ...base, charset: 204 });
      expect(wb.getResolvedFontRecord(index)?.name).toBe('Aptos Narrow');
    } finally {
      wb.dispose();
    }
  });
});
