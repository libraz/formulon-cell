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

  it('names a theme-linked font by the profile script whatever its charset', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'zh', platform: 'mac' });
    try {
      const base = wb.getFontRecord(0);
      if (!base) throw new Error('font 0 missing');
      const index = wb.addFontRecord({ ...base, charset: 204 });
      expect(wb.getResolvedFontRecord(index)?.name).toBe('等线');
    } finally {
      wb.dispose();
    }
  });
});

describe('theme-linked fonts without a charset', () => {
  it("take the profile language's script face", async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'en' });
    try {
      const base = wb.getFontRecord(0);
      if (!base) throw new Error('font 0 missing');
      const index = wb.addFontRecord({ ...base, name: 'Calibri', scheme: 2, charset: null });
      expect(wb.getResolvedFontRecord(index)?.name).toBe('Aptos Narrow');
      expect(wb.setSpreadsheetProfileId('windows-ja_JP')).toBe(true);
      expect(wb.getResolvedFontRecord(index)?.name).toBe('游ゴシック');
    } finally {
      wb.dispose();
    }
  });
});
