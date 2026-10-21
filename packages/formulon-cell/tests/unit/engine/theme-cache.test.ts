import { describe, expect, it, vi } from 'vitest';

import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

const nativeOf = (handle: WorkbookHandle): { getTheme: () => unknown } =>
  (handle as unknown as { wb: { getTheme: () => unknown } }).wb;

describe('workbook theme cache', () => {
  it('reads the engine theme once for repeated theme-linked font resolution', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      const font = wb.getFontRecord(0);
      expect(font?.scheme === 1 || font?.scheme === 2).toBe(true);
      const spy = vi.spyOn(nativeOf(wb), 'getTheme');
      const first = wb.getResolvedFontRecord(0);
      for (let i = 0; i < 50; i += 1) wb.getResolvedFontRecord(0);
      expect(spy.mock.calls.length).toBeLessThanOrEqual(1);
      expect(wb.getResolvedFontRecord(0)).toEqual(first);
    } finally {
      wb.dispose();
    }
  });

  it('drops the cached theme when theme fonts or colours change', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      const theme = wb.getWorkbookTheme();
      expect(theme).not.toBeNull();
      if (!theme) return;
      const fonts = { ...theme.fonts, minorEastAsian: 'Meiryo', minorLatin: 'Meiryo' };
      expect(wb.setWorkbookThemeFonts(fonts)).toBe(true);
      expect(wb.getWorkbookTheme()?.fonts.minorLatin).toBe('Meiryo');
      const colors = [...theme.colors];
      colors[4] = 0xff112233;
      expect(wb.setWorkbookThemeColors(colors)).toBe(true);
      expect(wb.getWorkbookTheme()?.colors[4]).toBe(0xff112233);
      expect(wb.getWorkbookTheme()?.fonts.minorLatin).toBe('Meiryo');
    } finally {
      wb.dispose();
    }
  });

  it('returns a copy callers may mutate without poisoning the cache', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const theme = wb.getWorkbookTheme();
      if (!theme) return;
      const before = theme.fonts.minorLatin;
      theme.fonts.minorLatin = 'Mutated';
      expect(wb.getWorkbookTheme()?.fonts.minorLatin).toBe(before);
    } finally {
      wb.dispose();
    }
  });
});
