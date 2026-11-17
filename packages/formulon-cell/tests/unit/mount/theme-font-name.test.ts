import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/index.js';

describe('theme-linked font names on a mounted workbook', () => {
  let sheet: MountedStubSheet | undefined;
  afterEach(() => {
    sheet?.workbook.dispose();
    sheet?.dispose();
    sheet = undefined;
  });

  it('follow the engine face after a locale switch', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
    const base = workbook.getFontRecord(0);
    if (!base) throw new Error('font 0 missing');
    const index = workbook.addFontRecord({ ...base, name: 'Calibri', scheme: 2, charset: null });
    sheet = await mountStubSheet({ workbook, locale: 'en' });
    expect(sheet.instance.setLocale('ja')).toBe(true);
    expect(workbook.getFontRecord(index)?.name).toBe('游ゴシック');
    expect(workbook.getResolvedFontRecord(index)?.name).toBe('游ゴシック');
  });
});
