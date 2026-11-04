import { afterEach, describe, expect, it } from 'vitest';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Mac formula bar labels', () => {
  let sheet: MountedStubSheet | null = null;
  afterEach(() => {
    sheet?.dispose();
    sheet = null;
  });

  const label = (selector: string): string | null | undefined =>
    sheet?.host.querySelector(selector)?.getAttribute('aria-label');

  it('uses the desktop accessibility names on the Mac host', async () => {
    sheet = await mountStubSheet({ locale: 'ja', ui: { platform: 'mac' } });
    expect(label('.fc-host__formulabar-action--cancel')).toBe('数式をキャンセルします');
    expect(label('.fc-host__formulabar-action--accept')).toBe('数式を入力します');
    expect(label('.fc-host__formulabar-expand')).toBe('数式バーの展開');
    expect(label('.fc-host__formulabar-fx')).toBe('数式パレットを開きます');
  });

  it('keeps the shared wording elsewhere', async () => {
    sheet = await mountStubSheet({ locale: 'ja' });
    expect(label('.fc-host__formulabar-action--cancel')).toBe(
      sheet.instance.i18n.strings.a11y.cancelFormulaEdit,
    );
  });
});
