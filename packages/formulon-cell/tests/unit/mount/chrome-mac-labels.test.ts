import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const formulaBarCss = readFileSync(join(root, 'src/styles/core/surface/formulabar.css'), 'utf8');

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

  it.each(['mac', 'default'] as const)(
    'draws the fx glyph outside the button text on %s',
    async (platform) => {
      sheet = await mountStubSheet({ locale: 'en', ui: { platform } });
      const fx = sheet.host.querySelector<HTMLButtonElement>('.fc-host__formulabar-fx');
      expect(fx?.textContent).toBe('');
      expect(fx?.getAttribute('aria-label')).toBeTruthy();
      expect(formulaBarCss).toMatch(/\.fc-host__formulabar-fx::before\s*\{[^}]*content: "ƒx";/);
    },
  );

  it('keeps the shared wording elsewhere', async () => {
    sheet = await mountStubSheet({ locale: 'ja' });
    expect(label('.fc-host__formulabar-action--cancel')).toBe(
      sheet.instance.i18n.strings.a11y.cancelFormulaEdit,
    );
  });
});
