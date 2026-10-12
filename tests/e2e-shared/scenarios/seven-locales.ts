import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

const LOCALE_PROFILES = [
  { id: 'en', label: 'English', ribbonHome: 'Home', profileToken: 'en_US' },
  { id: 'ja', label: '日本語', ribbonHome: 'ホーム', profileToken: 'ja_JP' },
  { id: 'de', label: 'Deutsch', ribbonHome: 'Start', profileToken: 'de_DE' },
  { id: 'fr', label: 'Français', ribbonHome: 'Accueil', profileToken: 'fr_FR' },
  { id: 'zh', label: '中文', ribbonHome: '开始', profileToken: 'zh_CN' },
  { id: 'ko', label: '한국어', ribbonHome: '홈', profileToken: 'ko_KR' },
  { id: 'th', label: 'ไทย', ribbonHome: 'หน้าแรก', profileToken: 'th_TH' },
] as const;

/** Seven-language smoke path: switch the demo selector and verify that the
 * wrapper chrome, core i18n controller, ribbon labels, and engine profile all
 * move together. The scenario deliberately uses native labels so it also
 * catches a selector that falls back to short two-letter buttons. */
export async function runSevenLocalesScenario(
  page: Page,
  platform: 'mac' | 'default' = 'default',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  const consoleErrors = sp.collectConsoleErrors();

  await sp.mount({ locale: 'en-US', fixture: 'empty', platform });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('=DOLLAR(-1234.567)');

  const panelToggle = page.locator('.demo__account .demo__share').nth(1);
  await panelToggle.click();
  const panel = page.locator('aside.demo__panel');
  await expect(panel).toBeVisible();
  const localeSelect = panel.getByRole('combobox');
  const localeOptions = panel.getByRole('listbox').getByRole('option');
  await expect(localeSelect).toHaveCount(1);
  await localeSelect.click();
  await expect(localeOptions).toHaveText(LOCALE_PROFILES.map((profile) => profile.label));
  await localeSelect.press('Escape');
  await expect(localeOptions.first()).toBeHidden();
  await localeSelect.press('End');
  await expect(localeSelect).toHaveAttribute(
    'aria-activedescendant',
    (await localeOptions.last().getAttribute('id')) ?? '',
  );
  await localeSelect.press('Escape');

  const nativeSelect = page.locator('.fc-host__locale-switcher select');
  await expect(nativeSelect).toHaveCount(0);
  await page.evaluate(() => {
    (
      window as unknown as {
        __fcInst: { setFeatures: (flags: { localeSwitcher: boolean }) => void };
      }
    ).__fcInst.setFeatures({ localeSwitcher: true });
  });
  await expect(nativeSelect).toBeVisible();
  await expect(nativeSelect.locator('option')).toHaveCount(7);

  for (const [index, profile] of LOCALE_PROFILES.entries()) {
    if (index % 2 === 0) {
      await nativeSelect.selectOption(profile.id);
    } else {
      await localeSelect.click();
      await localeOptions.filter({ hasText: profile.label }).click();
    }
    await expect(nativeSelect).toHaveValue(profile.id);
    await expect(localeSelect).toHaveAttribute('data-value', profile.id);
    await expect(localeSelect).toHaveText(profile.label);
    await expect(page.locator('html')).toHaveAttribute('lang', profile.id);

    await expect
      .poll(
        () =>
          page.evaluate(
            () =>
              (window as unknown as { __fcInst?: { i18n?: { locale?: string } } }).__fcInst?.i18n
                ?.locale,
          ),
        { message: `core i18n locale should switch to ${profile.id}` },
      )
      .toBe(profile.id);

    await expect(
      page.getByRole('tab', { name: profile.ribbonHome, exact: true }).first(),
    ).toBeVisible();

    // Labels measured from the Mac Excel 16.113.4 accessibility tree.
    if (platform === 'mac' && profile.id === 'fr') {
      await expect(page.getByRole('tab', { name: 'Insérer', exact: true })).toBeVisible();
      await expect(page.getByRole('tab', { name: 'Automatisation', exact: true })).toBeVisible();
    }
    if (platform === 'mac' && profile.id === 'th') {
      await expect(page.getByRole('tab', { name: 'อัตโนมัติ', exact: true })).toBeVisible();
    }

    await expect
      .poll(
        () =>
          page.evaluate(() => {
            const instance = (
              window as unknown as {
                __fcInst?: {
                  workbook?: { spreadsheetProfileId?: () => string | null };
                };
              }
            ).__fcInst;
            return instance?.workbook?.spreadsheetProfileId?.() ?? null;
          }),
        { message: `engine profile should follow ${profile.id}` },
      )
      .toBe(`${platform === 'mac' ? 'mac' : 'windows'}-${profile.profileToken}`);

    const expectedDollar = {
      en: '($1,234.57)',
      ja: '¥-1,235',
      de: '-1.234,57 €',
      fr: '(1 234,57 €)',
      zh: '(¥1,234.57)',
      ko: '(₩1,235)',
      th: '(฿1,234.57)',
    }[profile.id];
    await expect
      .poll(() =>
        page.evaluate(() => {
          const instance = (
            window as unknown as {
              __fcInst?: {
                workbook: {
                  getValue: (addr: { sheet: number; row: number; col: number }) => {
                    kind: string;
                    value?: string;
                  };
                };
                store: {
                  getState: () => {
                    data: { cells: Map<string, { value: { kind: string; value?: string } }> };
                  };
                };
              };
            }
          ).__fcInst;
          if (!instance) throw new Error('Missing spreadsheet instance');
          return {
            engine: instance.workbook.getValue({ sheet: 0, row: 0, col: 0 }).value,
            displayed: instance.store.getState().data.cells.get('0:0:0')?.value.value,
          };
        }),
      )
      .toEqual({ engine: expectedDollar, displayed: expectedDollar });
  }

  await page.evaluate(() => {
    (
      window as unknown as {
        __fcInst: { setFeatures: (flags: { localeSwitcher: boolean }) => void };
      }
    ).__fcInst.setFeatures({ localeSwitcher: false });
  });
  await expect(nativeSelect).toHaveCount(0);

  expect(consoleErrors.read(), 'expected no console errors while switching locales').toEqual([]);
}
