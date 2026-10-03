import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

async function mountMac(page: Page): Promise<SpreadsheetPage> {
  const sp = new SpreadsheetPage(page);
  await page.goto('/?fixture=empty&locale=en&platform=mac');
  await sp.waitForReady();
  await sp.expectNoStub();
  await expect(page.locator('.fc-host')).toHaveAttribute('data-fc-platform', 'mac');
  return sp;
}

export async function runMacEditShortcutScenario(page: Page): Promise<void> {
  const sp = await mountMac(page);
  await sp.typeIntoActiveCell('original');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Control+u');
  const editor = page.locator('.fc-host__editor');
  await expect(editor).toBeVisible();
  await expect(editor).toHaveValue('original');
  await editor.fill('changed');
  await page.keyboard.press('Escape');
  await expect(editor).toBeHidden();
  expect(await sp.formulaBarValue()).toBe('original');
  await page.keyboard.press('Control+u');
  await editor.fill('committed');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('committed');
  // Rebinding the engine changes DOM listener registration order. Editing
  // must still win over the Windows underline route after either rebind.
  for (const rebind of ['features', 'locale']) {
    await page.evaluate((reason) => {
      const inst = (
        window as unknown as {
          __fcInst: {
            setFeatures(flags: { shortcuts: boolean }): void;
            i18n: { setLocale(locale: string): void };
          };
        }
      ).__fcInst;
      if (reason === 'features') {
        inst.setFeatures({ shortcuts: false });
        inst.setFeatures({ shortcuts: true });
      } else {
        inst.i18n.setLocale('ja');
      }
    }, rebind);
    await page.locator('.fc-host').focus();
    await page.keyboard.press('Control+u');
    await expect(editor).toHaveValue('committed');
    await page.keyboard.press('Escape');
    await expect(page.locator('[data-ribbon-command="underline"]')).not.toHaveAttribute(
      'aria-pressed',
      'true',
    );
  }
  await expect(page.locator('[data-ribbon-command="underline"]')).not.toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await page.keyboard.press('Meta+u');
  await expect(page.locator('[data-ribbon-command="underline"]')).toHaveAttribute(
    'aria-pressed',
    'true',
  );
}

export async function runMacReferenceShortcutScenario(page: Page): Promise<void> {
  const sp = await mountMac(page);
  await sp.focusHost();
  await page.keyboard.type('=A2');
  const editor = page.locator('.fc-host__editor');
  await page.keyboard.press('Meta+t');
  await expect(editor).toHaveValue('=$A$2');
  await page.keyboard.press('Meta+t');
  await expect(editor).toHaveValue('=A$2');
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  const formula = page.locator('.fc-host__formulabar-input');
  await formula.focus();
  await formula.evaluate((el: HTMLTextAreaElement) => {
    el.setSelectionRange(el.value.length, el.value.length);
  });
  await page.keyboard.press('Meta+t');
  await expect(formula).toHaveValue('=$A2');
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await page.keyboard.press('Escape');
  expect(await sp.formulaBarValue()).toBe('=A$2');
}

export async function runMacPasteSpecialShortcutScenario(page: Page): Promise<void> {
  const sp = await mountMac(page);
  await sp.typeIntoActiveCell('42');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Meta+c');
  // The clipboard command writes asynchronously; wait for the copy outline.
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (
          window as unknown as {
            __fcInst: { store: { getState(): { ui: { copyRange: unknown } } } };
          }
        ).__fcInst;
        return Boolean(inst.store.getState().ui.copyRange);
      }),
    )
    .toBe(true);
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Meta+Control+v');
  await expect(page.getByRole('dialog')).toHaveCount(1);
  await expect(page.getByRole('dialog')).toContainText(/Paste Special/i);
  await page.keyboard.press('Escape');
  expect(await sp.formulaBarValue()).toBe('');
}
