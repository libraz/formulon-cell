import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';

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
      const inst = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      if (reason === 'features') {
        for (const shortcuts of [false, true]) {
          inst.update({
            ui: { ...inst.options.ui, features: { ...inst.options.ui?.features, shortcuts } },
          });
        }
      } else {
        inst.update({ locale: 'ja' });
      }
    }, rebind);
    await page.locator('.fc-host__grid').focus();
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
        const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
        const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
        return Boolean(store.getState().ui.copyRange);
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
