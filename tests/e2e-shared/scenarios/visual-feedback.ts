import { expect, type Locator, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

const themes = ['paper', 'ink', 'contrast'] as const;

async function setTheme(page: Page, theme: string): Promise<void> {
  await page.evaluate((theme) => {
    (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.setTheme(theme);
  }, theme);
}

async function openMenu(page: Page): Promise<Locator> {
  await page.locator('.fc-host').first().focus();
  await page.keyboard.press('Shift+F10');
  const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  await expect(menu).toBeVisible();
  return menu;
}

async function expectVisibleFocus(mark: Locator): Promise<void> {
  const focus = await mark.evaluate((el) => {
    const style = getComputedStyle(el);
    return { style: style.outlineStyle, width: Number.parseFloat(style.outlineWidth) };
  });
  expect(focus.style).not.toBe('none');
  expect(focus.width).toBeGreaterThanOrEqual(2);
}

export async function runDirectionFocusFeedbackScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale: 'ja' });
  await sp.expectNoStub();
  for (const theme of themes) {
    await setTheme(page, theme);
    for (const kind of ['insert', 'delete'] as const) {
      const menu = await openMenu(page);
      await menu.locator(`[data-fc-action="${kind}Cells"]`).click();
      const dialog = page.locator('.fc-cellshift');
      const initial = dialog.locator('input:checked');
      await page.keyboard.press('Tab');
      await expect(initial).toBeFocused();
      await expectVisibleFocus(dialog.locator('input:focus + .fc-cellshift__radio'));
      await page.keyboard.press('ArrowUp');
      const focused = dialog.locator('input:focus');
      await expect(focused).toBeChecked();
      await expectVisibleFocus(dialog.locator('input:focus + .fc-cellshift__radio'));
      await page.keyboard.press('Tab');
      await expect(dialog.locator('.fc-cellshift__button--secondary')).toBeFocused();
      await page.keyboard.press('Enter');
      await expect(dialog).toHaveCount(0);
      await expect(page.locator('.fc-host').first()).toBeFocused();
    }
  }
}

export async function runMiniToolbarFeedbackScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale: 'en' });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('17');
  await sp.focusHost();
  await page.mouse.move(1090, 0);
  for (const theme of themes) {
    await setTheme(page, theme);
    let menu = await openMenu(page);
    const mini = menu.locator('.fc-ctxmenu__mini');
    for (const action of [
      'bold',
      'italic',
      'underline',
      'alignLeft',
      'alignCenter',
      'alignRight',
    ]) {
      await expect(mini.locator(`[data-fc-action="${action}"]`)).toHaveAttribute(
        'aria-pressed',
        'false',
      );
    }
    for (const action of ['bold', 'italic', 'underline', 'alignCenter']) {
      const button = mini.locator(`[data-fc-action="${action}"]`);
      const before = await button.evaluate((el) => getComputedStyle(el).backgroundColor);
      await button.click();
      await expect(menu).toBeHidden();
      menu = await openMenu(page);
      await expect(button).toHaveAttribute('aria-pressed', 'true');
      expect(await button.evaluate((el) => getComputedStyle(el).backgroundColor)).not.toBe(before);
      await button.focus();
      await expectVisibleFocus(button);
      await page.keyboard.press('Escape');
      await sp.shortcut('z');
      menu = await openMenu(page);
      await expect(button).toHaveAttribute('aria-pressed', 'false');
    }
    await expect(mini.locator('[data-fc-action="borders"]')).not.toHaveAttribute('aria-pressed');
    await expect(mini.locator('[data-fc-action="formatCells"]')).not.toHaveAttribute(
      'aria-pressed',
    );
    await page.keyboard.press('Escape');
    await page.evaluate(() => {
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.setPolicy({
        readOnly: true,
        defaultOperation: 'deny',
        selection: true,
        copy: true,
        batchDenied: 'reject',
      });
    });
    menu = await openMenu(page);
    const disabled = menu.locator('.fc-ctxmenu__mini [data-fc-action="bold"]');
    await expect(disabled).toBeDisabled();
    const background = await disabled.evaluate((el) => getComputedStyle(el).backgroundColor);
    await disabled.hover();
    const muted = await disabled.evaluate((el) => {
      const style = getComputedStyle(el);
      return { background: style.backgroundColor, opacity: Number(style.opacity) };
    });
    expect(muted.background).toBe(background);
    expect(muted.opacity).toBeLessThan(1);
    await page.keyboard.press('Escape');
    await page.evaluate(() => {
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.setPolicy();
    });
    await page.mouse.move(1090, 0);
  }
}
