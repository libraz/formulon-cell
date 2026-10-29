import { expect, type Locator, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type PopupKind = 'cell' | 'sheet' | 'status';

async function openPopup(page: Page, kind: PopupKind): Promise<Locator> {
  if (kind === 'cell') {
    await page
      .locator('.fc-host')
      .first()
      .click({ button: 'right', position: { x: 200, y: 200 } });
    return page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
  }
  if (kind === 'sheet') {
    await page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').click({ button: 'right' });
    await expect(page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)')).toBeHidden();
    return page.locator('.fc-sheetmenu');
  }
  await page.locator('.fc-host__statusbar').click({ button: 'right' });
  await expect(page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)')).toBeHidden();
  return page.locator('.fc-statusbar__chooser');
}

export async function runPopupFocusScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();
  for (const kind of ['cell', 'sheet', 'status'] as const) {
    for (const key of ['Tab', 'Shift+Tab']) {
      const menu = await openPopup(page, kind);
      await expect(menu).toBeVisible();
      await page.keyboard.press(key);
      await expect(menu, `${kind}: ${key} dismisses`).toBeHidden();
    }
    const menu = await openPopup(page, kind);
    await expect(menu).toBeVisible();
    if (kind === 'cell') {
      const parent = menu.locator('[data-fc-submenu="sortMenu"]');
      const submenu = page.locator('.fc-ctxmenu__sub');
      await expect(parent).toBeEnabled();
      await parent.focus();
      await page.keyboard.press('ArrowRight');
      await expect(submenu).toBeVisible();
      await expect(submenu.locator(':focus')).toHaveCount(1);
      await page.keyboard.press('ArrowLeft');
      await expect(submenu).toBeHidden();
      await expect(menu).toBeVisible();
      await expect(parent).toBeFocused();
      await page.keyboard.press('ArrowRight');
      await expect(submenu).toBeVisible();
      await page.keyboard.press('Escape');
      await expect(submenu).toBeHidden();
      await expect(menu).toBeVisible();
      await expect(parent).toBeFocused();
    }
    const formulaInput = page.locator('.fc-host__formulabar-input').first();
    await formulaInput.focus();
    await expect(menu, `${kind}: focus leaves`).toBeHidden();
    await page.keyboard.press('ArrowDown');
    await expect(formulaInput).toBeFocused();

    await openPopup(page, kind);
    await page.evaluate(() => {
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.openFormatDialog();
    });
    const dialog = page.locator('.fc-fmtdlg[aria-label="Format Cells"]');
    await expect(menu, `${kind}: a dialog takes focus`).toBeHidden();
    await expect(dialog).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dialog).toBeHidden();
  }
}

export async function runPopupViewportScenario(page: Page): Promise<void> {
  await page.setViewportSize({ width: 390, height: 300 });
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();
  for (const kind of ['sheet', 'status', 'cell'] as const) {
    await page.evaluate(() => window.scrollTo(0, 0));
    let menu: Locator;
    if (kind === 'cell') {
      // Keep the pointer out so keyboard scrolling cannot open a hover submenu.
      await page.mouse.move(389, 0);
      await page
        .locator('.fc-host')
        .first()
        .evaluate((host) => (host as HTMLElement).focus({ preventScroll: true }));
      await page.keyboard.press('Shift+F10');
      menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
    } else {
      menu = await openPopup(page, kind);
    }
    await expect(menu).toBeVisible();
    const bounds = await menu.evaluate((element) => {
      const rect = element.getBoundingClientRect();
      return {
        left: rect.left,
        top: rect.top,
        right: rect.right,
        bottom: rect.bottom,
        width: innerWidth,
        height: innerHeight,
      };
    });
    expect.soft(bounds.left, `${kind} ${JSON.stringify(bounds)}`).toBeGreaterThanOrEqual(0);
    expect.soft(bounds.top, `${kind} ${JSON.stringify(bounds)}`).toBeGreaterThanOrEqual(0);
    expect
      .soft(bounds.right, `${kind} ${JSON.stringify(bounds)}`)
      .toBeLessThanOrEqual(bounds.width);
    expect
      .soft(bounds.bottom, `${kind} ${JSON.stringify(bounds)}`)
      .toBeLessThanOrEqual(bounds.height);
    await page.keyboard.press('End');
    await expect(menu, `${kind}: keyboard scrolling keeps menu open`).toBeVisible();
    const focusBounds = await menu.evaluate((element) => {
      const active = element.ownerDocument.activeElement;
      if (!(active instanceof HTMLElement) || !element.contains(active))
        throw new Error('Menu lost focus');
      const item = active.getBoundingClientRect();
      const panel = element.getBoundingClientRect();
      return { top: item.top, bottom: item.bottom, panelTop: panel.top, panelBottom: panel.bottom };
    });
    expect
      .soft(focusBounds.top, `${kind} ${JSON.stringify(focusBounds)}`)
      .toBeGreaterThanOrEqual(focusBounds.panelTop);
    expect
      .soft(focusBounds.bottom, `${kind} ${JSON.stringify(focusBounds)}`)
      .toBeLessThanOrEqual(focusBounds.panelBottom);
    await page.keyboard.press('Escape');
    await expect(menu).toBeHidden();
  }
}
