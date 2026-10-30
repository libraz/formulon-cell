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

async function highlightedRootItems(menu: Locator): Promise<string[]> {
  return menu.locator('.fc-ctxmenu__item').evaluateAll((items) =>
    items
      .filter((item) => {
        const background = getComputedStyle(item).backgroundColor;
        return background !== 'transparent' && background !== 'rgba(0, 0, 0, 0)';
      })
      .map(
        (item) =>
          (item as HTMLElement).dataset.fcSubmenu ?? (item as HTMLElement).dataset.fcAction ?? '',
      ),
  );
}

export async function runSubmenuHighlightScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  for (const platform of ['default', 'mac'] as const) {
    await sp.mount({ locale: 'en', platform });
    await sp.expectNoStub();
    await sp.focusHost();
    await page.mouse.move(1270, 0);
    await page.keyboard.press('Shift+F10');
    const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
    const parent = menu.locator('[data-fc-submenu="sortMenu"]');
    const sibling = menu.locator('[data-fc-action="clear"]');
    const submenu = page.locator('.fc-ctxmenu__sub');
    await parent.hover();
    await expect(submenu).toBeVisible();
    await sibling.hover();
    expect(await highlightedRootItems(menu), `${platform}: hover after submenu`).toEqual(['clear']);

    await parent.hover();
    await expect(submenu).toBeVisible();
    // The stationary pointer remains over the parent while the keyboard moves.
    await page.keyboard.press('Home');
    await expect(menu.locator('[data-fc-action="cut"]')).toBeFocused();
    expect(await highlightedRootItems(menu), `${platform}: keyboard after hover`).toEqual(['cut']);
    await expect(submenu).toBeHidden();
    await page.keyboard.press('End');
    const last = menu.locator('.fc-ctxmenu__item:not(:disabled)').last();
    await expect(last).toBeFocused();
    const lastId = await last.evaluate((item) => item.dataset.fcSubmenu ?? item.dataset.fcAction);
    expect(await highlightedRootItems(menu)).toEqual([lastId]);

    await sibling.hover();
    await parent.hover();
    await expect(submenu).toBeVisible();
    const child = submenu.locator('.fc-ctxmenu__item:not(:disabled)').first();
    await child.hover();
    expect(await highlightedRootItems(menu), `${platform}: parent of active child`).toEqual([
      'sortMenu',
    ]);
    await page.keyboard.press('ArrowLeft');
    await expect(submenu).toBeHidden();
    await expect(parent).toBeFocused();
    expect(await highlightedRootItems(menu)).toEqual(['sortMenu']);
    await page.keyboard.press('Escape');
    await expect(menu).toBeHidden();
    const moduleUrl = `/@fs${
      new URL('../../../packages/formulon-cell/src/interact/object-menu.ts', import.meta.url)
        .pathname
    }`;
    await page.evaluate(async (moduleUrl) => {
      const host = document.querySelector<HTMLElement>('.fc-host');
      if (!host) throw new Error('Missing host');
      const { createObjectMenu } = await import(moduleUrl);
      createObjectMenu(host).open(
        80,
        220,
        [
          {
            id: 'first',
            label: 'First object action',
            run: () => {
              host.dataset.objectAction = 'first';
            },
          },
          {
            id: 'second',
            label: 'Second object action',
            run: () => {
              host.dataset.objectAction = 'second';
            },
          },
          {
            id: 'blocked',
            label: 'Unavailable object action',
            disabled: true,
            run: () => {
              host.dataset.objectAction = 'blocked';
            },
          },
        ],
        { label: 'Object actions', restoreFocusTo: host },
      );
    }, moduleUrl);
    const objectMenu = page.locator('.fc-ctxmenu--object');
    await expect(objectMenu.locator('[data-fc-action="first"]')).toBeFocused();
    expect(await highlightedRootItems(objectMenu)).toEqual(['first']);
    await objectMenu.locator('[data-fc-action="second"]').hover();
    await expect(objectMenu.locator('[data-fc-action="second"]')).toBeFocused();
    expect(await highlightedRootItems(objectMenu)).toEqual(['second']);
    await page.keyboard.press('Home');
    expect(await highlightedRootItems(objectMenu)).toEqual(['first']);
    await page.keyboard.press('End');
    expect(await highlightedRootItems(objectMenu)).toEqual(['second']);
    await objectMenu.locator('[data-fc-action="blocked"]').hover();
    await expect(objectMenu).toBeFocused();
    expect(await highlightedRootItems(objectMenu)).toEqual([]);
    await page.keyboard.press('Enter');
    await page.keyboard.press('Space');
    await expect(objectMenu).toBeVisible();
    await expect(page.locator('.fc-host').first()).not.toHaveAttribute('data-object-action');
    await page.keyboard.press('ArrowDown');
    await expect(objectMenu.locator('[data-fc-action="first"]')).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await expect(objectMenu.locator('[data-fc-action="second"]')).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(objectMenu).toHaveCount(0);
    await expect(page.locator('.fc-host').first()).toHaveAttribute('data-object-action', 'second');
    await expect(page.locator('.fc-host').first()).toBeFocused();
  }
}

export async function runMacSubmenuLayoutScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  for (const locale of ['en', 'ja', 'de', 'fr']) {
    for (const width of [1280, 390]) {
      await page.setViewportSize({ width, height: 800 });
      await sp.mount({ locale, platform: 'mac' });
      await sp.expectNoStub();
      await page.mouse.move(width - 1, 0);
      await page.locator('.fc-host').first().focus();
      await page.keyboard.press('Shift+F10');
      const menu = page.locator('.fc-ctxmenu:not(.fc-ctxmenu__sub)');
      const submenu = page.locator('.fc-ctxmenu__sub');
      for (const id of ['sortMenu', 'pasteSpecialMenu']) {
        await menu.locator(`[data-fc-submenu="${id}"]`).focus();
        await page.keyboard.press('ArrowRight');
        await expect(submenu).toBeVisible();
        const bounds = await submenu.evaluate((element) => {
          const box = element.getBoundingClientRect();
          return { left: box.left, right: box.right, width: box.width, viewport: innerWidth };
        });
        expect(bounds.left, `${locale}/${width}/${id}`).toBeGreaterThanOrEqual(0);
        expect(bounds.right, `${locale}/${width}/${id}`).toBeLessThanOrEqual(bounds.viewport);
        if (id === 'sortMenu') {
          const rootWidth = await menu.evaluate((element) => element.getBoundingClientRect().width);
          expect(bounds.width, `${locale}/${width}: short submenu fits its content`).toBeLessThan(
            rootWidth,
          );
        }
        if (width === 1280) {
          const clipped = await submenu
            .locator('.fc-ctxmenu__label')
            .evaluateAll((labels) =>
              labels
                .filter((label) => label.scrollWidth > label.clientWidth + 1)
                .map((label) => label.textContent),
            );
          expect(clipped, `${locale}/${id}: complete labels on desktop`).toEqual([]);
        }
        await page.keyboard.press('ArrowLeft');
        await expect(submenu).toBeHidden();
      }
      await page.keyboard.press('Escape');
      await expect(menu).toBeHidden();
    }
  }
}
