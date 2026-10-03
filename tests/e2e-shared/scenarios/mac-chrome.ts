import { expect, type Locator, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

/** Geometry and surface assertions for the explicit macOS Excel profile. */
export async function runMacChromeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();

  await expect(page.locator('.demo[data-fc-platform="mac"]')).toHaveCount(1);
  await expect(page.locator('.fc-host[data-fc-platform="mac"]')).toHaveCount(1);

  const tabs = await page
    .locator('[data-ribbon-tab]')
    .evaluateAll((items) => items.map((item) => item.getAttribute('data-ribbon-tab')));
  expect(tabs).toEqual([
    'home',
    'insert',
    'draw',
    'pageLayout',
    'formulas',
    'data',
    'review',
    'view',
    'automate',
  ]);
  await expect(page.locator('[data-ribbon-tab="file"]')).toHaveCount(0);
  await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);

  const geometry = await page.evaluate(() => {
    const measure = (selector: string): { height: number; width: number } => {
      const rect = document.querySelector<HTMLElement>(selector)?.getBoundingClientRect();
      return { height: rect?.height ?? 0, width: rect?.width ?? 0 };
    };
    return {
      titlebar: measure('.fc-tb__titlebar'),
      tabs: measure('.fc-tb__ribbon-tabs'),
      ribbon: measure('.fc-tb__ribbon:not([hidden])'),
      formulabar: measure('.fc-host__formulabar'),
      namebox: measure('.fc-host__formulabar-tag'),
      sheetbar: measure('.fc-host__sheetbar'),
      statusbar: measure('.fc-host__statusbar'),
    };
  });
  expect(geometry.titlebar.height).toBe(40);
  expect(geometry.tabs.height).toBe(30);
  expect(geometry.ribbon.height).toBeGreaterThanOrEqual(72);
  expect(geometry.ribbon.height).toBeLessThanOrEqual(78);
  expect(geometry.formulabar.height).toBe(28);
  expect(geometry.namebox.width).toBe(87);
  expect(geometry.sheetbar.height).toBe(26);
  expect(geometry.statusbar.height).toBe(28);

  const ribbon = page.locator('.fc-tb__ribbon:not([hidden])').first();
  expect(await ribbon.evaluate((el) => el.scrollWidth)).toBeGreaterThanOrEqual(
    await ribbon.evaluate((el) => el.clientWidth),
  );

  const backstageButton = page.locator('.demo__brand-mark').first();
  await expect(backstageButton).toHaveAttribute('aria-label', 'File');
  await backstageButton.click();
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Open', exact: true })).toBeVisible();
  await backstage.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(backstage).toBeHidden();
  await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);
}

/** Visible labels and the final command must survive compact ribbon layouts. */
export async function runMacRibbonLayoutScenario(page: Page): Promise<void> {
  const tabs = [
    'home',
    'insert',
    'draw',
    'pageLayout',
    'formulas',
    'data',
    'review',
    'view',
    'automate',
  ];
  for (const locale of ['en', 'ja'] as const) {
    const sp = new SpreadsheetPage(page);
    await sp.mount({ platform: 'mac', locale, fixture: 'empty' });
    await sp.expectNoStub();
    for (const mode of ['full', 'singleLine']) {
      if (mode === 'singleLine') {
        await page.locator('[data-ribbon-toggle]').click();
        await page.locator('[data-ribbon-display-option="singleLine"]').click();
      }
      for (const width of [1440, 1024, 768, 390]) {
        await page.setViewportSize({ width, height: 900 });
        for (const tab of tabs) {
          await page.locator(`[data-ribbon-tab="${tab}"]`).click();
          const panel = page.locator(`[data-ribbon-panel="${tab}"]`);
          const issues = await panel.evaluate((el) => {
            const errors: string[] = [];
            for (const label of el.querySelectorAll<HTMLElement>(
              'button > span:not(.fc-tb__rb-icon):not(.fc-tb__rb-split-chevron)',
            )) {
              if (!label.textContent?.trim() || !label.getBoundingClientRect().height) continue;
              if (
                label.scrollWidth > label.clientWidth + 1 ||
                label.scrollHeight > label.clientHeight + 1
              )
                errors.push(`clipped label: ${label.textContent}`);
            }
            el.scrollLeft = el.scrollWidth;
            const bounds = el.getBoundingClientRect();
            const toggle = el.parentElement
              ?.querySelector('[data-ribbon-toggle]')
              ?.getBoundingClientRect();
            const controls = Array.from(
              el.querySelectorAll<HTMLElement>(
                '.fc-tb__ribbon-tools > button, .fc-tb__ribbon-tools > .fc-tb__rb-select',
              ),
            );
            for (const control of controls) {
              const rect = control.getBoundingClientRect();
              if (!rect.height) continue;
              if (rect.top < bounds.top - 1 || rect.bottom > bounds.bottom + 1)
                errors.push(`vertical clipping: ${control.dataset.ribbonCommand}`);
              if (
                toggle &&
                rect.right > bounds.left &&
                rect.left < bounds.right &&
                rect.left < toggle.right &&
                rect.right > toggle.left &&
                rect.top < toggle.bottom &&
                rect.bottom > toggle.top
              )
                errors.push(`toggle overlap: ${control.dataset.ribbonCommand}`);
            }
            return errors;
          });
          expect(issues, `${locale}/${mode}/${width}/${tab}`).toEqual([]);
        }
      }
    }
  }
}

/** Display-mode state, tab-rail reachability, and collapsed peek behavior. */
export async function runMacRibbonDisplayScenario(page: Page): Promise<void> {
  const tabIds = [
    'home',
    'insert',
    'draw',
    'pageLayout',
    'formulas',
    'data',
    'review',
    'view',
    'automate',
  ] as const;
  const modes = ['singleLine', 'tabsOnly', 'full'] as const;
  const widths = [390, 768] as const;
  type MacRibbonTab = (typeof tabIds)[number];
  type MacRibbonDisplayMode = (typeof modes)[number];

  const expectSelectedTab = async (tabId: MacRibbonTab): Promise<void> => {
    await expect(page.locator('[data-ribbon-tab][aria-selected="true"]')).toHaveCount(1);
    await expect(page.locator(`[data-ribbon-tab="${tabId}"][aria-selected="true"]`)).toHaveCount(1);
  };

  const expectActiveTab = async (tabId: MacRibbonTab): Promise<void> => {
    await expectSelectedTab(tabId);
    await expect(page.locator(`[data-ribbon-panel="${tabId}"]`)).toBeVisible();
    await expect(page.locator('.fc-tb__ribbon:not([hidden])')).toHaveCount(1);
  };

  const expectTabLabels = async (): Promise<void> => {
    const issues = await page.locator('.fc-tb__ribbon-tabs').evaluate((rail) => {
      const errors: string[] = [];
      for (const tab of rail.querySelectorAll<HTMLButtonElement>('[data-ribbon-tab]')) {
        const text = tab.textContent?.trim() ?? '';
        const rect = tab.getBoundingClientRect();
        const range = document.createRange();
        range.selectNodeContents(tab);
        const textRect = range.getBoundingClientRect();
        if (!text) errors.push('empty tab label');
        if (!rect.height) errors.push(`hidden tab label: ${text}`);
        if (tab.scrollWidth > tab.clientWidth + 1 || tab.scrollHeight > tab.clientHeight + 1) {
          errors.push(`scroll-clipped tab label: ${text}`);
        }
        if (
          textRect.width > rect.width + 1 ||
          textRect.height > rect.height + 1 ||
          textRect.left < rect.left - 1 ||
          textRect.right > rect.right + 1 ||
          textRect.top < rect.top - 1 ||
          textRect.bottom > rect.bottom + 1
        ) {
          errors.push(`geometry-clipped tab label: ${text}`);
        }
      }
      return errors;
    });
    expect(issues).toEqual([]);
  };

  const expectTabInRail = async (tabId: MacRibbonTab): Promise<void> => {
    const visible = await page.locator(`[data-ribbon-tab="${tabId}"]`).evaluate((tab) => {
      const rail = tab.parentElement?.getBoundingClientRect();
      const rect = tab.getBoundingClientRect();
      return Boolean(
        rail &&
          rect.height > 0 &&
          rect.left >= rail.left - 1 &&
          rect.right <= rail.right + 1 &&
          rect.top >= rail.top - 1 &&
          rect.bottom <= rail.bottom + 1,
      );
    });
    expect(visible, `${tabId} tab should be visible in the scrolled tab rail`).toBe(true);
  };

  const expectGridStable = async (before: {
    x: number;
    y: number;
    width: number;
    height: number;
  }): Promise<void> => {
    const after = await page.locator('.fc-host__grid').boundingBox();
    if (!after) throw new Error('worksheet grid lost its layout box');
    for (const key of ['x', 'y', 'width', 'height'] as const) {
      expect(
        Math.abs(after[key] - before[key]),
        `grid ${key} changed during ribbon peek`,
      ).toBeLessThanOrEqual(1);
    }
  };

  for (const locale of ['en', 'ja'] as const) {
    const sp = new UserJourneyPage(page);
    await sp.mount({ platform: 'mac', locale, fixture: 'empty' });
    await sp.expectNoStub();
    let expectedActiveTab: MacRibbonTab = tabIds[0];

    const shell = page.locator('.fc-tb__ribbon-shell');
    const toggle = page.locator('[data-ribbon-toggle]');
    const openDisplayMenu = async (expectedMode: MacRibbonDisplayMode) => {
      await toggle.click();
      const menu = page.locator('.fc-tb__ribbon-display-menu');
      await expect(menu).toBeVisible();
      const menuBounds = await menu.evaluate((element) => {
        const rect = element.getBoundingClientRect();
        return {
          left: rect.left,
          top: rect.top,
          right: rect.right,
          bottom: rect.bottom,
          viewportWidth: window.innerWidth,
          viewportHeight: window.innerHeight,
        };
      });
      expect(menuBounds.left, `${locale}/${expectedMode} menu left`).toBeGreaterThanOrEqual(-1);
      expect(menuBounds.top, `${locale}/${expectedMode} menu top`).toBeGreaterThanOrEqual(-1);
      expect(menuBounds.right, `${locale}/${expectedMode} menu right`).toBeLessThanOrEqual(
        menuBounds.viewportWidth + 1,
      );
      expect(menuBounds.bottom, `${locale}/${expectedMode} menu bottom`).toBeLessThanOrEqual(
        menuBounds.viewportHeight + 1,
      );
      await expect(menu.locator('[data-ribbon-display-option][aria-checked="true"]')).toHaveCount(
        1,
      );
      await expect(menu.locator(`[data-ribbon-display-option="${expectedMode}"]`)).toHaveAttribute(
        'aria-checked',
        'true',
      );
      return menu;
    };

    for (const mode of modes) {
      for (const width of widths) {
        await page.setViewportSize({ width, height: 900 });
        const currentModeValue = await shell.getAttribute('data-ribbon-display-mode');
        if (!currentModeValue || !modes.includes(currentModeValue as MacRibbonDisplayMode)) {
          throw new Error(`unexpected Mac ribbon display mode: ${currentModeValue}`);
        }
        const currentMode = currentModeValue as MacRibbonDisplayMode;
        await expectSelectedTab(expectedActiveTab);
        if (currentMode === 'tabsOnly') {
          await expect(page.locator(`[data-ribbon-panel="${expectedActiveTab}"]`)).toBeHidden();
        } else {
          await expect(page.locator(`[data-ribbon-panel="${expectedActiveTab}"]`)).toBeVisible();
        }

        let menu = await openDisplayMenu(currentMode);
        if (currentMode === mode) {
          await page.keyboard.press('Escape');
        } else {
          await menu.locator(`[data-ribbon-display-option="${mode}"]`).click();
        }
        await expect(menu).toHaveCount(0);
        await expect(shell).toHaveAttribute('data-ribbon-display-mode', mode);
        menu = await openDisplayMenu(mode);
        await page.keyboard.press('Escape');
        await expect(menu).toHaveCount(0);
        await expectTabLabels();

        const firstTab = page.locator(`[data-ribbon-tab="${tabIds[0]}"]`);
        const lastTab = page.locator(`[data-ribbon-tab="${tabIds[tabIds.length - 1]}"]`);
        const collapsedGrid =
          mode === 'tabsOnly' ? await page.locator('.fc-host__grid').boundingBox() : null;
        const requireCollapsedGrid = (): NonNullable<typeof collapsedGrid> => {
          if (!collapsedGrid) throw new Error('worksheet grid is missing in collapsed mode');
          return collapsedGrid;
        };
        await firstTab.scrollIntoViewIfNeeded();
        await firstTab.focus();
        await expect(firstTab).toBeFocused();
        await firstTab.click();
        await expect(firstTab).toBeFocused();
        await expectTabInRail(tabIds[0]);
        expectedActiveTab = tabIds[0];

        if (mode === 'tabsOnly') {
          await expect(page.locator(`[data-ribbon-panel="${tabIds[0]}"]`)).toBeVisible();
          await expectGridStable(requireCollapsedGrid());
          await page.keyboard.press('Escape');
          await expect(page.locator('.fc-tb__ribbon-shell--peek')).toHaveCount(0);
          await expect(page.locator(`[data-ribbon-panel="${tabIds[0]}"]`)).toBeHidden();
          await expectSelectedTab(tabIds[0]);
          await expectGridStable(requireCollapsedGrid());
        } else {
          await expectActiveTab(tabIds[0]);
        }

        await lastTab.scrollIntoViewIfNeeded();
        await lastTab.focus();
        await expect(lastTab).toBeFocused();
        await lastTab.click();
        await expect(lastTab).toBeFocused();
        await expectTabInRail(tabIds[tabIds.length - 1]);
        expectedActiveTab = tabIds[tabIds.length - 1];

        if (mode === 'tabsOnly') {
          await expect(page.locator('.fc-tb__ribbon-shell--peek')).toHaveCount(1);
          await expectActiveTab(tabIds[tabIds.length - 1]);
          await expectGridStable(requireCollapsedGrid());
          await page.keyboard.press('Escape');
          await expect(page.locator('.fc-tb__ribbon-shell--peek')).toHaveCount(0);
          await expect(
            page.locator(`[data-ribbon-panel="${tabIds[tabIds.length - 1]}"]`),
          ).toBeHidden();
          await expectSelectedTab(tabIds[tabIds.length - 1]);
          await expectGridStable(requireCollapsedGrid());
        } else {
          await expectActiveTab(tabIds[tabIds.length - 1]);
        }
      }
    }
    await sp.expectNoConsoleErrors();
  }
}

/** Display-menu placement under narrow, transformed ribbon hosts. */
export async function runMacRibbonDisplayPlacementScenario(page: Page): Promise<void> {
  type DisplayMode = 'full' | 'singleLine' | 'tabsOnly' | 'autoHide';
  const modes: readonly DisplayMode[] = ['singleLine', 'tabsOnly', 'autoHide', 'full'];
  const hosts = [
    { name: 'scaled-320', viewportWidth: 320, scale: 2, logicalWidth: 160 },
    { name: 'scaled-160', viewportWidth: 160, scale: 2, logicalWidth: 80 },
    { name: 'unscaled-160', viewportWidth: 160, scale: 1, logicalWidth: 160 },
  ] as const;

  const configureHost = async (host: (typeof hosts)[number]) => {
    return page.locator('.fc-tb__ribbon-shell').evaluate(
      (shell, { logicalWidth, scale }) => {
        // The toolbar mount host is `display: contents`; use the nearest real
        // layout wrapper so the transform creates a measurable fixed block.
        const ancestor = shell.closest<HTMLElement>('.fc-tb__sheet-col');
        if (!(ancestor instanceof HTMLElement)) {
          throw new Error('Mac ribbon shell has no measurable host ancestor');
        }
        // Deliberate embed stress: this is the real shell ancestor, with a
        // logical width that makes its transformed toggle land at 320/160px.
        ancestor.style.width = `${logicalWidth}px`;
        ancestor.style.minWidth = '0px';
        ancestor.style.maxWidth = `${logicalWidth}px`;
        ancestor.style.transformOrigin = 'top left';
        ancestor.style.transform = scale === 1 ? 'none' : `scale(${scale})`;
        const ancestorRect = ancestor.getBoundingClientRect();
        if (ancestorRect.width <= 0 || getComputedStyle(ancestor).display === 'contents') {
          throw new Error('Mac ribbon stress wrapper did not produce a measurable layout box');
        }
        const toggle = shell.querySelector<HTMLElement>('[data-ribbon-toggle]');
        const toggleRect = toggle?.getBoundingClientRect();
        return {
          ancestorWidth: ancestorRect.width,
          toggleLeft: toggleRect?.left ?? 0,
          toggleRight: toggleRect?.right ?? 0,
        };
      },
      { logicalWidth: host.logicalWidth, scale: host.scale },
    );
  };

  const assertMenuGeometry = async (menu: Locator, label: string) => {
    const geometry = await menu.evaluate((element) => {
      const rect = element.getBoundingClientRect();
      const options = Array.from(
        element.querySelectorAll<HTMLElement>('[data-ribbon-display-option]'),
      ).map((option) => {
        const optionRect = option.getBoundingClientRect();
        return {
          left: optionRect.left,
          top: optionRect.top,
          right: optionRect.right,
          bottom: optionRect.bottom,
          width: optionRect.width,
          height: optionRect.height,
        };
      });
      return {
        left: rect.left,
        top: rect.top,
        right: rect.right,
        bottom: rect.bottom,
        viewportWidth: window.innerWidth,
        viewportHeight: window.innerHeight,
        options,
      };
    });
    expect(geometry.left, `${label} menu left`).toBeGreaterThanOrEqual(0);
    expect(geometry.top, `${label} menu top`).toBeGreaterThanOrEqual(0);
    expect(geometry.right, `${label} menu right`).toBeLessThanOrEqual(geometry.viewportWidth);
    expect(geometry.bottom, `${label} menu bottom`).toBeLessThanOrEqual(geometry.viewportHeight);
    expect(geometry.options).toHaveLength(4);
    for (const [index, option] of geometry.options.entries()) {
      expect(option.width, `${label} option ${index} width`).toBeGreaterThan(0);
      expect(option.height, `${label} option ${index} height`).toBeGreaterThan(0);
      expect(option.left, `${label} option ${index} left`).toBeGreaterThanOrEqual(geometry.left);
      expect(option.top, `${label} option ${index} top`).toBeGreaterThanOrEqual(geometry.top);
      expect(option.right, `${label} option ${index} right`).toBeLessThanOrEqual(geometry.right);
      expect(option.bottom, `${label} option ${index} bottom`).toBeLessThanOrEqual(geometry.bottom);
      expect(option.left, `${label} option ${index} viewport left`).toBeGreaterThanOrEqual(0);
      expect(option.top, `${label} option ${index} viewport top`).toBeGreaterThanOrEqual(0);
      expect(option.right, `${label} option ${index} viewport right`).toBeLessThanOrEqual(
        geometry.viewportWidth,
      );
      expect(option.bottom, `${label} option ${index} viewport bottom`).toBeLessThanOrEqual(
        geometry.viewportHeight,
      );
    }
  };

  for (const host of hosts) {
    await page.setViewportSize({ width: host.viewportWidth, height: 900 });
    const sp = new UserJourneyPage(page);
    await sp.mount({ platform: 'mac', locale: 'en', fixture: 'empty' });
    await sp.expectNoStub();
    const configured = await configureHost(host);
    expect(configured.ancestorWidth, `${host.name} host width`).toBeCloseTo(
      host.logicalWidth * host.scale,
      0,
    );
    expect(configured.toggleLeft, `${host.name} toggle left`).toBeGreaterThanOrEqual(0);
    expect(configured.toggleRight, `${host.name} toggle right`).toBeLessThanOrEqual(
      host.viewportWidth,
    );

    const shell = page.locator('.fc-tb__ribbon-shell');
    const toggle = page.locator('[data-ribbon-toggle]');
    for (const targetMode of modes) {
      const currentMode = (await shell.getAttribute('data-ribbon-display-mode')) as DisplayMode;
      const menu = page.locator('.fc-tb__ribbon-display-menu');
      await toggle.click();
      await expect(menu).toBeVisible();
      await assertMenuGeometry(menu, `${host.name}/${currentMode}`);
      const option = menu.locator(`[data-ribbon-display-option="${targetMode}"]`);
      await expect(option).toBeVisible();
      await option.scrollIntoViewIfNeeded();
      if (targetMode === currentMode) await page.keyboard.press('Escape');
      else await option.click();
      await expect(menu).toHaveCount(0);
      await expect(shell).toHaveAttribute('data-ribbon-display-mode', targetMode);
    }
    await sp.expectNoConsoleErrors();
  }
}
