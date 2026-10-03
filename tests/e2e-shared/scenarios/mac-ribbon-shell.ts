import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacCollapsedRibbonScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();
  await page.locator('[data-ribbon-tab="home"]').dblclick();
  await expect(page.locator('[data-ribbon-panel="home"]')).toBeHidden();
  const collapsedGrid = await page.locator('.fc-host__grid').boundingBox();
  await page.locator('[data-ribbon-tab="home"]').click();
  await expect(page.locator('[data-ribbon-panel="home"]')).toBeVisible();
  expect(await page.locator('.fc-host__grid').boundingBox()).toEqual(collapsedGrid);
  await page.locator('[data-ribbon-tab="insert"]').click();
  await expect(page.locator('[data-ribbon-panel="insert"]')).toBeVisible();
  await sp.clickRibbon('mac.insert.shapes');
  await expect(page.locator('#menu-mac-insert-shapes')).toBeVisible();
  await page.locator('[data-ribbon-tab="formulas"]').click();
  await expect(page.locator('#menu-mac-insert-shapes')).toBeHidden();
  await expect(page.locator('[data-ribbon-panel="formulas"]')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('[data-ribbon-panel="formulas"]')).toBeHidden();
  await page.locator('[data-ribbon-tab="data"]').click();
  await expect(page.locator('[data-ribbon-panel="data"]')).toBeVisible();
  await page.locator('.fc-host__grid').click({ position: { x: 200, y: 160 } });
  await expect(page.locator('[data-ribbon-panel="data"]')).toBeHidden();
  // A double click pins the ribbon without losing the clicked tab.
  await page.locator('[data-ribbon-tab="view"]').dblclick();
  await expect(page.locator('[data-ribbon-panel="view"]')).toBeVisible();
  await page.locator('.fc-host__grid').click({ position: { x: 200, y: 160 } });
  await expect(page.locator('[data-ribbon-panel="view"]')).toBeVisible();
  await sp.expectNoConsoleErrors();
}

export async function runMacRibbonSwitchingScenario(
  page: Page,
  locale: 'en' | 'ja' = 'en',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ platform: 'mac', locale });
  await sp.expectNoStub();
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
  for (const id of [...tabs, ...[...tabs].reverse()]) {
    await page.locator(`[data-ribbon-tab="${id}"]`).click();
    await expect(page.locator('[data-ribbon-tab][aria-selected="true"]')).toHaveCount(1);
    await expect(page.locator(`[data-ribbon-tab="${id}"]`)).toHaveAttribute(
      'aria-selected',
      'true',
    );
    await expect(page.locator('.fc-tb__ribbon:not([hidden])')).toHaveCount(1);
    const panel = page.locator(`[data-ribbon-panel="${id}"]`);
    await expect(panel).toBeVisible();
    expect(
      await panel.locator('[data-ribbon-command], [data-ribbon-control]').count(),
    ).toBeGreaterThan(0);
  }
  // The tab list is a keyboard surface too; its active panel must follow focus.
  await page.locator('[data-ribbon-tab="home"]').focus();
  await page.keyboard.press('ArrowRight');
  await expect(page.locator('[data-ribbon-tab="insert"]')).toBeFocused();
  await expect(page.locator('[data-ribbon-panel="insert"]')).toBeVisible();
  await page.keyboard.press('End');
  await expect(page.locator('[data-ribbon-tab="automate"]')).toBeFocused();
  await expect(page.locator('[data-ribbon-panel="automate"]')).toBeVisible();
  await page.keyboard.press('Home');
  await expect(page.locator('[data-ribbon-tab="home"]')).toBeFocused();
  await expect(page.locator('[data-ribbon-panel="home"]')).toBeVisible();
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.locator('[data-ribbon-tab="formulas"]').click();
  const functionLibrary = page.locator(
    '[data-ribbon-panel="formulas"] .fc-tb__ribbon-group--function-library',
  );
  await expect(functionLibrary).toBeVisible();
  const libraryGeometry = await functionLibrary.evaluate((group) => {
    const panel = group.closest<HTMLElement>('.fc-tb__ribbon');
    const buttons = Array.from(
      group.querySelectorAll<HTMLButtonElement>(
        '.fc-tb__ribbon-tools > button[data-ribbon-command]',
      ),
    );
    const panelRect = panel?.getBoundingClientRect();
    const groupRect = group.getBoundingClientRect();
    return {
      groupWidth: groupRect.width,
      panelHeight: panelRect?.height ?? 0,
      panelLeft: panelRect?.left ?? 0,
      panelRight: panelRect?.right ?? 0,
      panelTop: panelRect?.top ?? 0,
      panelBottom: panelRect?.bottom ?? 0,
      buttons: buttons.map((button) => {
        const rect = button.getBoundingClientRect();
        const icon = button.querySelector<SVGElement>('.fc-tb__rb-icon')?.getBoundingClientRect();
        const labelElement = Array.from(button.children)
          .filter((child): child is HTMLElement => child instanceof HTMLElement)
          .find(
            (child) =>
              !child.classList.contains('fc-tb__rb-icon') &&
              !child.classList.contains('fc-tb__rb-split-chevron'),
          );
        const label = labelElement?.getBoundingClientRect();
        const labelFragments = (() => {
          if (!labelElement) return [];
          const range = document.createRange();
          range.selectNodeContents(labelElement);
          return Array.from(range.getClientRects())
            .filter((fragment) => fragment.width > 0 && fragment.height > 0)
            .map((fragment) => ({
              left: fragment.left,
              right: fragment.right,
              top: fragment.top,
              bottom: fragment.bottom,
            }));
        })();
        return {
          id: button.dataset.ribbonCommand,
          left: rect.left,
          right: rect.right,
          top: rect.top,
          bottom: rect.bottom,
          width: rect.width,
          iconBottom: icon?.bottom ?? 0,
          labelTop: label?.top ?? 0,
          labelFragments,
          hasMenu: button.getAttribute('aria-haspopup') === 'menu',
        };
      }),
    };
  });
  if (locale === 'ja') {
    expect(libraryGeometry.groupWidth).toBeGreaterThanOrEqual(502);
    expect(libraryGeometry.groupWidth).toBeLessThanOrEqual(506);
  } else {
    expect(libraryGeometry.groupWidth).toBeGreaterThanOrEqual(502);
  }
  expect(libraryGeometry.panelHeight).toBe(75);
  expect(libraryGeometry.buttons).toHaveLength(10);
  const libraryTop = libraryGeometry.buttons[0]?.top ?? 0;
  for (const [index, button] of libraryGeometry.buttons.entries()) {
    expect(Math.abs(button.top - libraryTop), button.id).toBeLessThanOrEqual(1);
    expect(button.bottom - button.top, button.id).toBeGreaterThanOrEqual(65);
    expect(button.bottom - button.top, button.id).toBeLessThanOrEqual(67);
    expect(button.iconBottom, button.id).toBeLessThanOrEqual(button.labelTop + 1);
    expect(button.top, button.id).toBeGreaterThanOrEqual(libraryGeometry.panelTop - 1);
    expect(button.bottom, button.id).toBeLessThanOrEqual(libraryGeometry.panelBottom + 1);
    if (locale === 'ja') {
      if (index === 0) {
        expect(button.width, button.id).toBeGreaterThanOrEqual(37);
        expect(button.width, button.id).toBeLessThanOrEqual(39);
      } else {
        expect(button.width, button.id).toBeGreaterThanOrEqual(49);
        expect(button.width, button.id).toBeLessThanOrEqual(51);
      }
    } else {
      expect(button.width, button.id).toBeGreaterThanOrEqual(index === 0 ? 38 : 50);
      if (index > 0) {
        const previous = libraryGeometry.buttons[index - 1];
        expect(button.left, button.id).toBeGreaterThanOrEqual((previous?.right ?? 0) - 1);
      }
    }
    expect(button.hasMenu, button.id).toBe(index > 0);
  }
  const labelsRequiringContainment = new Set([
    'mac.formulas.insertFunction',
    'mac.formulas.lookup',
    'mac.formulas.more',
  ]);
  const containmentButtons = libraryGeometry.buttons.filter((button) =>
    labelsRequiringContainment.has(button.id ?? ''),
  );
  expect(containmentButtons).toHaveLength(3);
  expect(containmentButtons.map((button) => button.id).sort()).toEqual(
    [...labelsRequiringContainment].sort(),
  );
  for (const button of containmentButtons) {
    expect(button.labelFragments.length, `${button.id} text fragments`).toBeGreaterThan(0);
    for (const fragment of button.labelFragments) {
      expect(fragment.left, `${button.id} label left`).toBeGreaterThanOrEqual(button.left - 1);
      expect(fragment.right, `${button.id} label right`).toBeLessThanOrEqual(button.right + 1);
      expect(fragment.top, `${button.id} label top`).toBeGreaterThanOrEqual(button.top - 1);
      expect(fragment.bottom, `${button.id} label bottom`).toBeLessThanOrEqual(button.bottom + 1);
      expect(fragment.left, `${button.id} panel left`).toBeGreaterThanOrEqual(
        libraryGeometry.panelLeft - 1,
      );
      expect(fragment.right, `${button.id} panel right`).toBeLessThanOrEqual(
        libraryGeometry.panelRight + 1,
      );
      expect(fragment.top, `${button.id} panel top`).toBeGreaterThanOrEqual(
        libraryGeometry.panelTop - 1,
      );
      expect(fragment.bottom, `${button.id} panel bottom`).toBeLessThanOrEqual(
        libraryGeometry.panelBottom + 1,
      );
    }
  }
  await page.locator('[data-ribbon-toggle]').click();
  await page.locator('[data-ribbon-display-option="singleLine"]').click();
  await page.setViewportSize({ width: 1440, height: 900 });
  await page.locator('[data-ribbon-tab="formulas"]').click();
  const singleLineGeometry = await functionLibrary.evaluate((group) => {
    const panel = group.closest<HTMLElement>('.fc-tb__ribbon');
    const buttons = Array.from(
      group.querySelectorAll<HTMLButtonElement>(
        '.fc-tb__ribbon-tools > button[data-ribbon-command]',
      ),
    );
    const panelRect = panel?.getBoundingClientRect();
    return {
      panelHeight: panelRect?.height ?? 0,
      buttons: buttons.map((button) => {
        const rect = button.getBoundingClientRect();
        const icon = button.querySelector<SVGElement>('.fc-tb__rb-icon')?.getBoundingClientRect();
        const label = Array.from(button.children)
          .filter((child): child is HTMLElement => child instanceof HTMLElement)
          .find(
            (child) =>
              !child.classList.contains('fc-tb__rb-icon') &&
              !child.classList.contains('fc-tb__rb-split-chevron'),
          )
          ?.getBoundingClientRect();
        const chevron = button
          .querySelector<HTMLElement>(':scope > .fc-tb__rb-split-chevron')
          ?.getBoundingClientRect();
        return {
          id: button.dataset.ribbonCommand,
          left: rect.left,
          right: rect.right,
          top: rect.top,
          bottom: rect.bottom,
          width: rect.width,
          height: rect.height,
          iconLeft: icon?.left ?? 0,
          iconTop: icon?.top ?? 0,
          iconWidth: icon?.width ?? 0,
          iconHeight: icon?.height ?? 0,
          labelLeft: label?.left ?? 0,
          labelRight: label?.right ?? 0,
          labelTop: label?.top ?? 0,
          labelBottom: label?.bottom ?? 0,
          chevronLeft: chevron?.left ?? 0,
          chevronRight: chevron?.right ?? 0,
          chevronTop: chevron?.top ?? 0,
          chevronBottom: chevron?.bottom ?? 0,
          hasMenu: button.getAttribute('aria-haspopup') === 'menu',
        };
      }),
    };
  });
  expect(singleLineGeometry.panelHeight).toBe(42);
  expect(singleLineGeometry.buttons).toHaveLength(10);
  const firstSingleLineButton = singleLineGeometry.buttons[0];
  const secondSingleLineButton = singleLineGeometry.buttons[1];
  expect(
    (secondSingleLineButton?.left ?? 0) - (firstSingleLineButton?.right ?? 0),
    'single-line Insert Function gap',
  ).toBeLessThanOrEqual(4);
  for (const button of singleLineGeometry.buttons) {
    expect(button.height, button.id).toBeGreaterThanOrEqual(29);
    expect(button.height, button.id).toBeLessThanOrEqual(31);
    expect(button.iconWidth, button.id).toBe(16);
    expect(button.iconHeight, button.id).toBe(16);
    const buttonCenter = (button.top + button.bottom) / 2;
    const iconCenter = (button.iconTop + button.iconTop + button.iconHeight) / 2;
    expect(Math.abs(iconCenter - buttonCenter), button.id).toBeLessThanOrEqual(3);
    const labelCenter = (button.labelTop + button.labelBottom) / 2;
    expect(Math.abs(labelCenter - buttonCenter), button.id).toBeLessThanOrEqual(3);
    if (button.hasMenu) {
      expect(button.chevronRight, button.id).toBeLessThanOrEqual(button.right + 1);
      expect(button.chevronLeft, button.id).toBeGreaterThanOrEqual(button.left - 1);
      expect(button.chevronTop, button.id).toBeGreaterThanOrEqual(button.top - 1);
      expect(button.chevronBottom, button.id).toBeLessThanOrEqual(button.bottom + 1);
      const chevronCenter = (button.chevronTop + button.chevronBottom) / 2;
      expect(Math.abs(chevronCenter - buttonCenter), button.id).toBeLessThanOrEqual(4);
    }
  }
  await page.locator('[data-ribbon-toggle]').click();
  await page.locator('[data-ribbon-display-option="full"]').click();
  await expect(page.locator('.fc-tb__ribbon-shell--full')).toHaveCount(1);
  // Compact groups must stay inside the ribbon in either display mode.
  for (const mode of ['full', 'singleLine']) {
    if (mode === 'singleLine') {
      await page.locator('[data-ribbon-toggle]').click();
      await page.locator('[data-ribbon-display-option="singleLine"]').click();
    }
    for (const width of [1440, 1024, 768, 390]) {
      await page.setViewportSize({ width, height: 900 });
      for (const id of tabs) {
        await page.locator(`[data-ribbon-tab="${id}"]`).click();
        const clipped = await page.locator(`[data-ribbon-panel="${id}"]`).evaluate((panel) => {
          const bounds = panel.getBoundingClientRect();
          return Array.from(
            panel.querySelectorAll<HTMLElement>(
              '.fc-tb__ribbon-tools > button, .fc-tb__ribbon-tools > .fc-tb__rb-dd, .fc-tb__ribbon-tools > .fc-tb__rb-color',
            ),
          )
            .filter((command) => {
              const rect = command.getBoundingClientRect();
              return (
                rect.height > 0 &&
                (rect.top < bounds.top - 1 ||
                  rect.bottom > bounds.bottom + 1 ||
                  Array.from(command.children).some((child) => {
                    const content = child.getBoundingClientRect();
                    return (
                      content.height > 0 &&
                      (content.top < bounds.top - 1 ||
                        content.bottom > bounds.bottom + 1 ||
                        content.left < rect.left - 1 ||
                        content.right > rect.right + 1)
                    );
                  }))
              );
            })
            .map((command) => command.dataset.ribbonCommand);
        });
        expect(clipped, `${id} controls clipped in ${mode} mode at ${width}px`).toEqual([]);
      }
    }
  }
}
