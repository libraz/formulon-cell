import { expect, type Page, type TestInfo } from '@playwright/test';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

const LOCALES = ['en', 'ja', 'de', 'fr', 'zh', 'ko', 'th'] as const;

export async function runLocaleVisualLayout(
  page: Page,
  info: TestInfo,
  platform: 'default' | 'mac',
  width: number,
): Promise<void> {
  await page.setViewportSize({ width, height: 800 });
  const spreadsheet = new SpreadsheetPage(page);
  await spreadsheet.mount({ locale: 'en', platform });
  await spreadsheet.expectNoStub();
  const shell = page.locator(`[data-fc-platform="${platform}"] > .fc-tb__ribbon-shell`).first();
  await expect(shell).toBeVisible();
  await page.evaluate(() => {
    const instance = (
      window as unknown as {
        __fcInst: { setFeatures: (options: { localeSwitcher: boolean }) => void };
      }
    ).__fcInst;
    instance.setFeatures({ localeSwitcher: true });
  });
  const localeSelect = page.locator('.fc-host__locale-switcher-select').first();
  for (const locale of LOCALES) {
    await localeSelect.selectOption(locale);
    await expect(shell.locator('..')).toHaveAttribute('lang', locale);
    const tabs = await shell
      .locator('[data-ribbon-tab]')
      .evaluateAll((nodes) =>
        nodes
          .map((node) => (node as HTMLElement).dataset.ribbonTab)
          .filter((tab) => tab !== 'file'),
      );
    for (const tab of tabs) {
      const trigger = shell.locator(`[data-ribbon-tab="${tab}"]`).first();
      await trigger.click();
      await expect(trigger).toHaveAttribute('aria-selected', 'true');
      const rail = shell.locator('.fc-tb__ribbon:not([hidden])').first();
      const geometry = await rail.evaluate((node) => ({
        clientWidth: node.clientWidth,
        scrollWidth: node.scrollWidth,
        overflow: getComputedStyle(node).overflowX,
      }));
      if (geometry.scrollWidth > geometry.clientWidth + 1) {
        expect(
          ['auto', 'scroll'],
          `${locale}/${platform}/${tab} commands must be reachable`,
        ).toContain(geometry.overflow);
      }
      const clippedChildren = await rail.evaluate((panel) => {
        const panelBounds = panel.getBoundingClientRect();
        const controls = Array.from(
          panel.querySelectorAll<HTMLElement>(
            '.fc-tb__ribbon-tools > button, .fc-tb__ribbon-tools > .fc-tb__rb-dd, .fc-tb__ribbon-tools > .fc-tb__rb-color',
          ),
        );
        return controls.flatMap((control) => {
          const controlBounds = control.getBoundingClientRect();
          if (controlBounds.height <= 0 || controlBounds.width <= 0) return [];
          const isDropdown = control.matches('.fc-tb__rb-dd, .fc-tb__rb-color');
          const fullLabel =
            control.getAttribute('title')?.trim() ||
            control.getAttribute('aria-label')?.trim() ||
            control
              .querySelector<HTMLElement>(':scope > button')
              ?.getAttribute('aria-label')
              ?.trim();
          const allowsEllipsis = isDropdown && Boolean(fullLabel);
          const violations: string[] = [];
          if (
            controlBounds.top < panelBounds.top - 1 ||
            controlBounds.bottom > panelBounds.bottom + 1
          ) {
            violations.push('control-vertical');
          }
          for (const child of Array.from(control.children)) {
            const childElement = child as HTMLElement;
            const childBounds = childElement.getBoundingClientRect();
            if (childBounds.height <= 0 || childBounds.width <= 0) continue;
            const verticalOverflow =
              childBounds.top < panelBounds.top - 1 ||
              childBounds.bottom > panelBounds.bottom + 1 ||
              childBounds.top < controlBounds.top - 1 ||
              childBounds.bottom > controlBounds.bottom + 1;
            if (verticalOverflow) violations.push('child-vertical');
            const horizontalOverflow =
              childBounds.left < controlBounds.left - 1 ||
              childBounds.right > controlBounds.right + 1;
            const style = getComputedStyle(childElement);
            const textEllipsis =
              childElement.scrollWidth > childElement.clientWidth + 1 &&
              style.textOverflow === 'ellipsis';
            if (horizontalOverflow && !(allowsEllipsis && textEllipsis)) {
              violations.push('child-horizontal');
            }
          }
          const command =
            control.dataset.ribbonCommand ?? control.dataset.ribbonControl ?? control.className;
          return violations.length ? [{ command, violations }] : [];
        });
      });
      expect(
        clippedChildren,
        `${locale}/${platform}/${tab} ribbon children must stay inside their controls`,
      ).toEqual([]);
      const commands = rail.locator(
        'button[data-ribbon-command]:visible:not(:disabled):not([aria-disabled="true"])',
      );
      if (await commands.count()) {
        const last = commands.last();
        await last.scrollIntoViewIfNeeded();
        await last.focus();
        const bounds = await last.boundingBox();
        const container = await rail.boundingBox();
        expect(bounds).not.toBeNull();
        expect(container).not.toBeNull();
        if (bounds && container) {
          expect(bounds.x).toBeGreaterThanOrEqual(container.x - 2);
          expect(bounds.x + bounds.width).toBeLessThanOrEqual(container.x + container.width + 2);
        }
      }
      if (platform === 'default' && tab === 'view') {
        const protect = rail.locator('button[data-ribbon-command="protect"]');
        await protect.evaluate((node) => (node as HTMLElement).focus({ preventScroll: true }));
        await protect.locator('.fc-tb__rb-split-chevron').click();
        const menu = rail.locator('[id="menu-protect-view"]');
        await expect(menu).toBeVisible();
        await expect(protect).toHaveAttribute('aria-expanded', 'true');
        await page.keyboard.press('Escape');
        await expect(menu).toBeHidden();
        await expect(protect).toHaveAttribute('aria-expanded', 'false');
        await expect(protect).toBeFocused();
      }
      const missingLabels = await rail
        .locator('[data-ribbon-command]:visible')
        .evaluateAll((nodes) =>
          nodes
            .filter((node) => {
              const control = node.matches('button, input, select, [role="button"]')
                ? node
                : node.querySelector(
                    'button, select, input:not([aria-hidden="true"]), [role="button"]',
                  );
              return control && !control.getAttribute('aria-label')?.trim();
            })
            .map((node) => (node as HTMLElement).dataset.ribbonCommand),
        );
      expect(missingLabels).toEqual([]);
      const iconFailures = await rail
        .locator('[data-ribbon-command]:visible')
        .evaluateAll((nodes) =>
          nodes.flatMap((node) =>
            Array.from(node.querySelectorAll<SVGSVGElement>('svg.fc-tb__rb-icon'))
              .filter((icon) => {
                const bounds = icon.getBoundingClientRect();
                return (
                  bounds.width < 8 ||
                  bounds.height < 8 ||
                  (icon.closest('.fc-tb__ribbon-group--font') !== null &&
                    Math.abs(bounds.width - bounds.height) > 1) ||
                  icon.getAttribute('aria-hidden') !== 'true'
                );
              })
              .map(() => (node as HTMLElement).dataset.ribbonCommand),
          ),
        );
      expect(iconFailures, `${locale}/${platform}/${tab} icons must retain their size`).toEqual([]);
    }
    await shell.locator('[data-ribbon-tab="home"]').click();
    if (locale === 'de') {
      await expect(shell.locator('[data-ribbon-command="bold"] svg text')).toHaveText('F');
      await expect(shell.locator('[data-ribbon-command="italic"] svg text')).toHaveText('K');
    }
    if (locale === 'fr') {
      await expect(shell.locator('[data-ribbon-command="bold"] svg text')).toHaveText('G');
      await expect(shell.locator('[data-ribbon-command="underline"] svg text')).toHaveText('S');
    }
    const paste = shell.locator('[data-ribbon-command="paste"]').first();
    await paste.scrollIntoViewIfNeeded();
    const label = await paste.getAttribute('aria-label');
    expect(label?.trim()).toBeTruthy();
    await paste.hover();
    const tip = page.locator('.fc-tb__screen-tip[role="tooltip"]:visible');
    await expect(tip).toBeVisible();
    await expect(tip).toContainText(label ?? '');
    const tipBounds = await tip.boundingBox();
    expect(tipBounds).not.toBeNull();
    if (tipBounds) {
      expect(tipBounds.x).toBeGreaterThanOrEqual(0);
      expect(tipBounds.y).toBeGreaterThanOrEqual(0);
      expect(tipBounds.x + tipBounds.width).toBeLessThanOrEqual(width);
      expect(tipBounds.y + tipBounds.height).toBeLessThanOrEqual(800);
    }
    await paste.focus();
    await expect(paste).toHaveAttribute('aria-describedby', (await tip.getAttribute('id')) ?? '');
    await page.keyboard.press('Escape');
    await expect(tip).toHaveCount(0);
    await shell.locator('[data-ribbon-command="fontFamily"] button').first().click();
    const fontOptions = page.locator('.fc-tb__rb-dd__list [role="option"]');
    await expect(fontOptions.first()).toBeVisible();
    const missingFontLabels = await fontOptions.evaluateAll((options) =>
      options.flatMap((option) => {
        const name = option.querySelector('.fc-tb__rb-dd__label')?.textContent?.trim();
        const title = option.getAttribute('title');
        const accessibleName = option.getAttribute('aria-label');
        const describedBy = option.getAttribute('aria-describedby');
        const tooltipText = describedBy
          ? option.ownerDocument.getElementById(describedBy)?.textContent
          : null;
        const fullText = title ?? tooltipText;
        return name && accessibleName?.includes(name) && fullText?.includes(accessibleName)
          ? []
          : [name];
      }),
    );
    expect(
      missingFontLabels,
      `${locale}/${platform} font candidates expose their full names`,
    ).toEqual([]);
    await page.keyboard.press('Escape');
    if (width === 1280) {
      await page.screenshot({ path: info.outputPath(`${platform}-${locale}-home.png`) });
    }
    for (const button of await page.locator('.demo__quick .demo__title-icon').all()) {
      expect(await button.getAttribute('title')).toBe(await button.getAttribute('aria-label'));
    }
  }
}
