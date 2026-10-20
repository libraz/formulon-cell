import { expect, type Page } from '@playwright/test';

import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMacHomeMenusScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac' });
  await sp.expectNoStub();

  const alignment = page.locator('[data-ribbon-panel="home"] .fc-tb__ribbon-group--alignment');
  await expect(alignment).toBeVisible();
  await expect(alignment.locator('[data-ribbon-command="wrap"]')).toHaveClass(/fc-tb__rb--wide/);
  await expect(alignment.locator('[data-ribbon-command="wrap"]')).toHaveClass(/fc-tb__rb--stacked/);
  await expect(alignment.locator('[data-ribbon-command="merge"]')).toHaveClass(/fc-tb__rb--wide/);
  await expect(alignment.locator('[data-ribbon-command="merge"]')).toHaveClass(
    /fc-tb__rb--stacked/,
  );
  await expect(alignment.locator('[data-ribbon-command="wrap"]')).toContainText('Wrap Text');
  await expect(alignment.locator('[data-ribbon-command="merge"]')).toContainText('Merge & Center');
  const alignmentLayout = await alignment.locator('.fc-tb__ribbon-tools').evaluate((tools) => {
    const style = getComputedStyle(tools);
    const command = (id: string): DOMRect =>
      tools.querySelector<HTMLElement>(`[data-ribbon-command="${id}"]`)?.getBoundingClientRect() ??
      new DOMRect();
    const wrap = command('wrap');
    const merge = command('merge');
    return {
      display: style.display,
      rows: style.gridTemplateRows,
      breakDisplay: getComputedStyle(
        tools.querySelector<HTMLElement>('[data-ribbon-command="alignment-row-2"]') as HTMLElement,
      ).display,
      wrapMergeColumn: Math.round(wrap.left),
      mergeColumn: Math.round(merge.left),
      wrapRow: Math.round(wrap.top),
      mergeRow: Math.round(merge.top),
      wrapLabelFits:
        (tools.querySelector<HTMLElement>('[data-ribbon-command="wrap"]')?.scrollWidth ?? 0) <=
        (tools.querySelector<HTMLElement>('[data-ribbon-command="wrap"]')?.clientWidth ?? 0),
      mergeLabelFits:
        (tools.querySelector<HTMLElement>('[data-ribbon-command="merge"]')?.scrollWidth ?? 0) <=
        (tools.querySelector<HTMLElement>('[data-ribbon-command="merge"]')?.clientWidth ?? 0),
    };
  });
  expect(alignmentLayout.display).toBe('grid');
  expect(alignmentLayout.rows).toBe('26px 26px');
  expect(alignmentLayout.breakDisplay).toBe('none');
  expect(alignmentLayout.wrapMergeColumn).toBe(alignmentLayout.mergeColumn);
  expect(alignmentLayout.mergeRow).toBeGreaterThan(alignmentLayout.wrapRow);
  expect(alignmentLayout.wrapLabelFits).toBe(true);
  expect(alignmentLayout.mergeLabelFits).toBe(true);

  const home = page.locator('[data-ribbon-panel="home"]');
  for (const [command, label] of [
    ['cut', 'Cut'],
    ['copy', 'Copy'],
    ['formatPainter', 'Format'],
    ['autosum', 'AutoSum'],
    ['fillHome', 'Fill'],
    ['clearFormat', 'Clear'],
  ] as const) {
    const button = home.locator(`.fc-tb__rb[data-ribbon-command="${command}"]`);
    await expect(button).toHaveClass(/fc-tb__rb--stacked/);
    await expect(button).toContainText(label);
  }
  for (const command of ['insertRows', 'deleteRows', 'formatCellsHome']) {
    const button = home.locator(`.fc-tb__rb[data-ribbon-command="${command}"]`);
    await expect(button).not.toHaveClass(/fc-tb__rb--stacked/);
    await expect(button.locator('xpath=ancestor::section[1]')).toHaveClass(
      /fc-tb__ribbon-group--tiles/,
    );
  }

  await sp.enter('A1', 'Formatted');
  await sp.goTo('A1');
  const underline = page.locator('.fc-tb__rb[data-ribbon-command="underline"]');
  await underline.locator('.fc-tb__rb-split-chevron').click();
  const menu = page.locator('#menu-underline');
  await expect(menu).toBeVisible();
  await menu.locator('[data-underline-action="double"]').click();
  const read = () =>
    page.evaluate(
      () =>
        (
          window as unknown as {
            __fcInst: {
              store: { getState(): { format: { formats: Map<string, { underline?: string }> } } };
            };
          }
        ).__fcInst.store
          .getState()
          .format.formats.get('0:0:0')?.underline,
    );
  await expect.poll(read).toBe('double');
  await page.locator('.fc-host').focus();
  await page.keyboard.press('Meta+z');
  await expect.poll(read).not.toBe('double');
  await page.locator('.fc-tb__rb[data-ribbon-command="merge"] .fc-tb__rb-split-chevron').click();
  await expect(page.locator('#menu-merge')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(page.locator('#menu-merge')).toBeHidden();
  await sp.expectNoConsoleErrors();
}
