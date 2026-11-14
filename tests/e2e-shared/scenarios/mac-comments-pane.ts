import { expect, type Page } from '@playwright/test';

import { selectRibbonTab } from '../pages/ribbon-tabs.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

const EMPTY = {
  en: 'There are no comments in this file.',
  ja: 'このファイル内にコメントはありません。',
} as const;

/** Review > Show Comments toggles the docked Comments pane and reads pressed while it is open.
 *  An empty book shows the desktop's empty-state line; a posted comment is listed with its cell;
 *  the pane's close button unpresses the toggle. */
export async function runMacShowCommentsPaneScenario(
  page: Page,
  locale: keyof typeof EMPTY,
): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ platform: 'mac', locale, fixture: 'empty' });
  await page.setViewportSize({ width: 1440, height: 900 });
  await sp.expectNoStub();
  await selectRibbonTab(page, 'review');
  const toggle = page
    .locator('[data-ribbon-panel="review"] [data-ribbon-command="mac.review.showComments"]')
    .first();
  const pane = page.locator('.fc-host__taskpane-dock .fc-commentspane');
  const expectPressed = async (pressed: boolean) => {
    await expect(toggle).toHaveAttribute('aria-pressed', String(pressed));
    if (pressed) await expect(toggle).toHaveClass(/fc-tb__rb--active/);
    else await expect(toggle).not.toHaveClass(/fc-tb__rb--active/);
  };

  await expectPressed(false);
  await expect(pane).toBeHidden();
  await toggle.click();
  await expect(pane).toBeVisible();
  await expectPressed(true);
  await expect(pane.locator('.fc-commentspane__empty')).toHaveText(EMPTY[locale]);
  await expect(pane.locator('.fc-commentspane__item')).toHaveCount(0);
  await toggle.click();
  await expect(pane).toBeHidden();
  await expectPressed(false);

  await sp.goTo('C4');
  await sp.clickRibbon('mac.review.newComment');
  const card = page.locator('.fc-cmtnote').first();
  await expect(card).toBeVisible();
  await card.locator('textarea').fill('Check the total');
  await card.locator('.fc-cmtnote__send').click();
  await page.keyboard.press('Escape');
  await expect(card).toBeHidden();

  await toggle.click();
  await expect(pane).toBeVisible();
  await expectPressed(true);
  await expect(pane.locator('.fc-commentspane__empty')).toHaveCount(0);
  const item = pane.locator('.fc-commentspane__item');
  await expect(item).toHaveCount(1);
  await expect(item.locator('.fc-commentspane__cell')).toHaveText('C4');
  await expect(item).toContainText('Check the total');

  await pane.locator('.fc-commentspane__close').click();
  await expect(pane).toBeHidden();
  await expectPressed(false);
  await sp.expectNoConsoleErrors();
}
