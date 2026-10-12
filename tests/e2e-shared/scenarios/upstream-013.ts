import { expect, type Page } from '@playwright/test';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runNativeR1C1Scenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ fixture: 'empty' });
  await sp.expectNoStub();
  await sp.enter('A1', '5');
  await sp.enter('B2', '7');
  await sp.enter('C3', '=A1+$B$2');
  await sp.goTo('C3');
  await page.locator('[data-ribbon-tab="view"]').click();
  await sp.clickRibbon('viewR1C1');
  const formulaBar = page.locator('.fc-host__formulabar-input').first();
  await expect(formulaBar).toHaveValue('=R[-2]C[-2]+R2C2');
  await formulaBar.fill('=R[-2]C[-2]+R2C2+1');
  await formulaBar.press('Enter');
  await expect.poll(() => sp.readValue('C3')).toEqual({ kind: 'number', value: 13 });
  await expect.poll(() => sp.readFormula('C3')).toBe('=A1+$B$2+1');
  await sp.goTo('R3C3');
  await sp.clickRibbon('viewR1C1');
  await expect(formulaBar).toHaveValue('=A1+$B$2+1');
  await sp.expectNoConsoleErrors();
}

type NativeProbe = {
  workbook: {
    getTheme(): { colors: number[]; fonts: { minorLatin: string } };
    getThreadedComments(sheet: number): { text: string; done: boolean; parentId: string }[];
    listDrawingObjects(sheet: number): { objectId: number }[];
    getImage(sheet: number, objectId: number): { bytes: Uint8Array } | null;
  };
};

const readNative = (page: Page) =>
  page.evaluate(() => {
    const instance = (window as Window & { __fcInst: NativeProbe }).__fcInst;
    const objects = instance.workbook.listDrawingObjects(0);
    return {
      theme: instance.workbook.getTheme(),
      comments: instance.workbook.getThreadedComments(0),
      imageSizes: objects.map(
        (object) => instance.workbook.getImage(0, object.objectId)?.bytes.length ?? 0,
      ),
    };
  });

export async function runNativeThreadThemeScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ fixture: 'empty', platform: 'mac' });
  await sp.expectNoStub();
  await sp.goTo('B3');
  await page.locator('[data-ribbon-tab="review"]').click();
  await sp.clickRibbon('mac.review.newComment');
  const dialog = page.locator('.fc-cmtnote').first();
  await expect(dialog).toBeVisible();
  await dialog.locator('textarea').fill('Review this total');
  await dialog.locator('.fc-cmtnote__btn--primary').click();
  await expect
    .poll(() => readNative(page))
    .toMatchObject({
      comments: [{ text: 'Review this total', parentId: '' }],
    });
  await sp.clickRibbon('mac.review.newComment');
  await dialog.locator('.fc-cmtnote__reply').click();
  await dialog.locator('textarea').fill('Confirmed');
  await dialog.locator('.fc-cmtnote__btn--primary').click();
  await expect
    .poll(async () => (await readNative(page)).comments.map((c) => c.text))
    .toEqual(['Review this total', 'Confirmed']);
  await sp.clickRibbon('mac.review.newComment');
  await dialog.locator('.fc-cmtnote__resolve').click();
  await page.keyboard.press('Escape');
  await expect.poll(async () => (await readNative(page)).comments[0]?.done).toBe(true);
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  await sp.clickRibbon('mac.page.theme');
  await page.locator('[data-ribbon-command="mac.page.theme.gray"]').click();
  await expect.poll(async () => (await readNative(page)).theme.fonts.minorLatin).toBe('Arial');
  await sp.expectNoConsoleErrors();
}

export async function runNativeImageSaveScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount({ fixture: 'empty' });
  await sp.expectNoStub();
  await sp.goTo('B3');
  await page.locator('[data-ribbon-tab="review"]').click();
  await sp.openRibbonSplitMenu('newCommentReview', 'menu-review-new-comment');
  await page.locator('#menu-review-new-comment [data-comment-action="thread"]').click();
  const comment = page.locator('.fc-cmtnote').first();
  await comment.locator('textarea').fill('Persist this thread');
  await comment.locator('.fc-cmtnote__btn--primary').click();
  await sp.goTo('C4');
  await page.locator('[data-ribbon-tab="insert"]').click();
  await sp.clickRibbon('pictureInsert');
  const [chooser] = await Promise.all([
    page.waitForEvent('filechooser'),
    page.locator('#menu-picture-insert [data-picture-insert="device"]').click(),
  ]);
  const png = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64',
  );
  await chooser.setFiles({ name: 'pixel.png', mimeType: 'image/png', buffer: png });
  await expect.poll(async () => (await readNative(page)).imageSizes).toEqual([png.length]);
  await expect(page.locator('.fc-illustration img')).toHaveCount(1);
  await page.locator('[data-ribbon-tab="pageLayout"]').click();
  await sp.clickRibbon('pageTheme');
  await page.locator('#menu-page-theme [data-workbook-theme-action="gray"]').click();
  await expect.poll(async () => (await readNative(page)).theme.fonts.minorLatin).toBe('Arial');
  const saved = await sp.saveDownload();
  await sp.openDownload(saved);
  await expect
    .poll(async () => (await readNative(page)).comments.map((c) => c.text))
    .toEqual(['Persist this thread']);
  await expect.poll(async () => (await readNative(page)).imageSizes).toEqual([png.length]);
  await expect(page.locator('.fc-illustration img')).toHaveCount(1);
  await expect.poll(async () => (await readNative(page)).theme.fonts.minorLatin).toBe('Arial');
  await sp.expectNoConsoleErrors();
}
