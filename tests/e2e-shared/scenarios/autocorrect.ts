import { expect, type Page } from '@playwright/test';

import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

type InstanceWindow = Window & { __fcInst: SpreadsheetInstance };

/** Typed entry and the text AutoCorrect leaves in the cell. */
const CASES = [
  // Han and kana entries are replaced mid-text, as those scripts put no spaces between words.
  ['A1', '今日はこんにちわ', '今日はこんにちは'],
  ['A2', 'こんにちわ世界', 'こんにちは世界'],
  ['A3', '御社様', '貴社様'],
  // A Latin entry keeps the word boundary.
  ['A4', 'teh cat', 'the cat'],
  ['A5', 'xteh cat', 'xteh cat'],
  ['A6', '日本teh', '日本teh'],
] as const;

/** AutoCorrect replaces an entry written in Han or kana wherever it is typed, and a Latin entry
 *  only as a whole word. */
export async function runAutoCorrectScriptBoundaryScenario(page: Page): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'ja' });
  await sheet.expectNoStub();
  await page.evaluate(() =>
    (window as unknown as InstanceWindow).__fcInst.update({
      autoCorrect: {
        entries: {
          ja: [
            { from: 'こんにちわ', to: 'こんにちは' },
            { from: '御社', to: '貴社' },
            { from: 'teh', to: 'the' },
          ],
        },
      },
    }),
  );
  for (const [ref, typed, stored] of CASES) {
    // Typed text without a key event does not start an entry, so F2 opens the editor first.
    await sheet.goTo(ref);
    await page.keyboard.press('F2');
    await page.keyboard.type(typed);
    await page.keyboard.press('Enter');
    await expect.poll(() => sheet.readValue(ref), typed).toEqual({ kind: 'text', value: stored });
  }
  await sheet.expectNoConsoleErrors();
}
