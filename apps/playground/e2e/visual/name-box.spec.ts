import { expect, test } from '@playwright/test';

import { mountVisualPage } from './helpers.js';

test('@visual name box — defined-name dropdown', async ({ page }) => {
  await mountVisualPage(page, '/?theme=light&locale=en');
  await page.evaluate(() => {
    const inst = (
      window as unknown as {
        __fcInst?: {
          workbook: {
            definedNames: () => Iterable<{ name: string; formula: string }>;
          };
        };
      }
    ).__fcInst;
    if (!inst) throw new Error('missing spreadsheet instance');
    inst.workbook.definedNames = function* () {
      yield { name: 'Sales_Q1', formula: 'Sheet1!$B$2:$C$4' };
      yield { name: 'Totals', formula: 'Sheet1!$E$2:$E$8' };
    };
  });

  const nameBox = page.locator('.fc-host__formulabar-tag');
  await nameBox.focus();
  await page.keyboard.press('Alt+ArrowDown');
  const menu = page.locator('.fc-namebox-menu');
  await expect(menu).toBeVisible();
  await expect(menu).toHaveScreenshot('name-box-defined-name-dropdown.png', {
    maxDiffPixels: 80,
    animations: 'disabled',
  });
});
