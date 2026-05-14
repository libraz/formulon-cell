import { test } from '@playwright/test';

import { SpreadsheetPage } from '../../../tests/e2e-shared/pages/SpreadsheetPage.js';

test('DEBUG: fill scenario', async ({ page }) => {
  const sp = new SpreadsheetPage(page);
  await sp.mount();

  const readState = async () =>
    page.evaluate(() => {
      // biome-ignore lint/suspicious/noExplicitAny: debug
      const inst = (window as any).__fcInst;
      const s = inst?.store?.getState?.();
      return s
        ? {
            active: s.selection.active,
            range: s.selection.range,
            editor: s.ui.editor.kind,
          }
        : null;
    });

  // biome-ignore lint/suspicious/noConsole: debug
  console.log('initial', await readState());

  await sp.typeIntoActiveCell('seed');
  // biome-ignore lint/suspicious/noConsole: debug
  console.log('post-type', await readState());

  await page.keyboard.press('ArrowUp');
  // biome-ignore lint/suspicious/noConsole: debug
  console.log('post-up', await readState());

  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  // biome-ignore lint/suspicious/noConsole: debug
  console.log('post-shift-down x2', await readState());

  await sp.shortcut('d');
  // biome-ignore lint/suspicious/noConsole: debug
  console.log('post-mod-d', await readState());

  await page.keyboard.press('Escape');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  const fb = await page.locator('.fc-host__formulabar-input').inputValue();
  // biome-ignore lint/suspicious/noConsole: debug
  console.log('post-walk', await readState(), 'fb=', fb);
});
