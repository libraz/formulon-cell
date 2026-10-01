import { expect, type Page } from '@playwright/test';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runMergedStructureCopyScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('A2', '7');
  await sp.enter('D7', '=SUM($A$2:$A$4)');
  await sp.goTo('A2:B4');
  await sp.clickRibbon('bold');
  const mergeMenu = await sp.openRibbonSplitMenu('merge', 'menu-merge');
  await mergeMenu.locator('[data-merge-action="mergeCells"]').click();
  await expect.poll(() => sp.readMerge('A2')).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 });
  await sp.goTo('D3');
  await sp.insertRowsOrColumns('rows');
  await expect.poll(() => sp.readMerge('A2')).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 4, c1: 1 });
  await expect.poll(() => sp.readFormula('D8')).toBe('=SUM($A$2:$A$5)');
  await sp.goTo('A2');
  await sp.shortcut('c');
  await sp.goTo('F2');
  await sp.shortcut('v');
  await expect.poll(() => sp.readMerge('F2')).toEqual({ sheet: 0, r0: 1, c0: 5, r1: 4, c1: 6 });
  await expect.poll(() => sp.readValue('F2')).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readFormat('G5')).toMatchObject({ bold: true });
  await sp.shortcut('z');
  await expect.poll(() => sp.readMerge('F2')).toBeNull();
  await sp.shortcut('y');
  await expect.poll(() => sp.readMerge('F2')).not.toBeNull();
  await page.keyboard.press('Escape');
  await sp.goTo('D2');
  await sp.deleteRowsOrColumns('rows');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readMerge('A2')).toEqual({ sheet: 0, r0: 1, c0: 0, r1: 3, c1: 1 });
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readFormula('D8')).toBe('=SUM($A$2:$A$5)');
  await sp.expectNoConsoleErrors();
}

export async function runCrossSheetCutReferencesScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.renameSelectedSheet('Source');
  await page.locator('.fc-host__sheetbar-add').click();
  await sp.renameSelectedSheet('Target');
  await sp.chooseSheet('Source');
  await sp.enter('A1', '7');
  await sp.enter('B1', '=A1');
  await sp.enter('C1', '=B1');
  await sp.goTo('B1');
  await sp.clickRibbon('bold');
  await sp.goTo('B1');
  await sp.shortcut('x');
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (
          window as Window & {
            __fcInst?: { store: { getState(): { ui: { copyMode: string | null } } } };
          }
        ).__fcInst;
        return inst?.store.getState().ui.copyMode;
      }),
    )
    .toBe('cut');
  expect(await sp.readValue('B1')).toEqual({ kind: 'number', value: 7 });
  await sp.chooseSheet('Target');
  await sp.goTo('D3');
  await sp.shortcut('v');
  await expect.poll(() => sp.readFormula('D3', 1)).toBe('=Source!A1');
  await expect.poll(() => sp.readFormula('C1', 0)).toBe('=Target!D3');
  await expect.poll(() => sp.readValue('D3', 1)).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readValue('B1', 0)).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readFormat('D3', 1)).toMatchObject({ bold: true });
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('B1', 0)).toBe('=A1');
  await expect.poll(() => sp.readFormula('C1', 0)).toBe('=B1');
  await expect.poll(() => sp.readValue('D3', 1)).toEqual({ kind: 'blank' });
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('D3', 1)).toBe('=Source!A1');
  await expect.poll(() => sp.readFormula('C1', 0)).toBe('=Target!D3');
  await sp.expectNoConsoleErrors();
}

export async function runPartialRangeCutScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.renameSelectedSheet('Source');
  await page.locator('.fc-host__sheetbar-add').click();
  await sp.renameSelectedSheet('Target');
  await sp.chooseSheet('Source');
  await sp.enter('A1', '1');
  await sp.enter('A2', '2');
  await sp.enter('A3', '3');
  await sp.enter('F1', '=SUM(A1:A3)');
  await sp.enter('F2', '=SUM(Target!A1:A3)');
  await sp.goTo('A1');
  await sp.shortcut('x');
  await sp.chooseSheet('Target');
  await sp.goTo('D3');
  await sp.shortcut('v');
  await expect.poll(() => sp.readFormula('F1', 0)).toBe('=SUM(A2:A3)');
  await expect.poll(() => sp.readFormula('F2', 0)).toBe('=SUM(Target!A1:A3)');
  await expect.poll(() => sp.readValue('F1', 0)).toEqual({ kind: 'number', value: 5 });
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('F1', 0)).toBe('=SUM(A1:A3)');
  await expect.poll(() => sp.readValue('A1', 0)).toEqual({ kind: 'number', value: 1 });
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('F1', 0)).toBe('=SUM(A2:A3)');
  await expect.poll(() => sp.readValue('D3', 1)).toEqual({ kind: 'number', value: 1 });
  await sp.expectNoConsoleErrors();
}
