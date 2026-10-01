import { expect, type Page } from '@playwright/test';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

async function clickColumnHeader(
  page: Page,
  col: number,
  button: 'left' | 'right' = 'left',
): Promise<void> {
  const position = await page.evaluate((targetCol) => {
    const instance = (
      window as Window & {
        __fcInst?: {
          store: {
            getState(): {
              layout: {
                headerColWidth: number;
                headerRowHeight: number;
                defaultColWidth: number;
                colWidths: Map<number, number>;
              };
              viewport: { zoom: number };
            };
          };
        };
      }
    ).__fcInst;
    if (!instance) throw new Error('window.__fcInst is not available');
    const { layout, viewport } = instance.store.getState();
    const zoom = viewport.zoom;
    let x = layout.headerColWidth * zoom;
    for (let index = 0; index < targetCol; index += 1)
      x += (layout.colWidths.get(index) ?? layout.defaultColWidth) * zoom;
    x += ((layout.colWidths.get(targetCol) ?? layout.defaultColWidth) * zoom) / 2;
    return { x, y: (layout.headerRowHeight * zoom) / 2 };
  }, col);
  await page.locator('.fc-host__canvas').first().click({ position, button });
}

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

export async function runDirectionalFillHistoryScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('A1', 'Item 1');
  await sp.enter('A2', 'old');
  await sp.goTo('A1');
  await sp.clickRibbon('bold');
  await sp.goTo('A1:A3');
  await sp.shortcut('d');
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'text', value: 'Item 1' });
  await expect.poll(() => sp.readFormat('A3')).toMatchObject({ bold: true });
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'text', value: 'old' });
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'blank' });
  expect((await sp.readFormat('A3')) ?? {}).not.toMatchObject({ bold: true });
  await sp.shortcut('y');
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'text', value: 'Item 1' });
  await expect.poll(() => sp.readFormat('A3')).toMatchObject({ bold: true });
  await sp.expectNoConsoleErrors();
}

export async function runSortFormulaHistoryScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('A1', 'Key');
  await sp.enter('B1', 'Calculation');
  for (const [ref, value] of [
    ['A2', '3'],
    ['A3', '1'],
    ['A4', '2'],
    ['D2', '100'],
    ['D3', '200'],
    ['D4', '300'],
  ]) {
    await sp.enter(ref as string, value as string);
  }
  for (const row of [2, 3, 4]) await sp.enter(`B${row}`, `=A${row}+$D$2+D${row}`);
  await sp.enter('F2', '=A2');
  await sp.goTo('B3');
  await sp.clickRibbon('bold');
  await sp.goTo('A1:B4');
  await page.locator('[data-ribbon-tab="data"]').click();
  await sp.clickRibbon('sortAsc');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'number', value: 1 });
  await expect.poll(() => sp.readFormula('B2')).toBe('=A2+$D$2+D2');
  await expect.poll(() => sp.readValue('B2')).toEqual({ kind: 'number', value: 201 });
  await expect.poll(() => sp.readFormat('B2')).toMatchObject({ bold: true });
  expect(await sp.readFormula('F2')).toBe('=A2');
  await sp.goTo('A1:B4');
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'number', value: 3 });
  await expect.poll(() => sp.readFormat('B3')).toMatchObject({ bold: true });
  expect((await sp.readFormat('B2')) ?? {}).not.toMatchObject({ bold: true });
  await sp.shortcut('y');
  await expect.poll(() => sp.readValue('A2')).toEqual({ kind: 'number', value: 1 });
  await expect.poll(() => sp.readFormat('B2')).toMatchObject({ bold: true });
  await sp.expectNoConsoleErrors();
}

export async function runCellBandReferencesScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.renameSelectedSheet('Source');
  await page.locator('.fc-host__sheetbar-add').click();
  await sp.renameSelectedSheet('Target');
  await sp.enter('D1', '=Source!$A$2');
  await sp.enter('D2', '=SUM(Source!A1:A3)');
  await sp.enter('D3', '=A2');
  await sp.chooseSheet('Source');
  for (const row of [1, 2, 3]) await sp.enter(`A${row}`, String(row));
  await sp.enter('C1', '=Source!$A$2');
  await sp.enter('C2', '=SUM(A1:A3)');
  await sp.goTo('A2');
  await sp.insertRowsOrColumns('cells');
  await page.locator('.fc-cellshift__button--primary').click();
  await expect.poll(() => sp.readFormula('D1', 1)).toBe('=Source!$A$3');
  await expect.poll(() => sp.readFormula('D2', 1)).toBe('=SUM(Source!A1:A4)');
  expect(await sp.readFormula('D3', 1)).toBe('=A2');
  await sp.goTo('A3');
  await sp.deleteRowsOrColumns('cells');
  await page.locator('.fc-cellshift__button--primary').click();
  await expect.poll(() => sp.readFormula('D1', 1)).toBe('=Source!#REF!');
  await expect.poll(() => sp.readFormula('D2', 1)).toBe('=SUM(Source!A1:A3)');
  await sp.goTo('A3');
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('D1', 1)).toBe('=Source!$A$3');
  await expect.poll(() => sp.readValue('A3')).toEqual({ kind: 'number', value: 2 });
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('D1', 1)).toBe('=Source!#REF!');
  await sp.expectNoConsoleErrors();
}

export async function runRepeatedPasteScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('A1', '1');
  await sp.enter('B1', '=A1*2');
  await sp.enter('A2', '3');
  await sp.enter('B2', '=A2*2');
  await sp.goTo('B2');
  await sp.clickRibbon('bold');
  await sp.goTo('A1:B2');
  await sp.shortcut('c');
  await sp.goTo('D1:G4');
  await sp.shortcut('v');
  await expect.poll(() => sp.readFormula('G4')).toBe('=F4*2');
  await expect.poll(() => sp.readValue('G4')).toEqual({ kind: 'number', value: 6 });
  await expect.poll(() => sp.readFormat('G4')).toMatchObject({ bold: true });
  await expect(page.locator('.fc-host__formulabar-tag').first()).toHaveValue('D1:G4');
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('G4')).toEqual({ kind: 'blank' });
  expect((await sp.readFormat('G4')) ?? {}).not.toMatchObject({ bold: true });
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('G4')).toBe('=F4*2');
  await expect.poll(() => sp.readFormat('G4')).toMatchObject({ bold: true });
  await page.keyboard.press('Escape');
  await sp.enter('J1', '9');
  await sp.goTo('K2:L3');
  const mergeMenu = await sp.openRibbonSplitMenu('merge', 'menu-merge');
  await mergeMenu.locator('[data-merge-action="mergeCells"]').click();
  const targetMerge = { sheet: 0, r0: 1, c0: 10, r1: 2, c1: 11 };
  await expect.poll(() => sp.readMerge('K2')).toEqual(targetMerge);
  await sp.goTo('J1');
  await sp.shortcut('c');
  await sp.goTo('K2');
  await sp.shortcut('v');
  await expect.poll(() => sp.readValue('K2')).toEqual({ kind: 'number', value: 9 });
  expect(await sp.readMerge('K2')).toEqual(targetMerge);
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('K2')).toEqual({ kind: 'blank' });
  expect(await sp.readMerge('K2')).toEqual(targetMerge);
  await sp.shortcut('y');
  await expect.poll(() => sp.readValue('K2')).toEqual({ kind: 'number', value: 9 });
  expect(await sp.readMerge('K2')).toEqual(targetMerge);
  await sp.expectNoConsoleErrors();
}

export async function runWholeColumnCopyInsertScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('A5', '7');
  await sp.enter('B5', '=A5*2');
  await sp.enter('D2', 'old');
  await sp.goTo('A5');
  await sp.clickRibbon('bold');
  await clickColumnHeader(page, 0);
  await sp.shortcut('c');
  await clickColumnHeader(page, 3);
  await sp.shortcut('v');
  await expect.poll(() => sp.readValue('D5')).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readValue('D2')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readFormat('D5')).toMatchObject({ bold: true });
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('D2')).toEqual({ kind: 'text', value: 'old' });
  await expect.poll(() => sp.readValue('D5')).toEqual({ kind: 'blank' });
  await page.keyboard.press('Escape');
  await clickColumnHeader(page, 1);
  await sp.shortcut('c');
  await clickColumnHeader(page, 3, 'right');
  await page.locator('.fc-ctxmenu [data-fc-action="insertCopiedCells"]').click();
  await expect.poll(() => sp.readFormula('D5')).toBe('=C5*2');
  await expect.poll(() => sp.readValue('D1')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readValue('E2')).toEqual({ kind: 'text', value: 'old' });
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('D2')).toEqual({ kind: 'text', value: 'old' });
  await expect.poll(() => sp.readFormula('D5')).toBeNull();
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('D5')).toBe('=C5*2');
  await sp.expectNoConsoleErrors();
}

async function clickRowHeader(
  page: Page,
  row: number,
  button: 'left' | 'right' = 'left',
): Promise<void> {
  const position = await page.evaluate((targetRow) => {
    const instance = (
      window as Window & {
        __fcInst?: {
          store: {
            getState(): {
              layout: {
                headerColWidth: number;
                headerRowHeight: number;
                defaultRowHeight: number;
                rowHeights: Map<number, number>;
              };
              viewport: { zoom: number };
            };
          };
        };
      }
    ).__fcInst;
    if (!instance) throw new Error('window.__fcInst is not available');
    const { layout, viewport } = instance.store.getState();
    const zoom = viewport.zoom;
    let y = layout.headerRowHeight * zoom;
    for (let index = 0; index < targetRow; index += 1)
      y += (layout.rowHeights.get(index) ?? layout.defaultRowHeight) * zoom;
    y += ((layout.rowHeights.get(targetRow) ?? layout.defaultRowHeight) * zoom) / 2;
    return { x: (layout.headerColWidth * zoom) / 2, y };
  }, row);
  await page.locator('.fc-host__canvas').first().click({ position, button });
}

export async function runWholeRowCopyInsertScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('E1', '7');
  await sp.enter('E2', '=E1*2');
  await sp.enter('A4', 'old');
  await sp.goTo('E1');
  await sp.clickRibbon('bold');
  await clickRowHeader(page, 0);
  await sp.shortcut('c');
  await clickRowHeader(page, 3);
  await sp.shortcut('v');
  await expect.poll(() => sp.readValue('E4')).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readValue('A4')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readFormat('E4')).toMatchObject({ bold: true });
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('A4')).toEqual({ kind: 'text', value: 'old' });
  await expect.poll(() => sp.readValue('E4')).toEqual({ kind: 'blank' });
  await page.keyboard.press('Escape');
  await clickRowHeader(page, 1);
  await sp.shortcut('c');
  await clickRowHeader(page, 3, 'right');
  await page.locator('.fc-ctxmenu [data-fc-action="insertCopiedCells"]').click();
  await expect.poll(() => sp.readFormula('E4')).toBe('=E3*2');
  await expect.poll(() => sp.readValue('A4')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readValue('A5')).toEqual({ kind: 'text', value: 'old' });
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('A4')).toEqual({ kind: 'text', value: 'old' });
  await expect.poll(() => sp.readFormula('E4')).toBeNull();
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('E4')).toBe('=E3*2');
  await sp.expectNoConsoleErrors();
}

export async function runWholeBandInsertDeleteScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('B5', '7');
  await sp.enter('F1', '=B5');
  await clickColumnHeader(page, 1);
  await sp.shortcut('Shift+Equal');
  await expect.poll(() => sp.readValue('C5')).toEqual({ kind: 'number', value: 7 });
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await expect.poll(() => sp.readFormula('G1')).toBe('=C5');
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('B5')).toEqual({ kind: 'number', value: 7 });
  await clickColumnHeader(page, 1);
  await sp.shortcut('Minus');
  await expect.poll(() => sp.readValue('B5')).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readFormula('E1')).toBe('=#REF!');
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('F1')).toBe('=B5');
  await clickColumnHeader(page, 1);
  await sp.insertRowsOrColumns('cells');
  await expect.poll(() => sp.readValue('C5')).toEqual({ kind: 'number', value: 7 });
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await sp.shortcut('z');

  await sp.enter('E2', '11');
  await clickRowHeader(page, 1);
  await sp.shortcut('Shift+Equal');
  await expect.poll(() => sp.readValue('E3')).toEqual({ kind: 'number', value: 11 });
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('E2')).toEqual({ kind: 'number', value: 11 });
  await clickRowHeader(page, 1);
  await sp.shortcut('Minus');
  await expect.poll(() => sp.readValue('E2')).toEqual({ kind: 'blank' });
  await sp.shortcut('z');
  await clickRowHeader(page, 1);
  await sp.insertRowsOrColumns('cells');
  await expect.poll(() => sp.readValue('E3')).toEqual({ kind: 'number', value: 11 });
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('E2')).toEqual({ kind: 'number', value: 11 });
  await sp.expectNoConsoleErrors();
}

export async function runCopiedBandInsertRoutesScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('E5', '7');
  await sp.enter('D5', '=$E5');
  await clickColumnHeader(page, 3);
  await sp.shortcut('c');
  await sp.goTo('B5');
  await sp.shortcut('Shift+Equal');
  await expect.poll(() => sp.readFormula('B5')).toBe('=$F5');
  await expect.poll(() => sp.readFormula('E5')).toBe('=$F5');
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('D5')).toBe('=$E5');
  await expect.poll(() => sp.readFormula('B5')).toBeNull();
  await page.keyboard.press('Escape');

  await sp.enter('J5', '11');
  await sp.enter('J4', '=J$5');
  await clickRowHeader(page, 3);
  await sp.shortcut('c');
  await sp.goTo('A2');
  await sp.insertRowsOrColumns('cells');
  await expect.poll(() => sp.readFormula('J2')).toBe('=J$6');
  await expect.poll(() => sp.readFormula('J5')).toBe('=J$6');
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('J4')).toBe('=J$5');
  await expect.poll(() => sp.readFormula('J2')).toBeNull();
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('J2')).toBe('=J$6');
  await sp.expectNoConsoleErrors();
}

export async function runCutBandInsertRoutesScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.enter('B5', '2');
  await sp.enter('C5', '3');
  await sp.enter('E5', '7');
  await sp.enter('D5', '=$E5');
  await sp.enter('J10', '=D5');
  await sp.goTo('D5');
  await sp.clickRibbon('bold');
  await clickColumnHeader(page, 3);
  await sp.shortcut('x');
  await sp.goTo('B5');
  await sp.shortcut('Shift+Equal');
  await expect.poll(() => sp.readFormula('B5')).toBe('=$E5');
  await expect.poll(() => sp.readFormula('J10')).toBe('=B5');
  await expect.poll(() => sp.readValue('C5')).toEqual({ kind: 'number', value: 2 });
  await expect.poll(() => sp.readValue('D5')).toEqual({ kind: 'number', value: 3 });
  await expect.poll(() => sp.readFormat('B5')).toMatchObject({ bold: true });
  await expect(page.locator('.fc-cellshift')).toHaveCount(0);
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('D5')).toBe('=$E5');
  await expect.poll(() => sp.readFormula('J10')).toBe('=D5');
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('B5')).toBe('=$E5');
  await page.keyboard.press('Escape');

  await sp.enter('J1', '11');
  await sp.enter('J2', '=J$1');
  await sp.enter('J3', '3');
  await sp.enter('J4', '4');
  await sp.enter('K10', '=J2');
  await clickRowHeader(page, 1);
  await sp.shortcut('x');
  await sp.goTo('A5');
  await sp.insertRowsOrColumns('cells');
  await expect.poll(() => sp.readFormula('J4')).toBe('=J$1');
  await expect.poll(() => sp.readValue('J2')).toEqual({ kind: 'number', value: 3 });
  await expect.poll(() => sp.readValue('J3')).toEqual({ kind: 'number', value: 4 });
  await expect.poll(() => sp.readFormula('K10')).toBe('=J4');
  await sp.shortcut('z');
  await expect.poll(() => sp.readFormula('J2')).toBe('=J$1');
  await expect.poll(() => sp.readFormula('K10')).toBe('=J2');
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('J4')).toBe('=J$1');
  await clickRowHeader(page, 1);
  await sp.shortcut('x');
  await sp.shortcut('Shift+Equal');
  await expect.poll(() => sp.readValue('J2')).toEqual({ kind: 'number', value: 3 });
  const copyMode = () =>
    page.evaluate(() => {
      const inst = (
        window as Window & {
          __fcInst?: { store: { getState(): { ui: { copyMode: string | null } } } };
        }
      ).__fcInst;
      return inst?.store.getState().ui.copyMode;
    });
  await expect.poll(copyMode).toBe('cut');
  await clickRowHeader(page, 2);
  await sp.shortcut('Shift+Equal');
  await expect.poll(copyMode).toBeNull();
  await expect.poll(() => sp.readValue('J2')).toEqual({ kind: 'number', value: 3 });
  await sp.expectNoConsoleErrors();
}

export async function runCrossSheetCutBandInsertScenario(page: Page): Promise<void> {
  const sp = new UserJourneyPage(page);
  await sp.mount();
  await sp.expectNoStub();
  await sp.renameSelectedSheet('Source');
  await page.locator('.fc-host__sheetbar-add').click();
  await sp.renameSelectedSheet('Target');
  await sp.enter('D5', '9');
  await sp.chooseSheet('Source');
  await sp.enter('B5', '7');
  await sp.enter('C5', '8');
  await sp.enter('J10', '=B5');
  await clickColumnHeader(page, 1);
  await sp.shortcut('x');
  await sp.chooseSheet('Target');
  await clickColumnHeader(page, 3, 'right');
  await page.locator('.fc-ctxmenu [data-fc-action="insertCopiedCells"]').click();
  await expect.poll(() => sp.readValue('D5', 1)).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readValue('E5', 1)).toEqual({ kind: 'number', value: 9 });
  await expect.poll(() => sp.readValue('B5', 0)).toEqual({ kind: 'blank' });
  await expect.poll(() => sp.readValue('C5', 0)).toEqual({ kind: 'number', value: 8 });
  await expect.poll(() => sp.readFormula('J10', 0)).toBe('=Target!D5');
  await sp.shortcut('z');
  await expect.poll(() => sp.readValue('B5', 0)).toEqual({ kind: 'number', value: 7 });
  await expect.poll(() => sp.readValue('D5', 1)).toEqual({ kind: 'number', value: 9 });
  await expect.poll(() => sp.readFormula('J10', 0)).toBe('=B5');
  await sp.shortcut('y');
  await expect.poll(() => sp.readFormula('J10', 0)).toBe('=Target!D5');
  await sp.expectNoConsoleErrors();
}
