import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

async function rowGeometry(page: Page, row: number) {
  return page.evaluate(
    ({ row }) => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const { layoutForView } = w.__fcProbes['render/geometry'];
      const state = w.__fcProbes['mount/internals']
        .spreadsheetInternals(w.__fcInst)
        .store.getState();
      const layout = layoutForView(state);
      return {
        logicalHeight: layout.rowHeights.get(row) ?? layout.defaultRowHeight,
        explicitRows: state.layout.rowHeights.size,
      };
    },
    { row },
  );
}

export async function runAutomaticRowheight(
  page: Page,
  platform: 'default' | 'mac',
  capture?: () => Promise<void>,
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await sheet.enter('A1', 'Tall text stays inside its row');
  await sheet.goTo('A1');
  const original = await rowGeometry(page, 0);
  const other = await rowGeometry(page, 1);
  const chooseFontSize = async (points: number) => {
    await page.locator('[data-ribbon-command="fontSize"] button').first().click();
    await page.locator(`.fc-tb__rb-dd__opt[data-value="${points}"]`).click();
  };

  await chooseFontSize(24);
  await expect
    .poll(async () => (await rowGeometry(page, 0)).logicalHeight)
    .toBeGreaterThan(original.logicalHeight);
  const enlarged = await rowGeometry(page, 0);
  expect(enlarged.explicitRows).toBe(0);
  expect((await rowGeometry(page, 1)).logicalHeight).toBe(other.logicalHeight);
  await capture?.();

  // Pointer geometry and the visible grid must agree after the row grows.
  const next = await gridCellCenter(page, 1, 0);
  await page.mouse.click(next.x, next.y);
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.selection.get().active
          .row,
    ),
  ).toBe(1);

  await sheet.shortcut('z');
  expect((await rowGeometry(page, 0)).logicalHeight).toBe(original.logicalHeight);
  await sheet.shortcut('Shift+z');
  expect((await rowGeometry(page, 0)).logicalHeight).toBe(enlarged.logicalHeight);
  await sheet.goTo('A1');
  await chooseFontSize(12);
  expect((await rowGeometry(page, 0)).logicalHeight).toBeLessThan(enlarged.logicalHeight);

  await page.evaluate(() => {
    const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
    const { setRowsHeight } = w.__fcProbes['commands/row-col-layout'];
    const { store, controller } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
    setRowsHeight(store, controller.history, 0, 0, 24, handle);
  });
  await chooseFontSize(36);
  expect((await rowGeometry(page, 0)).logicalHeight).toBe(24);

  await page
    .locator('.fc-host__grid')
    .first()
    .click({ position: { x: 10, y: 10 } });
  await chooseFontSize(18);
  expect((await rowGeometry(page, 100)).logicalHeight).toBeGreaterThan(original.logicalHeight);
  expect((await rowGeometry(page, 0)).logicalHeight).toBe(24);
  expect((await rowGeometry(page, 100)).explicitRows).toBe(1);
  await sheet.shortcut('z');
  await expect
    .poll(async () => {
      const instanceState = await page.evaluate(() => {
        const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
        const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
        const state = store.getState();
        // The resolved font record has no facade member; read it from the engine handle.
        const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
        const effective = handle.getEffectiveStyle(0, 100, 0);
        const xf = effective && handle.getCellXf(effective.xfIndex);
        return {
          nativePoints: xf && handle.getResolvedFontRecord(xf.fontIndex)?.sizePt,
          sheetPatch: state.format.sheetFontOverrides?.get(0)?.patch ?? null,
          axisOverride: state.format.sheetAxisOverrides?.has(0) ?? false,
        };
      });
      return {
        height: (await rowGeometry(page, 100)).logicalHeight,
        ...instanceState,
      };
    })
    .toEqual({
      height: original.logicalHeight,
      nativePoints: platform === 'mac' ? 12 : 11,
      sheetPatch: null,
      axisOverride: false,
    });
  expect((await rowGeometry(page, 0)).logicalHeight).toBe(24);
  await sheet.shortcut('Shift+z');
  const worksheetHeight = (await rowGeometry(page, 100)).logicalHeight;
  await sheet.goTo('A2');
  await page.keyboard.press('Shift+Space');
  await chooseFontSize(11);
  expect((await rowGeometry(page, 1)).logicalHeight).toBe(original.logicalHeight);
  expect((await rowGeometry(page, 100)).logicalHeight).toBe(worksheetHeight);
  expect((await rowGeometry(page, 1)).explicitRows).toBe(1);
  await sheet.shortcut('z');
  await expect.poll(async () => (await rowGeometry(page, 1)).logicalHeight).toBe(worksheetHeight);
  await sheet.shortcut('z');
  await expect
    .poll(async () => (await rowGeometry(page, 100)).logicalHeight)
    .toBe(original.logicalHeight);
  await sheet.expectNoConsoleErrors();
}
