import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

async function rowGeometry(page: Page, row: number) {
  return page.evaluate(
    ({ row }) => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const { layoutForView } = (window as unknown as { __fcProbes: E2eProbes }).__fcProbes[
        'render/geometry'
      ];
      const state = instance.store.getState();
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
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().selection
          .active.row,
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
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const { setRowsHeight } = (window as unknown as { __fcProbes: E2eProbes }).__fcProbes[
      'commands/row-col-layout'
    ];
    setRowsHeight(instance.store, instance.history, 0, 0, 24, instance.workbook);
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
        const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
        const state = instance.store.getState();
        const effective = instance.workbook.getEffectiveStyle(0, 100, 0);
        const xf = effective && instance.workbook.getCellXf(effective.xfIndex);
        return {
          nativePoints: xf && instance.workbook.getResolvedFontRecord(xf.fontIndex)?.sizePt,
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
