import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

export async function runAxisFormatStructure(
  page: Page,
  locale: string,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale, platform });
  await sheet.expectNoStub();
  // New workbooks default to 11 pt on the default profile and 12 pt on Mac.
  const baseSize = platform === 'mac' ? 12 : 11;
  const header = async (axis: 'row' | 'col', index: number) => {
    const box = await page.locator('.fc-host__grid').first().boundingBox();
    if (!box) throw new Error('Missing grid');
    const cell = await gridCellCenter(page, axis === 'row' ? index : 0, axis === 'col' ? index : 0);
    await page.mouse.click(
      axis === 'row' ? box.x + 10 : cell.x,
      axis === 'row' ? cell.y : box.y + 10,
    );
  };
  const chooseSize = async (size: number) => {
    await page.locator('[data-ribbon-command="fontSize"] button').first().click();
    await page.locator(`.fc-tb__rb-dd__opt[data-value="${size}"]`).click();
  };
  const expectSizes = async (axis: 'row' | 'col', indices: number[], sizes: number[]) => {
    await expect
      .poll(() =>
        page.evaluate(
          ({ axis, indices }) => {
            const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
            const wb = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
            return indices.map((index) => {
              const style = wb.getEffectiveStyle(
                0,
                axis === 'row' ? index : 1048575,
                axis === 'col' ? index : 16383,
              );
              const xf = style && wb.getCellXf(style.xfIndex);
              return xf && wb.getFontRecord(xf.fontIndex)?.sizePt;
            });
          },
          { axis, indices },
        ),
      )
      .toEqual(sizes);
  };
  const history = async (action: 'undo' | 'redo') => {
    expect(
      await page.evaluate((action) => {
        const { history } = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
        const available = action === 'undo' ? history.canUndo : history.canRedo;
        history[action]();
        return available;
      }, action),
    ).toBe(true);
  };
  for (const axis of ['row', 'col'] as const) {
    for (const operation of ['insert', 'delete'] as const) {
      await header(axis, 3);
      await chooseSize(18);
      await expectSizes(axis, [2, 3, 4], [baseSize, 18, baseSize]);
      await header(axis, 2);
      const kind = axis === 'row' ? 'rows' : 'cols';
      if (operation === 'insert') await sheet.insertRowsOrColumns(kind);
      else await sheet.deleteRowsOrColumns(kind);
      const shifted = operation === 'insert' ? 4 : 2;
      await expectSizes(axis, [shifted, 3], [18, baseSize]);
      await header(axis === 'row' ? 'col' : 'row', 6);
      await chooseSize(24);
      await expectSizes(axis, [shifted, 3], [18, baseSize]);
      await history('undo');
      await history('undo');
      await expectSizes(axis, [2, 3, 4], [baseSize, 18, baseSize]);
      await history('redo');
      await expectSizes(axis, [shifted, 3], [18, baseSize]);
      await history('undo');
      await history('undo');
      await expectSizes(axis, [2, 3, 4], [baseSize, baseSize, baseSize]);
    }
  }
  await sheet.expectNoConsoleErrors();
}
