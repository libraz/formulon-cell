import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runSelectionFormat(
  page: Page,
  locale: string,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new SpreadsheetPage(page);
  const errors = sheet.collectConsoleErrors();
  await sheet.mount({ locale, platform });
  await sheet.expectNoStub();
  // New workbooks default to 11 pt on the default profile and 12 pt on Mac.
  const baseSize = platform === 'mac' ? 12 : 11;
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const sheet0 = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      sheet0.setText('A1', 'outside');
      sheet0.setText('B1', 'selected');
      sheet0.setText('B2', 'hole');
    });
    instance.history.clear();
  });
  const native = (addresses: [number, number][]) =>
    page.evaluate((addresses) => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const handle = w.__fcProbes['api/workbook'].workbookHandleOf(w.__fcInst.workbook);
      return addresses.map(([row, col]) => {
        const style = handle.getEffectiveStyle(0, row, col);
        const xf = style && handle.getCellXf(style.xfIndex);
        const font = xf && handle.getFontRecord(xf.fontIndex);
        return {
          size: font?.sizePt,
          family: font?.name,
          bold: font?.bold === true,
          align: xf?.horizontal,
        };
      });
    }, addresses);
  const expectSizes = async (addresses: [number, number][], sizes: number[]) => {
    await expect
      .poll(async () => (await native(addresses)).map((format) => format.size))
      .toEqual(sizes);
  };
  const chooseSize = async (size: number) => {
    await page.locator('[data-ribbon-command="fontSize"] button').first().click();
    await page.locator(`.fc-tb__rb-dd__opt[data-value="${size}"]`).click();
    await expect(page.locator('[data-ribbon-select="fontSize"] .fc-tb__rb-dd__value')).toHaveText(
      String(size),
    );
  };
  const historyStep = (action: 'undo' | 'redo') =>
    page.evaluate((action) => {
      const { history } = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const available = action === 'undo' ? history.canUndo : history.canRedo;
      history[action]();
      return available;
    }, action);
  const undo = async () => {
    expect(await historyStep('undo')).toBe(true);
  };
  const header = async (
    axis: 'row' | 'col',
    index: number,
    modifier?: 'Shift' | 'Control' | 'Meta',
  ) => {
    const box = await page.locator('.fc-host__grid').first().boundingBox();
    if (!box) throw new Error('Missing grid');
    const cell = await gridCellCenter(page, axis === 'row' ? index : 0, axis === 'col' ? index : 0);
    if (modifier) await page.keyboard.down(modifier);
    try {
      await page.mouse.click(
        axis === 'row' ? box.x + 10 : cell.x,
        axis === 'row' ? cell.y : box.y + 10,
      );
    } finally {
      if (modifier) await page.keyboard.up(modifier);
    }
  };
  const clickCell = async (row: number, col: number, modifier?: 'Shift' | 'Control' | 'Meta') => {
    const point = await gridCellCenter(page, row, col);
    if (modifier) await page.keyboard.down(modifier);
    try {
      await page.mouse.click(point.x, point.y);
    } finally {
      if (modifier) await page.keyboard.up(modifier);
    }
  };

  await header('col', 1);
  await chooseSize(18);
  await expectSizes(
    [
      [0, 1],
      [1048575, 1],
      [0, 0],
      [1048575, 2],
    ],
    [18, 18, baseSize, baseSize],
  );
  await undo();
  await expectSizes(
    [
      [0, 1],
      [1048575, 1],
    ],
    [baseSize, baseSize],
  );
  expect(await historyStep('redo')).toBe(true);
  await expectSizes([[1048575, 1]], [18]);
  await undo();

  await header('row', 2);
  await chooseSize(18);
  await expectSizes(
    [
      [2, 16383],
      [1, 16383],
      [3, 16383],
    ],
    [18, baseSize, baseSize],
  );
  await undo();
  await expectSizes([[2, 16383]], [baseSize]);

  for (const axis of ['row', 'col'] as const) {
    await header(axis, 1);
    await header(axis, axis === 'row' ? 8 : 3, 'Shift');
    await chooseSize(20);
    const points: [number, number][] =
      axis === 'row'
        ? [
            [1, 16383],
            [4, 16383],
            [8, 16383],
            [0, 16383],
            [9, 16383],
          ]
        : [
            [1048575, 1],
            [1048575, 2],
            [1048575, 3],
            [1048575, 0],
            [1048575, 4],
          ];
    await expectSizes(points, [20, 20, 20, baseSize, baseSize]);
    await undo();
    await expectSizes(points, [baseSize, baseSize, baseSize, baseSize, baseSize]);
    for (const modifier of ['Control', 'Meta'] as const) {
      await header(axis, 1);
      await header(axis, 3, modifier);
      await chooseSize(18);
      const bands: [number, number][] =
        axis === 'row'
          ? [
              [1, 16383],
              [3, 16383],
              [2, 16383],
            ]
          : [
              [1048575, 1],
              [1048575, 3],
              [1048575, 2],
            ];
      await expectSizes(bands, [18, 18, baseSize]);
      await undo();
      await expectSizes(bands, [baseSize, baseSize, baseSize]);
      await header(axis, 1);
      if (axis === 'row') await header(axis, 8, 'Shift');
      await clickCell(1, 1, modifier);
      await chooseSize(18);
      const perforated: [number, number][] =
        axis === 'row'
          ? [
              [1, 16383],
              [8, 16383],
              [1, 1],
              [0, 16383],
            ]
          : [
              [0, 1],
              [1048575, 1],
              [1, 1],
              [1048575, 2],
            ];
      await expectSizes(perforated, [18, 18, baseSize, baseSize]);
      await undo();
      await expectSizes(perforated, [baseSize, baseSize, baseSize, baseSize]);
    }
  }

  for (const modifier of ['Control', 'Meta'] as const) {
    await clickCell(0, 0);
    await clickCell(2, 2, modifier);
    await chooseSize(24);
    await expectSizes(
      [
        [0, 0],
        [2, 2],
        [1, 1],
      ],
      [24, 24, baseSize],
    );
    await undo();
    await page.locator('.fc-host__grid').first().focus();
    await page.keyboard.press('Shift+ArrowRight');
    await chooseSize(24);
    await expectSizes(
      [
        [0, 0],
        [2, 2],
        [2, 3],
        [1, 1],
      ],
      [24, 24, 24, baseSize],
    );
    await undo();
    await clickCell(0, 0);
    await clickCell(2, 2, 'Shift');
    await clickCell(1, 1, modifier);
    await chooseSize(24);
    await expectSizes(
      [
        [0, 0],
        [2, 2],
        [1, 1],
      ],
      [24, 24, baseSize],
    );
    await undo();
  }
  await clickCell(0, 0);
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await chooseSize(24);
  await expectSizes(
    [
      [0, 0],
      [2, 2],
      [1, 1],
      [3, 3],
    ],
    [24, 24, 24, baseSize],
  );
  await undo();

  for (const axis of ['row', 'col'] as const) {
    await header(axis, 1);
    await page.locator('[data-ribbon-command="formatCellsHome"]').first().click();
    await page.locator('#menu-format-cells [data-cell-format="dialog"]').click();
    const dialog = page
      .locator('.fc-fmtdlg:not([hidden])')
      .filter({ has: page.locator('button[data-fc-tab="font"]') })
      .first();
    await expect(dialog).toBeVisible();
    await dialog.locator('button[data-fc-tab="font"]').click();
    await dialog.locator('.fc-fmtdlg__font-size-row input').fill('16');
    await dialog.locator('.fc-fmtdlg__btn--primary').click();
    await expect(dialog).toBeHidden();
    const points: [number, number][] =
      axis === 'row'
        ? [
            [1, 16383],
            [2, 16383],
          ]
        : [
            [1048575, 1],
            [1048575, 2],
          ];
    await expectSizes(points, [16, baseSize]);
    await undo();
    await expectSizes(points, [baseSize, baseSize]);
  }

  await header('col', 1);
  await page.locator('[data-ribbon-command="fontFamily"] button').first().click();
  await page.locator('.fc-tb__rb-dd__opt[data-value="Arial"]').click();
  await expect
    .poll(async () =>
      (
        await native([
          [0, 1],
          [1048575, 1],
        ])
      ).map((format) => format.family),
    )
    .toEqual(['Arial', 'Arial']);
  await page.locator('[data-ribbon-command="bold"]').first().click();
  await expect
    .poll(async () =>
      (
        await native([
          [0, 1],
          [1048575, 1],
          [1048575, 0],
        ])
      ).map((format) => format.bold),
    )
    .toEqual([true, true, false]);
  const saved = await page.evaluate(async () => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const { WorkbookHandle } = (window as unknown as { __fcProbes: E2eProbes }).__fcProbes[
      'engine/workbook-handle'
    ];
    const workbook = await WorkbookHandle.loadBytes(instance.workbook.save().bytes);
    try {
      const addresses: [number, number][] = [
        [0, 1],
        [1048575, 1],
        [1048575, 0],
      ];
      return addresses.map(([row, col]) => {
        const effective = workbook.getEffectiveStyle(0, row, col);
        const xf = effective && workbook.getCellXf(effective.xfIndex);
        const font = xf && workbook.getFontRecord(xf.fontIndex);
        return { size: font?.sizePt, family: font?.name, bold: font?.bold === true };
      });
    } finally {
      workbook.dispose();
    }
  });
  expect(saved.slice(0, 2)).toEqual([
    { size: baseSize, family: 'Arial', bold: true },
    { size: baseSize, family: 'Arial', bold: true },
  ]);
  expect(saved[2]).toMatchObject({ size: baseSize, bold: false });
  expect(
    await page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      return w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst).store.getState()
        .format.formats.size;
    }),
  ).toBeLessThan(100);
  expect(errors.read()).toEqual([]);
}
