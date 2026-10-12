import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
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
  const seeded = await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    return instance.applyChanges([
      { addr: { sheet: 0, row: 0, col: 0 }, input: 'outside' },
      { addr: { sheet: 0, row: 0, col: 1 }, input: 'selected' },
      { addr: { sheet: 0, row: 1, col: 1 }, input: 'hole' },
    ]).status;
  });
  expect(seeded).toBe('applied');
  const native = (addresses: [number, number][]) =>
    page.evaluate((addresses) => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      return addresses.map(([row, col]) => {
        const style = instance.workbook.getEffectiveStyle(0, row, col);
        const xf = style && instance.workbook.getCellXf(style.xfIndex);
        const font = xf && instance.workbook.getFontRecord(xf.fontIndex);
        return {
          size: font?.size,
          family: font?.name,
          bold: font?.bold === true,
          align: xf?.horizontalAlign,
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
  const undo = async () => {
    expect(
      await page.evaluate(() =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
      ),
    ).toBe(true);
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
    [18, 18, 11, 11],
  );
  await undo();
  await expectSizes(
    [
      [0, 1],
      [1048575, 1],
    ],
    [11, 11],
  );
  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.redo(),
    ),
  ).toBe(true);
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
    [18, 11, 11],
  );
  await undo();
  await expectSizes([[2, 16383]], [11]);

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
    await expectSizes(points, [20, 20, 20, 11, 11]);
    await undo();
    await expectSizes(points, [11, 11, 11, 11, 11]);
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
      await expectSizes(bands, [18, 18, 11]);
      await undo();
      await expectSizes(bands, [11, 11, 11]);
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
      await expectSizes(perforated, [18, 18, 11, 11]);
      await undo();
      await expectSizes(perforated, [11, 11, 11, 11]);
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
      [24, 24, 11],
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
      [24, 24, 24, 11],
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
      [24, 24, 11],
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
    [24, 24, 24, 11],
  );
  await undo();

  for (const axis of ['row', 'col'] as const) {
    await header(axis, 1);
    await page.locator('[data-ribbon-command="formatCellsHome"]').first().click();
    await page.locator('#menu-format-cells [data-cell-format="dialog"]').click();
    const dialog = page.locator('[class="fc-fmtdlg"]').first();
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
    await expectSizes(points, [16, 11]);
    await undo();
    await expectSizes(points, [11, 11]);
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
  const moduleUrl = `/@fs${new URL('../../../packages/formulon-cell/src/engine/workbook-handle.ts', import.meta.url).pathname}`;
  const saved = await page.evaluate(async (moduleUrl) => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const { WorkbookHandle } = await import(moduleUrl);
    const workbook = await WorkbookHandle.loadBytes(instance.workbook.save());
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
        return { size: font?.size, family: font?.name, bold: font?.bold === true };
      });
    } finally {
      workbook.dispose();
    }
  }, moduleUrl);
  expect(saved.slice(0, 2)).toEqual([
    { size: 11, family: 'Arial', bold: true },
    { size: 11, family: 'Arial', bold: true },
  ]);
  expect(saved[2]).toMatchObject({ size: 11, bold: false });
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().format
          .formats.size,
    ),
  ).toBeLessThan(100);
  expect(errors.read()).toEqual([]);
}
