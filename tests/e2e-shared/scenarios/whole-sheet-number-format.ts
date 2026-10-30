import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type PaintProbe = Window & { __fcPaintedText: string[] };

export async function runWholeSheetNumberFormat(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  await page.addInitScript(() => {
    const painted: string[] = [];
    (window as unknown as PaintProbe).__fcPaintedText = painted;
    const original = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (...args) {
      if (this.canvas.classList.contains('fc-host__canvas')) painted.push(args[0]);
      return Reflect.apply(original, this, args);
    };
  });
  const sheet = new SpreadsheetPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  const moduleUrl = `/@fs${
    new URL('../../../packages/formulon-cell/src/store/store.ts', import.meta.url).pathname
  }`;
  await page.evaluate(async (url) => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const { mutators } = await import(url);
    const inputs = ['1234.5', "'00123", 'abc', '=SUM(A1,1)', '="label"'];
    const result = instance.applyChanges(
      inputs.map((input, row) => ({ addr: { sheet: 0, row, col: 0 }, input })),
    );
    if (result.status !== 'applied') throw new Error('Mixed values were not seeded');
    mutators.setColWidth(instance.store, 0, 160);
    mutators.setCellFormat(
      instance.store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'percent', decimals: 1 },
        bold: true,
      },
    );
  }, moduleUrl);
  await page
    .locator('.fc-host__grid')
    .first()
    .click({ position: { x: 10, y: 10 } });
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState()
            .selection.range,
      ),
    )
    .toEqual({ sheet: 0, r0: 0, c0: 0, r1: 1048575, c1: 16383 });
  await sheet.shortcut('1');
  const dialog = page.getByRole('dialog', { name: 'Format Cells', exact: true });
  await expect(dialog).toBeVisible();
  await dialog.locator('[data-fc-tab="font"][role="tab"]').click();
  await expect(dialog.locator('input[data-fc-check="bold"]')).toHaveJSProperty(
    'indeterminate',
    true,
  );
  await dialog.locator('[data-fc-tab="number"][role="tab"]').click();
  await dialog.locator('[data-fc-cat="fixed"]').click();
  await dialog.locator('.fc-fmtdlg__cat-controls input[type="number"]').fill('2');
  await page.evaluate(() => {
    (window as unknown as PaintProbe).__fcPaintedText.length = 0;
  });
  await dialog.locator('.fc-fmtdlg__btn--primary').click();
  await expect(dialog).toBeHidden();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const drawn = (window as unknown as PaintProbe).__fcPaintedText;
        return ['1234.50', '00123', 'abc', '1235.50', 'label'].filter((text) =>
          drawn.includes(text),
        );
      }),
    )
    .toEqual(['1234.50', '00123', 'abc', '1235.50', 'label']);
  expect(
    await page.evaluate(() => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const wb = instance.workbook;
      return {
        text: wb.getValue({ sheet: 0, row: 1, col: 0 }),
        formula: wb.cellFormula({ sheet: 0, row: 3, col: 0 }),
        farFormat: wb.getEffectiveStyle(0, 1048575, 16383)?.numFmtCode,
        formatCount: instance.store.getState().format.formats.size,
        bold: instance.store.getState().format.formats.get('0:0:0')?.bold,
      };
    }),
  ).toEqual({
    text: { kind: 'text', value: '00123' },
    formula: '=SUM(A1,1)',
    farFormat: '0.00',
    formatCount: 1,
    bold: true,
  });
  await sheet.shortcut('z');
  expect(
    await page.evaluate(() => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      return instance.workbook.getEffectiveStyle(0, 1048575, 16383)?.numFmtCode;
    }),
  ).toBe('General');
}
