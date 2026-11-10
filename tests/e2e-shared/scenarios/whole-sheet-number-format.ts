import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type PaintedRun = { text: string; y: number; end: number };
type PaintProbe = Window & { __fcPaintedText: PaintedRun[] };

export async function runWholeSheetNumberFormat(
  page: Page,
  platform: 'default' | 'mac',
): Promise<void> {
  await page.addInitScript(() => {
    const painted: PaintedRun[] = [];
    (window as unknown as PaintProbe).__fcPaintedText = painted;
    const original = CanvasRenderingContext2D.prototype.fillText;
    CanvasRenderingContext2D.prototype.fillText = function (...args) {
      if (this.canvas.classList.contains('fc-host__canvas')) {
        const [text, x, y] = args;
        const width = this.measureText(text).width;
        const start =
          this.textAlign === 'right' ? x - width : this.textAlign === 'center' ? x - width / 2 : x;
        // Mac lays text out one glyph per call; join calls that continue the previous run.
        const last = painted[painted.length - 1];
        if (last && last.y === y && Math.abs(start - last.end) <= 1.5) {
          last.text += text;
          last.end = start + width;
        } else {
          painted.push({ text, y, end: start + width });
        }
      }
      return Reflect.apply(original, this, args);
    };
  });
  const sheet = new SpreadsheetPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const probes = (window as unknown as { __fcProbes: E2eProbes }).__fcProbes;
    const { mutators } = probes['store/store'];
    const { store } = probes['mount/internals'].spreadsheetInternals(instance);
    const sheet0 = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      sheet0.setNumber(0, 0, 1234.5);
      sheet0.setText(1, 0, '00123');
      sheet0.setText(2, 0, 'abc');
      sheet0.setFormula(3, 0, '=SUM(A1,1)');
      sheet0.setFormula(4, 0, '="label"');
    });
    instance.history.clear();
    mutators.setColWidth(store, 0, 160);
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        numFmt: { kind: 'percent', decimals: 1 },
        bold: true,
      },
    );
  });
  await page
    .locator('.fc-host__grid')
    .first()
    .click({ position: { x: 10, y: 10 } });
  await expect
    .poll(() =>
      page.evaluate(() =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.selection.get(),
      ),
    )
    .toMatchObject({
      sheet: 0,
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 1048575, lastCol: 16383 }],
    });
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
        const drawn = (window as unknown as PaintProbe).__fcPaintedText.map((run) => run.text);
        return ['1234.50', '00123', 'abc', '1235.50', 'label'].filter((text) =>
          drawn.includes(text),
        );
      }),
    )
    .toEqual(['1234.50', '00123', 'abc', '1235.50', 'label']);
  expect(
    await page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const instance = w.__fcInst;
      const sheet0 = instance.workbook.sheet(0);
      const probes = w.__fcProbes;
      const handle = probes['api/workbook'].workbookHandleOf(instance.workbook);
      const { store } = probes['mount/internals'].spreadsheetInternals(instance);
      return {
        text: sheet0.getValue(1, 0),
        formula: sheet0.getFormula(3, 0),
        farFormat: handle.getEffectiveStyle(0, 1048575, 16383)?.numberFormat,
        formatCount: store.getState().format.formats.size,
        bold: store.getState().format.formats.get('0:0:0')?.bold,
      };
    }),
  ).toEqual({
    text: '00123',
    formula: '=SUM(A1,1)',
    farFormat: '0.00',
    formatCount: 1,
    bold: true,
  });
  await sheet.shortcut('z');
  expect(
    await page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      return w.__fcProbes['api/workbook']
        .workbookHandleOf(w.__fcInst.workbook)
        .getEffectiveStyle(0, 1048575, 16383)?.numberFormat;
    }),
  ).toBe('General');
}
