import { expect, type Page } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runPrintPdfSmokeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await page.evaluate(() => {
    const win = window as unknown as { __fcInst?: SpreadsheetInstance; __fcProbes: E2eProbes };
    const inst = win.__fcInst;
    if (!inst) return;
    inst.workbook.batch(() => {
      const sheet = inst.workbook.sheet(0);
      sheet.setText(0, 0, 'PDF Title Row');
      sheet.setText(1, 0, 'PDF Title Col');
      sheet.setText(1, 1, 'PDF Body 1');
      sheet.setText(1, 2, 'PDF Body 2');
      sheet.setText(2, 1, 'PDF Body 3');
    });
    inst.history.clear();
    // Layout and page setup are UI state with no public write route.
    const { store } = win.__fcProbes['mount/internals'].spreadsheetInternals(inst);
    (
      store as unknown as {
        setState(updater: (state: Record<string, unknown>) => Record<string, unknown>): void;
      }
    ).setState((state) => {
      const layout = state.layout as {
        colWidths: Map<number, number>;
        rowHeights: Map<number, number>;
        defaultColWidth: number;
        defaultRowHeight: number;
      };
      const pageSetup = state.pageSetup as { setupBySheet: Map<number, Record<string, unknown>> };
      const colWidths = new Map(layout.colWidths);
      colWidths.set(0, 100);
      colWidths.set(1, 360);
      colWidths.set(2, 360);
      const rowHeights = new Map(layout.rowHeights);
      rowHeights.set(0, 100);
      rowHeights.set(1, 360);
      rowHeights.set(2, 360);
      const setupBySheet = new Map(pageSetup.setupBySheet);
      setupBySheet.set(0, {
        ...(setupBySheet.get(0) ?? {}),
        printArea: 'B2:C3',
        printTitleRows: '1:1',
        printTitleCols: 'A:A',
        margins: { top: 1, right: 1, bottom: 1, left: 1 },
        fitWidth: 0,
        fitHeight: 0,
      });
      return {
        ...state,
        layout: { ...layout, colWidths, rowHeights },
        pageSetup: { ...pageSetup, setupBySheet },
      };
    });
  });

  await page.locator('[data-ribbon-tab="file"]').first().click();
  const backstage = page.locator('.fc-tb__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await backstage.getByRole('button', { name: 'Print', exact: true }).first().click();
  await expect(backstage.locator('[data-demo-print-preview]')).toBeVisible();
  await expect(backstage.frameLocator('.fc-tb__print-frame').locator('body')).toContainText(
    'PDF Body 1',
  );

  const srcdoc = await backstage.locator('.fc-tb__print-frame').getAttribute('srcdoc');
  expect(srcdoc).toContain('PDF Title Row');
  expect(srcdoc).toContain('PDF Title Col');
  expect(srcdoc).toContain('PDF Body 1');
  expect(srcdoc).toContain('fc-print__area--break');
  expect(srcdoc).toContain('fc-print__title-col-cell');
  expect(srcdoc?.match(/PDF Title Row/g)?.length ?? 0).toBeGreaterThan(1);
  expect(srcdoc?.match(/PDF Title Col/g)?.length ?? 0).toBeGreaterThan(1);

  const printPage = await page.context().newPage();
  try {
    await printPage.setContent(srcdoc ?? '', { waitUntil: 'load' });
    await printPage.emulateMedia({ media: 'print' });
    const pdf = await printPage.pdf({ format: 'A4', printBackground: true });
    expect(pdf.byteLength).toBeGreaterThan(1_000);
    expect(pdf.subarray(0, 5).toString()).toBe('%PDF-');
  } finally {
    await printPage.close();
  }
}
