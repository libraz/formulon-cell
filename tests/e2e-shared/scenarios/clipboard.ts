import type { Page } from '@playwright/test';
import { expect } from '@playwright/test';

import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';
import { expectSystemClipboardText } from '../pages/system-clipboard.js';

type InstanceWindow = Window & { __fcInst?: SpreadsheetInstance; __fcProbes?: E2eProbes };

/** Reads a cell of the first sheet as the engine handle's tagged value. */
const readCellValue = (page: Page, row: number, col: number) =>
  page.evaluate(
    ([r, c]: [number, number]) => {
      const w = window as InstanceWindow;
      if (!w.__fcInst || !w.__fcProbes) return undefined;
      // The tagged cell value has no facade equivalent; read it from the engine handle.
      return w.__fcProbes['api/workbook']
        .workbookHandleOf(w.__fcInst.workbook)
        .getValue({ sheet: 0, row: r, col: c });
    },
    [row, col] as [number, number],
  );

const seedCellValue = (page: Page, row: number, col: number, value: string | number) =>
  page.evaluate(
    ([r, c, v]) => {
      const inst = (window as InstanceWindow).__fcInst;
      if (!inst) return;
      const sheet = inst.workbook.sheet(0);
      inst.workbook.batch(() => {
        if (typeof v === 'number') sheet.setNumber(r as number, c as number, v);
        else sheet.setText(r as number, c as number, v as string);
      });
      inst.history.clear();
    },
    [row, col, value],
  );

const selectCell = (page: Page, ref: string) =>
  page.evaluate((target) => (window as InstanceWindow).__fcInst?.selection.set(target), ref);

/** C01 — internal copy/paste round-trip via Mod+C/V keyboard shortcuts.
 *  Canvas content isn't queryable, so we re-select each cell and read the
 *  formula bar to confirm the value landed. */
export async function runCopyPasteScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  // Seed A1=alpha; cursor advances to A2 on Enter.
  await sp.typeIntoActiveCell('alpha');
  // Step back to A1, copy.
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('c');
  await expectSystemClipboardText(page, 'alpha');

  // Navigate to A3 and paste. Active cell after paste is the paste anchor.
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await sp.shortcut('v');

  // Mod+V reads navigator.clipboard asynchronously, so poll the formula bar.
  await expect.poll(() => sp.formulaBarValue()).toBe('alpha');
}

/** C02 — Mod+X cut → paste removes the source.
 *  Source becomes empty after paste; destination holds the value. */
export async function runCutPasteScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await sp.typeIntoActiveCell('beta');
  await page.keyboard.press('ArrowUp');
  await sp.shortcut('x');
  await expectSystemClipboardText(page, 'beta');

  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await sp.shortcut('v');

  // Destination has the value (paste is async — see C01).
  await expect.poll(() => sp.formulaBarValue()).toBe('beta');

  // ...and the source (A1) is empty.
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowUp');
  expect(await sp.formulaBarValue()).toBe('');
}

/** C04 — multi-cell Mod+V paste must undo as one transaction. */
export async function runMultiCellPasteUndoScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const readCell = (row: number, col: number) => readCellValue(page, row, col);
  const seedCell = (row: number, col: number, value: string | number) =>
    seedCellValue(page, row, col, value);

  await seedCell(0, 0, 'old-a');
  await seedCell(0, 1, 10);
  await seedCell(1, 0, 'old-b');
  await seedCell(1, 1, 20);
  await selectCell(page, 'A1');

  await expect
    .poll(() => readCell(0, 0))
    .toEqual({
      kind: 'text',
      value: 'old-a',
    });
  await expect
    .poll(() => readCell(1, 1))
    .toEqual({
      kind: 'number',
      value: 20,
    });

  await page.evaluate(() => navigator.clipboard.writeText('foo\t42\r\nbar\t99'));
  await sp.focusHost();
  await sp.shortcut('v');

  await expect
    .poll(() => readCell(0, 0))
    .toEqual({
      kind: 'text',
      value: 'foo',
    });
  await expect
    .poll(() => readCell(1, 1))
    .toEqual({
      kind: 'number',
      value: 99,
    });

  await sp.shortcut('z');

  await expect
    .poll(() => readCell(0, 0))
    .toEqual({
      kind: 'text',
      value: 'old-a',
    });
  await expect
    .poll(() => readCell(0, 1))
    .toEqual({
      kind: 'number',
      value: 10,
    });
  await expect
    .poll(() => readCell(1, 0))
    .toEqual({
      kind: 'text',
      value: 'old-b',
    });
  await expect
    .poll(() => readCell(1, 1))
    .toEqual({
      kind: 'number',
      value: 20,
    });

  await sp.shortcut('y');

  await expect
    .poll(() => readCell(0, 0))
    .toEqual({
      kind: 'text',
      value: 'foo',
    });
  await expect
    .poll(() => readCell(0, 1))
    .toEqual({
      kind: 'number',
      value: 42,
    });
  await expect
    .poll(() => readCell(1, 0))
    .toEqual({
      kind: 'text',
      value: 'bar',
    });
  await expect
    .poll(() => readCell(1, 1))
    .toEqual({
      kind: 'number',
      value: 99,
    });
}

/** C05 — ribbon Paste must use the same transaction semantics as Mod+V. */
export async function runRibbonPasteUndoScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const readCell = (row: number, col: number) => readCellValue(page, row, col);
  const seedCell = (row: number, col: number, value: string | number) =>
    seedCellValue(page, row, col, value);

  await seedCell(0, 0, 'old-a');
  await seedCell(0, 1, 10);
  await seedCell(1, 0, 'old-b');
  await seedCell(1, 1, 20);
  await selectCell(page, 'A1');

  await page.evaluate(() => navigator.clipboard.writeText('foo\t42\r\nbar\t99'));
  await page.locator('[data-ribbon-command="paste"]').click();

  await expect
    .poll(() => readCell(0, 0))
    .toEqual({
      kind: 'text',
      value: 'foo',
    });
  await expect
    .poll(() => readCell(1, 1))
    .toEqual({
      kind: 'number',
      value: 99,
    });

  await sp.shortcut('z');

  await expect
    .poll(() => readCell(0, 0))
    .toEqual({
      kind: 'text',
      value: 'old-a',
    });
  await expect
    .poll(() => readCell(0, 1))
    .toEqual({
      kind: 'number',
      value: 10,
    });
  await expect
    .poll(() => readCell(1, 0))
    .toEqual({
      kind: 'text',
      value: 'old-b',
    });
  await expect
    .poll(() => readCell(1, 1))
    .toEqual({
      kind: 'number',
      value: 20,
    });

  await sp.shortcut('y');

  await expect
    .poll(() => readCell(0, 0))
    .toEqual({
      kind: 'text',
      value: 'foo',
    });
  await expect
    .poll(() => readCell(0, 1))
    .toEqual({
      kind: 'number',
      value: 42,
    });
  await expect
    .poll(() => readCell(1, 0))
    .toEqual({
      kind: 'text',
      value: 'bar',
    });
  await expect
    .poll(() => readCell(1, 1))
    .toEqual({
      kind: 'number',
      value: 99,
    });
}

/** A copied font family with an ampersand and a quote reads back unchanged through the
 *  browser's CSS parser, which decodes the escapes the HTML clipboard writes. */
export async function runHtmlFontFamilyEscapeScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();
  const families = await page.evaluate(() => {
    const w = window as InstanceWindow;
    if (!w.__fcInst || !w.__fcProbes) return null;
    const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
    w.__fcProbes['store/store'].mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      { fontFamily: 'R&D "Mono"' },
    );
    const html = w.__fcProbes['commands/clipboard/html'].encodeHtml(store.getState(), {
      sheet: 0,
      r0: 0,
      c0: 0,
      r1: 0,
      c1: 0,
    });
    const td = new DOMParser().parseFromString(html, 'text/html').querySelector('td');
    const reference = document.createElement('td');
    reference.style.fontFamily = '"R&D \\"Mono\\""';
    return {
      escaped: html.includes('\\26 '),
      read: td?.style.fontFamily,
      reference: reference.style.fontFamily,
    };
  });
  expect(families?.escaped).toBe(true);
  expect(families?.read).toBe(families?.reference);
  expect(families?.read).toContain('R&D');
}
