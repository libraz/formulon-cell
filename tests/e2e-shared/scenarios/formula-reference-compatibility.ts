import { expect, type Page } from '@playwright/test';
import type { E2eProbes } from '../../../apps/demo-shared/e2e-probes.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/api/spreadsheet/instance.js';
import { gridCellCenter } from '../pages/SpreadsheetPage.js';
import { UserJourneyPage } from '../pages/UserJourneyPage.js';

// coverwise generate, strength=2, seed=42: all 12 parameter pairs covered.
export const formulaDataEdgeCases = [
  { surface: 'inline', platform: 'default', style: 'a1' },
  { surface: 'bar', platform: 'default', style: 'r1c1' },
  { surface: 'bar', platform: 'mac', style: 'a1' },
  { surface: 'inline', platform: 'mac', style: 'r1c1' },
] as const;

export async function runFormulaDataEdgeNavigation(
  page: Page,
  scenario: (typeof formulaDataEdgeCases)[number],
): Promise<void> {
  const { surface, platform, style } = scenario;
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const sheet0 = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      sheet0.setFormula(1, 1, '=9');
      sheet0.setNumber(1, 2, 1);
      sheet0.setNumber(1, 3, 2);
      sheet0.setNumber(1, 4, 3);
      sheet0.setNumber(1, 7, 4);
    });
    instance.history.clear();
  });
  const input = page.locator(
    surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input',
  );
  const begin = async () => {
    await page.evaluate(() => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: false } }));
    });
    await sheet.goTo('B2');
    await page.evaluate((r1c1) => {
      const w = window as unknown as { __fcInst: SpreadsheetInstance; __fcProbes: E2eProbes };
      const { store } = w.__fcProbes['mount/internals'].spreadsheetInternals(w.__fcInst);
      store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1 } }));
    }, style === 'r1c1');
    if (surface === 'inline') await page.keyboard.press('F2');
    else await input.focus();
    await input.fill('=SUM()+7');
    await input.evaluate((el) => (el as HTMLTextAreaElement).setSelectionRange(5, 5));
  };
  const modifier = platform === 'mac' ? 'Meta' : 'Control';
  const undo = platform === 'mac' ? 'Meta+z' : 'Control+z';
  const redo = platform === 'mac' ? 'Meta+Shift+z' : 'Control+y';
  const first = style === 'r1c1' ? 'RC[3]' : 'E2';
  const expanded = style === 'r1c1' ? 'RC[3]:RC[6]' : 'E2:H2';

  await begin();
  await page.keyboard.press('F2');
  await page.keyboard.press(`${modifier}+ArrowRight`);
  await expect(input).toHaveValue(`=SUM(${first})+7`);
  await page.keyboard.press(`${modifier}+Shift+ArrowRight`);
  await expect(input).toHaveValue(`=SUM(${expanded})+7`);
  expect(await sheet.readFormula('B2')).toBe('=9');
  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.selection.get(),
    ),
  ).toMatchObject({ sheet: 0, active: { row: 1, col: 1 } });
  await page.keyboard.press(undo);
  await expect(input).toHaveValue('=SUM()+7');
  await page.keyboard.press(redo);
  await expect(input).toHaveValue(`=SUM(${expanded})+7`);
  await page.keyboard.press('Enter');
  await expect.poll(() => sheet.readFormula('B2')).toBe('=SUM(E2:H2)+7');
  expect(await sheet.readValue('B2')).toEqual({ kind: 'number', value: 14 });

  await begin();
  const picked = await gridCellCenter(page, 1, 7);
  await page.mouse.click(picked.x, picked.y);
  await expect(input).toHaveValue(`=SUM(${style === 'r1c1' ? 'RC[6]' : 'H2'})+7`);
  await page.keyboard.press(`${modifier}+ArrowLeft`);
  await expect(input).toHaveValue(`=SUM(${first})+7`);
  await page.keyboard.press(undo);
  await expect(input).toHaveValue('=SUM()+7');
  await page.keyboard.press(redo);
  await expect(input).toHaveValue(`=SUM(${first})+7`);
  await page.keyboard.press('Escape');
  expect(await sheet.readFormula('B2')).toBe('=SUM(E2:H2)+7');
  await sheet.expectNoConsoleErrors();
}

export async function runStructuredReferenceCompletion(
  page: Page,
  surface: 'inline' | 'bar',
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new UserJourneyPage(page);
  await sheet.mount({ locale: 'en', platform });
  await sheet.expectNoStub();
  await page.evaluate(() => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    const headers = ['A]B', 'A,B', "O'Brien", 'Revenue', 'Region', 'Revenue', '金額', 'Net'];
    const values = [7, 11, 13, 17, 23, 29, 31, 37];
    const sheet0 = instance.workbook.sheet(0);
    instance.workbook.batch(() => {
      headers.forEach((header, col) => {
        sheet0.setText(0, col, header);
        sheet0.setNumber(1, col, values[col] ?? 0);
      });
    });
    instance.history.clear();
    for (const table of [
      { name: 'Sales', range: 'A1:D2', columns: headers.slice(0, 4) },
      { name: 'Data', range: 'E1:E2', columns: ['Region'] },
      { name: 'Sales.Data', range: 'F1:F2', columns: ['Revenue'] },
      { name: '売上', range: 'G1:G2', columns: ['金額'] },
      { name: '\\Ledger', range: 'H1:H2', columns: ['Net'] },
    ]) {
      sheet0.addTable(table);
    }
  });
  const input = page.locator(
    surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input',
  );
  // `stored` is the read-back form: a column-only reference collapses its inner brackets.
  const cases = [
    {
      prefix: 'Sales[A',
      label: 'A]B',
      reference: "Sales[[A']B]]",
      stored: "Sales[A']B]",
      value: 8,
    },
    {
      prefix: "Sales[A']",
      label: 'A]B',
      reference: "Sales[[A']B]]",
      stored: "Sales[A']B]",
      value: 8,
    },
    { prefix: 'Sales[A', label: 'A,B', reference: 'Sales[[A,B]]', stored: 'Sales[A,B]', value: 12 },
    {
      prefix: "Sales[O''",
      label: "O'Brien",
      reference: "Sales[[O''Brien]]",
      stored: "Sales[O''Brien]",
      value: 14,
    },
    { prefix: 'Sales.Data[Re', label: 'Revenue', reference: 'Sales.Data[Revenue]', value: 30 },
    { prefix: '売上[金', label: '金額', reference: '売上[金額]', value: 32 },
    { prefix: '\\Ledger[Ne', label: 'Net', reference: '\\Ledger[Net]', value: 38 },
  ];
  for (const entry of cases) {
    await sheet.goTo('J2');
    if (surface === 'inline') await page.keyboard.press('F2');
    else await input.focus();
    const draft = `=SUM(${entry.prefix})+1`;
    await input.fill(draft);
    await input.evaluate(
      (el, caret) => (el as HTMLTextAreaElement).setSelectionRange(caret, caret),
      5 + entry.prefix.length,
    );
    const option = page
      .locator('.fc-autocomplete__item')
      .filter({ has: page.getByText(entry.label, { exact: true }) });
    await expect(option).toHaveCount(1);
    await option.click();
    await page.keyboard.press('Tab');
    await expect(input).toHaveValue(`=SUM(${entry.reference})+1`);
    await page.keyboard.press(platform === 'mac' ? 'Meta+z' : 'Control+z');
    await expect(input).toHaveValue(draft);
    await page.keyboard.press(platform === 'mac' ? 'Meta+Shift+z' : 'Control+y');
    await expect(input).toHaveValue(`=SUM(${entry.reference})+1`);
    await page.keyboard.press('Enter');
    await expect
      .poll(() => sheet.readFormula('J2'))
      .toBe(`=SUM(${entry.stored ?? entry.reference})+1`);
    expect(await sheet.readValue('J2')).toEqual({ kind: 'number', value: entry.value });
  }
  await sheet.expectNoConsoleErrors();
}
