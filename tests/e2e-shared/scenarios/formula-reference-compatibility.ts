import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
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
    instance.applyChanges([
      { addr: { sheet: 0, row: 1, col: 1 }, input: '=9' },
      { addr: { sheet: 0, row: 1, col: 2 }, input: '1' },
      { addr: { sheet: 0, row: 1, col: 3 }, input: '2' },
      { addr: { sheet: 0, row: 1, col: 4 }, input: '3' },
      { addr: { sheet: 0, row: 1, col: 7 }, input: '4' },
    ]);
  });
  const input = page.locator(
    surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input',
  );
  const begin = async () => {
    await page.evaluate(() => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: false } }));
    });
    await sheet.goTo('B2');
    await page.evaluate((r1c1) => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1 } }));
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
    await page.evaluate(
      () =>
        (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.store.getState().selection
          .active,
    ),
  ).toEqual({ sheet: 0, row: 1, col: 1 });
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
    instance.applyChanges(
      headers.flatMap((header, col) => [
        { addr: { sheet: 0, row: 0, col }, input: header },
        { addr: { sheet: 0, row: 1, col }, input: String(values[col]) },
      ]),
    );
    for (const table of [
      { name: 'Sales', ref: 'A1:D2', columns: headers.slice(0, 4) },
      { name: 'Data', ref: 'E1:E2', columns: ['Region'] },
      { name: 'Sales.Data', ref: 'F1:F2', columns: ['Revenue'] },
      { name: '売上', ref: 'G1:G2', columns: ['金額'] },
      { name: '\\Ledger', ref: 'H1:H2', columns: ['Net'] },
    ]) {
      if (instance.workbook.createTable({ ...table, sheetIndex: 0 }) < 0)
        throw new Error(`Table fixture rejected: ${table.name}`);
    }
  });
  const input = page.locator(
    surface === 'inline' ? '.fc-host__editor' : '.fc-host__formulabar-input',
  );
  const cases = [
    { prefix: 'Sales[A', label: 'A]B', reference: "Sales[[A']B]]", value: 8 },
    { prefix: "Sales[A']", label: 'A]B', reference: "Sales[[A']B]]", value: 8 },
    { prefix: 'Sales[A', label: 'A,B', reference: 'Sales[[A,B]]', value: 12 },
    { prefix: "Sales[O''", label: "O'Brien", reference: "Sales[[O''Brien]]", value: 14 },
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
    await expect.poll(() => sheet.readFormula('J2')).toBe(`=SUM(${entry.reference})+1`);
    expect(await sheet.readValue('J2')).toEqual({ kind: 'number', value: entry.value });
  }
  await sheet.expectNoConsoleErrors();
}
