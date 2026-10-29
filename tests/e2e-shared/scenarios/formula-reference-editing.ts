import { expect, type Page } from '@playwright/test';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { gridCellCenter, SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export type FormulaEditEntry = 'f2' | 'double-click' | 'formula-bar';

export async function runFormulaReferenceEditing(
  page: Page,
  entry: FormulaEditEntry,
  platform: 'default' | 'mac',
): Promise<void> {
  const sheet = new SpreadsheetPage(page);
  const errors = sheet.collectConsoleErrors();
  await sheet.mount({ platform });
  await sheet.expectNoStub();
  const original = '=SUM(B1:B2)+D1';
  await page.evaluate((formula) => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.applyChanges([
      { addr: { sheet: 0, row: 0, col: 0 }, input: formula },
      { addr: { sheet: 0, row: 0, col: 1 }, input: '2' },
      { addr: { sheet: 0, row: 1, col: 1 }, input: '3' },
      { addr: { sheet: 0, row: 0, col: 3 }, input: '7' },
      { addr: { sheet: 0, row: 3, col: 2 }, input: '11' },
      { addr: { sheet: 0, row: 4, col: 3 }, input: '13' },
    ]);
  }, original);
  const clickCell = async (row: number, col: number) => {
    const point = await gridCellCenter(page, row, col);
    await page.mouse.click(point.x, point.y);
  };
  const readSource = () =>
    page.evaluate(() => {
      const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
      const addr = { sheet: 0, row: 0, col: 0 };
      return {
        formula: instance.workbook.cellFormula(addr),
        value: instance.workbook.getValue(addr),
      };
    });
  const input = page
    .locator(entry === 'formula-bar' ? '.fc-host__formulabar-input' : '.fc-host__editor')
    .first();
  const begin = async () => {
    await clickCell(0, 0);
    if (entry === 'f2') await page.keyboard.press('F2');
    else if (entry === 'double-click') {
      const point = await gridCellCenter(page, 0, 0);
      await page.mouse.dblclick(point.x, point.y);
    } else await input.focus();
    await expect(input).toBeFocused();
  };
  const select = async (start: number, end: number) => {
    await input.evaluate(
      (element, { start, end }) => {
        (element as HTMLTextAreaElement).setSelectionRange(start, end);
      },
      { start, end },
    );
  };

  await begin();
  await expect(input).toHaveValue(original);
  await select(5, 10); // Only B1:B2, retaining SUM(...) and +D1.
  await clickCell(3, 2);
  await expect(input).toHaveValue('=SUM(C4)+D1');
  expect((await readSource()).formula).toBe(original);
  const from = await gridCellCenter(page, 3, 2);
  const to = await gridCellCenter(page, 4, 3);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await page.mouse.up();
  await expect(input).toHaveValue('=SUM(C4:D5)+D1');
  await page.keyboard.press('Enter');
  await expect
    .poll(readSource)
    .toEqual({ formula: '=SUM(C4:D5)+D1', value: { kind: 'number', value: 31 } });

  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
    ),
  ).toBe(true);
  expect((await readSource()).formula).toBe(original);

  await begin();
  await select(8, 10); // Replace just B2, preserving the B1: range prefix throughout the drag.
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await expect(input).toHaveValue('=SUM(B1:C4)+D1');
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await expect(input).toHaveValue('=SUM(B1:C4:D5)+D1');
  await page.mouse.up();
  await page.keyboard.press('Enter');
  await expect.poll(readSource).toEqual({
    formula: '=SUM(B1:C4:D5)+D1',
    value: { kind: 'number', value: 43 },
  });
  expect(
    await page.evaluate(() =>
      (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst.undo(),
    ),
  ).toBe(true);

  await begin();
  await select(8, 10);
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  const reverse = await gridCellCenter(page, 4, 0);
  await page.mouse.move(reverse.x, reverse.y, { steps: 5 });
  await expect(input).toHaveValue('=SUM(B1:A4:C5)+D1');
  await page.mouse.up();
  await page.keyboard.press('Escape');
  expect((await readSource()).formula).toBe(original);

  await begin();
  await select(8, 8); // A caret inside B2 replaces the complete reference range.
  await clickCell(3, 2);
  await expect(input).toHaveValue('=SUM(C4)+D1');
  const reference = await gridCellCenter(page, 4, 3);
  await page.mouse.dblclick(reference.x, reference.y);
  await expect(input).toHaveValue('=SUM(D5)+D1');
  await expect(input).toBeFocused();
  expect((await readSource()).formula).toBe(original);
  await page.keyboard.press('Escape');

  if (entry !== 'formula-bar') {
    await begin();
    await select(0, original.length);
    await page.keyboard.press('F9');
    await expect(input).toHaveValue('=12');
    expect((await readSource()).formula).toBe(original);
    await page.keyboard.press('Escape');
  }

  await begin();
  await select(0, original.length);
  await clickCell(3, 2);
  await expect(input).toHaveValue('=C4');
  await page.keyboard.press('Escape');
  expect((await readSource()).formula).toBe(original);

  await begin();
  await select(0, original.length);
  await clickCell(3, 2);
  await page.keyboard.press('Enter');
  await expect.poll(readSource).toEqual({ formula: '=C4', value: { kind: 'number', value: 11 } });
  // The same pointer route must replace relative R1C1 references during a drag.
  await page.evaluate((formula) => {
    const instance = (window as unknown as { __fcInst: SpreadsheetInstance }).__fcInst;
    instance.applyChanges([{ addr: { sheet: 0, row: 0, col: 0 }, input: formula }]);
    instance.store.setState((state) => ({ ...state, ui: { ...state.ui, r1c1: true } }));
  }, original);
  await begin();
  const r1c1 = await input.inputValue();
  expect(r1c1).toBe('=SUM(RC[1]:R[1]C[1])+RC[3]');
  await select(r1c1.indexOf('(') + 1, r1c1.indexOf(')'));
  await clickCell(3, 2);
  await expect(input).toHaveValue('=SUM(R[3]C[2])+RC[3]');
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await page.mouse.up();
  await expect(input).toHaveValue('=SUM(R[3]C[2]:R[4]C[3])+RC[3]');
  await page.keyboard.press('Enter');
  await expect.poll(readSource).toEqual({
    formula: '=SUM(C4:D5)+D1',
    value: { kind: 'number', value: 31 },
  });
  expect(errors.read()).toEqual([]);
}
