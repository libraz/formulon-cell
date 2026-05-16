import { expect, type Page, test } from '@playwright/test';

import {
  runCopyPasteScenario,
  runCutPasteScenario,
  runMultiCellPasteUndoScenario,
  runRibbonPasteUndoScenario,
} from '../../../tests/e2e-shared/scenarios/clipboard.js';
import { SpreadsheetPage } from '../../../tests/e2e-shared/pages/SpreadsheetPage.js';

test('C01 (playground): Mod+C/V round-trips a cell value', async ({ page }) => {
  await runCopyPasteScenario(page);
});

test('C02 (playground): Mod+X clears the source after paste', async ({ page }) => {
  await runCutPasteScenario(page);
});

test('C04 (playground): multi-cell Mod+V paste undoes as one transaction', async ({ page }) => {
  await runMultiCellPasteUndoScenario(page);
});

test('C05 (playground): ribbon Paste undoes a multi-cell paste as one transaction', async ({
  page,
}) => {
  await runRibbonPasteUndoScenario(page);
});

type CellAddr = { sheet: number; row: number; col: number };

async function selectRange(
  page: Page,
  range: { r0: number; c0: number; r1: number; c1: number },
): Promise<void> {
  await page.evaluate((selection) => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            setState: (
              updater: (state: {
                selection: {
                  active: CellAddr;
                  anchor: CellAddr;
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                  extraRanges?: unknown[];
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    const active = { sheet: 0, row: selection.r0, col: selection.c0 };
    inst?.store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        active,
        anchor: active,
        range: { sheet: 0, ...selection },
        extraRanges: [],
      },
    }));
  }, range);
}

async function seedRibbonPasteSource(page: Page): Promise<void> {
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            setNumber: (addr: CellAddr, value: number) => void;
            setText: (addr: CellAddr, value: string) => void;
            setFormula: (addr: CellAddr, formula: string) => void;
            recalc: () => void;
          };
          store: {
            setState: (
              updater: (state: {
                format: { formats: Map<string, Record<string, unknown>> };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    if (!inst) return;
    inst.workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
    inst.workbook.setFormula({ sheet: 0, row: 0, col: 1 }, '=A1+1');
    inst.workbook.setText({ sheet: 0, row: 1, col: 0 }, 'x');
    inst.workbook.setNumber({ sheet: 0, row: 1, col: 1 }, 4);
    inst.workbook.setText({ sheet: 0, row: 0, col: 9 }, 'keep');
    inst.workbook.recalc();
    inst.store.setState((state) => {
      const formats = new Map(state.format.formats);
      for (const row of [0, 1]) {
        for (const col of [0, 1]) {
          formats.set(`0:${row}:${col}`, {
            bold: true,
            fill: '#ffff00',
            numFmt: '#,##0.00',
          });
        }
      }
      return { ...state, format: { formats } };
    });
  });
}

async function readCell(
  page: Page,
  row: number,
  col: number,
): Promise<{ kind?: string; value?: string | number; formula?: string | null }> {
  return page.evaluate(
    ({ row, col }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            workbook: {
              getValue: (addr: CellAddr) => { kind: string; value?: string | number };
              cellFormula: (addr: CellAddr) => string | null;
            };
          }
        | undefined;
      const addr = { sheet: 0, row, col };
      const value = inst?.workbook.getValue(addr);
      return {
        kind: value?.kind,
        value: value?.value,
        formula: inst?.workbook.cellFormula(addr) ?? null,
      };
    },
    { row, col },
  );
}

async function readFormat(
  page: Page,
  row: number,
  col: number,
): Promise<Record<string, unknown> | undefined> {
  return page.evaluate(
    ({ row, col }) => {
      const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
        | {
            store: {
              getState: () => { format: { formats: Map<string, Record<string, unknown>> } };
            };
          }
        | undefined;
      return inst?.store.getState().format.formats.get(`0:${row}:${col}`);
    },
    { row, col },
  );
}

async function clickPasteMenuAction(page: Page, action: string): Promise<void> {
  await page.locator('[data-ribbon-command="paste"] .demo__rb-split-chevron').click();
  await page.locator(`#menu-paste [data-paste-action="${action}"]`).click();
}

test('C06 (playground): ribbon Paste menu applies values, formulas, formats, and transpose', async ({
  page,
}) => {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await seedRibbonPasteSource(page);
  await selectRange(page, { r0: 0, c0: 0, r1: 1, c1: 1 });
  await page.locator('[data-ribbon-command="copy"]').click();

  await selectRange(page, { r0: 0, c0: 3, r1: 1, c1: 4 });
  await clickPasteMenuAction(page, 'values');
  await expect.poll(() => readCell(page, 0, 3)).toMatchObject({ kind: 'number', value: 1 });
  await expect
    .poll(() => readCell(page, 0, 4))
    .toMatchObject({ kind: 'number', value: 2, formula: null });
  expect((await readFormat(page, 0, 3))?.fill).toBeUndefined();
  await sp.shortcut('z');
  await expect.poll(() => readCell(page, 0, 3)).toMatchObject({ kind: 'blank' });
  await sp.shortcut('y');
  await expect.poll(() => readCell(page, 0, 4)).toMatchObject({ kind: 'number', value: 2 });

  await selectRange(page, { r0: 2, c0: 3, r1: 3, c1: 4 });
  await clickPasteMenuAction(page, 'values-and-numfmt');
  await expect
    .poll(() => readCell(page, 2, 4))
    .toMatchObject({ kind: 'number', value: 2, formula: null });
  await expect.poll(() => readFormat(page, 2, 4)).toMatchObject({ numFmt: '#,##0.00' });
  expect((await readFormat(page, 2, 4))?.fill).toBeUndefined();
  expect((await readFormat(page, 2, 4))?.bold).toBeUndefined();
  await sp.shortcut('z');
  await expect.poll(() => readCell(page, 2, 4)).toMatchObject({ kind: 'blank' });
  expect((await readFormat(page, 2, 4))?.numFmt).toBeUndefined();
  await sp.shortcut('y');
  await expect.poll(() => readFormat(page, 2, 4)).toMatchObject({ numFmt: '#,##0.00' });

  await selectRange(page, { r0: 0, c0: 6, r1: 1, c1: 7 });
  await clickPasteMenuAction(page, 'formulas');
  await expect.poll(() => readCell(page, 0, 6)).toMatchObject({ kind: 'blank' });
  await expect
    .poll(() => readCell(page, 0, 7))
    .toMatchObject({ kind: 'number', value: 1, formula: '=G1+1' });
  await sp.shortcut('z');
  await expect.poll(() => readCell(page, 0, 7)).toMatchObject({ kind: 'blank' });
  await sp.shortcut('y');
  await expect.poll(() => readCell(page, 0, 7)).toMatchObject({ formula: '=G1+1' });

  await selectRange(page, { r0: 2, c0: 6, r1: 3, c1: 7 });
  await clickPasteMenuAction(page, 'formulas-and-numfmt');
  await expect.poll(() => readCell(page, 2, 6)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readFormat(page, 2, 6)).toMatchObject({ numFmt: '#,##0.00' });
  expect((await readFormat(page, 2, 6))?.fill).toBeUndefined();
  await expect
    .poll(() => readCell(page, 2, 7))
    .toMatchObject({ kind: 'number', value: 1, formula: '=G3+1' });
  await expect.poll(() => readFormat(page, 2, 7)).toMatchObject({ numFmt: '#,##0.00' });
  await sp.shortcut('z');
  await expect.poll(() => readCell(page, 2, 7)).toMatchObject({ kind: 'blank' });
  expect((await readFormat(page, 2, 7))?.numFmt).toBeUndefined();
  await sp.shortcut('y');
  await expect.poll(() => readCell(page, 2, 7)).toMatchObject({ formula: '=G3+1' });

  await selectRange(page, { r0: 0, c0: 9, r1: 1, c1: 10 });
  await clickPasteMenuAction(page, 'formats');
  await expect.poll(() => readCell(page, 0, 9)).toMatchObject({ kind: 'text', value: 'keep' });
  await expect.poll(() => readFormat(page, 0, 9)).toMatchObject({
    bold: true,
    fill: '#ffff00',
    numFmt: '#,##0.00',
  });
  await sp.shortcut('z');
  expect((await readFormat(page, 0, 9))?.fill).toBeUndefined();
  await sp.shortcut('y');
  await expect.poll(() => readFormat(page, 0, 9)).toMatchObject({ fill: '#ffff00' });

  await selectRange(page, { r0: 4, c0: 0, r1: 5, c1: 1 });
  await clickPasteMenuAction(page, 'transpose');
  await expect.poll(() => readCell(page, 4, 0)).toMatchObject({ kind: 'number', value: 1 });
  await expect.poll(() => readCell(page, 5, 0)).toMatchObject({ kind: 'number', value: 2 });
  await expect.poll(() => readCell(page, 4, 1)).toMatchObject({ kind: 'text', value: 'x' });
  await expect.poll(() => readCell(page, 5, 1)).toMatchObject({ kind: 'number', value: 4 });
  await sp.shortcut('z');
  await expect.poll(() => readCell(page, 4, 0)).toMatchObject({ kind: 'blank' });
  await expect.poll(() => readCell(page, 5, 1)).toMatchObject({ kind: 'blank' });
});

test('C07 (playground): ribbon Paste Special dialog applies skip blanks and transpose', async ({
  page,
}) => {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            setText: (addr: CellAddr, value: string) => void;
            cells: (sheet: number) => Iterable<{
              addr: CellAddr;
              value: unknown;
              formula: string | null;
            }>;
            recalc: () => void;
          };
          store: {
            setState: (
              updater: (state: { data: { cells: Map<string, unknown> } }) => unknown,
            ) => void;
          };
        }
      | undefined;
    if (!inst) return;
    inst.workbook.setText({ sheet: 0, row: 0, col: 1 }, 'new-bottom');
    inst.workbook.setText({ sheet: 0, row: 0, col: 4 }, 'old-top');
    inst.workbook.setText({ sheet: 0, row: 1, col: 4 }, 'old-bottom');
    inst.workbook.recalc();
    const cells = new Map<string, { value: unknown; formula: string | null }>();
    for (const cell of inst.workbook.cells(0)) {
      cells.set(`${cell.addr.sheet}:${cell.addr.row}:${cell.addr.col}`, {
        value: cell.value,
        formula: cell.formula,
      });
    }
    inst.store.setState((state) => ({ ...state, data: { ...state.data, cells } }));
  });

  await selectRange(page, { r0: 0, c0: 0, r1: 0, c1: 1 });
  await page.locator('[data-ribbon-command="copy"]').click();

  await selectRange(page, { r0: 0, c0: 4, r1: 0, c1: 4 });
  await page.locator('[data-ribbon-command="paste"] .demo__rb-split-chevron').click();
  await page.locator('#menu-paste [data-paste-action="dialog"]').click();

  const dialog = page.getByRole('dialog', { name: /Paste Special|形式を選択して貼り付け/ });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('radio', { name: /^(Values|値)$/ }).check();
  await dialog.getByRole('checkbox', { name: /Skip blanks|空白セルを無視する/ }).check();
  await dialog.getByRole('checkbox', { name: /Transpose|行\/列の入れ替え/ }).check();
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();

  await expect.poll(() => readCell(page, 0, 4)).toMatchObject({
    kind: 'text',
    value: 'old-top',
  });
  await expect.poll(() => readCell(page, 1, 4)).toMatchObject({
    kind: 'text',
    value: 'new-bottom',
  });

  await sp.shortcut('z');
  await expect.poll(() => readCell(page, 0, 4)).toMatchObject({
    kind: 'text',
    value: 'old-top',
  });
  await expect.poll(() => readCell(page, 1, 4)).toMatchObject({
    kind: 'text',
    value: 'old-bottom',
  });
});
