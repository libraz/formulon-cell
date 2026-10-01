import { expect, type Page } from '@playwright/test';

import type { WorkbookHandle } from '../../../packages/formulon-cell/src/engine/workbook-handle.js';
import type { SpreadsheetInstance } from '../../../packages/formulon-cell/src/mount/types.js';
import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

type DemoWindow = Window & { __fcInst?: SpreadsheetInstance };

export async function runExcelDataBarAppearanceScenario(
  page: Page,
  locale: 'en' | 'ja',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('-10');
  await page.keyboard.type('20');
  await page.keyboard.press('Enter');
  await expect
    .poll(() =>
      page.evaluate(() => {
        const wb = (window as DemoWindow).__fcInst?.workbook;
        return [0, 1].map((row) => wb?.getValue({ sheet: 0, row, col: 0 }));
      }),
    )
    .toEqual([
      { kind: 'number', value: -10 },
      { kind: 'number', value: 20 },
    ]);
  await openRule(page, 'data-bar');
  const dialog = page.locator('.fc-conddlg');
  await dialog.locator('.fc-conddlg__form input[type="text"]').first().fill('A1:A2');
  const axis = dialog.locator('[data-cf-bar-axis-position]');
  const negative = dialog.locator('[data-cf-bar-negative-color]');
  const borderStyle = dialog.locator('[data-cf-bar-border-style]');
  const border = dialog.locator('[data-cf-bar-border-color]');
  const negativeBorder = dialog.locator('[data-cf-bar-negative-border-color]');
  const axisColor = dialog.locator('[data-cf-bar-axis-color]');
  await expect(axis).toHaveValue('automatic');
  await expect(negative).toHaveValue('#ff0000');
  await expect(borderStyle).toHaveValue('none');
  await expect(border).toBeDisabled();
  await expect(negativeBorder).toBeDisabled();
  await axis.selectOption('middle');
  await negative.fill('#cc0000');
  await borderStyle.selectOption('solid');
  await border.fill('#003366');
  await negativeBorder.fill('#990000');
  await axisColor.fill('#006600');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  await expect
    .poll(() => readRules(page))
    .toMatchObject([
      {
        kind: 'data-bar',
        axisPosition: 'middle',
        negativeColor: '#cc0000',
        borderColor: '#003366',
        negativeBorderColor: '#990000',
        axisColor: '#006600',
      },
    ]);
  await openRule(page, 'data-bar', true);
  await expect(axis).toHaveValue('middle');
  await expect(negative).toHaveValue('#cc0000');
  await expect(borderStyle).toHaveValue('solid');
  await expect(border).toHaveValue('#003366');
  await expect(negativeBorder).toHaveValue('#990000');
  await expect(axisColor).toHaveValue('#006600');
  await axis.selectOption('none');
  await expect(axisColor).toBeDisabled();
  await dialog
    .getByRole('button', {
      name: locale === 'ja' ? 'ルールを保存' : 'Save rule',
      exact: true,
    })
    .click();
  await expect(dialog).toBeHidden();
  await saveAndReload(page);
  await openRule(page, 'data-bar', true);
  await expect(axis).toHaveValue('none');
  await expect(negative).toHaveValue('#cc0000');
  await expect(borderStyle).toHaveValue('solid');
  await expect(border).toHaveValue('#003366');
  await expect(negativeBorder).toHaveValue('#990000');
  await expect(axisColor).toHaveValue('#006600');
  await expect(axisColor).toBeDisabled();
  for (const width of [1280, 390]) {
    await page.setViewportSize({ width, height: 800 });
    for (const input of [negative, border, negativeBorder, axisColor]) {
      await expect(input).toBeVisible();
      const geometry = await input.evaluate((element) => {
        const bounds = element.closest('.fc-conddlg__body')?.getBoundingClientRect();
        if (!bounds) throw new Error('Conditional formatting body unavailable');
        const rect = element.getBoundingClientRect();
        return {
          fits: rect.left >= bounds.left && rect.right <= bounds.right,
          label: element.getAttribute('aria-label'),
          width: innerWidth,
          left: rect.left,
          right: rect.right,
          bodyLeft: bounds.left,
          bodyRight: bounds.right,
        };
      });
      expect(geometry.fits, JSON.stringify(geometry)).toBe(true);
    }
  }
  await expectOneManagedRule(page);
}

async function openRule(page: Page, kind: 'data-bar' | 'icon-set', edit = false): Promise<void> {
  await page.evaluate(
    ({ kind: ruleKind, edit: isEdit }) => {
      const inst = (window as DemoWindow).__fcInst;
      if (!inst) throw new Error('Spreadsheet instance unavailable');
      inst.openConditionalDialog(
        isEdit ? { mode: 'edit', editIndex: 0 } : { mode: 'new', kind: ruleKind },
      );
    },
    { kind, edit },
  );
  await expect(page.locator('.fc-conddlg')).toBeVisible();
}

async function readRules(page: Page) {
  return page.evaluate(() => {
    const inst = (window as DemoWindow).__fcInst;
    if (!inst) throw new Error('Spreadsheet instance unavailable');
    return inst.store.getState().conditional.rules;
  });
}

async function saveAndReload(page: Page): Promise<void> {
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as DemoWindow).__fcInst?.workbook.getConditionalFormats(0).length,
      ),
    )
    .toBe(1);
  await page.evaluate(async () => {
    const inst = (window as DemoWindow).__fcInst;
    if (!inst) throw new Error('Spreadsheet instance unavailable');
    const ctor = inst.workbook.constructor as typeof WorkbookHandle;
    const next = await ctor.loadBytes(inst.workbook.save());
    if (next.isStub) throw new Error('Reload must use the real WASM engine');
    await inst.setWorkbook(next);
  });
}

async function expectOneManagedRule(page: Page): Promise<void> {
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    const inst = (window as DemoWindow).__fcInst;
    if (!inst) throw new Error('Spreadsheet instance unavailable');
    inst.openCfRulesDialog();
  });
  const manager = page.locator('.fc-cfrulesdlg');
  await expect(manager).toBeVisible();
  await expect(manager.locator('.fc-cfrulesdlg__table tbody tr')).toHaveCount(1);
  await expect(manager.locator('.fc-cfrulesdlg__table tbody tr')).toHaveAttribute(
    'data-rule-source',
    'engine',
  );
}

export async function runExcelDataBarBoundsScenario(
  page: Page,
  locale: 'en' | 'ja',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale });
  await sp.expectNoStub();
  await sp.typeIntoActiveCell('50');
  await page.keyboard.press('ArrowUp');
  await openRule(page, 'data-bar');
  const dialog = page.locator('.fc-conddlg');
  await expect(dialog.locator('.fc-conddlg__style-row select')).toHaveValue('data-bar');
  const minType = dialog.locator('[data-cf-bar-min-type]');
  const maxType = dialog.locator('[data-cf-bar-max-type]');
  const minValue = dialog.locator('[data-cf-bar-min-value]');
  const maxValue = dialog.locator('[data-cf-bar-max-value]');
  const direction = dialog.locator('[data-cf-bar-direction]');
  await expect(minType).toHaveValue('min');
  await expect(maxType).toHaveValue('max');
  await expect(minValue).toBeHidden();
  await expect(maxValue).toBeHidden();
  await minType.selectOption('number');
  await maxType.selectOption('number');
  await minValue.fill('20');
  await maxValue.fill('80');
  await direction.selectOption('right-to-left');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  const expected = {
    kind: 'data-bar',
    min: { kind: 'number', value: 20 },
    max: { kind: 'number', value: 80 },
    direction: 'right-to-left',
  };
  await expect.poll(() => readRules(page)).toMatchObject([expected]);
  await openRule(page, 'data-bar', true);
  await expect(minType).toHaveValue('number');
  await expect(maxType).toHaveValue('number');
  await expect(minValue).toHaveValue('20');
  await expect(maxValue).toHaveValue('80');
  await expect(direction).toHaveValue('right-to-left');
  await minValue.fill('30');
  await direction.selectOption('left-to-right');
  await dialog
    .getByRole('button', { name: locale === 'ja' ? 'ルールを保存' : 'Save rule', exact: true })
    .click();
  await expect(dialog).toBeHidden();
  const edited = { ...expected, min: { kind: 'number', value: 30 }, direction: 'left-to-right' };
  await expect.poll(() => readRules(page)).toMatchObject([edited]);
  await saveAndReload(page);
  await expect.poll(() => readRules(page)).toMatchObject([edited]);
  await openRule(page, 'data-bar', true);
  await expect(minValue).toHaveValue('30');
  await expect(maxValue).toHaveValue('80');
  await expect(direction).toHaveValue('left-to-right');
  await expectOneManagedRule(page);
}

export async function runExcelIconComparisonScenario(
  page: Page,
  locale: 'en' | 'ja',
): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount({ locale });
  await sp.expectNoStub();
  await openRule(page, 'icon-set');
  const dialog = page.locator('.fc-conddlg');
  await expect.soft(dialog.locator('.fc-conddlg__style-row select')).toHaveValue('icon-set');
  for (const width of [1280, 390]) {
    await page.setViewportSize({ width, height: 800 });
    await expect
      .poll(() =>
        dialog.evaluate((element) => {
          const body = element.querySelector('.fc-conddlg__body');
          if (!body) throw new Error('Conditional formatting body unavailable');
          const bounds = body.getBoundingClientRect();
          const visibleInputs = Array.from(element.querySelectorAll('[data-cf-icon-value]')).filter(
            (input) => input.getClientRects().length > 0,
          );
          return {
            visibleCount: visibleInputs.length,
            allWithin: visibleInputs.every((input) => {
              const rect = input.getBoundingClientRect();
              return rect.left >= bounds.left && rect.right <= bounds.right;
            }),
          };
        }),
      )
      .toEqual({ visibleCount: 2, allWithin: true });
  }
  await page.setViewportSize({ width: 1280, height: 800 });
  const firstOperator = dialog.locator('[data-cf-icon-operator="0"]');
  const secondOperator = dialog.locator('[data-cf-icon-operator="1"]');
  await expect(firstOperator).toHaveValue('>=');
  await expect(secondOperator).toHaveValue('>=');
  await firstOperator.selectOption('>');
  await dialog.locator('[data-cf-icon-type="0"]').selectOption('number');
  await dialog.locator('[data-cf-icon-value="0"]').fill('20');
  await dialog.locator('[data-cf-icon-type="1"]').selectOption('number');
  await dialog.locator('[data-cf-icon-value="1"]').fill('80');
  await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  await expect(dialog).toBeHidden();
  const expected = {
    kind: 'icon-set',
    thresholds: [
      { kind: 'number', value: 20, gte: false },
      { kind: 'number', value: 80 },
    ],
  };
  await expect.poll(() => readRules(page)).toMatchObject([expected]);
  await openRule(page, 'icon-set', true);
  await expect(firstOperator).toHaveValue('>');
  await expect(secondOperator).toHaveValue('>=');
  await firstOperator.selectOption('>=');
  await secondOperator.selectOption('>');
  await dialog
    .getByRole('button', { name: locale === 'ja' ? 'ルールを保存' : 'Save rule', exact: true })
    .click();
  await expect(dialog).toBeHidden();
  const edited = {
    kind: 'icon-set',
    thresholds: [
      { kind: 'number', value: 20 },
      { kind: 'number', value: 80, gte: false },
    ],
  };
  await expect.poll(() => readRules(page)).toMatchObject([edited]);
  await saveAndReload(page);
  await expect.poll(() => readRules(page)).toMatchObject([edited]);
  await openRule(page, 'icon-set', true);
  await expect(firstOperator).toHaveValue('>=');
  await expect(secondOperator).toHaveValue('>');
  await expectOneManagedRule(page);
}
