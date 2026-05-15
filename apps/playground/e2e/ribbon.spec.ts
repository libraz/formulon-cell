import { expect, type Page, test } from '@playwright/test';

const ribbonTabs = [
  {
    id: 'file',
    label: 'File',
    commands: ['pageSetup', 'print', 'links', 'formatCells', 'gotoSpecial'],
  },
  {
    id: 'home',
    label: 'Home',
    commands: [
      'paste',
      'cut',
      'copy',
      'formatPainter',
      'clearFormat',
      'fontFamily',
      'fontSize',
      'fontGrow',
      'fontShrink',
      'font-row-2',
      'bold',
      'italic',
      'underline',
      'strike',
      'borders',
      'borderPreset',
      'borderStyle',
      'fontColor',
      'fillColor',
      'top',
      'middle',
      'alignment-row-2',
      'alignL',
      'alignC',
      'alignR',
      'wrap',
      'merge',
      'general',
      'number-row-2',
      'currency',
      'percent',
      'comma',
      'decDown',
      'decUp',
      'conditional',
      'cellStyles',
      'rules',
      'insertRows',
      'deleteRows',
      'insertCols',
      'deleteCols',
      'formatCellsHome',
      'autosum',
      'undoHome',
      'redoHome',
      'sortAscHome',
      'filterHome',
      'findHome',
      'gotoSpecialHome',
    ],
  },
  {
    id: 'insert',
    label: 'Insert',
    commands: [
      'pivotTableInsert',
      'formatTableInsert',
      'namedRangesInsert',
      'removeDupesInsert',
      'chartInsert',
      'hyperlinkInsert',
      'linksInsert',
      'commentInsert',
      'fxInsert',
    ],
  },
  { id: 'draw', label: 'Draw', commands: ['drawPen', 'drawErase'] },
  {
    id: 'pageLayout',
    label: 'Page Layout',
    commands: [
      'marginsPreset',
      'orientationPreset',
      'paperSizePreset',
      'pageSetupAdvanced',
      'printPageLayout',
    ],
  },
  {
    id: 'formulas',
    label: 'Formulas',
    commands: [
      'fx',
      'autosumFormula',
      'sum',
      'avg',
      'namedRanges',
      'precedents',
      'dependents',
      'clearArrows',
      'recalcNow',
      'calcOptions',
      'watch',
    ],
  },
  {
    id: 'data',
    label: 'Data',
    commands: ['filter', 'sortAsc', 'sortDesc', 'removeDupes', 'linksData', 'hideRows', 'hideCols'],
  },
  {
    id: 'review',
    label: 'Review',
    commands: [
      'spellingReview',
      'translateReview',
      'newCommentReview',
      'findReview',
      'protectReview',
      'accessibility',
    ],
  },
  {
    id: 'view',
    label: 'View',
    commands: ['watchView', 'freeze', 'zoom75', 'zoom100', 'zoom125', 'protect'],
  },
  { id: 'automate', label: 'Automate', commands: ['script'] },
  { id: 'acrobat', label: 'Acrobat', commands: ['addIn', 'pdf'] },
] as const;

async function mount(page: Page): Promise<void> {
  await page.goto('/');
  await page.waitForSelector('.fc-host', { state: 'attached', timeout: 30_000 });
  await page.waitForFunction(
    () => {
      const host = document.querySelector('.fc-host') as HTMLElement | null;
      const state = host?.dataset.fcEngineState;
      return state === 'ready' || state === 'ready-stub';
    },
    { timeout: 30_000 },
  );
}

async function closeDialog(page: Page): Promise<void> {
  await page
    .getByRole('button', { name: /^(Cancel|Close)$/ })
    .last()
    .click();
}

test('R01: ribbon tabs switch visible panels and render expected commands', async ({ page }) => {
  await mount(page);

  for (const tab of ribbonTabs) {
    await page.getByRole('tab', { name: tab.label, exact: true }).click();

    const visiblePanel = page.locator('.demo__ribbon:not([hidden])');
    await expect(visiblePanel).toHaveCount(1);
    await expect(visiblePanel).toHaveAttribute('data-ribbon-panel', tab.id);
    await expect
      .poll(() =>
        visiblePanel
          .locator('[data-ribbon-command]')
          .evaluateAll((nodes) => nodes.map((node) => node.getAttribute('data-ribbon-command'))),
      )
      .toEqual([...tab.commands]);
  }

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="formatCellsHome"]')).toHaveAttribute(
    'aria-keyshortcuts',
    /Control\+1/,
  );
  await expect(page.locator('[data-ribbon-command="findHome"]')).toHaveAttribute(
    'aria-keyshortcuts',
    /Control\+F/,
  );

  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="namedRanges"]')).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+F3',
  );
  await expect(page.locator('[data-ribbon-command="fx"]')).toHaveAttribute(
    'aria-keyshortcuts',
    'Shift+F3',
  );
});

test('R02: Home font controls render and apply formatting', async ({ page }) => {
  await mount(page);

  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    'Aptos',
  );
  await expect(page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__value')).toHaveText(
    '11',
  );
  await expect(page.locator('[data-ribbon-select="borderPreset"]')).toBeVisible();
  await expect(page.locator('[data-ribbon-select="borderStyle"]')).toBeVisible();
  await expect(page.locator('[data-ribbon-command="fontColor"] input[type="color"]')).toHaveValue(
    '#201f1e',
  );
  await expect(page.locator('[data-ribbon-command="fillColor"] input[type="color"]')).toHaveValue(
    '#ffffff',
  );

  await page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__btn').click();
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__list')).toBeVisible();
  await page.locator('[data-ribbon-select="fontFamily"] [data-value="Arial"]').click();
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    'Arial',
  );

  await page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__btn').click();
  await page.locator('[data-ribbon-select="fontSize"] [data-value="14"]').click();
  await expect(page.locator('[data-ribbon-select="fontSize"] .demo__rb-dd__value')).toHaveText(
    '14',
  );

  const activeFormat = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
              format: { formats: Map<string, { fontFamily?: string; fontSize?: number }> };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    const active = state?.selection.active;
    return active ? state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`) : null;
  });

  expect(activeFormat).toMatchObject({ fontFamily: 'Arial', fontSize: 14 });
});

test('R03: routed ribbon commands open dialogs and mutate workbook state', async ({ page }) => {
  await mount(page);

  await page.getByRole('tab', { name: 'File', exact: true }).click();
  await page.locator('[data-ribbon-command="pageSetup"]').click();
  await expect(page.locator('.fc-pgsetup')).toBeVisible();
  await closeDialog(page);

  await page.locator('[data-ribbon-command="links"]').click();
  await expect(page.locator('.fc-extlinkdlg')).toBeVisible();
  await closeDialog(page);

  await page.locator('[data-ribbon-command="gotoSpecial"]').click();
  await expect(page.locator('.fc-goto')).toBeVisible();
  await closeDialog(page);

  await page.getByRole('tab', { name: 'Home', exact: true }).click();
  await page.locator('[data-ribbon-command="rules"]').click();
  await expect(page.locator('.fc-cfrulesdlg')).toBeVisible();
  await closeDialog(page);

  await page.getByRole('tab', { name: 'Draw', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="drawPen"]')).toBeEnabled();
  await expect(page.locator('[data-ribbon-command="drawErase"]')).toBeEnabled();
  await page.locator('[data-ribbon-command="drawPen"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | {
              store: {
                getState: () => {
                  selection: { active: { sheet: number; row: number; col: number } };
                  format: { formats: Map<string, { borders?: Record<string, unknown> }> };
                };
              };
            }
          | undefined;
        const state = inst?.store.getState();
        const active = state?.selection.active;
        return active
          ? (state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`)?.borders ??
              null)
          : null;
      }),
    )
    .toMatchObject({ top: { style: 'thin' }, right: { style: 'thin' } });
  await page.locator('[data-ribbon-command="drawErase"]').click();
  await expect
    .poll(() =>
      page.evaluate(() => {
        const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
          | {
              store: {
                getState: () => {
                  selection: { active: { sheet: number; row: number; col: number } };
                  format: { formats: Map<string, { borders?: Record<string, unknown> }> };
                };
              };
            }
          | undefined;
        const state = inst?.store.getState();
        const active = state?.selection.active;
        return active
          ? (state?.format.formats.get(`${active.sheet}:${active.row}:${active.col}`)?.borders ??
              null)
          : null;
      }),
    )
    .toMatchObject({ top: false, right: false, bottom: false, left: false });

  await page.getByRole('tab', { name: 'Insert', exact: true }).click();
  await page.locator('[data-ribbon-command="formatTableInsert"]').click();
  await page.locator('[data-ribbon-command="chartInsert"]').click();
  await expect(page.locator('.fc-chart')).toBeVisible();

  const objectCounts = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              tables: { tables: unknown[] };
              charts: { charts: unknown[] };
            };
          };
        }
      | undefined;
    const state = inst?.store.getState();
    return {
      tables: state?.tables.tables.length ?? 0,
      charts: state?.charts.charts.length ?? 0,
    };
  });
  expect(objectCounts.tables).toBeGreaterThan(0);
  expect(objectCounts.charts).toBeGreaterThan(0);

  await page.getByRole('tab', { name: 'Formulas', exact: true }).click();
  await page.locator('[data-ribbon-command="watch"]').click();
  await expect(page.locator('.fc-host__watchdock')).toBeVisible();

  await page.locator('[data-ribbon-command="fx"]').click();
  await expect(page.locator('.fc-fxdialog')).toBeVisible();
  await closeDialog(page);

  await page.getByRole('tab', { name: 'Review', exact: true }).click();
  for (const command of ['spellingReview', 'translateReview', 'accessibility']) {
    await expect(page.locator(`[data-ribbon-command="${command}"]`)).toBeEnabled();
    await page.locator(`[data-ribbon-command="${command}"]`).click();
    await expect(page.getByRole('dialog')).toBeVisible();
    await closeDialog(page);
  }

  await page.getByRole('tab', { name: 'Automate', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="script"]')).toBeEnabled();
  await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            setText: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          };
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
            };
            setState: (
              fn: (state: {
                selection: {
                  active: { sheet: number; row: number; col: number };
                  anchor: { sheet: number; row: number; col: number };
                  range: { sheet: number; r0: number; c0: number; r1: number; c1: number };
                };
              }) => unknown,
            ) => void;
          };
        }
      | undefined;
    const active = inst?.store.getState().selection.active;
    if (inst && active) {
      inst.workbook.setText(active, 'script value');
      inst.workbook.setText(
        { sheet: active.sheet, row: active.row, col: active.col + 1 },
        'second value',
      );
      inst.store.setState((state) => ({
        ...state,
        selection: {
          ...state.selection,
          active,
          anchor: active,
          range: {
            sheet: active.sheet,
            r0: active.row,
            c0: active.col,
            r1: active.row,
            c1: active.col + 1,
          },
        },
      }));
    }
  });
  await page.locator('[data-ribbon-command="script"]').click();
  await expect(page.getByRole('dialog')).toBeVisible();
  await page.getByRole('textbox', { name: 'Command' }).fill('uppercase');
  await page.getByRole('button', { name: 'Run' }).click();
  const scriptResult = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          workbook: {
            getValue: (addr: { sheet: number; row: number; col: number }) => {
              kind: string;
              value?: string;
            };
          };
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
            };
          };
        }
      | undefined;
    const active = inst?.store.getState().selection.active;
    return inst && active
      ? [
          inst.workbook.getValue(active),
          inst.workbook.getValue({ sheet: active.sheet, row: active.row, col: active.col + 1 }),
        ]
      : null;
  });
  expect(scriptResult).toMatchObject([
    { kind: 'text', value: 'SCRIPT VALUE' },
    { kind: 'text', value: 'SECOND VALUE' },
  ]);
  const undoResult = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          undo: () => boolean;
          workbook: {
            getValue: (addr: { sheet: number; row: number; col: number }) => {
              kind: string;
              value?: string;
            };
          };
          store: {
            getState: () => {
              selection: { active: { sheet: number; row: number; col: number } };
            };
          };
        }
      | undefined;
    const active = inst?.store.getState().selection.active;
    const ok = inst?.undo() ?? false;
    return inst && active
      ? {
          ok,
          values: [
            inst.workbook.getValue(active),
            inst.workbook.getValue({ sheet: active.sheet, row: active.row, col: active.col + 1 }),
          ],
        }
      : null;
  });
  expect(undoResult).toMatchObject({
    ok: true,
    values: [
      { kind: 'text', value: 'script value' },
      { kind: 'text', value: 'second value' },
    ],
  });

  await page.getByRole('tab', { name: 'Acrobat', exact: true }).click();
  await expect(page.locator('[data-ribbon-command="addIn"]')).toBeEnabled();
  await page.locator('[data-ribbon-command="addIn"]').click();
  await expect(page.getByRole('dialog')).toBeVisible();
  await closeDialog(page);

  await page.getByRole('tab', { name: 'View', exact: true }).click();
  await page.locator('[data-ribbon-command="zoom125"]').click();
  const zoom = await page.evaluate(() => {
    const inst = (window as Window & { __fcInst?: unknown }).__fcInst as
      | {
          store: {
            getState: () => {
              viewport: { zoom: number };
            };
          };
        }
      | undefined;
    return inst?.store.getState().viewport.zoom;
  });
  expect(zoom).toBe(1.25);
});

test('R04: ribbon tabs support Excel-style arrow, Home, and End keyboard navigation', async ({
  page,
}) => {
  await mount(page);

  const home = page.getByRole('tab', { name: 'Home', exact: true });
  const insert = page.getByRole('tab', { name: 'Insert', exact: true });
  const acrobat = page.getByRole('tab', { name: 'Acrobat', exact: true });
  const file = page.getByRole('tab', { name: 'File', exact: true });

  await home.focus();
  await expect(home).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(insert).toBeFocused();
  await expect(insert).toHaveAttribute('aria-selected', 'true');
  await expect(page.locator('.demo__ribbon:not([hidden])')).toHaveAttribute(
    'data-ribbon-panel',
    'insert',
  );

  await page.keyboard.press('End');
  await expect(acrobat).toBeFocused();
  await expect(acrobat).toHaveAttribute('aria-selected', 'true');

  await page.keyboard.press('Home');
  await expect(file).toBeFocused();
  await expect(file).toHaveAttribute('aria-selected', 'true');

  await page.keyboard.press('ArrowLeft');
  await expect(acrobat).toBeFocused();
  await expect(acrobat).toHaveAttribute('tabindex', '0');
  await expect(home).toHaveAttribute('tabindex', '-1');
});

test('R05: ribbon dropdowns support keyboard selection and Escape dismissal', async ({ page }) => {
  await mount(page);

  const fontButton = page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__btn');
  await fontButton.focus();
  await page.keyboard.press('ArrowDown');

  const list = page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__list');
  await expect(list).toBeVisible();
  await expect(fontButton).toHaveAttribute('aria-expanded', 'true');
  await expect(
    page.locator('[data-ribbon-select="fontFamily"] [data-value="Aptos"]'),
  ).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(
    page.locator('[data-ribbon-select="fontFamily"] [data-value="Calibri"]'),
  ).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(list).toBeHidden();
  await expect(fontButton).toBeFocused();
  await expect(page.locator('[data-ribbon-select="fontFamily"] .demo__rb-dd__value')).toHaveText(
    'Calibri',
  );

  await page.keyboard.press('ArrowDown');
  await expect(list).toBeVisible();
  await page.keyboard.press('End');
  await expect(
    page.locator('[data-ribbon-select="fontFamily"] [data-value="Consolas"]'),
  ).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(list).toBeHidden();
  await expect(fontButton).toBeFocused();
  await expect(fontButton).toHaveAttribute('aria-expanded', 'false');
});

test('R06: ribbon split menus support menu keyboard navigation and focus return', async ({
  page,
}) => {
  await mount(page);

  const borders = page.locator('#btn-borders');
  const borderMenu = page.locator('#menu-borders');
  await borders.click();
  await expect(borderMenu).toBeVisible();
  await expect(borderMenu.getByRole('menuitem', { name: 'All borders' })).toBeFocused();
  await page.keyboard.press('End');
  await expect(borderMenu.getByRole('menuitem', { name: 'More borders…' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(borderMenu).toBeHidden();
  await expect(borders).toBeFocused();

  await page.getByRole('tab', { name: 'Data', exact: true }).click();
  const sort = page.locator('#btn-sort');
  const sortMenu = page.locator('#menu-sort');
  await sort.click();
  await expect(sortMenu).toBeVisible();
  await page.keyboard.press('ArrowDown');
  await expect(sortMenu.getByRole('menuitem', { name: 'Sort Z → A' })).toBeFocused();
  await page.keyboard.press('Home');
  await expect(sortMenu.getByRole('menuitem', { name: 'Sort A → Z' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(sortMenu).toBeHidden();
  await expect(sort).toBeFocused();
});
