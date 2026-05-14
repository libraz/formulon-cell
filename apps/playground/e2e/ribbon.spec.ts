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
