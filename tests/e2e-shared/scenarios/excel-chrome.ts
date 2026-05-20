import { expect, type Page } from '@playwright/test';

import { SpreadsheetPage } from '../pages/SpreadsheetPage.js';

export async function runExcelChromeBackstageSearchScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await sp.mount();
  await sp.expectNoStub();

  const fileTab = page.locator('[data-ribbon-tab="file"]').first();
  await expect(fileTab).toBeVisible();
  await fileTab.click();

  const backstage = page.locator('.demo__backstage[role="dialog"]').first();
  await expect(backstage).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Info', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'New', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Open', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Save', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Save As', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Print', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Share', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Export', exact: true })).toBeVisible();
  await expect(backstage.getByRole('button', { name: 'Options', exact: true })).toBeVisible();

  await page.evaluate(() => {
    const win = window as unknown as {
      __fcInst?: {
        print?: (mode?: 'print' | 'pdf') => void;
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
        };
      };
      __fcPrintModes?: string[];
    };
    win.__fcPrintModes = [];
    if (win.__fcInst) {
      win.__fcInst.workbook?.setText?.({ sheet: 0, row: 0, col: 0 }, 'Preview Cell');
      win.__fcInst.print = (mode = 'print') => {
        win.__fcPrintModes?.push(mode);
      };
    }
  });

  await backstage.getByRole('button', { name: 'Print', exact: true }).first().click();
  await expect(backstage.locator('[data-demo-print-preview]')).toBeVisible();
  await expect(backstage.frameLocator('.demo__print-frame').locator('body')).toContainText(
    'Preview Cell',
  );
  await expect(backstage.getByRole('button', { name: 'Export to PDF', exact: true })).toBeVisible();
  await backstage
    .locator('[data-demo-print-preview]')
    .getByRole('button', { name: 'Print' })
    .click();
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as unknown as { __fcPrintModes?: string[] }).__fcPrintModes ?? [],
      ),
    )
    .toContain('print');

  await backstage.getByRole('button', { name: 'Export', exact: true }).first().click();
  await expect
    .poll(() =>
      page.evaluate(
        () => (window as unknown as { __fcPrintModes?: string[] }).__fcPrintModes ?? [],
      ),
    )
    .toEqual(['print', 'pdf']);

  await backstage.getByRole('button', { name: 'Options', exact: true }).click();
  await expect(page.getByRole('complementary', { name: 'Options panel' })).toBeVisible();

  await backstage.getByRole('button', { name: /^Page Setup\b/ }).click();
  const pageSetup = page.getByRole('dialog', { name: 'Page Setup' });
  await expect(pageSetup).toBeVisible();
  await pageSetup.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(pageSetup).toBeHidden();

  await backstage.getByRole('button', { name: 'Info', exact: true }).click();
  await backstage.getByRole('button', { name: /^Edit Links\b/ }).click();
  const externalLinks = page.getByRole('dialog', { name: 'External Links' });
  await expect(externalLinks).toBeVisible();
  await externalLinks.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(externalLinks).toBeHidden();

  await backstage.getByRole('button', { name: 'Close', exact: true }).click();
  await expect(backstage).toBeHidden();
  await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);

  await page.keyboard.press('F6');
  await expect(page.locator('.demo__quick button[aria-label="Save"]')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('[data-ribbon-tab="home"]')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('.fc-host__formulabar-tag')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('.fc-host')).toBeFocused();
  await page.keyboard.press('F6');
  await expect(page.locator('.fc-host__statusbar')).toBeFocused();
  await page.keyboard.press('Shift+F6');
  await expect(page.locator('.fc-host')).toBeFocused();

  const search = page.getByRole('combobox', { name: 'Search commands' });
  await page.keyboard.press('Alt+Q');
  await expect(search).toBeFocused();
  await search.fill('training');

  const helpResult = page
    .locator('.demo__command-item')
    .filter({ hasText: 'Help and training' })
    .first();
  await expect(helpResult).toBeVisible();
  await search.press('ArrowDown');
  await expect(search).toHaveAttribute('aria-activedescendant', 'demo-search-option-0');
  await expect(page.locator('#demo-search-option-0')).toHaveAttribute('aria-selected', 'true');
  await search.press('Enter');

  await expect(page.locator('[data-ribbon-tab="help"][aria-selected="true"]')).toHaveCount(1);

  await search.fill('coming soon');
  const disabledHelpResult = page
    .locator('.demo__command-item')
    .filter({ hasText: 'Coming soon' })
    .first();
  await expect(disabledHelpResult).toBeVisible();
  await expect(disabledHelpResult).toHaveAttribute('aria-disabled', 'true');
  await expect(disabledHelpResult).toHaveAttribute('data-disabled-reason', 'Coming soon');
  await expect(disabledHelpResult).toContainText('Help');
  await expect(disabledHelpResult).toContainText('Coming soon');
}

export async function runExcelChromeTableStyleGalleryScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.setViewportSize({ width: 1440, height: 520 });
  await sp.mount();
  await sp.expectNoStub();

  await page.locator('[data-ribbon-tab="home"]').click();
  await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);

  const tableButton = page.locator('[data-ribbon-command="formatTableHome"]').first();
  await expect(tableButton).toBeVisible();
  await expect(tableButton).toHaveAttribute('aria-haspopup', 'menu');
  await expect(tableButton).toHaveAttribute('data-ribbon-activation', 'gallery');
  await tableButton.click();

  const menu = page.locator('#menu-table-style-home').first();
  await expect(menu).toBeVisible();
  await expect(tableButton).toHaveAttribute('aria-expanded', 'true');
  await expect(menu.locator('.app__tablestyle-heading')).toHaveCount(3);
  await expect(menu.locator('.app__tablestyle-swatch')).toHaveCount(63);

  const layout = await menu.evaluate((el) => {
    const rect = el.getBoundingClientRect();
    const styles = getComputedStyle(el);
    return {
      bottom: rect.bottom,
      clientHeight: el.clientHeight,
      maxHeight: styles.maxHeight,
      overflowY: styles.overflowY,
      scrollHeight: el.scrollHeight,
      viewportHeight: window.innerHeight,
    };
  });
  expect(layout.overflowY).toBe('auto');
  expect(layout.maxHeight).not.toBe('none');
  expect(layout.clientHeight).toBeLessThanOrEqual(340);
  expect(layout.scrollHeight).toBeGreaterThan(layout.clientHeight);
  expect(layout.bottom).toBeLessThanOrEqual(layout.viewportHeight);

  await menu.evaluate((el) => {
    el.scrollTop = el.scrollHeight;
  });
  await expect(menu.locator('[data-table-style-footer="new-table-style"]')).toBeVisible();
  await expect(menu.locator('[data-table-style-footer="new-pivot-style"]')).toBeVisible();
}

export async function runExcelChromeCreateTableDialogScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.setViewportSize({ width: 1440, height: 520 });
  await sp.mount();
  await sp.expectNoStub();

  await page.evaluate(() => {
    const w = window as unknown as {
      __fcInst?: {
        workbook?: {
          setText?: (addr: { sheet: number; row: number; col: number }, value: string) => void;
          setNumber?: (addr: { sheet: number; row: number; col: number }, value: number) => void;
        };
        store?: {
          setState?: (fn: (state: Record<string, unknown>) => Record<string, unknown>) => void;
        };
      };
    };
    const inst = w.__fcInst;
    if (!inst?.workbook || !inst.store?.setState) {
      throw new Error('window.__fcInst with workbook/store APIs is required');
    }
    inst.workbook.setText?.({ sheet: 0, row: 0, col: 0 }, 'Name');
    inst.workbook.setText?.({ sheet: 0, row: 0, col: 1 }, 'Qty');
    inst.workbook.setText?.({ sheet: 0, row: 1, col: 0 }, 'Alpha');
    inst.workbook.setNumber?.({ sheet: 0, row: 1, col: 1 }, 1);
    inst.workbook.setText?.({ sheet: 0, row: 2, col: 0 }, 'Beta');
    inst.workbook.setNumber?.({ sheet: 0, row: 2, col: 1 }, 2);
    inst.store.setState((state) => ({
      ...state,
      selection: {
        ...(state.selection as Record<string, unknown>),
        active: { sheet: 0, row: 0, col: 0 },
        range: { sheet: 0, r0: 0, c0: 0, r1: 2, c1: 1 },
      },
    }));
  });

  await page.locator('[data-ribbon-tab="insert"]').click();
  await expect(page.locator('[data-ribbon-tab="insert"][aria-selected="true"]')).toHaveCount(1);

  const tableButton = page.locator('[data-ribbon-command="formatTableInsert"]').first();
  await expect(tableButton).toBeVisible();
  await expect(tableButton).toHaveAttribute('data-ribbon-activation', 'dialog');
  await tableButton.click();

  const dialog = page.getByRole('dialog', { name: 'Create Table' });
  await expect(dialog).toBeVisible();
  const rangePicker = dialog.locator('[data-range-picker-kind="table-range"]');
  await expect(rangePicker).toBeVisible();
  await expect(dialog.getByText('Specify the data range to convert to a table.')).toBeVisible();
  await expect(dialog.getByLabel('Specify the data range to convert to a table.')).toHaveValue(
    'Sheet1!$A$1:$B$3',
  );
  await expect(dialog.getByLabel('My table has headers')).toBeChecked();
  await expect(dialog.getByRole('button', { name: 'Cancel' })).toBeVisible();
  await expect(dialog.getByRole('button', { name: 'OK' })).toBeVisible();
  await rangePicker.click();
  await expect(rangePicker).toHaveAttribute('aria-pressed', 'true');
  await expect(dialog).toHaveClass(/fc-fmtdlg--range-picking/);
  await expect(dialog.locator('.fc-range-picker--picking')).toBeVisible();
  const canvasBox = await page.locator('.fc-host__canvas').first().boundingBox();
  expect(canvasBox, 'canvas must be laid out').not.toBeNull();
  if (!canvasBox) throw new Error('canvas not measured');
  const passThrough = await page.evaluate(
    ({ x, y }) => {
      const el = document.elementFromPoint(x, y) as HTMLElement | null;
      return {
        closestHost: !!el?.closest('.fc-host'),
        closestDialog: !!el?.closest('.fc-fmtdlg'),
      };
    },
    {
      x: Math.round(canvasBox.x + canvasBox.width / 2),
      y: Math.round(canvasBox.y + canvasBox.height / 2),
    },
  );
  expect(passThrough.closestHost).toBe(true);
  expect(passThrough.closestDialog).toBe(false);
  await page.keyboard.press('Escape');
  await expect(rangePicker).toHaveAttribute('aria-pressed', 'false');
  await expect(dialog).not.toHaveClass(/fc-fmtdlg--range-picking/);
  await dialog.getByRole('button', { name: 'Cancel' }).click();
  await expect(dialog).toBeHidden();
}

export async function runExcelChromeConditionalFormattingMenuScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.setViewportSize({ width: 1440, height: 520 });
  await sp.mount();
  await sp.expectNoStub();

  await page.locator('[data-ribbon-tab="home"]').click();
  await expect(page.locator('[data-ribbon-tab="home"][aria-selected="true"]')).toHaveCount(1);

  const conditionalButton = page.locator('[data-ribbon-command="conditional"]').first();
  await expect(conditionalButton).toBeVisible();
  await expect(conditionalButton).toHaveAttribute('data-ribbon-activation', 'gallery');
  await conditionalButton.click();

  const menu = page.locator('#menu-conditional').first();
  await expect(menu).toBeVisible();
  await expect(conditionalButton).toHaveAttribute('aria-expanded', 'true');

  const topLevel = await menu.evaluate((el) =>
    Array.from(el.children)
      .filter((child) => child.classList.contains('app__menu-item'))
      .map((child) => {
        const button = child as HTMLElement;
        return {
          text: button.textContent?.trim() ?? '',
          submenu: button.dataset.cfSubmenu ?? null,
          action: button.dataset.cfAction ?? null,
          hasIcon: button.querySelector('.app__cf-icon') !== null,
          hasCaret: button.querySelector('.app__menu-item__caret') !== null,
        };
      }),
  );
  expect(topLevel).toEqual([
    {
      text: 'Highlight Cells Rules',
      submenu: 'highlight',
      action: 'submenu-highlight',
      hasIcon: true,
      hasCaret: true,
    },
    {
      text: 'Top/Bottom Rules',
      submenu: 'topBottom',
      action: 'submenu-topBottom',
      hasIcon: true,
      hasCaret: true,
    },
    {
      text: 'Data Bars',
      submenu: 'dataBar',
      action: 'submenu-dataBar',
      hasIcon: true,
      hasCaret: true,
    },
    {
      text: 'Color Scales',
      submenu: 'colorScale',
      action: 'submenu-colorScale',
      hasIcon: true,
      hasCaret: true,
    },
    {
      text: 'Icon Sets',
      submenu: 'iconSet',
      action: 'submenu-iconSet',
      hasIcon: true,
      hasCaret: true,
    },
    {
      text: 'New Rule...',
      submenu: null,
      action: 'new-rule',
      hasIcon: true,
      hasCaret: false,
    },
    {
      text: 'Clear Rules',
      submenu: 'clear',
      action: 'submenu-clear',
      hasIcon: true,
      hasCaret: true,
    },
    {
      text: 'Manage Rules...',
      submenu: null,
      action: 'manage',
      hasIcon: true,
      hasCaret: false,
    },
  ]);

  const dataBarsTrigger = menu.locator('[data-cf-submenu="dataBar"]').first();
  await dataBarsTrigger.hover();
  const dataBarsPanel = menu.locator('[data-cf-panel="dataBar"]').first();
  await expect(dataBarsPanel).toBeVisible();
  await expect(dataBarsPanel.locator('.app__menu-heading')).toHaveText([
    'Gradient Fill',
    'Solid Fill',
  ]);
  await expect(dataBarsPanel.locator('.app__cf-choice')).toHaveCount(12);
  await expect(dataBarsPanel.locator('[data-cf-action="new-rule"]')).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(conditionalButton).toHaveAttribute('aria-expanded', 'false');
}

export async function runExcelChromeHomeDenseRibbonScenario(page: Page): Promise<void> {
  const sp = new SpreadsheetPage(page);
  await page.setViewportSize({ width: 1440, height: 520 });
  await sp.mount();
  await sp.expectNoStub();

  await page.locator('[data-ribbon-tab="home"]').click();
  const homePanel = page.locator('[data-ribbon-panel="home"]').first();
  await expect(homePanel).toBeVisible();

  const stylesGroup = homePanel.locator('.demo__ribbon-group--styles').first();
  const cellsGroup = homePanel.locator('.demo__ribbon-group--cells').first();
  const editingGroup = homePanel.locator('.demo__ribbon-group--editing').first();
  await expect(stylesGroup).toHaveClass(/demo__ribbon-group--tiles/);
  await expect(cellsGroup).toHaveClass(/demo__ribbon-group--stacked/);
  await expect(editingGroup).toHaveClass(/demo__ribbon-group--mixed/);

  for (const command of ['insertRows', 'deleteRows', 'formatCellsHome']) {
    await expect(cellsGroup.locator(`[data-ribbon-command="${command}"]`)).toHaveClass(
      /demo__rb--stacked/,
    );
  }
  for (const command of ['autosum', 'fillHome', 'clearFormat']) {
    await expect(editingGroup.locator(`[data-ribbon-command="${command}"]`)).toHaveClass(
      /demo__rb--stacked/,
    );
  }
  for (const command of ['sortFilterHome', 'findHome']) {
    await expect(editingGroup.locator(`[data-ribbon-command="${command}"]`)).not.toHaveClass(
      /demo__rb--stacked/,
    );
  }

  const layout = await homePanel.evaluate((panel) => {
    const panelRect = panel.getBoundingClientRect();
    const overflowing: Array<{ command: string; bottom: number; panelBottom: number }> = [];
    for (const button of panel.querySelectorAll<HTMLElement>('[data-ribbon-command]')) {
      const rect = button.getBoundingClientRect();
      if (rect.bottom > panelRect.bottom + 1) {
        overflowing.push({
          command: button.dataset.ribbonCommand ?? '',
          bottom: rect.bottom,
          panelBottom: panelRect.bottom,
        });
      }
    }
    return {
      panelClientWidth: panel.clientWidth,
      panelScrollWidth: panel.scrollWidth,
      overflowing,
    };
  });
  expect(layout.overflowing).toEqual([]);
  expect(layout.panelScrollWidth).toBeLessThanOrEqual(layout.panelClientWidth + 1);
}
