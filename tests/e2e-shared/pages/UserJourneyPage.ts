import type { Locator, Page } from '@playwright/test';
import { expect } from '@playwright/test';

import { SpreadsheetPage } from './SpreadsheetPage.js';

export interface CellAddress {
  sheet: number;
  row: number;
  col: number;
}

export interface CellValue {
  kind: string;
  value?: unknown;
}

export interface CellFormat {
  borders?: Record<string, unknown>;
  [key: string]: unknown;
}

export interface MergeRange {
  sheet: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

const columnNumber = (letters: string): number => {
  let value = 0;
  for (const letter of letters.toUpperCase()) value = value * 26 + letter.charCodeAt(0) - 64;
  return value - 1;
};

export const a1Address = (ref: string, sheet = 0): CellAddress => {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref.trim());
  if (!match) throw new Error(`Expected a single-cell A1 address, got ${ref}`);
  const letters = match[1];
  const row = match[2];
  if (!letters || !row) throw new Error(`Invalid A1 address: ${ref}`);
  return { sheet, row: Number(row) - 1, col: columnNumber(letters) };
};

type WorkbookProbe = {
  workbook: {
    getValue(addr: CellAddress): CellValue;
    cellFormula(addr: CellAddress): string | null;
    sheetCount: number;
    sheetName(index: number): string;
  };
  store: {
    getState(): {
      data: { sheetIndex: number };
      format: { formats: Map<string, CellFormat> };
      merges: { byAnchor: Map<string, MergeRange> };
    };
  };
};

const ensureProbe = (page: Page): Promise<void> =>
  page.evaluate(() => {
    if (!(window as Window & { __fcInst?: WorkbookProbe }).__fcInst) {
      throw new Error('window.__fcInst is not available');
    }
  });

/** User-facing spreadsheet actions shared by the React and Vue journeys. */
export class UserJourneyPage extends SpreadsheetPage {
  private readonly consoleErrors = this.collectConsoleErrors();

  async expectNoConsoleErrors(): Promise<void> {
    expect(this.consoleErrors.read(), 'user journey should not emit console errors').toEqual([]);
  }

  async goTo(ref: string): Promise<void> {
    const nameBox = this.page.locator('.fc-host__formulabar-tag').first();
    await nameBox.fill(ref);
    await nameBox.press('Enter');
    await expect(this.page.locator('.fc-host').first()).toBeFocused();
    let displayedRef = ref;
    const range = /^([A-Z]+\d+):([A-Z]+\d+)$/i.exec(ref);
    if (range?.[1] && range[2]) {
      const sheetIndex = await this.page.evaluate(() => {
        const instance = (window as Window & { __fcInst?: WorkbookProbe }).__fcInst;
        if (!instance) throw new Error('window.__fcInst is not available');
        return instance.store.getState().data.sheetIndex;
      });
      const merge = await this.readMerge(range[1], sheetIndex);
      const first = a1Address(range[1]);
      const last = a1Address(range[2]);
      if (
        merge &&
        merge.r0 === first.row &&
        merge.c0 === first.col &&
        merge.r1 === last.row &&
        merge.c1 === last.col
      )
        displayedRef = range[1];
    }
    await expect.poll(() => nameBox.inputValue(), { timeout: 2_000 }).toBe(displayedRef);
  }

  async enter(ref: string, value: string): Promise<void> {
    await this.goTo(ref);
    await this.page.keyboard.type(value);
    await this.page.keyboard.press('Enter');
  }

  async readValue(ref: string, sheet = 0): Promise<CellValue> {
    const addr = a1Address(ref, sheet);
    await ensureProbe(this.page);
    return this.page.evaluate(
      ({ addr: target }) => {
        const instance = (window as Window & { __fcInst?: WorkbookProbe }).__fcInst;
        if (!instance) throw new Error('window.__fcInst is not available');
        return instance.workbook.getValue(target);
      },
      { addr },
    );
  }

  async readFormula(ref: string, sheet = 0): Promise<string | null> {
    const addr = a1Address(ref, sheet);
    await ensureProbe(this.page);
    return this.page.evaluate(
      ({ addr: target }) => {
        const instance = (window as Window & { __fcInst?: WorkbookProbe }).__fcInst;
        if (!instance) throw new Error('window.__fcInst is not available');
        return instance.workbook.cellFormula(target);
      },
      { addr },
    );
  }

  async readFormat(ref: string, sheet = 0): Promise<CellFormat | null> {
    const addr = a1Address(ref, sheet);
    await ensureProbe(this.page);
    return this.page.evaluate(
      ({ addr: target }) => {
        const instance = (window as Window & { __fcInst?: WorkbookProbe }).__fcInst;
        if (!instance) throw new Error('window.__fcInst is not available');
        const format = instance.store
          .getState()
          .format.formats.get(`${target.sheet}:${target.row}:${target.col}`);
        return format ? { ...format } : null;
      },
      { addr },
    );
  }

  async readMerge(ref: string, sheet = 0): Promise<MergeRange | null> {
    const addr = a1Address(ref, sheet);
    await ensureProbe(this.page);
    return this.page.evaluate(
      ({ addr: target }) => {
        const instance = (window as Window & { __fcInst?: WorkbookProbe }).__fcInst;
        if (!instance) throw new Error('window.__fcInst is not available');
        const merge = instance.store
          .getState()
          .merges.byAnchor.get(`${target.sheet}:${target.row}:${target.col}`);
        return merge ? { ...merge } : null;
      },
      { addr },
    );
  }

  async sheetCount(): Promise<number> {
    await ensureProbe(this.page);
    return this.page.evaluate(() => {
      const instance = (window as Window & { __fcInst?: WorkbookProbe }).__fcInst;
      if (!instance) throw new Error('window.__fcInst is not available');
      return instance.workbook.sheetCount;
    });
  }

  async sheetName(index: number): Promise<string> {
    await ensureProbe(this.page);
    return this.page.evaluate((sheetIndex) => {
      const instance = (window as Window & { __fcInst?: WorkbookProbe }).__fcInst;
      if (!instance) throw new Error('window.__fcInst is not available');
      return instance.workbook.sheetName(sheetIndex);
    }, index);
  }

  async chooseSheet(name: string): Promise<void> {
    const tab = this.page.locator('.fc-host__sheetbar-tab').filter({ hasText: name }).first();
    await expect(tab).toBeVisible();
    await tab.click();
    await expect(tab).toHaveAttribute('aria-selected', 'true');
  }

  async renameSelectedSheet(name: string): Promise<void> {
    const tab = this.page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').first();
    await tab.focus();
    await this.page.keyboard.press('F2');
    const input = this.page.locator('.fc-host__sheetbar-rename').first();
    await expect(input).toBeVisible();
    await input.fill(name);
    await input.press('Enter');
    await expect(input).toBeHidden();
    await expect(
      this.page.locator('.fc-host__sheetbar-tab[aria-selected="true"]').first(),
    ).toHaveText(name);
  }

  async clickRibbon(command: string): Promise<void> {
    const button = this.page.locator(`[data-ribbon-command="${command}"]`).first();
    await expect(button).toBeVisible();
    await button.click();
  }

  async openFormatDialog(): Promise<Locator> {
    await this.clickRibbon('formatCellsHome');
    const menu = this.page.locator('#menu-format-cells').first();
    await expect(menu).toBeVisible();
    await menu.locator('[data-cell-format="dialog"]').click();
    const dialog = this.page.locator('[class="fc-fmtdlg"]').first();
    await expect(dialog).toBeVisible();
    return dialog;
  }

  async openHyperlinkDialog(): Promise<Locator> {
    await this.shortcut('k');
    const dialog = this.page.locator('.fc-hldlg').first();
    await expect(dialog).toBeVisible();
    return dialog;
  }

  async openRibbonSplitMenu(command: string, menuId: string): Promise<Locator> {
    const button = this.page.locator(`[data-ribbon-command="${command}"]`).first();
    await expect(button).toBeVisible();
    await button.locator('.fc-tb__rb-split-chevron').click();
    const menu = this.page.locator(`#${menuId}`).first();
    await expect(menu).toBeVisible();
    return menu;
  }

  async insertRowsOrColumns(kind: 'cells' | 'rows' | 'cols'): Promise<void> {
    await this.clickRibbon('insertRows');
    const menu = this.page.locator('#menu-insert-cells').first();
    await expect(menu).toBeVisible();
    await menu.locator(`[data-cell-insert="${kind}"]`).click();
  }

  async deleteRowsOrColumns(kind: 'cells' | 'rows' | 'cols'): Promise<void> {
    await this.clickRibbon('deleteRows');
    const menu = this.page.locator('#menu-delete-cells').first();
    await expect(menu).toBeVisible();
    await menu.locator(`[data-cell-delete="${kind}"]`).click();
  }

  async insertCopiedCells(): Promise<void> {
    await this.page
      .locator('.fc-host__canvas')
      .first()
      .click({ button: 'right', position: { x: 200, y: 200 } });
    const contextMenu = this.page.locator('.fc-ctxmenu').first();
    await expect(contextMenu).toBeVisible();
    await contextMenu.locator('[data-fc-action="insertCopiedCells"]').click();
    const dialog = this.page.locator('.fc-insertcopied').first();
    await expect(dialog).toBeVisible();
    const right = dialog.locator('input[type="radio"][value="right"]');
    await right.locator('..').click();
    await expect(right).toBeChecked();
    await dialog.getByRole('button', { name: 'OK', exact: true }).click();
  }

  async saveDownload() {
    const downloadPromise = this.page.waitForEvent('download');
    await this.page.getByRole('button', { name: 'Save', exact: true }).click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.xlsx$/);
    expect(await download.failure()).toBeNull();
    const path = await download.path();
    expect(path, 'Save should produce an xlsx download').not.toBeNull();
    if (!path) throw new Error('Save did not expose a local download path');
    return path;
  }

  async openDownload(path: string): Promise<void> {
    await this.page.locator('[data-ribbon-tab="file"]').click();
    const backstage = this.page.locator('.fc-tb__backstage[role="dialog"]').first();
    await expect(backstage).toBeVisible();
    const [chooser] = await Promise.all([
      this.page.waitForEvent('filechooser'),
      backstage.getByRole('button', { name: 'Open', exact: true }).first().click(),
    ]);
    await chooser.setFiles(path);
    await backstage.getByRole('button', { name: 'Close', exact: true }).click();
    await expect(backstage).toBeHidden();
  }
}
