import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { Workbook } from '@libraz/formulon';
import { afterEach, describe, expect, it } from 'vitest';
import { hydrateSheetPrFromEngine } from '../../../src/engine/sheet-pr-sync.js';
import { SheetKind } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import { patchZipText } from './xlsx-zip-patch.js';

const fixture = (): Uint8Array =>
  new Uint8Array(readFileSync(join(__dirname, '../../fixtures/charts/chart_chartsheet.xlsx')));

const rawSheet = (wb: WorkbookHandle, index: number) =>
  (wb as unknown as { wb: Workbook }).wb.sheet(index);

describe('chartsheet tab colour and protection', () => {
  const opened: WorkbookHandle[] = [];
  afterEach(() => {
    for (const wb of opened.splice(0)) wb.dispose();
  });

  const open = async (bytes: Uint8Array = fixture()): Promise<WorkbookHandle> => {
    const wb = await WorkbookHandle.loadBytes(bytes);
    opened.push(wb);
    return wb;
  };

  it('sets a tab colour through the host path and keeps it across save and reload', async () => {
    const wb = await open();
    expect(wb.sheetKind(0)).toBe(SheetKind.Chartsheet);
    const store = createSpreadsheetStore();
    wb.attachStore(store);
    mutators.setSheetTabColor(store, 0, '#c00000');
    expect(wb.getSheetPr(0)?.tabColor).toEqual({ rgb: 'FFC00000' });

    const loaded = await open(wb.save());
    expect(loaded.sheetKind(0)).toBe(SheetKind.Chartsheet);
    expect(loaded.getSheetPr(0)?.tabColor).toEqual({ rgb: 'FFC00000' });
    const fresh = createSpreadsheetStore();
    hydrateSheetPrFromEngine(loaded, fresh);
    expect(fresh.getState().layout.sheetTabColors.get(0)).toBe('#c00000');
  });

  it('sets a tab colour on the engine chartsheet and reads it back', async () => {
    const wb = await open();
    const sheet = rawSheet(wb, 0);
    expect(sheet.kind).toBe(SheetKind.Chartsheet);
    sheet.setTabColor('#00B050');
    expect(sheet.getTabColor()).not.toBeNull();
  });

  it('reads no protection from a chartsheet part without <sheetProtection>', async () => {
    const wb = await open();
    expect(rawSheet(wb, 0).getChartsheetProtection()).toBeNull();
  });

  it('reads the protection a chartsheet part states', async () => {
    const bytes = patchZipText(fixture(), 'xl/chartsheets/sheet1.xml', (xml) => {
      expect(xml).toContain('</sheetViews>');
      return xml.replace(
        '</sheetViews>',
        '</sheetViews><sheetProtection content="1" objects="1"/>',
      );
    });
    const wb = await open(bytes);
    const protection = rawSheet(wb, 0).getChartsheetProtection();
    expect(protection?.allowEditContent).toBe(false);
    expect(protection?.allowEditObjects).toBe(false);
  });
});
