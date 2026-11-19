import { CfRuleType, SparklineType, ValidationType } from '@libraz/formulon';
import { mount, type SpreadsheetInstance } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];

const mountSheet = async (): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en' });
  live.push(instance);
  return instance;
};

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

/** Runs undo then redo, checking the state predicate after each. */
const roundTrip = (
  sheet: SpreadsheetInstance,
  before: () => void,
  after: () => void,
  run: () => void,
): void => {
  sheet.history.clear();
  before();
  run();
  after();
  expect(sheet.history.canUndo).toBe(true);
  sheet.history.undo();
  before();
  expect(sheet.history.canUndo).toBe(false);
  sheet.history.redo();
  after();
};

describe('facade journal records', () => {
  it('journals conditional format rules', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    roundTrip(
      sheet,
      () => {
        expect(s.listCfRules()).toHaveLength(0);
        expect(store.getState().conditional.rules).toHaveLength(0);
      },
      () => {
        expect(s.listCfRules()).toHaveLength(1);
        expect(store.getState().conditional.rules).toHaveLength(1);
      },
      () => s.addCfRule({ type: CfRuleType.Expression, ranges: 'A1:A3', formula1: '=A1>1' }),
    );
  });

  it('journals data validation', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    const stored = (): number =>
      [...store.getState().format.formats.values()].filter((f) => f.validation).length +
      (store.getState().format.rangeValidations?.length ?? 0);
    roundTrip(
      sheet,
      () => {
        expect(s.listValidations()).toHaveLength(0);
        expect(stored()).toBe(0);
      },
      () => {
        expect(s.listValidations()).toHaveLength(1);
        expect(stored()).toBeGreaterThan(0);
      },
      () =>
        s.addValidation({
          ranges: 'A1:A3',
          type: ValidationType.Whole,
          formula1: '1',
          formula2: '9',
        }),
    );
  });

  it('journals sheet protection', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    sheet.history.clear();
    void s.protect().then(() => {});
    await Promise.resolve();
    await new Promise((r) => setTimeout(r, 0));
    expect(s.getSheetProtection()).not.toBeNull();
    expect(store.getState().protection.protectedSheets.has(0)).toBe(true);
    sheet.history.undo();
    expect(s.getSheetProtection()).toBeNull();
    expect(store.getState().protection.protectedSheets.has(0)).toBe(false);
    sheet.history.redo();
    expect(s.getSheetProtection()).not.toBeNull();
    expect(store.getState().protection.protectedSheets.has(0)).toBe(true);
  });

  it('journals protected ranges', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    roundTrip(
      sheet,
      () => {
        expect(s.listProtectedRanges()).toHaveLength(0);
        expect(store.getState().protection.allowedEditRanges).toHaveLength(0);
      },
      () => {
        expect(s.listProtectedRanges()).toHaveLength(1);
        expect(store.getState().protection.allowedEditRanges).toHaveLength(1);
      },
      () => s.addProtectedRange({ name: 'Edit', ranges: 'B2:C3' }),
    );
  });

  it('journals workbook protection', async () => {
    const sheet = await mountSheet();
    const wb = sheet.workbook;
    const { store } = spreadsheetInternals(sheet);
    roundTrip(
      sheet,
      () => {
        expect(wb.getWorkbookProtection()?.lockStructure ?? false).toBe(false);
        expect(store.getState().protection.workbookStructure).toBeUndefined();
      },
      () => {
        expect(wb.getWorkbookProtection()?.lockStructure).toBe(true);
        expect(store.getState().protection.workbookStructure).toBeDefined();
      },
      () => wb.setWorkbookProtection({ lockStructure: true }),
    );
  });

  it('journals print settings', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    roundTrip(
      sheet,
      () => {
        expect(s.getPrintArea()).toBeNull();
        expect(s.listRowBreaks()).toHaveLength(0);
        expect(store.getState().pageSetup.setupBySheet.get(0)?.printArea ?? null).toBeNull();
      },
      () => {
        expect(s.getPrintArea()).toHaveLength(1);
        expect(s.listRowBreaks()).toHaveLength(1);
        expect(store.getState().pageSetup.setupBySheet.get(0)?.printArea).toBeTruthy();
      },
      () =>
        sheet.workbook.batch(() => {
          s.setPrintArea('A1:C10');
          s.setRowBreak(5, {});
        }),
    );
  });

  it('journals the tab colour', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    roundTrip(
      sheet,
      () => {
        expect(s.getTabColor()).toBeNull();
        expect(store.getState().layout.sheetTabColors.has(0)).toBe(false);
      },
      () => {
        expect(s.getTabColor()?.argb).toBe(0xffff0000);
        expect(store.getState().layout.sheetTabColors.get(0)).toBe('#ff0000');
      },
      () => s.setTabColor('#ff0000'),
    );
  });

  it('journals sheet visibility', async () => {
    const sheet = await mountSheet();
    sheet.workbook.addSheet('Second');
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    roundTrip(
      sheet,
      () => {
        expect(s.getVisibility()).toBe(0);
        expect(store.getState().layout.hiddenSheets.has(0)).toBe(false);
      },
      () => {
        expect(s.getVisibility()).toBe(1);
        expect(store.getState().layout.hiddenSheets.has(0)).toBe(true);
      },
      () => s.setVisibility(1),
    );
  });

  it('journals the sheet properties XML', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const xml = '<sheetPr><tabColor rgb="FF00FF00"/></sheetPr>';
    roundTrip(
      sheet,
      () => expect(s.ooxml.getSheetPrXml() ?? '').not.toContain('00FF00'),
      () => expect(s.ooxml.getSheetPrXml()).toContain('00FF00'),
      () => s.ooxml.setSheetPrXml(xml),
    );
  });

  it('journals defined names', async () => {
    const sheet = await mountSheet();
    const wb = sheet.workbook;
    roundTrip(
      sheet,
      () => expect(wb.getDefinedName('Total')).toBeNull(),
      () => expect(wb.getDefinedName('Total')?.formula).toContain('$A$1'),
      () => wb.setDefinedName({ name: 'Total', formula: '=Sheet1!$A$1' }),
    );
  });

  it('journals notes', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    const noted = (): boolean => !!store.getState().format.formats.get('0:1:1')?.comment;
    roundTrip(
      sheet,
      () => {
        expect(s.listComments()).toHaveLength(0);
        expect(noted()).toBe(false);
      },
      () => {
        expect(s.listComments()).toHaveLength(1);
        expect(noted()).toBe(true);
      },
      () => s.setComment('B2', { text: 'hello' }),
    );
  });

  it('journals people and threaded comments', async () => {
    const sheet = await mountSheet();
    const wb = sheet.workbook;
    const s = wb.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    let id = '';
    roundTrip(
      sheet,
      () => {
        expect(wb.listPeople()).toHaveLength(0);
        expect(store.getState().threadedComments.persons).toHaveLength(0);
      },
      () => {
        expect(wb.listPeople()).toHaveLength(1);
        expect(store.getState().threadedComments.persons).toHaveLength(1);
      },
      () => {
        id = wb.addPerson({ displayName: 'Ada' });
      },
    );
    roundTrip(
      sheet,
      () => {
        expect(s.listThreadedComments()).toHaveLength(0);
        expect(store.getState().threadedComments.bySheet.get(0)?.length ?? 0).toBe(0);
      },
      () => {
        expect(s.listThreadedComments()).toHaveLength(2);
        expect(store.getState().threadedComments.bySheet.get(0)?.length ?? 0).toBeGreaterThan(0);
      },
      () =>
        wb.batch(() => {
          const root = s.addThreadedComment({ text: 'q', personId: id, cell: 'C3' });
          s.addThreadedComment({ text: 'a', personId: id, parentId: root });
        }),
    );
  });

  it('journals sparkline groups', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    roundTrip(
      sheet,
      () => {
        expect(s.listSparklineGroups()).toHaveLength(0);
        expect(store.getState().sparkline.sparklines.size).toBe(0);
      },
      () => {
        expect(s.listSparklineGroups()).toHaveLength(1);
        expect(store.getState().sparkline.sparklines.size).toBe(1);
      },
      () =>
        s.addSparklineGroup({
          type: SparklineType.Column,
          sparklines: [{ source: 'A1:C1', location: 'D1' }],
        }),
    );
  });

  it('journals the sheet autofilter and the rows it hides', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    const { store } = spreadsheetInternals(sheet);
    s.setValues('A1', [['h'], ['a'], ['b'], ['a']]);
    sheet.history.clear();
    roundTrip(
      sheet,
      () => {
        expect(s.getAutoFilter()).toBeNull();
        expect(store.getState().ui.filterRanges).toHaveLength(0);
      },
      () => {
        expect(s.getAutoFilter()).not.toBeNull();
        expect(store.getState().ui.filterRanges).toHaveLength(1);
      },
      () => s.setAutoFilter({ range: 'A1:A4' }),
    );
  });

  it('records nothing for view and calc writes', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    sheet.history.clear();
    s.updateView({ zoomPct: 150 });
    s.setSelection({ active: 'B2' });
    sheet.workbook.setCalcMode(1);
    sheet.workbook.updateIterative({ maxIterations: 50 });
    expect(sheet.history.canUndo).toBe(false);
  });
});
