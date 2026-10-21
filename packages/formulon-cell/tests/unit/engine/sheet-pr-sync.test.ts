import { describe, expect, it } from 'vitest';
import { hydrateSheetPrFromEngine } from '../../../src/engine/sheet-pr-sync.js';
import { mergeSheetPrFragment, readSheetPr } from '../../../src/engine/sheet-pr-xml.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('sheetPr fragment editing', () => {
  it('adds a tab colour and keeps the rest of sheetPr verbatim', () => {
    const xml =
      '<sheetPr codeName="Sheet1" filterMode="1"><outlinePr applyStyles="1"/><pageSetUpPr fitToPage="1"/></sheetPr>';
    const out = mergeSheetPrFragment(xml, { tabColor: { rgb: 'FFC00000' } });
    expect(out).toBe(
      '<sheetPr codeName="Sheet1" filterMode="1"><tabColor rgb="FFC00000"/><outlinePr applyStyles="1"/><pageSetUpPr fitToPage="1"/></sheetPr>',
    );
  });

  it('sets and clears the summary position without touching other outlinePr attributes', () => {
    const xml = '<sheetPr><tabColor theme="4" tint="0.4"/><outlinePr applyStyles="1"/></sheetPr>';
    const set = mergeSheetPrFragment(xml, { summaryBelow: false, summaryRight: false });
    expect(set).toBe(
      '<sheetPr><tabColor theme="4" tint="0.4"/><outlinePr applyStyles="1" summaryBelow="0" summaryRight="0"/></sheetPr>',
    );
    expect(readSheetPr(set ?? '')).toMatchObject({ summaryBelow: false, summaryRight: false });
    const cleared = mergeSheetPrFragment(set ?? '', { summaryBelow: true, summaryRight: true });
    expect(cleared).toBe(
      '<sheetPr><tabColor theme="4" tint="0.4"/><outlinePr applyStyles="1"/></sheetPr>',
    );
  });

  it('returns an empty fragment when nothing is left to state', () => {
    expect(
      mergeSheetPrFragment('<sheetPr><tabColor rgb="FF00FF00"/></sheetPr>', { tabColor: null }),
    ).toBe('');
    expect(mergeSheetPrFragment('', { summaryBelow: true })).toBe('');
  });

  it('reads theme tab colours and the summary defaults', () => {
    expect(readSheetPr('<sheetPr><tabColor theme="5" tint="-0.25"/></sheetPr>')).toEqual({
      tabColor: { theme: 5, tint: -0.25 },
      summaryBelow: true,
      summaryRight: true,
    });
  });
});

describe('sheetPr persistence through the engine', () => {
  it('saves tab colour and outline summary from the store and loads them back', async () => {
    const wb = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      expect(wb.capabilities.sheetPrXml).toBe(true);
      expect(wb.addSheet('Second')).toBe(1);
      const store = createSpreadsheetStore();
      wb.attachStore(store);
      mutators.setSheetTabColor(store, 1, '#c00000');
      mutators.setOutlineSummary(store, 1, { below: false });
      mutators.setOutlineSummary(store, 0, { right: false });

      loaded = await WorkbookHandle.loadBytes(wb.save());
      expect(loaded.getSheetPr(0)).toMatchObject({
        tabColor: null,
        summaryBelow: true,
        summaryRight: false,
      });
      expect(loaded.getSheetPr(1)).toMatchObject({
        tabColor: { rgb: 'FFC00000' },
        summaryBelow: false,
        summaryRight: true,
      });

      const fresh = createSpreadsheetStore();
      hydrateSheetPrFromEngine(loaded, fresh);
      const { layout } = fresh.getState();
      expect(layout.sheetTabColors.get(1)).toBe('#c00000');
      expect(layout.outlineSummary?.get(1)).toEqual({ below: false, right: true });
      expect(layout.outlineSummary?.get(0)).toEqual({ below: true, right: false });

      mutators.setSheetTabColor(store, 1, null);
      expect(wb.getSheetPr(1)?.tabColor).toBeNull();
    } finally {
      loaded?.dispose();
      wb.dispose();
    }
  });

  it('keeps the other sheetPr content when the tab colour changes', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const raw = (wb as unknown as { wb: { getSheetSheetPrXml(s: number): { xml: string } } }).wb;
      expect(wb.setSheetPageSetup(0, { fitToPage: true, fitToWidth: 1, fitToHeight: 1 })).toBe(
        true,
      );
      expect(raw.getSheetSheetPrXml(0).xml).toContain('fitToPage');
      expect(wb.setSheetPr(0, { tabColor: { rgb: 'FF00B050' } })).toBe(true);
      expect(raw.getSheetSheetPrXml(0).xml).toContain('fitToPage');
      expect(wb.getSheetPr(0)?.tabColor).toEqual({ rgb: 'FF00B050' });
    } finally {
      wb.dispose();
    }
  });
});
