// @vitest-environment node

import { describe, expect, it } from 'vitest';
import { hydrateLayoutFromEngine, rehydrateAxisLayout } from '../../../src/engine/layout-sync.js';
import { engineProfileIdFor } from '../../../src/engine/spreadsheet-profile.js';
import { defaultFontForLocale, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type RawStatus = { ok: boolean };
type RawWorkbook = {
  getSheetColumns(sheet: number): {
    status: RawStatus;
    columns: ReadonlyArray<{ first: number; last: number; width: number }>;
  };
  getSheetFormatDefaults(sheet: number): {
    status: RawStatus;
    baseColWidth: number;
    defaultRowHeight: number;
    hasDefaultRowHeight: boolean;
  };
  getRowHeightPt(sheet: number, row: number): { status: RawStatus; value: number };
};

const rawWorkbook = (handle: WorkbookHandle): RawWorkbook =>
  (handle as unknown as { wb: RawWorkbook }).wb;

const PX_PER_PT = 96 / 72;

describe('engineProfileIdFor', () => {
  it.each([
    ['ja', 'mac', 'mac-365-ja_JP'],
    ['ja-JP', 'mac', 'mac-365-ja_JP'],
    ['en', 'mac', 'mac-365-en_US'],
    ['en-GB', 'mac', 'mac-365-en_US'],
    ['zh-Hant-TW', 'mac', 'mac-365-zh_CN'],
    ['ko_KR', 'default', 'win-365-ko_KR'],
    ['eo', 'mac', 'mac-365-en_US'],
    ['', 'default', 'win-365-en_US'],
  ] as const)('maps %s on %s to %s', (locale, platform, id) => {
    expect(engineProfileIdFor(locale, platform)).toBe(id);
  });
});

describe('defaultFontForLocale', () => {
  it('keeps the 11 pt default-profile fonts', () => {
    expect(defaultFontForLocale('ja')).toEqual({ name: '游ゴシック', size: 11 });
    expect(defaultFontForLocale('en')).toEqual({ name: 'Calibri', size: 11 });
    expect(defaultFontForLocale(undefined)).toEqual({ name: 'Calibri', size: 11 });
    expect(defaultFontForLocale('ja', 'default')).toEqual({ name: '游ゴシック', size: 11 });
  });

  it.each([
    ['ja-JP', '游ゴシック'],
    ['zh', '等线'],
    ['ko', '맑은 고딕'],
    ['th', 'Tahoma'],
    ['en', 'Aptos Narrow'],
    ['de', 'Aptos Narrow'],
    ['eo', 'Aptos Narrow'],
  ])('names the Mac %s family at 12 pt', (locale, name) => {
    expect(defaultFontForLocale(locale, 'mac')).toEqual({ name, size: 12 });
  });
});

describe('WorkbookHandle.createDefault platform', () => {
  it('creates a Mac ja book from the engine profile and copies its defaults to new sheets', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      expect(wb.geometryPlatform).toBe('mac');
      expect(wb.workbookDefaultFont).toMatchObject({ name: '游ゴシック', size: 12 });
      const raw = rawWorkbook(wb);
      expect(raw.getRowHeightPt(0, 0).value).toBe(20);
      expect(wb.getSheetDefaultLayoutPx(0)).toEqual({
        defaultColWidth: 75 * PX_PER_PT,
        defaultRowHeight: 20 * PX_PER_PT,
      });

      const added = wb.addSheet();
      expect(added).toBe(1);
      const defaults = raw.getSheetFormatDefaults(1);
      expect(defaults.baseColWidth).toBe(10);
      expect(defaults.hasDefaultRowHeight).toBe(true);
      expect(defaults.defaultRowHeight).toBe(20);
      expect(raw.getRowHeightPt(1, 0).value).toBe(20);
      expect(wb.getSheetDefaultLayoutPx(1)).toEqual(wb.getSheetDefaultLayoutPx(0));
    } finally {
      wb.dispose();
    }
  });

  it('creates a Mac en book with Aptos Narrow 12, 16 pt rows and 65 pt columns', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'en', platform: 'mac' });
    try {
      expect(wb.workbookDefaultFont).toMatchObject({ name: 'Aptos Narrow', size: 12 });
      expect(wb.getSheetDefaultLayoutPx(0)).toEqual({
        defaultColWidth: 65 * PX_PER_PT,
        defaultRowHeight: 16 * PX_PER_PT,
      });
    } finally {
      wb.dispose();
    }
  });

  it('writes Mac column widths as the raw width Mac Excel saves and reads them back', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      const raw = rawWorkbook(wb);
      for (const [pt, savedRaw] of [
        [64, 9.140625],
        [75, 10.7109375],
        [145, 20.7109375],
      ] as const) {
        expect(wb.setColumnWidth(0, 2, 2, pt * PX_PER_PT)).toBe(true);
        const column = raw.getSheetColumns(0).columns.find((entry) => entry.first === 2);
        expect(column?.width).toBe(savedRaw);
        expect(wb.getColumnLayouts(0).find((entry) => entry.first === 2)?.width).toBe(
          pt * PX_PER_PT,
        );
      }
    } finally {
      wb.dispose();
    }
  });

  it('seeds the Windows default row height on default-profile books', async () => {
    const en = await WorkbookHandle.createDefault({ locale: 'en' });
    try {
      expect(en.geometryPlatform).toBe('default');
      expect(en.workbookDefaultFont).toMatchObject({ name: 'Calibri', size: 11 });
      const raw = rawWorkbook(en);
      expect(raw.getSheetFormatDefaults(0).defaultRowHeight).toBe(15);
      expect(raw.getRowHeightPt(0, 0).value).toBe(15);
      expect(en.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBe(20);
      expect(en.getSheetDefaultLayoutPx(0)?.defaultColWidth).toBeCloseTo(64, 1);
      expect(en.addSheet()).toBe(1);
      expect(raw.getRowHeightPt(1, 0).value).toBe(15);
      expect(raw.getSheetFormatDefaults(1).baseColWidth).toBe(8);
    } finally {
      en.dispose();
    }

    const ja = await WorkbookHandle.createDefault({ locale: 'ja' });
    try {
      expect(rawWorkbook(ja).getRowHeightPt(0, 0).value).toBe(18.75);
      expect(ja.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBe(25);
      // 游ゴシック 11, base 8: 69 px rounded up to 72 px (8.38 characters).
      expect(ja.getSheetDefaultLayoutPx(0)?.defaultColWidth).toBe(72);
    } finally {
      ja.dispose();
    }
  });

  it('adds bare sheets to a loaded book', async () => {
    const created = await WorkbookHandle.createDefault({ locale: 'en' });
    const loaded = await WorkbookHandle.loadBytes(created.save());
    created.dispose();
    try {
      expect(loaded.addSheet()).toBe(1);
      const raw = rawWorkbook(loaded);
      expect(raw.getSheetFormatDefaults(0).hasDefaultRowHeight).toBe(true);
      expect(raw.getSheetFormatDefaults(1).hasDefaultRowHeight).toBe(false);
    } finally {
      loaded.dispose();
    }
  });

  it('re-projects axis sizes under the other platform without touching the view', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      expect(wb.setColumnWidth(0, 1, 1, 84 * PX_PER_PT)).toBe(true);
      const store = createSpreadsheetStore();
      hydrateLayoutFromEngine(wb, store, 0);
      expect(store.getState().layout.defaultColWidth).toBe(100);
      expect(store.getState().layout.colWidths.get(1)).toBe(112);
      store.setState((s) => ({
        ...s,
        layout: { ...s.layout, freezeRows: 2 },
        viewport: { ...s.viewport, zoom: 1.5 },
        ui: { ...s.ui, showGridLines: false },
      }));

      wb.geometryPlatform = 'default';
      rehydrateAxisLayout(wb, store, 0);
      let state = store.getState();
      // Windows reads base 10 with 游ゴシック 12 (9 px MDW): 97 px rounded up to 104 px.
      expect(state.layout.defaultColWidth).toBe(104);
      expect(state.layout.defaultRowHeight).toBeCloseTo(80 / 3, 8);
      expect(state.layout.colWidths.get(1)).not.toBe(112);
      expect(state.layout.freezeRows).toBe(2);
      expect(state.viewport.zoom).toBe(1.5);
      expect(state.ui.showGridLines).toBe(false);

      wb.geometryPlatform = 'mac';
      rehydrateAxisLayout(wb, store, 0);
      state = store.getState();
      expect(state.layout.defaultColWidth).toBe(100);
      expect(state.layout.colWidths.get(1)).toBe(112);
    } finally {
      wb.dispose();
    }
  });

  it('degrades to the platform font and size defaults on the stub', async () => {
    const mac = await WorkbookHandle.createDefault({
      locale: 'ja',
      platform: 'mac',
      preferStub: true,
    });
    try {
      expect(mac.geometryPlatform).toBe('mac');
      expect(mac.workbookDefaultFont).toMatchObject({ name: '游ゴシック', size: 12 });
      expect(mac.sheetCount).toBe(1);
      expect(mac.addSheet()).toBe(1);
      const store = createSpreadsheetStore();
      hydrateLayoutFromEngine(mac, store, 0);
      expect(store.getState().layout.defaultColWidth).toBe(100);
      expect(store.getState().layout.defaultRowHeight).toBeCloseTo(80 / 3, 8);
      mac.geometryPlatform = 'default';
      rehydrateAxisLayout(mac, store, 0);
      expect(store.getState().layout.defaultColWidth).toBe(64);
      expect(store.getState().layout.defaultRowHeight).toBe(20);
    } finally {
      mac.dispose();
    }

    const fallback = await WorkbookHandle.createDefault({ locale: 'en', preferStub: true });
    try {
      expect(fallback.geometryPlatform).toBe('default');
      expect(fallback.workbookDefaultFont).toMatchObject({ name: 'Calibri', size: 11 });
    } finally {
      fallback.dispose();
    }
  });
});
