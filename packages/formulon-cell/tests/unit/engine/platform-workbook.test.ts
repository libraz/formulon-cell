// @vitest-environment node

import { FontScheme } from '@libraz/formulon';
import { describe, expect, it, vi } from 'vitest';
import { hydrateLayoutFromEngine, rehydrateAxisLayout } from '../../../src/engine/layout-sync.js';
import { loadFormulon } from '../../../src/engine/loader.js';
import { engineProfileIdFor } from '../../../src/engine/spreadsheet-profile.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

type RawSheet = {
  listColumnOverrides(): ReadonlyArray<{ first: number; last: number; widthChars: number | null }>;
  getFormatDefaults(): {
    baseColumnWidthChars: number;
    defaultRowHeightPt: number | null;
  };
  updateFormatDefaults(patch: object): void;
  getRowHeightPt(row: number): number;
};
type RawWorkbook = { sheet(index: number): RawSheet };

const rawWorkbook = (handle: WorkbookHandle): RawWorkbook =>
  (handle as unknown as { wb: RawWorkbook }).wb;

const PX_PER_PT = 96 / 72;

describe('engineProfileIdFor', () => {
  it.each([
    ['ja', 'mac', 'mac-365-ja_JP'],
    ['ja-JP', 'mac', 'mac-365-ja_JP'],
    ['en', 'mac', 'mac-365-en_US'],
    ['en-GB', 'mac', 'mac-365-en_US'],
    ['zh-Hant-TW', 'mac', 'mac-365-zh_TW'],
    ['zh-TW', 'default', 'win-365-zh_TW'],
    ['zh_HK', 'mac', 'mac-365-zh_TW'],
    ['zh-Hant', 'mac', 'mac-365-zh_TW'],
    ['zh-Hans-HK', 'mac', 'mac-365-zh_CN'],
    ['zh-CN', 'mac', 'mac-365-zh_CN'],
    ['zh', 'default', 'win-365-zh_CN'],
    ['ko_KR', 'default', 'win-365-ko_KR'],
    ['eo', 'mac', 'mac-365-en_US'],
    ['', 'default', 'win-365-en_US'],
  ] as const)('maps %s on %s to %s', (locale, platform, id) => {
    expect(engineProfileIdFor(locale, platform)).toBe(id);
  });
});

describe('WorkbookHandle.createDefault profile font', () => {
  // Font 0 as each engine profile states it: linked to the theme's minor font, drawn as that face.
  it.each([
    ['en', 'default', 'Aptos Narrow', 11],
    ['ja', 'default', '游ゴシック', 11],
    ['de', 'default', 'Aptos Narrow', 11],
    ['fr', 'default', 'Aptos Narrow', 11],
    ['zh', 'default', '等线', 11],
    ['ko', 'default', '맑은 고딕', 11],
    ['th', 'default', 'Tahoma', 11],
    ['ja-JP', 'mac', '游ゴシック', 12],
    ['zh', 'mac', '等线', 12],
    ['zh-TW', 'mac', '新細明體', 12],
    ['zh-Hans-HK', 'mac', '等线', 12],
    ['ko', 'mac', '맑은 고딕', 12],
    ['th', 'mac', 'Tahoma', 12],
    ['en', 'mac', 'Aptos Narrow', 12],
    ['de', 'mac', 'Aptos Narrow', 12],
  ] as const)('gives a %s %s book %s %d pt', async (locale, platform, name, size) => {
    const wb = await WorkbookHandle.createDefault({ locale, platform });
    try {
      expect(wb.getFontRecord(0)).toMatchObject({ name, sizePt: size, scheme: FontScheme.Minor });
      expect(wb.workbookDisplayFont).toEqual({ name, size });
    } finally {
      wb.dispose();
    }
  });

  it('rejects when the profile cannot load, rather than making a book without it', async () => {
    const module = await loadFormulon();
    const spy = vi.spyOn(module.Workbook, 'createForProfile').mockImplementation(() => {
      throw new Error('profile unavailable');
    });
    try {
      for (const platform of ['default', 'mac'] as const) {
        await expect(WorkbookHandle.createDefault({ locale: 'ja', platform })).rejects.toThrow(
          'profile unavailable',
        );
      }
    } finally {
      spy.mockRestore();
    }
  });
});

describe('WorkbookHandle.createDefault platform', () => {
  it('creates a Mac ja book from the engine profile and copies its defaults to new sheets', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      expect(wb.geometryPlatform).toBe('mac');
      expect(wb.workbookDefaultFont).toMatchObject({ name: '游ゴシック', size: 12 });
      const raw = rawWorkbook(wb);
      expect(raw.sheet(0).getRowHeightPt(0)).toBe(20);
      // 游ゴシック 12: digit advance 0.56 x 12 pt rounds to 7, so 10 x 7 + 5 = 75 pt columns.
      expect(wb.getSheetDefaultLayoutPx(0)).toEqual({
        defaultColWidth: 75 * PX_PER_PT,
        defaultRowHeight: 20 * PX_PER_PT,
        headerDigitAdvance: expect.closeTo(6.72, 8),
      });

      const added = wb.addSheet();
      expect(added).toBe(1);
      const defaults = raw.sheet(1).getFormatDefaults();
      expect(defaults.baseColumnWidthChars).toBe(10);
      expect(defaults.defaultRowHeightPt).not.toBeNull();
      expect(defaults.defaultRowHeightPt).toBe(20);
      expect(raw.sheet(1).getRowHeightPt(0)).toBe(20);
      expect(wb.getSheetDefaultLayoutPx(1)).toEqual(wb.getSheetDefaultLayoutPx(0));
    } finally {
      wb.dispose();
    }
  });

  it('creates a Mac en book with Aptos Narrow 12, 16 pt rows and 65 pt columns', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'en', platform: 'mac' });
    try {
      expect(wb.workbookDefaultFont).toMatchObject({ name: 'Aptos Narrow', size: 12 });
      // Aptos Narrow 12: digit advance 0.52 x 12 pt rounds to 6, so 10 x 6 + 5 = 65 pt columns.
      expect(wb.getSheetDefaultLayoutPx(0)).toEqual({
        defaultColWidth: 65 * PX_PER_PT,
        defaultRowHeight: 16 * PX_PER_PT,
        headerDigitAdvance: expect.closeTo(6.24, 8),
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
        const column = raw
          .sheet(0)
          .listColumnOverrides()
          .find((entry) => entry.first === 2);
        expect(column?.widthChars).toBe(savedRaw);
        expect(wb.getColumnLayouts(0).find((entry) => entry.first === 2)?.width).toBe(
          pt * PX_PER_PT,
        );
      }
    } finally {
      wb.dispose();
    }
  });

  it.each([
    ['en', 15, 20, 64],
    ['de', 15, 20, 64],
    // 游ゴシック 11, base 8: 69 px rounded up to 72 px (8.38 characters).
    ['ja', 18.75, 25, 72],
  ])(
    'takes a Windows %s book row height of %d pt from the profile, on added sheets too',
    async (locale, pt, px, colPx) => {
      const wb = await WorkbookHandle.createDefault({ locale });
      try {
        expect(wb.geometryPlatform).toBe('default');
        const raw = rawWorkbook(wb);
        expect(raw.sheet(0).getFormatDefaults().defaultRowHeightPt).toBe(pt);
        expect(raw.sheet(0).getRowHeightPt(0)).toBe(pt);
        expect(wb.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBe(px);
        expect(wb.getSheetDefaultLayoutPx(0)?.defaultColWidth).toBeCloseTo(colPx, 0);
        expect(wb.addSheet()).toBe(1);
        expect(raw.sheet(1).getRowHeightPt(0)).toBe(pt);
        expect(raw.sheet(1).getFormatDefaults().baseColumnWidthChars).toBe(8);
      } finally {
        wb.dispose();
      }
    },
  );

  it('keeps a Mac sheet default row height only when customHeight pins it', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      const raw = rawWorkbook(wb).sheet(0);
      const stated = { baseColumnWidthChars: 10, defaultRowHeightPt: 30 };
      raw.updateFormatDefaults({ ...stated, customHeight: false });
      // 游ゴシック 12 recomputes to 20 pt, as Mac Excel does on load.
      expect(wb.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBeCloseTo(20 * PX_PER_PT, 8);
      raw.updateFormatDefaults({ ...stated, customHeight: true });
      expect(wb.getSheetDefaultLayoutPx(0)?.defaultRowHeight).toBeCloseTo(30 * PX_PER_PT, 8);
    } finally {
      wb.dispose();
    }
  });

  it('adds bare sheets to a loaded book', async () => {
    const created = await WorkbookHandle.createDefault({ locale: 'en' });
    const loaded = await WorkbookHandle.loadBytes(created.save());
    created.dispose();
    try {
      expect(loaded.addSheet()).toBe(1);
      const raw = rawWorkbook(loaded);
      expect(raw.sheet(0).getFormatDefaults().defaultRowHeightPt).not.toBeNull();
      expect(raw.sheet(1).getFormatDefaults().defaultRowHeightPt).toBeNull();
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

  it('reads the layout platform from the profile, and moves the profile when it is set', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'ja', platform: 'mac' });
    try {
      expect(wb.spreadsheetProfileId()).toBe('mac-ja_JP');
      expect(wb.geometryPlatform).toBe('mac');
      expect(wb.setSpreadsheetProfileId('windows-ja_JP')).toBe(true);
      expect(wb.geometryPlatform).toBe('default');
      wb.geometryPlatform = 'mac';
      expect(wb.spreadsheetProfileId()).toBe('mac-ja_JP');
    } finally {
      wb.dispose();
    }
  });

  it('makes a plain book on the stub, which has no profiles', async () => {
    const plain = await WorkbookHandle.createDefault({ preferStub: true });
    const mac = await WorkbookHandle.createDefault({
      locale: 'ja',
      platform: 'mac',
      preferStub: true,
    });
    try {
      expect(mac.geometryPlatform).toBe('mac');
      expect(mac.workbookDefaultFont).toEqual(plain.workbookDefaultFont);
      expect(mac.sheetCount).toBe(1);
      expect(mac.addSheet()).toBe(1);
    } finally {
      mac.dispose();
      plain.dispose();
    }
  });
});
