import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import * as cellStyleCommands from '../../../src/commands/cell-styles.js';
import {
  activeCellStyleId,
  applyCellStyle,
  applyCellStyleByName,
  applyCellStyleToSelection,
  CELL_STYLES,
  type CellStyleId,
  createCellStyleFromActiveFormat,
  customCellStyleId,
  getCellStyle,
  listCustomCellStyles,
  mergeCellStylesFromWorkbook,
} from '../../../src/commands/cell-styles.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import { syncCellFormatsToEngine } from '../../../src/engine/cell-format-sync.js';
import {
  cellStyleKeysByXfId,
  collectNamedStyles,
} from '../../../src/engine/cell-style-writeback.js';
import { addrKey, WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/index.js';

const macStyleOracle = JSON.parse(
  readFileSync(
    resolve(dirname(fileURLToPath(import.meta.url)), '../../fixtures/excel-mac-style-groups.json'),
    'utf8',
  ),
) as {
  styles: readonly {
    name: string;
    galleryId: string | null;
    includedGroups: readonly string[];
  }[];
};

const namedStyleWorkbook = (): WorkbookHandle =>
  ({
    getNamedCellStyles: () => [
      {
        index: 0,
        name: 'Normal',
        xfId: 0,
        builtinId: 0,
        iLevel: 0,
        customBuiltin: false,
      },
      {
        index: 1,
        name: 'Imported Review',
        xfId: 2,
        builtinId: -1,
        iLevel: 0,
        customBuiltin: false,
      },
    ],
    getCellStyleXf: (xfId: number) =>
      xfId === 2
        ? {
            fontIndex: 1,
            fillIndex: 1,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 2,
            verticalAlign: 1,
            wrapText: true,
          }
        : {
            fontIndex: 0,
            fillIndex: 0,
            borderIndex: 0,
            numFmtId: 0,
            horizontalAlign: 0,
            verticalAlign: 2,
            wrapText: false,
          },
    getFontRecord: (fontIndex: number) => ({
      name: 'Calibri',
      size: 11,
      bold: fontIndex === 1,
      italic: false,
      strike: false,
      underline: 0,
      colorArgb: fontIndex === 1 ? 0xff006100 : 0xff000000,
    }),
    getFillRecord: (fillIndex: number) => ({
      pattern: fillIndex === 1 ? 1 : 0,
      fgArgb: fillIndex === 1 ? 0xffc6efce : 0,
      bgArgb: 0,
    }),
    getBorderRecord: () => ({
      left: { style: 0, colorArgb: 0xff000000 },
      right: { style: 0, colorArgb: 0xff000000 },
      top: { style: 0, colorArgb: 0xff000000 },
      bottom: { style: 0, colorArgb: 0xff000000 },
      diagonal: { style: 0, colorArgb: 0xff000000 },
      diagonalUp: false,
      diagonalDown: false,
    }),
    getNumFmtCode: () => null,
  }) as unknown as WorkbookHandle;

describe('CELL_STYLES', () => {
  it('contains the spreadsheet-flavored presets', () => {
    const ids = CELL_STYLES.map((s) => s.id);
    expect(ids).toContain('normal');
    expect(ids).toContain('heading1');
    expect(ids).toContain('good');
    expect(ids).toContain('checkCell');
    expect(ids).toContain('explanatoryText');
    expect(ids).toContain('accent1');
    expect(ids).toContain('accent6_20');
    expect(ids).toContain('currency');
  });

  it('matches every observed Mac gallery include flag', () => {
    const observed = new Map(
      macStyleOracle.styles.map((style) => {
        const galleryId =
          style.galleryId ??
          `accent${style.name.at(-1)}_${style.name.startsWith('40%') ? '40' : '60'}`;
        return [galleryId, [...style.includedGroups].sort()] as const;
      }),
    );
    expect(observed.size).toBe(47);
    for (const style of CELL_STYLES) {
      expect([...(style.includedGroups ?? []).slice().sort()]).toEqual(observed.get(style.id));
    }
  });

  it('maps the twelve previously unregistered Mac accent rows', () => {
    const expectedIds: CellStyleId[] = [
      ...Array.from({ length: 6 }, (_, index) => `accent${index + 1}_40` as CellStyleId),
      ...Array.from({ length: 6 }, (_, index) => `accent${index + 1}_60` as CellStyleId),
    ];
    const missing = macStyleOracle.styles.filter((style) => style.galleryId === null);
    expect(macStyleOracle.styles).toHaveLength(47);
    expect(missing.map((style) => style.name)).toEqual([
      ...expectedIds.slice(0, 6).map((id) => `40% - アクセント ${id[6]}`),
      ...expectedIds.slice(6).map((id) => `60% - アクセント ${id[6]}`),
    ]);
    const byId = new Map(CELL_STYLES.map((style) => [style.id, style]));
    expect(expectedIds.map((id) => byId.get(id)?.builtinId)).toEqual([
      31, 35, 39, 43, 47, 51, 32, 36, 40, 44, 48, 52,
    ]);
    for (const style of missing) {
      const id = (
        style.name.startsWith('40%')
          ? `accent${style.name.at(-1)}_40`
          : `accent${style.name.at(-1)}_60`
      ) as CellStyleId;
      expect(byId.get(id)?.includedGroups).toEqual(['font', 'fill']);
    }
  });

  it('provides the observed Mac fallback palette without changing defaults', () => {
    expect(cellStyleCommands.cellStyleFallbackProfileForPlatform?.('mac')).toBe('excel365Mac');
    expect(cellStyleCommands.cellStyleFallbackProfileForPlatform?.('windows')).toBe('default');
    expect(cellStyleCommands.cellStyleFallbackFormat?.('accent1_40', 'default')).toMatchObject({
      color: '#000000',
      fill: '#b4c7e7',
    });
    expect(cellStyleCommands.cellStyleFallbackFormat?.('accent1_40', 'excel365Mac')).toMatchObject({
      color: '#000000',
      fill: '#83cceb',
    });
    expect(cellStyleCommands.cellStyleFallbackFormat?.('accent1', 'excel365Mac')).toMatchObject({
      color: '#ffffff',
      fill: '#156082',
    });
    const observed: readonly (readonly [CellStyleId, string, string])[] = [
      ['accent1', '#ffffff', '#156082'],
      ['accent1_20', '#000000', '#c0e6f5'],
      ['accent1_40', '#000000', '#83cceb'],
      ['accent1_60', '#000000', '#44b3e1'],
      ['accent2', '#ffffff', '#e97132'],
      ['accent2_20', '#000000', '#fbe2d5'],
      ['accent2_40', '#000000', '#f7c7ac'],
      ['accent2_60', '#000000', '#f1a983'],
      ['accent3', '#ffffff', '#196b24'],
      ['accent3_20', '#000000', '#c1f0c8'],
      ['accent3_40', '#000000', '#83e28e'],
      ['accent3_60', '#000000', '#47d359'],
      ['accent4', '#ffffff', '#0f9ed5'],
      ['accent4_20', '#000000', '#caedfb'],
      ['accent4_40', '#000000', '#94dcf8'],
      ['accent4_60', '#000000', '#61cbf3'],
      ['accent5', '#ffffff', '#a02b93'],
      ['accent5_20', '#000000', '#f2ceef'],
      ['accent5_40', '#000000', '#e49edd'],
      ['accent5_60', '#000000', '#d86dcd'],
      ['accent6', '#ffffff', '#4ea72e'],
      ['accent6_20', '#000000', '#daf2d0'],
      ['accent6_40', '#000000', '#b5e6a2'],
      ['accent6_60', '#000000', '#8ed973'],
    ];
    for (const [id, color, fill] of observed) {
      expect(cellStyleCommands.cellStyleFallbackFormat?.(id, 'excel365Mac')).toMatchObject({
        color,
        fill,
      });
    }
    const nonAccentObserved: readonly (readonly [CellStyleId, string, string?])[] = [
      ['title', '#0e2841'],
      ['heading1', '#0e2841'],
      ['heading2', '#0e2841'],
      ['heading3', '#0e2841'],
      ['heading4', '#0e2841'],
      ['checkCell', '#ffffff', '#a5a5a5'],
    ];
    for (const [id, color, fill] of nonAccentObserved) {
      expect(cellStyleCommands.cellStyleFallbackFormat?.(id, 'excel365Mac')).toMatchObject({
        color,
        ...(fill === undefined ? {} : { fill }),
      });
    }
    const portableDefaults: readonly (readonly [CellStyleId, string, string?])[] = [
      ['title', '#1f4e79'],
      ['heading1', '#1f4e79'],
      ['heading2', '#1f4e79'],
      ['heading3', '#1f4e79'],
      ['heading4', '#1f4e79'],
      ['checkCell', '#375623', '#a9d08e'],
    ];
    for (const [id, color, fill] of portableDefaults) {
      expect(cellStyleCommands.cellStyleFallbackFormat?.(id, 'default')).toMatchObject({
        color,
        ...(fill === undefined ? {} : { fill }),
      });
    }
    expect(cellStyleCommands.cellStyleFallbackFormat?.('note', 'excel365Mac')).toMatchObject({
      color: '#333333',
      fill: '#ffffcc',
    });
    expect(
      cellStyleCommands.cellStyleFallbackFormat?.('heading1', 'excel365Mac').borders,
    ).toMatchObject({ bottom: { style: 'medium', color: '#1f4e79' } });
  });

  it('keeps a single canonical profile when materialized cells disagree', () => {
    const store = createSpreadsheetStore();
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        cellStyle: 'accent1_40',
        color: '#000000',
        fill: '#83cceb',
      },
    );
    const macRegistration = collectNamedStyles(store.getState()).find(
      (style) => style.key === 'accent1_40',
    );
    expect(macRegistration?.format).toMatchObject({ color: '#000000', fill: '#83cceb' });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 1 },
      {
        cellStyle: 'accent1_40',
        color: '#000000',
        fill: '#b4c7e7',
      },
    );
    const registration = collectNamedStyles(store.getState()).find(
      (style) => style.key === 'accent1_40',
    );
    expect(registration?.format).toMatchObject({ color: '#000000', fill: '#b4c7e7' });

    const edited = createSpreadsheetStore();
    mutators.setCellFormat(
      edited,
      { sheet: 0, row: 0, col: 0 },
      {
        cellStyle: 'accent1_40',
        color: '#123456',
        fill: '#83cceb',
      },
    );
    expect(
      collectNamedStyles(edited.getState()).find((style) => style.key === 'accent1_40')?.format,
    ).toMatchObject({ color: '#000000', fill: '#b4c7e7' });
  });

  it('keeps the Mac palette across real catalog publication, reapplication, and save/load', async () => {
    const workbook = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      expect(workbook.isStub).toBe(false);
      const store = createSpreadsheetStore();
      const history = new History();
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      const context = {
        getWorkbook: () => workbook,
        getFallbackProfile: () => 'excel365Mac' as const,
      };
      expect(applyCellStyleToSelection(store, history, 'accent1_40', context)).toBe(true);
      expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
        cellStyle: 'accent1_40',
        color: '#000000',
        fill: '#83cceb',
      });
      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      const first = workbook.getNamedCellStyles().find((style) => style.builtinId === 31);
      expect(first).toBeDefined();
      if (!first) return;
      const firstXf = workbook.getCellStyleXf(first.xfId);
      expect(firstXf).toBeDefined();
      if (!firstXf) return;
      expect(workbook.getFillRecord(firstXf.fillIndex)?.fgArgb).toBe(0xff83cceb);
      expect(workbook.getFontRecord(firstXf.fontIndex)?.colorArgb).toBe(0xff000000);

      mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
      expect(applyCellStyleToSelection(store, history, 'accent1_40', context)).toBe(true);
      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      const second = workbook.getNamedCellStyles().find((style) => style.builtinId === 31);
      expect(second?.xfId).toBe(first.xfId);
      expect(
        workbook.getFillRecord(workbook.getCellStyleXf(second?.xfId ?? -1)?.fillIndex ?? -1)
          ?.fgArgb,
      ).toBe(0xff83cceb);

      loaded = await WorkbookHandle.loadBytes(workbook.save());
      expect(loaded.isStub).toBe(false);
      const loadedStyle = loaded.getNamedCellStyles().find((style) => style.builtinId === 31);
      expect(loadedStyle).toBeDefined();
      if (!loadedStyle) return;
      const loadedXf = loaded.getCellStyleXf(loadedStyle.xfId);
      expect(loadedXf).toBeDefined();
      if (!loadedXf) return;
      expect(loaded.getFillRecord(loadedXf.fillIndex)?.fgArgb).toBe(0xff83cceb);
      expect(loaded.getFontRecord(loadedXf.fontIndex)?.colorArgb).toBe(0xff000000);
      expect(cellStyleKeysByXfId(loaded).get(loadedStyle.xfId)).toBe('accent1_40');
      expect(loaded.getCellXf(loaded.getCellXfIndex(0, 0, 0) ?? -1)?.xfId).toBe(loadedStyle.xfId);
      expect(loaded.getCellXf(loaded.getCellXfIndex(0, 0, 1) ?? -1)?.xfId).toBe(loadedStyle.xfId);
    } finally {
      loaded?.dispose();
      workbook.dispose();
    }
  });

  it('keeps the Mac Check Cell palette across first, second, F4, and save/load', async () => {
    const workbook = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.getNamedCellStyles().find((style) => style.builtinId === 23)).toBeUndefined();
      const store = createSpreadsheetStore();
      const history = new History();
      const context = {
        getWorkbook: () => workbook,
        getFallbackProfile: () => 'excel365Mac' as const,
      };
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 0 },
        {
          fontFamily: 'Aptos Display',
          fontSize: 22,
          italic: true,
          strike: true,
          underline: true,
          align: 'center',
          vAlign: 'bottom',
          wrap: true,
        },
      );
      expect(applyCellStyleToSelection(store, history, 'checkCell', context)).toBe(true);
      expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
        cellStyle: 'checkCell',
        bold: true,
        color: '#ffffff',
        fill: '#a5a5a5',
        align: 'center',
        vAlign: 'bottom',
        wrap: true,
      });
      const firstFormat = store.getState().format.formats.get('0:0:0');
      expect(firstFormat?.fontFamily).toBeUndefined();
      expect(firstFormat?.fontSize).toBeUndefined();
      expect(firstFormat?.italic).toBeUndefined();
      expect(firstFormat?.strike).toBeUndefined();
      expect(firstFormat?.underline).toBeUndefined();

      const expectNativeCheckCell = (columns: readonly number[], expectedStyleXfId?: number) => {
        const published = workbook.getNamedCellStyles().find((style) => style.builtinId === 23);
        expect(published).toBeDefined();
        if (!published) throw new Error('Check Cell named style was not published');
        if (expectedStyleXfId !== undefined) expect(published.xfId).toBe(expectedStyleXfId);
        const styleXf = workbook.getCellStyleXf(published.xfId);
        expect(styleXf).toBeDefined();
        if (!styleXf) throw new Error('Check Cell style XF was not published');
        expect(workbook.getFontRecord(styleXf.fontIndex)?.colorArgb).toBe(0xffffffff);
        expect(workbook.getFillRecord(styleXf.fillIndex)?.fgArgb).toBe(0xffa5a5a5);
        for (const col of columns) {
          const cellXf = workbook.getCellXf(workbook.getCellXfIndex(0, 0, col) ?? -1);
          expect(cellXf).toMatchObject({
            xfId: published.xfId,
            fontIndex: styleXf.fontIndex,
            fillIndex: styleXf.fillIndex,
            borderIndex: styleXf.borderIndex,
          });
        }
        return {
          styleXfId: published.xfId,
          fontIndex: styleXf.fontIndex,
          fillIndex: styleXf.fillIndex,
          borderIndex: styleXf.borderIndex,
        };
      };

      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      const firstNative = expectNativeCheckCell([0]);

      mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
      expect(applyCellStyleToSelection(store, history, 'checkCell', context)).toBe(true);
      expect(store.getState().format.formats.get('0:0:1')).toMatchObject({
        cellStyle: 'checkCell',
        color: '#ffffff',
        fill: '#a5a5a5',
      });
      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      expectNativeCheckCell([0, 1], firstNative.styleXfId);

      mutators.setRange(store, { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 });
      expect(history.repeatLast()).toBe(true);
      expect(store.getState().format.formats.get('0:0:2')).toMatchObject({
        cellStyle: 'checkCell',
        color: '#ffffff',
        fill: '#a5a5a5',
      });
      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      expectNativeCheckCell([0, 1, 2], firstNative.styleXfId);

      loaded = await WorkbookHandle.loadBytes(workbook.save());
      const loadedStyle = loaded.getNamedCellStyles().find((style) => style.builtinId === 23);
      expect(loadedStyle).toBeDefined();
      if (!loadedStyle) return;
      const loadedXf = loaded.getCellStyleXf(loadedStyle.xfId);
      expect(loadedXf).toBeDefined();
      if (!loadedXf) return;
      expect(loaded.getFontRecord(loadedXf.fontIndex)?.colorArgb).toBe(0xffffffff);
      expect(loaded.getFillRecord(loadedXf.fillIndex)?.fgArgb).toBe(0xffa5a5a5);
      expect(cellStyleKeysByXfId(loaded).get(loadedStyle.xfId)).toBe('checkCell');
      for (const col of [0, 1, 2]) {
        expect(loaded.getCellXf(loaded.getCellXfIndex(0, 0, col) ?? -1)).toMatchObject({
          xfId: loadedStyle.xfId,
          fontIndex: loadedXf.fontIndex,
          fillIndex: loadedXf.fillIndex,
          borderIndex: loadedXf.borderIndex,
        });
      }
    } finally {
      loaded?.dispose();
      workbook.dispose();
    }
  });

  it('samples the current owner profile again for F4', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    let profile: 'default' | 'excel365Mac' = 'excel365Mac';
    const context = { getFallbackProfile: () => profile };
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(applyCellStyleToSelection(store, history, 'accent1_40', context)).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.fill).toBe('#83cceb');
    profile = 'default';
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    expect(history.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get('0:0:1')?.fill).toBe('#b4c7e7');
  });

  it('prefers raw builtin-23 selectors over Mac Check Cell fallback colors', async () => {
    const workbook = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      expect(workbook.isStub).toBe(false);
      const normal = workbook.getCellStyleXf(0);
      const defaultFont = workbook.getFontRecord(0);
      const defaultFill = workbook.getFillRecord(0);
      expect(normal).toBeDefined();
      expect(defaultFont).toBeDefined();
      expect(defaultFill).toBeDefined();
      if (!normal || !defaultFont || !defaultFill) return;
      const themedFont = workbook.addFontRecord({
        ...defaultFont,
        colorArgb: 0xff123456,
        color: {
          kind: 2,
          rgb: defaultFont.color?.rgb ?? 0,
          theme: 4,
          tint: 0.2,
          indexed: defaultFont.color?.indexed ?? 0,
        },
      });
      const themedFill = workbook.addFillRecord({
        ...defaultFill,
        pattern: 1,
        fgArgb: 0xff654321,
        fg: {
          kind: 2,
          rgb: defaultFill.fg?.rgb ?? 0,
          theme: 5,
          tint: 0.4,
          indexed: defaultFill.fg?.indexed ?? 0,
        },
      });
      const styleXfId = workbook.addCellStyleXfRecord({
        ...normal,
        fontIndex: themedFont,
        fillIndex: themedFill,
      });
      expect(workbook.setNamedCellStyle('Check Cell', styleXfId, 23)).toBe(true);
      expect(
        cellStyleCommands.resolveCellStyleFormat('checkCell', workbook, 'excel365Mac'),
      ).toMatchObject({
        color: '#123456',
        fill: '#654321',
      });

      const store = createSpreadsheetStore();
      const history = new History();
      const context = {
        getWorkbook: () => workbook,
        getFallbackProfile: () => 'excel365Mac' as const,
      };
      for (const col of [0, 1]) {
        mutators.setRange(store, { sheet: 0, r0: 0, c0: col, r1: 0, c1: col });
        expect(applyCellStyleToSelection(store, history, 'checkCell', context)).toBe(true);
        expect(store.getState().format.formats.get(`0:0:${col}`)).toMatchObject({
          cellStyle: 'checkCell',
          color: '#123456',
          fill: '#654321',
        });
      }
      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      const published = workbook.getNamedCellStyles().find((style) => style.builtinId === 23);
      expect(published?.xfId).toBe(styleXfId);
      const publishedXf = workbook.getCellStyleXf(published?.xfId ?? -1);
      expect(publishedXf).toBeDefined();
      if (!publishedXf) return;
      const firstStyleXfComponents = {
        fontIndex: publishedXf.fontIndex,
        fillIndex: publishedXf.fillIndex,
        borderIndex: publishedXf.borderIndex,
      };
      expect(workbook.getFontRecord(publishedXf.fontIndex)?.color).toMatchObject({
        kind: 2,
        theme: 4,
        tint: 0.2,
      });
      expect(workbook.getFillRecord(publishedXf.fillIndex)?.fg).toMatchObject({
        kind: 2,
        theme: 5,
        tint: 0.4,
      });
      const captureCellXf = (col: number) => {
        const cellXf = workbook.getCellXf(workbook.getCellXfIndex(0, 0, col) ?? -1);
        expect(cellXf).toBeDefined();
        if (!cellXf) throw new Error(`raw Check Cell cell XF missing for column ${col}`);
        return {
          xfId: cellXf.xfId,
          fontIndex: cellXf.fontIndex,
          fillIndex: cellXf.fillIndex,
          borderIndex: cellXf.borderIndex,
        };
      };
      const firstCellXfs = [captureCellXf(0), captureCellXf(1)];
      expect(firstCellXfs[0]).toMatchObject({ xfId: styleXfId });
      expect(firstCellXfs[1]).toMatchObject({ xfId: styleXfId });

      mutators.setRange(store, { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 });
      expect(history.repeatLast()).toBe(true);
      expect(store.getState().format.formats.get('0:0:2')).toMatchObject({
        cellStyle: 'checkCell',
        color: '#123456',
        fill: '#654321',
      });
      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      const afterRepeat = workbook.getNamedCellStyles().find((style) => style.builtinId === 23);
      expect(afterRepeat?.xfId).toBe(styleXfId);
      const afterRepeatXf = workbook.getCellStyleXf(afterRepeat?.xfId ?? -1);
      expect(afterRepeatXf).toMatchObject(firstStyleXfComponents);
      expect(captureCellXf(0)).toEqual(firstCellXfs[0]);
      expect(captureCellXf(1)).toEqual(firstCellXfs[1]);
      expect(captureCellXf(2)).toMatchObject({
        xfId: styleXfId,
        ...firstStyleXfComponents,
      });

      loaded = await WorkbookHandle.loadBytes(workbook.save());
      const loadedStyle = loaded.getNamedCellStyles().find((style) => style.builtinId === 23);
      expect(loadedStyle?.xfId).toBe(styleXfId);
      expect(cellStyleKeysByXfId(loaded).get(loadedStyle?.xfId ?? -1)).toBe('checkCell');
      const loadedXf = loaded.getCellStyleXf(loadedStyle?.xfId ?? -1);
      expect(loadedXf).toMatchObject(firstStyleXfComponents);
      expect(loaded.getFontRecord(loadedXf?.fontIndex ?? -1)?.color).toMatchObject({
        kind: 2,
        theme: 4,
        tint: 0.2,
      });
      expect(loaded.getFillRecord(loadedXf?.fillIndex ?? -1)?.fg).toMatchObject({
        kind: 2,
        theme: 5,
        tint: 0.4,
      });
      for (const col of [0, 1, 2]) {
        expect(loaded.getCellXf(loaded.getCellXfIndex(0, 0, col) ?? -1)).toMatchObject({
          xfId: loadedStyle?.xfId,
          ...firstStyleXfComponents,
        });
      }
    } finally {
      loaded?.dispose();
      workbook.dispose();
    }
  });

  it('replaces the Mac accent font and fill groups without inventing font defaults', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        italic: true,
        fontFamily: 'Aptos Display',
        fontSize: 22,
        fill: '#ffffff',
        color: '#ff0000',
      },
    );
    expect(
      applyCellStyleToSelection(store, history, 'accent1_40', {
        getFallbackProfile: () => 'excel365Mac',
      }),
    ).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
      color: '#000000',
      fill: '#83cceb',
    });
    const format = store.getState().format.formats.get('0:0:0');
    expect(format?.bold).toBeUndefined();
    expect(format?.italic).toBeUndefined();
    expect(format?.fontFamily).toBeUndefined();
    expect(format?.fontSize).toBeUndefined();
  });

  it('prefers raw builtin-31 selectors over RGB fallbacks and round-trips them', async () => {
    const workbook = await WorkbookHandle.createDefault();
    let loaded: WorkbookHandle | null = null;
    try {
      expect(workbook.isStub).toBe(false);
      const normal = workbook.getCellStyleXf(0);
      const defaultFont = workbook.getFontRecord(0);
      const defaultFill = workbook.getFillRecord(0);
      expect(normal).toBeDefined();
      expect(defaultFont).toBeDefined();
      expect(defaultFill).toBeDefined();
      if (!normal || !defaultFont || !defaultFill) return;
      const themedFont = workbook.addFontRecord({
        ...defaultFont,
        colorArgb: 0xff123456,
        color: {
          kind: 2,
          rgb: defaultFont.color?.rgb ?? 0,
          theme: 4,
          tint: 0.2,
          indexed: defaultFont.color?.indexed ?? 0,
        },
      });
      const themedFill = workbook.addFillRecord({
        ...defaultFill,
        pattern: 1,
        fgArgb: 0xff654321,
        fg: {
          kind: 2,
          rgb: defaultFill.fg?.rgb ?? 0,
          theme: 5,
          tint: 0.4,
          indexed: defaultFill.fg?.indexed ?? 0,
        },
      });
      const styleXfId = workbook.addCellStyleXfRecord({
        ...normal,
        fontIndex: themedFont,
        fillIndex: themedFill,
      });
      expect(workbook.setNamedCellStyle('40% - Accent 1', styleXfId, 31)).toBe(true);
      expect(
        cellStyleCommands.resolveCellStyleFormat('accent1_40', workbook, 'excel365Mac'),
      ).toMatchObject({
        color: '#123456',
        fill: '#654321',
      });

      const store = createSpreadsheetStore();
      mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
      expect(
        applyCellStyleToSelection(store, new History(), 'accent1_40', {
          getWorkbook: () => workbook,
          getFallbackProfile: () => 'excel365Mac',
        }),
      ).toBe(true);
      syncCellFormatsToEngine(workbook, store, 0, { strict: true });
      const published = workbook.getNamedCellStyles().find((style) => style.builtinId === 31);
      expect(published?.xfId).toBe(styleXfId);
      const publishedXf = workbook.getCellStyleXf(published?.xfId ?? -1);
      expect(workbook.getFontRecord(publishedXf?.fontIndex ?? -1)?.color).toMatchObject({
        kind: 2,
        theme: 4,
        tint: 0.2,
      });
      expect(workbook.getFillRecord(publishedXf?.fillIndex ?? -1)?.fg).toMatchObject({
        kind: 2,
        theme: 5,
        tint: 0.4,
      });

      loaded = await WorkbookHandle.loadBytes(workbook.save());
      const loadedStyle = loaded.getNamedCellStyles().find((style) => style.builtinId === 31);
      expect(loadedStyle?.xfId).toBe(styleXfId);
      const loadedXf = loaded.getCellStyleXf(loadedStyle?.xfId ?? -1);
      expect(loaded.getFontRecord(loadedXf?.fontIndex ?? -1)?.color).toMatchObject({
        kind: 2,
        theme: 4,
        tint: 0.2,
      });
      expect(loaded.getFillRecord(loadedXf?.fillIndex ?? -1)?.fg).toMatchObject({
        kind: 2,
        theme: 5,
        tint: 0.4,
      });
    } finally {
      loaded?.dispose();
      workbook.dispose();
    }
  });

  it('maps above-cap localized built-ins on read without authoring a false ordinal', () => {
    const workbook = {
      capabilities: { cellStyles: true },
      getNamedCellStyles: () => [
        {
          index: 4,
          name: '60% - アクセント 6',
          xfId: 9,
          builtinId: 52,
          iLevel: 0,
          customBuiltin: false,
        },
      ],
    } as unknown as WorkbookHandle;
    expect(cellStyleKeysByXfId(workbook).get(9)).toBe('accent6_60');
    const store = createSpreadsheetStore();
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'accent6_60' });
    const registration = collectNamedStyles(store.getState()).find(
      (style) => style.key === 'accent6_60',
    );
    expect(registration).toMatchObject({ builtinId: 52 });
  });

  it('keeps native font discriminator bits in portable fallbacks', () => {
    expect(getCellStyle('title')?.format.bold).toBe(false);
    expect(getCellStyle('warning')?.format.italic).toBe(false);
    expect(getCellStyle('linkedCell')?.format.italic).toBe(false);
    expect(getCellStyle('calculation')?.format.italic).toBe(false);
    expect(getCellStyle('explanatoryText')?.builtinId).toBe(53);
    expect(getCellStyle('heading3')?.format).toMatchObject({ bold: true, fontSize: 11 });
    expect(getCellStyle('heading4')?.format).toMatchObject({
      bold: true,
      italic: false,
      fontSize: 11,
    });
  });
});

describe('getCellStyle', () => {
  it('returns the matching def', () => {
    expect(getCellStyle('good')?.format.fill).toBe('#c6efce');
    expect(getCellStyle('accent1')?.format.fill).toBe('#4472c4');
    expect(getCellStyle('accent4_20')?.format.fill).toBe('#fff2cc');
  });

  it('returns undefined for unknown ids', () => {
    // @ts-expect-error — testing invalid id rejection
    expect(getCellStyle('not-a-style')).toBeUndefined();
  });
});

describe('applyCellStyle', () => {
  it('writes the named-style fields onto every cell in range', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 }, 'good');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    expect(fmt?.fill).toBe('#c6efce');
    expect(fmt?.color).toBe('#006100');
    expect(fmt?.cellStyle).toBe('good');
    const fmtCorner = store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }));
    expect(fmtCorner?.fill).toBe('#c6efce');
    expect(fmtCorner?.cellStyle).toBe('good');
  });

  it('clears every format field for the "normal" preset', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'good');
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'normal');
    const fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 0, col: 0 }));
    // Either the entry is gone, or all visible style fields are undefined.
    if (fmt) {
      expect(fmt.fill).toBeUndefined();
      expect(fmt.color).toBeUndefined();
      expect(fmt.bold).toBeUndefined();
      expect(fmt.cellStyle).toBeUndefined();
    }
  });

  it('applies Excel-style accent and explanatory presets', () => {
    const store = createSpreadsheetStore();
    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'accent5_20');
    let fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#1f4e79', fill: '#ddebf7' });

    applyCellStyle(store, null, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 }, 'explanatoryText');
    fmt = store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }));
    expect(fmt).toMatchObject({ color: '#7f7f7f', italic: true });
  });

  it('no-ops on unknown style id', () => {
    const store = createSpreadsheetStore();
    const before = store.getState();
    // @ts-expect-error — testing invalid id rejection
    applyCellStyle(store, null, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'not-a-style');
    expect(store.getState()).toBe(before);
  });

  it('creates a named style from the active cell format and records history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 0, col: 0 },
      {
        bold: true,
        fill: '#c6efce',
        color: '#006100',
      },
    );

    expect(
      createCellStyleFromActiveFormat(
        store,
        history,
        { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 },
        'Review OK',
        { include: { fill: false } },
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
    ).toMatchObject({
      cellStyle: 'Review OK',
      bold: true,
      color: '#006100',
    });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 }))?.fill,
    ).toBeUndefined();
    expect(listCustomCellStyles(store.getState())).toMatchObject([
      {
        id: customCellStyleId('Review OK'),
        label: 'Review OK',
        format: {
          bold: true,
          color: '#006100',
        },
      },
    ]);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        customCellStyleId('Review OK'),
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 })),
    ).toMatchObject({
      cellStyle: 'Review OK',
      bold: true,
      color: '#006100',
    });
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 }))?.fill,
    ).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 2 })),
    ).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(listCustomCellStyles(store.getState())).toEqual([]);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 1, col: 1 })),
    ).toBeUndefined();
  });

  it('merges visible workbook named styles into the session custom style registry', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    expect(mergeCellStylesFromWorkbook(store, history, namedStyleWorkbook())).toEqual({
      imported: 1,
      skipped: 1,
    });
    expect(listCustomCellStyles(store.getState())).toMatchObject([
      {
        id: customCellStyleId('Imported Review'),
        label: 'Imported Review',
        format: {
          bold: true,
          color: '#006100',
          fill: '#c6efce',
          align: 'center',
          vAlign: 'middle',
          wrap: true,
        },
      },
    ]);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 3, c0: 3, r1: 3, c1: 3 },
        customCellStyleId('Imported Review'),
      ),
    ).toBe(true);
    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 3, col: 3 })),
    ).toMatchObject({
      cellStyle: 'Imported Review',
      bold: true,
      fill: '#c6efce',
    });
    expect(history.undo()).toBe(true);
    expect(history.undo()).toBe(true);
    expect(listCustomCellStyles(store.getState())).toEqual([]);
  });

  it('repeats a gallery style onto whatever is selected when F4 fires', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    applyCellStyle(store, history, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, 'good');

    mutators.setRange(store, { sheet: 0, r0: 4, c0: 2, r1: 4, c1: 2 });
    expect(history.repeatLast()).toBe(true);

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 4, col: 2 })),
    ).toMatchObject({ cellStyle: 'good' });
  });

  it('repeats a custom style by name rather than redefining it', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.setRangeFormat(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }, { bold: true });
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    expect(
      createCellStyleFromActiveFormat(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        'Brand',
      ),
    ).toBe(true);

    mutators.setRange(store, { sheet: 0, r0: 2, c0: 0, r1: 2, c1: 0 });
    expect(history.repeatLast()).toBe(true);

    expect(
      store.getState().format.formats.get(addrKey({ sheet: 0, row: 2, col: 0 })),
    ).toMatchObject({ cellStyle: 'Brand', bold: true });
    expect(listCustomCellStyles(store.getState())).toHaveLength(1);
  });

  it('applies a built-in style across the complete primary-plus-extra union with group replacement', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        extraRanges: [
          { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 2 },
          { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 2 },
        ],
      },
    }));
    for (const col of [0, 1, 2]) {
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col },
        {
          bold: true,
          italic: true,
          fill: '#123456',
          numFmt: { kind: 'percent', decimals: 1 },
          locked: false,
          comment: `keep-${col}`,
          hyperlink: `https://example.test/${col}`,
        },
      );
    }
    mutators.setCellFormat(
      store,
      { sheet: 0, row: 2, col: 2 },
      {
        bold: true,
        fill: '#123456',
        numFmt: { kind: 'percent', decimals: 1 },
        locked: false,
        comment: 'keep-hole',
      },
    );

    expect(
      applyCellStyleToSelection(store, history, 'good', {
        origin: 'instanceApi',
        commandId: 'cellStyles',
      }),
    ).toBe(true);

    for (const addr of [
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 1 },
      { sheet: 0, row: 0, col: 2 },
    ]) {
      expect(store.getState().format.formats.get(addrKey(addr))).toMatchObject({
        color: '#006100',
        fill: '#c6efce',
        numFmt: { kind: 'percent', decimals: 1 },
        locked: false,
      });
      expect(store.getState().format.formats.get(addrKey(addr))?.bold).toBeUndefined();
      expect(store.getState().format.formats.get(addrKey(addr))?.comment).toMatch(/keep/);
    }
    expect(store.getState().format.formats.has('0:1:1')).toBe(false);
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(history.redo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.fill).toBe('#c6efce');
  });

  it('resolves the active merge anchor and preserves unresolved style keys', () => {
    const store = createSpreadsheetStore();
    mutators.mergeRange(store, { sheet: 0, r0: 2, c0: 2, r1: 2, c1: 3 });
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { cellStyle: 'good' });
    mutators.setActive(store, { sheet: 0, row: 2, col: 3 });
    expect(activeCellStyleId(store.getState())).toBe('good');

    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { cellStyle: 'Imported Review' });
    expect(activeCellStyleId(store.getState())).toBe('Imported Review');
    mutators.setCellFormat(store, { sheet: 0, row: 2, col: 2 }, { cellStyle: undefined });
    expect(activeCellStyleId(store.getState())).toBe('normal');
  });

  it('uses the observed group boundaries for Note and Heading 3', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    const addr = { sheet: 0, row: 0, col: 0 };
    mutators.setCellFormat(store, addr, {
      bold: true,
      color: '#123456',
      fill: '#abcdef',
      numFmt: { kind: 'percent', decimals: 2 },
      locked: false,
      borders: {
        top: { style: 'thick', color: '#ff0000' },
        bottom: { style: 'thin', color: '#00ff00' },
      },
    });

    expect(applyCellStyleToSelection(store, history, 'note')).toBe(true);
    let format = store.getState().format.formats.get(addrKey(addr));
    expect(format).toMatchObject({
      bold: true,
      color: '#123456',
      fill: '#ffffcc',
      numFmt: { kind: 'percent', decimals: 2 },
      locked: false,
    });
    expect(format?.borders).toBeUndefined();

    expect(applyCellStyleToSelection(store, history, 'heading3')).toBe(true);
    format = store.getState().format.formats.get(addrKey(addr));
    expect(format).toMatchObject({ bold: true, color: '#1f4e79', fill: '#ffffcc' });
    expect(format?.borders).toBeUndefined();
  });

  it('uses the same replacement groups through the explicit range API', () => {
    const store = createSpreadsheetStore();
    const addr = { sheet: 0, row: 1, col: 1 };
    mutators.setCellFormat(store, addr, {
      bold: true,
      color: '#123456',
      fill: '#abcdef',
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      borders: { top: { style: 'thick' } },
    });

    applyCellStyle(store, null, { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 }, 'note');
    const format = store.getState().format.formats.get(addrKey(addr));
    expect(format).toMatchObject({
      bold: true,
      color: '#123456',
      fill: '#ffffcc',
      numFmt: { kind: 'currency', decimals: 2, symbol: '$' },
      cellStyle: 'note',
    });
    expect(format?.borders).toBeUndefined();
  });

  it('keeps explicit range styling scoped and respects protection and the materialization cap', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:Scoped',
      label: 'Scoped',
      format: { bold: true },
    });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { fill: '#abcdef' });
    store.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        extraRanges: [{ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }],
      },
    }));

    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('Scoped'),
      ),
    ).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(store.getState().format.formats.get('0:0:1')?.fill).toBe('#abcdef');

    const locked = createSpreadsheetStore();
    const lockedHistory = new History();
    mutators.upsertCustomCellStyle(locked, {
      id: 'custom:Scoped',
      label: 'Scoped',
      format: { bold: true },
    });
    mutators.setSheetProtected(locked, 0, true);
    expect(
      applyCellStyleByName(
        locked,
        lockedHistory,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('Scoped'),
      ),
    ).toBe(false);
    expect(locked.getState().format.formats.size).toBe(0);
    expect(lockedHistory.repeatLast()).toBe(false);

    const huge = createSpreadsheetStore();
    expect(
      applyCellStyle(huge, null, { sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 1 }, 'good'),
    ).toBeUndefined();
    expect(huge.getState().format.formats.size).toBe(0);
  });

  it('authorizes the supplied range independently from the live selection', async () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:ScopedPolicy',
      label: 'ScopedPolicy',
      format: { bold: true },
    });
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 });
    mutators.addExtraRange(store, { sheet: 0, r0: 0, c0: 2, r1: 0, c1: 2 });
    const liveSelection = store.getState().selection;
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    const controller = new InteractionController({ store, getWb: () => workbook, history });
    const unregister = registerInteractionController(store, controller);
    const policy = fixedFormPolicy({ ranges: [{ sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }] });
    controller.setPolicy({ ...policy, operations: { ...policy.operations, format: true } });
    try {
      expect(
        applyCellStyleByName(
          store,
          history,
          { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
          customCellStyleId('ScopedPolicy'),
        ),
      ).toBe(false);
      expect(store.getState().format.formats.size).toBe(0);
      expect(history.repeatLast()).toBe(false);

      const targetPolicy = fixedFormPolicy({
        ranges: [{ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }],
      });
      controller.setPolicy({
        ...targetPolicy,
        operations: { ...targetPolicy.operations, format: true },
      });
      expect(
        applyCellStyleByName(
          store,
          history,
          { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 },
          customCellStyleId('ScopedPolicy'),
        ),
      ).toBe(true);
      expect(store.getState().format.formats.get('0:0:1')?.bold).toBe(true);
      expect(store.getState().format.formats.has('0:0:2')).toBe(false);
      expect(store.getState().selection).toEqual(liveSelection);
    } finally {
      unregister();
      controller.dispose();
      workbook.dispose();
    }
  });

  it('writes only unlocked cells in one protected-range history entry', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:ProtectedMixed',
      label: 'ProtectedMixed',
      format: { bold: true },
    });
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { locked: false });
    mutators.setSheetProtected(store, 0, true);
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        customCellStyleId('ProtectedMixed'),
      ),
    ).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({
      locked: false,
      bold: true,
    });
    expect(store.getState().format.formats.get('0:0:1')).toBeUndefined();
    expect(history.undo()).toBe(true);
    expect(store.getState().format.formats.get('0:0:0')).toMatchObject({ locked: false });
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBeUndefined();
  });

  it('rejects malformed, foreign-selection, and partial-merge plans without replacing repeat', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:StableRepeat',
      label: 'StableRepeat',
      format: { italic: true },
    });
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('StableRepeat'),
      ),
    ).toBe(true);
    mutators.setRange(store, { sheet: 0, r0: 1, c0: 0, r1: 1, c1: 0 });
    const selectionBefore = store.getState().selection;
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 2, c0: 0, r1: 1, c1: 0 },
        customCellStyleId('StableRepeat'),
      ),
    ).toBe(false);
    expect(store.getState().selection).toEqual(selectionBefore);
    expect(history.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get('0:1:0')?.italic).toBe(true);

    const foreignSelection = createSpreadsheetStore();
    mutators.upsertCustomCellStyle(foreignSelection, {
      id: 'custom:StableRepeat',
      label: 'StableRepeat',
      format: { italic: true },
    });
    foreignSelection.setState((state) => ({
      ...state,
      selection: {
        ...state.selection,
        extraRanges: [{ sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 }],
      },
    }));
    expect(
      applyCellStyleToSelection(foreignSelection, new History(), customCellStyleId('StableRepeat')),
    ).toBe(false);
    expect(foreignSelection.getState().format.formats.size).toBe(0);

    const merged = createSpreadsheetStore();
    mutators.upsertCustomCellStyle(merged, {
      id: 'custom:StableRepeat',
      label: 'StableRepeat',
      format: { italic: true },
    });
    mutators.mergeRange(merged, { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 });
    expect(
      applyCellStyleByName(
        merged,
        new History(),
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        customCellStyleId('StableRepeat'),
      ),
    ).toBe(false);
    expect(merged.getState().format.formats.size).toBe(0);
  });

  it('returns false for exactly 100001 cells without replacing history or repeat', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    mutators.upsertCustomCellStyle(store, {
      id: 'custom:TooWide',
      label: 'TooWide',
      format: { bold: true },
    });
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
        customCellStyleId('TooWide'),
      ),
    ).toBe(true);
    mutators.setRange(store, { sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 });
    expect(
      applyCellStyleByName(
        store,
        history,
        { sheet: 0, r0: 0, c0: 0, r1: 100_000, c1: 0 },
        customCellStyleId('TooWide'),
      ),
    ).toBe(false);
    expect(store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
    expect(history.canUndo()).toBe(true);
    expect(history.repeatLast()).toBe(true);
    expect(store.getState().format.formats.get('0:0:1')?.bold).toBe(true);
  });

  it('accepts a complete merge and rejects a partial merge through explicit APIs', () => {
    const complete = createSpreadsheetStore();
    const history = new History();
    const merge = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
    mutators.mergeRange(complete, merge);
    expect(applyCellStyleByName(complete, history, merge, customCellStyleId('Missing'))).toBe(
      false,
    );
    mutators.upsertCustomCellStyle(complete, {
      id: 'custom:Merged',
      label: 'Merged',
      format: { fill: '#c6efce' },
    });
    expect(applyCellStyleByName(complete, history, merge, customCellStyleId('Merged'))).toBe(true);
    expect(complete.getState().format.formats.get('0:0:0')?.fill).toBe('#c6efce');
    expect(complete.getState().format.formats.get('0:1:1')?.fill).toBe('#c6efce');

    const partial = createSpreadsheetStore();
    mutators.mergeRange(partial, merge);
    mutators.upsertCustomCellStyle(partial, {
      id: 'custom:Partial',
      label: 'Partial',
      format: { fill: '#ffc7ce' },
    });
    expect(
      applyCellStyleByName(
        partial,
        new History(),
        { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        customCellStyleId('Partial'),
      ),
    ).toBe(false);
    expect(partial.getState().format.formats.size).toBe(0);
  });
});
