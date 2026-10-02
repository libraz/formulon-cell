import { describe, expect, it, vi } from 'vitest';
import {
  applyCellStyleToSelection,
  cellStyleGroupMatches,
} from '../../../src/commands/cell-styles.js';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import { cellStyleKeysByXfId } from '../../../src/engine/cell-style-writeback.js';
import {
  computeNamedCellStyles,
  type NamedCellStylesView,
} from '../../../src/engine/cell-styles-meta.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

type EngineStyle = NonNullable<ReturnType<NamedCellStylesView['getCellStyle']>>;

const fullStyle = (overrides: Partial<EngineStyle> = {}): EngineStyle => ({
  name: 'Normal',
  xfId: 0,
  builtinId: 0,
  iLevel: 0,
  hidden: false,
  customBuiltin: false,
  ...overrides,
});

const view = (styles: (EngineStyle | null)[]): NamedCellStylesView => ({
  cellStyleCount: () => styles.length,
  getCellStyle: (i: number) => styles[i] ?? null,
});

describe('computeNamedCellStyles', () => {
  it('returns an empty list when count is zero', () => {
    expect(computeNamedCellStyles(view([]))).toEqual([]);
  });

  it('emits one entry per non-hidden style with its index', () => {
    expect(
      computeNamedCellStyles(
        view([
          fullStyle({ name: 'Normal', xfId: 0, builtinId: 0 }),
          fullStyle({ name: 'Heading 1', xfId: 1, builtinId: 16, iLevel: 1 }),
        ]),
      ),
    ).toEqual([
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
        name: 'Heading 1',
        xfId: 1,
        builtinId: 16,
        iLevel: 1,
        customBuiltin: false,
      },
    ]);
  });

  it('filters out hidden built-ins (the gallery hides those)', () => {
    const names = computeNamedCellStyles(
      view([
        fullStyle({ name: 'Normal' }),
        fullStyle({ name: 'Comma [0]', hidden: true }),
        fullStyle({ name: 'Heading 1' }),
      ]),
    ).map((s) => s.name);
    expect(names).toEqual(['Normal', 'Heading 1']);
  });

  it('skips slots where getCellStyle returns null (out-of-range guard)', () => {
    expect(computeNamedCellStyles(view([null, fullStyle({ name: 'Total' })]))).toHaveLength(1);
  });

  it('preserves the original index even when earlier slots are filtered', () => {
    const out = computeNamedCellStyles(
      view([
        fullStyle({ name: 'Normal' }),
        fullStyle({ name: 'Hidden', hidden: true }),
        fullStyle({ name: 'Heading 1', xfId: 5 }),
      ]),
    );
    expect(out.map((s) => s.index)).toEqual([0, 2]);
  });
});

describe('named-style engine publication', () => {
  it('does not publish a pending named style before strict cell-XF work succeeds', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'good' });
      const setNamedCellStyle = vi.spyOn(wb, 'setNamedCellStyle');
      vi.spyOn(wb, 'setCellXfIndex').mockReturnValue(false);

      expect(() => syncCellFormatsToEngine(wb, store, 0, { strict: true })).toThrow();
      expect(setNamedCellStyle).not.toHaveBeenCalled();
    } finally {
      wb.dispose();
    }
  });

  it('maps a localized built-in catalog entry by ordinal on the real engine', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'good' });
      syncCellFormatsToEngine(wb, store, 0);
      const good = wb.getNamedCellStyles().find((style) => style.builtinId === 26);
      expect(good).toBeDefined();
      if (!good) return;
      expect(wb.setNamedCellStyle('良い', good.xfId, 26)).toBe(true);
      expect(cellStyleKeysByXfId(wb).get(good.xfId)).toBe('good');
    } finally {
      wb.dispose();
    }
  });

  it('does not let a same-name custom row impersonate a missing built-in ordinal', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeDefined();
      if (!normal) return;
      const customXfId = wb.addCellStyleXfRecord({ ...normal, fontIndex: 0 });
      expect(wb.setNamedCellStyle('Good', customXfId, null)).toBe(true);

      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'good' });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });

      const goodRows = wb.getNamedCellStyles().filter((style) => style.name === 'Good');
      const good = goodRows.find((style) => style.builtinId === 26);
      expect(good).toBeDefined();
      expect(good?.xfId).not.toBe(customXfId);
    } finally {
      wb.dispose();
    }
  });

  it('upserts the correct built-in ordinal on the default publication path too', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeDefined();
      if (!normal) return;
      const customXfId = wb.addCellStyleXfRecord({ ...normal, fontIndex: 0 });
      expect(wb.setNamedCellStyle('Good', customXfId, null)).toBe(true);

      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'good' });
      syncCellFormatsToEngine(wb, store, 0);

      const good = wb.getNamedCellStyles().find((style) => style.builtinId === 26);
      expect(good).toBeDefined();
      expect(good?.xfId).not.toBe(customXfId);
    } finally {
      wb.dispose();
    }
  });

  it('treats omitted and explicit off values as the same style defaults', () => {
    expect(cellStyleGroupMatches({ underline: false }, {}, 'font')).toBe(true);
    expect(
      cellStyleGroupMatches({ borders: { top: false as never } }, { borders: {} }, 'border'),
    ).toBe(true);
  });

  it('retains a resolved nonzero Normal style XF for unstyled cells', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeDefined();
      if (!normal) return;
      const defaultFont = wb.getFontRecord(0);
      expect(defaultFont).toBeDefined();
      if (!defaultFont) return;
      const normalFont = wb.addFontRecord({ ...defaultFont, name: 'Normal Native' });
      const resolvedNormal = wb.addCellStyleXfRecord({ ...normal, fontIndex: normalFont });
      expect(resolvedNormal).toBeGreaterThan(0);
      expect(wb.setNamedCellStyle('Normal', resolvedNormal, 0)).toBe(true);
      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { bold: true });
      syncCellFormatsToEngine(wb, store, 0);
      expect(wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1)?.xfId).toBe(resolvedNormal);
    } finally {
      wb.dispose();
    }
  });

  it('surfaces a later catalog failure with the residual real-engine entry', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const store = createSpreadsheetStore();
      mutators.upsertCustomCellStyle(store, {
        id: 'custom:Probe One',
        label: 'Probe One',
        format: { bold: true },
      });
      mutators.upsertCustomCellStyle(store, {
        id: 'custom:Probe Two',
        label: 'Probe Two',
        format: { italic: true },
      });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'Probe One' });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { cellStyle: 'Probe Two' });
      const original = wb.setNamedCellStyle.bind(wb);
      let calls = 0;
      vi.spyOn(wb, 'setNamedCellStyle').mockImplementation((name, xfId, builtinId) => {
        calls += 1;
        return calls < 2 ? original(name, xfId, builtinId) : false;
      });

      expect(() => syncCellFormatsToEngine(wb, store, 0, { strict: true })).toThrow(AggregateError);
      expect(calls).toBe(2);
      expect(wb.getNamedCellStyles().some((style) => style.name === 'Probe One')).toBe(true);
    } finally {
      wb.dispose();
    }
  });

  it('replaces a stale custom native XF when a removed font member differs', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const staleFont = wb.getFontRecord(0);
      const normalXf = wb.getCellStyleXf(0);
      expect(staleFont).toBeDefined();
      expect(normalXf).toBeDefined();
      if (!staleFont || !normalXf) return;
      const staleFontIndex = wb.addFontRecord({ ...staleFont, bold: true, italic: true });
      const staleXfId = wb.addCellStyleXfRecord({ ...normalXf, fontIndex: staleFontIndex });
      expect(staleFontIndex).toBeGreaterThanOrEqual(0);
      expect(staleXfId).toBeGreaterThanOrEqual(0);
      expect(wb.setNamedCellStyle('Probe', staleXfId, null)).toBe(true);

      const store = createSpreadsheetStore();
      mutators.upsertCustomCellStyle(store, {
        id: 'custom:Probe',
        label: 'Probe',
        format: { bold: true },
      });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'Probe' });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });

      const probe = wb.getNamedCellStyles().find((style) => style.name === 'Probe');
      expect(probe).toBeDefined();
      if (!probe) return;
      expect(wb.getFontRecord(wb.getCellStyleXf(probe.xfId)?.fontIndex ?? -1)?.italic).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('reports both an unrestorable new style and a failed prior-style rollback', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const normalXf = wb.getCellStyleXf(0);
      expect(normalXf).toBeDefined();
      if (!normalXf) return;
      const existingXfId = wb.addCellStyleXfRecord({ ...normalXf, fontIndex: 0 });
      expect(existingXfId).toBeGreaterThanOrEqual(0);
      expect(wb.setNamedCellStyle('Probe Replace', existingXfId, null)).toBe(true);

      const store = createSpreadsheetStore();
      for (const [col, style] of [
        ['Probe New', { bold: true }],
        ['Probe Replace', { italic: true }],
        ['Probe Fail', { underline: true }],
      ] as const) {
        mutators.upsertCustomCellStyle(store, {
          id: `custom:${col}`,
          label: col,
          format: style,
        });
        mutators.setCellFormat(
          store,
          { sheet: 0, row: 0, col: col === 'Probe New' ? 0 : col === 'Probe Replace' ? 1 : 2 },
          { cellStyle: col },
        );
      }
      const original = wb.setNamedCellStyle.bind(wb);
      let calls = 0;
      vi.spyOn(wb, 'setNamedCellStyle').mockImplementation((name, xfId, builtinId) => {
        calls += 1;
        if (calls === 3 || calls === 4) return false;
        return original(name, xfId, builtinId);
      });

      let error: unknown;
      try {
        syncCellFormatsToEngine(wb, store, 0, { strict: true });
      } catch (caught) {
        error = caught;
      }
      expect(error).toBeInstanceOf(AggregateError);
      const details = error instanceof AggregateError ? error.errors.map(String).join('\n') : '';
      expect(details).toContain('unavailable for new entries');
      expect(details).toContain('rollback failed: Probe Replace');
    } finally {
      wb.dispose();
    }
  });

  it('preserves native selectors through a named-style sync and literal overrides only one group', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.isStub).toBe(false);
      const normalXf = wb.getCellStyleXf(0);
      const defaultFont = wb.getFontRecord(0);
      const defaultFill = wb.getFillRecord(0);
      const defaultBorder = wb.getBorderRecord(0);
      expect(normalXf).toBeDefined();
      expect(defaultFont).toBeDefined();
      expect(defaultFill).toBeDefined();
      expect(defaultBorder).toBeDefined();
      if (!normalXf || !defaultFont || !defaultFill || !defaultBorder) return;

      const themedFontIndex = wb.addFontRecord({
        ...defaultFont,
        colorArgb: 0xff006100,
        color: {
          kind: 2,
          rgb: defaultFont.color?.rgb ?? 0,
          theme: 4,
          tint: 0.2,
          indexed: defaultFont.color?.indexed ?? 0,
        },
      });
      const indexedFillIndex = wb.addFillRecord({
        ...defaultFill,
        pattern: 1,
        fgArgb: 0xffc6efce,
        fg: {
          kind: 3,
          rgb: defaultFill.fg?.rgb ?? 0,
          theme: defaultFill.fg?.theme ?? 0,
          tint: defaultFill.fg?.tint ?? 0,
          indexed: 42,
        },
      });
      const currentBorderIndex = wb.addBorderRecord({
        ...defaultBorder,
        top: {
          ...defaultBorder.top,
          style: 2,
          colorArgb: 0,
          color: {
            kind: 4,
            rgb: 0,
            theme: defaultBorder.top.color?.theme ?? 0,
            tint: defaultBorder.top.color?.tint ?? 0,
            indexed: defaultBorder.top.color?.indexed ?? 0,
          },
        },
      });
      const styleXfId = wb.addCellStyleXfRecord({
        ...normalXf,
        fontIndex: themedFontIndex,
        fillIndex: indexedFillIndex,
      });
      expect(styleXfId).toBeGreaterThanOrEqual(0);
      expect(wb.setNamedCellStyle('Good', styleXfId, 26)).toBe(true);

      const numFmtId = wb.addNumFmtCode('0.00');
      const currentXfId = wb.addXfRecord({
        ...normalXf,
        fontIndex: defaultFont === null ? 0 : wb.addFontRecord(defaultFont),
        fillIndex: defaultFill === null ? 0 : wb.addFillRecord(defaultFill),
        borderIndex: currentBorderIndex,
        numFmtId,
        horizontalAlign: 3,
        verticalAlign: 1,
        wrapText: true,
        justifyLastLine: false,
        hasAlignment: true,
        hasHorizontalAlign: true,
        hasVerticalAlign: true,
        hasWrapText: true,
        hasJustifyLastLine: true,
        textRotation: 37,
        indent: 2,
        relativeIndent: -1,
        shrinkToFit: true,
        readingOrder: 2,
        xfId: styleXfId,
      });
      expect(currentXfId).toBeGreaterThanOrEqual(0);
      expect(wb.setCellXfIndex(0, 0, 0, currentXfId)).toBe(true);

      const store = createSpreadsheetStore();
      hydrateCellFormatsFromEngine(wb, store, 0);
      expect(applyCellStyleToSelection(store, null, 'good', { getWorkbook: () => wb })).toBe(true);
      mutators.setCellFormat(
        store,
        { sheet: 0, row: 0, col: 0 },
        {
          color: '#006100',
          fill: '#c6efce',
          borders: { top: { style: 'medium' } },
          align: 'right',
          vAlign: 'middle',
          wrap: true,
          numFmt: { kind: 'fixed', decimals: 2 },
        },
      );
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const firstCellXf = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(firstCellXf).toMatchObject({
        fontIndex: themedFontIndex,
        fillIndex: indexedFillIndex,
        borderIndex: currentBorderIndex,
        numFmtId,
        xfId: styleXfId,
        horizontalAlign: 3,
        verticalAlign: 1,
        wrapText: true,
        hasAlignment: true,
        hasHorizontalAlign: true,
        hasVerticalAlign: true,
        hasWrapText: true,
        hasJustifyLastLine: true,
        textRotation: 37,
        indent: 2,
        relativeIndent: -1,
        shrinkToFit: true,
        readingOrder: 2,
      });
      expect(wb.getBorderRecord(firstCellXf?.borderIndex ?? -1)?.top.color?.kind).toBe(4);
      expect(wb.getNumFmtCode(firstCellXf?.numFmtId ?? -1)).toBe('0.00');

      // The best-effort branch must retain the same raw selectors while the
      // source workbook still carries their fallback metadata.
      syncCellFormatsToEngine(wb, store, 0);
      const defaultCellXf = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(defaultCellXf).toMatchObject({
        borderIndex: currentBorderIndex,
        numFmtId,
        xfId: styleXfId,
        textRotation: 37,
        indent: 2,
        relativeIndent: -1,
        shrinkToFit: true,
        readingOrder: 2,
        hasAlignment: true,
        hasHorizontalAlign: true,
        hasVerticalAlign: true,
        hasWrapText: true,
        hasJustifyLastLine: true,
      });
      expect(wb.getFontRecord(defaultCellXf?.fontIndex ?? -1)?.color?.kind).toBe(2);
      expect(wb.getFillRecord(defaultCellXf?.fillIndex ?? -1)?.fg?.kind).toBe(3);
      expect(wb.getBorderRecord(defaultCellXf?.borderIndex ?? -1)?.top.color?.kind).toBe(4);

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      const loadedGood = reloaded.getNamedCellStyles().find((style) => style.builtinId === 26);
      expect(loadedGood).toBeDefined();
      if (!loadedGood) return;
      const loadedStyleXf = reloaded.getCellStyleXf(loadedGood.xfId);
      expect(loadedStyleXf).toBeDefined();
      if (!loadedStyleXf) return;
      expect(reloaded.getFontRecord(loadedStyleXf.fontIndex)?.color).toMatchObject({
        kind: 2,
        theme: 4,
        tint: 0.2,
      });
      expect(reloaded.getFillRecord(loadedStyleXf.fillIndex)?.fg).toMatchObject({
        kind: 3,
        indexed: 42,
      });
      const loadedCellXf = reloaded.getCellXf(reloaded.getCellXfIndex(0, 0, 0) ?? -1);
      expect(loadedCellXf).toMatchObject({
        fontIndex: loadedStyleXf.fontIndex,
        fillIndex: loadedStyleXf.fillIndex,
        borderIndex: currentBorderIndex,
        numFmtId,
        xfId: loadedGood.xfId,
        horizontalAlign: 3,
        verticalAlign: 1,
        wrapText: true,
        hasAlignment: true,
        hasHorizontalAlign: true,
        hasVerticalAlign: true,
        hasWrapText: true,
        hasJustifyLastLine: true,
        textRotation: 37,
        indent: 2,
        relativeIndent: -1,
        shrinkToFit: true,
        readingOrder: 2,
      });
      expect(reloaded.getFontRecord(loadedCellXf?.fontIndex ?? -1)?.color).toMatchObject({
        kind: 2,
        theme: 4,
        tint: 0.2,
      });
      expect(reloaded.getFillRecord(loadedCellXf?.fillIndex ?? -1)?.fg).toMatchObject({
        kind: 3,
        indexed: 42,
      });
      expect(reloaded.getBorderRecord(loadedCellXf?.borderIndex ?? -1)?.top.color).toMatchObject({
        kind: 4,
      });
      expect(reloaded.getNumFmtCode(loadedCellXf?.numFmtId ?? -1)).toBe('0.00');

      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { color: '#000000' });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const overriddenCellXf = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(overriddenCellXf?.fillIndex).toBe(indexedFillIndex);
      expect(overriddenCellXf?.fontIndex).not.toBe(themedFontIndex);
      expect(wb.getFontRecord(overriddenCellXf?.fontIndex ?? -1)?.colorArgb).toBe(0xff000000);
    } finally {
      reloaded?.dispose();
      wb.dispose();
    }
  });

  it('preserves named raw rotation 255 until an explicit degree reset', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const normal = wb.getCellStyleXf(0);
      expect(normal).toBeDefined();
      if (!normal) return;
      const raw255Style = wb.addCellStyleXfRecord({
        ...normal,
        hasAlignment: true,
        hasHorizontalAlign: true,
        hasVerticalAlign: true,
        hasWrapText: true,
        hasJustifyLastLine: true,
        textRotation: 255,
        indent: 0,
        shrinkToFit: false,
        readingOrder: 0,
      });
      expect(wb.setNamedCellStyle('Normal', raw255Style, 0)).toBe(true);

      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'normal' });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const preserved = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(preserved?.textRotation).toBe(255);

      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { rotation: 0 });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });
      const reset = wb.getCellXf(wb.getCellXfIndex(0, 0, 0) ?? -1);
      expect(reset?.textRotation).toBe(0);
      expect(reset?.textRotation).not.toBe(255);
    } finally {
      wb.dispose();
    }
  });
});
