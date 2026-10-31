import { describe, expect, it, vi } from 'vitest';
import {
  applyCellStyleToSelection,
  cellStyleGroupMatches,
} from '../../../src/commands/cell-styles.js';
import {
  hydrateCellFormatsFromEngine,
  syncCellFormatsToEngine,
} from '../../../src/engine/cell-format-sync.js';
import {
  cellStyleKeysByXfId,
  planNamedCellStylesToEngine,
} from '../../../src/engine/cell-style-writeback.js';
import {
  computeNamedCellStyles,
  type NamedCellStylesView,
} from '../../../src/engine/cell-styles-meta.js';
import { argbToCssColor } from '../../../src/engine/format-writeback.js';
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

const importedNativeStyleFixture = async (hidden: boolean, customBuiltin = true) => {
  const wb = await WorkbookHandle.createDefault();
  const normalXf = wb.getCellStyleXf(0);
  const defaultFont = wb.getFontRecord(0);
  if (!normalXf || !defaultFont) throw new Error('Default style records are unavailable');
  const themeFontIndex = wb.addFontRecord({
    ...defaultFont,
    colorArgb: 0xff4472c4,
    color: { kind: 2, rgb: 0xff4472c4, theme: 10, tint: 0, indexed: 0 },
  });
  const styleXfId = wb.addCellStyleXfRecord({ ...normalXf, fontIndex: themeFontIndex });
  if (themeFontIndex < 0 || styleXfId < 0) throw new Error('Could not create themed style records');
  if (
    !wb.setNamedCellStyleRecord({
      name: 'LienWeb',
      xfId: styleXfId,
      builtinId: 8,
      iLevel: 0,
      hidden: false,
      customBuiltin,
    })
  ) {
    throw new Error('Could not create localized native style');
  }
  const cellXfId = wb.addXfRecord({ ...normalXf, fontIndex: themeFontIndex, xfId: styleXfId });
  if (cellXfId < 0 || !wb.setCellXfIndex(0, 0, 0, cellXfId)) {
    throw new Error('Could not format native-style fixture cell');
  }

  const store = createSpreadsheetStore();
  hydrateCellFormatsFromEngine(wb, store, 0);
  mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'LienWeb' });
  const styleIndex = Array.from({ length: wb.cellStyleCount() }, (_, index) => index).find(
    (index) => wb.getCellStyle(index)?.name === 'LienWeb',
  );
  if (styleIndex === undefined) throw new Error('Localized native style was not cataloged');
  const record = wb.getCellStyle(styleIndex);
  if (!record || !wb.setNamedCellStyleRecord({ ...record, hidden })) {
    throw new Error('Could not set localized native style visibility');
  }

  return { wb, store, styleXfId, themeFontIndex, styleIndex, record: { ...record, hidden } };
};

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
  it('publishes the strict Normal seed when the style-XF table is empty', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      vi.spyOn(wb, 'cellStyleCount').mockReturnValue(0);
      vi.spyOn(wb, 'cellStyleXfCount').mockReturnValue(0);
      vi.spyOn(wb, 'getCellStyle').mockReturnValue(null);
      const setNamedCellStyleRecord = vi.spyOn(wb, 'setNamedCellStyleRecord').mockReturnValue(true);
      const resolveStyleXf = vi.fn(() => 0);
      const plan = planNamedCellStylesToEngine(
        wb,
        createSpreadsheetStore().getState(),
        resolveStyleXf,
      );

      expect(plan.styleXfIds.get('normal')).toBe(0);
      expect(resolveStyleXf).toHaveBeenCalledExactlyOnceWith({});
      expect(setNamedCellStyleRecord).not.toHaveBeenCalled();

      plan.commit();

      expect(setNamedCellStyleRecord).toHaveBeenCalledExactlyOnceWith({
        name: 'Normal',
        xfId: 0,
        builtinId: 0,
      });
    } finally {
      wb.dispose();
    }
  });

  it.each([
    { path: 'strict staging', strict: true, hidden: false },
    { path: 'ordinary sync', strict: false, hidden: false },
    { path: 'strict staging', strict: true, hidden: true },
    { path: 'ordinary sync', strict: false, hidden: true },
  ])(
    'keeps an unchanged imported themed native built-in during $path (hidden: $hidden)',
    async ({ strict, hidden }) => {
      const fixture = await importedNativeStyleFixture(hidden);
      const { wb, store, styleXfId, themeFontIndex, styleIndex, record } = fixture;
      try {
        expect(wb.isStub).toBe(false);
        const originalStyleCount = wb.cellStyleCount();
        const originalStyleXfCount = wb.cellStyleXfCount();
        const originalXf = wb.getCellStyleXf(styleXfId);
        const originalFont = wb.getFontRecord(themeFontIndex);
        const setNamedCellStyleRecord = vi.spyOn(wb, 'setNamedCellStyleRecord');
        const addCellStyleXfRecord = vi.spyOn(wb, 'addCellStyleXfRecord');
        expect(originalXf).toBeDefined();
        expect(originalFont?.color).toMatchObject({ kind: 2, theme: 10 });

        syncCellFormatsToEngine(wb, store, 0, { strict });

        expect(wb.cellStyleCount()).toBe(originalStyleCount);
        expect(wb.cellStyleXfCount()).toBe(originalStyleXfCount);
        expect(setNamedCellStyleRecord).not.toHaveBeenCalled();
        expect(addCellStyleXfRecord).not.toHaveBeenCalled();
        expect(wb.getCellStyle(styleIndex)).toEqual(record);
        expect(wb.getCellStyleXf(styleXfId)).toEqual(originalXf);
        expect(wb.getFontRecord(themeFontIndex)?.color).toEqual(originalFont?.color);
      } finally {
        wb.dispose();
      }
    },
  );

  it.each([
    { path: 'strict staging', strict: true, hidden: false },
    { path: 'ordinary sync', strict: false, hidden: false },
    { path: 'strict staging', strict: true, hidden: true },
    { path: 'ordinary sync', strict: false, hidden: true },
  ])(
    'keeps imported native built-in metadata after a font-size edit during $path (hidden: $hidden)',
    async ({ strict, hidden }) => {
      const fixture = await importedNativeStyleFixture(hidden);
      const { wb, store, styleXfId, styleIndex } = fixture;
      try {
        expect(wb.isStub).toBe(false);
        const originalStyleCount = wb.cellStyleCount();
        const imported = store
          .getState()
          .format.customCellStyles?.find((style) => style.label === 'LienWeb');
        expect(imported).toBeDefined();
        if (!imported) throw new Error('Imported native style is unavailable');
        const editedFontSize = (imported.format.fontSize ?? 16) + 4;
        mutators.upsertCustomCellStyle(store, {
          ...imported,
          format: { ...imported.format, fontSize: editedFontSize },
        });

        syncCellFormatsToEngine(wb, store, 0, { strict });

        const updated = wb.getCellStyle(styleIndex);
        expect(updated).toMatchObject({
          name: 'LienWeb',
          builtinId: 8,
          iLevel: 0,
          hidden,
          customBuiltin: true,
        });
        expect(updated?.xfId).not.toBe(styleXfId);
        expect(wb.cellStyleCount()).toBe(originalStyleCount);
        const updatedXf = wb.getCellStyleXf(updated?.xfId ?? -1);
        expect(updatedXf).toBeDefined();
        const updatedFont = wb.getFontRecord(updatedXf?.fontIndex ?? -1);
        expect(updatedFont?.size).toBeCloseTo(editedFontSize * 0.75, 3);
      } finally {
        wb.dispose();
      }
    },
  );

  it('marks an edited native built-in custom and restores customBuiltin false on rollback', async () => {
    const fixture = await importedNativeStyleFixture(false, false);
    const { wb, store, styleIndex } = fixture;
    try {
      const previous = wb.getCellStyle(styleIndex);
      expect(previous).toMatchObject({ name: 'LienWeb', builtinId: 8, customBuiltin: false });
      const imported = store
        .getState()
        .format.customCellStyles?.find((style) => style.label === 'LienWeb');
      if (!previous || !imported) throw new Error('Imported native style fixture is incomplete');

      mutators.upsertCustomCellStyle(store, {
        ...imported,
        format: { ...imported.format, fontSize: (imported.format.fontSize ?? 16) + 4 },
      });
      mutators.upsertCustomCellStyle(store, {
        id: 'custom:Later Failure',
        label: 'Later Failure',
        format: { bold: true },
      });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { cellStyle: 'Later Failure' });

      const original = wb.setNamedCellStyleRecord.bind(wb);
      const records: Parameters<typeof wb.setNamedCellStyleRecord>[0][] = [];
      vi.spyOn(wb, 'setNamedCellStyleRecord').mockImplementation((record) => {
        records.push(record);
        return record.name === 'Later Failure' ? false : original(record);
      });

      expect(() => syncCellFormatsToEngine(wb, store, 0, { strict: true })).toThrow(AggregateError);

      const nativeWrites = records.filter((record) => record.name === 'LienWeb');
      expect(nativeWrites).toHaveLength(2);
      expect(nativeWrites[0]).toMatchObject({
        builtinId: 8,
        iLevel: 0,
        hidden: false,
        customBuiltin: true,
      });
      expect(nativeWrites[1]).toEqual(previous);
      expect(wb.getCellStyle(styleIndex)).toEqual(previous);
      expect(wb.getNamedCellStyles().some((style) => style.name === 'Later Failure')).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it('does not publish a pending named style before strict cell-XF work succeeds', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const store = createSpreadsheetStore();
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'good' });
      const setNamedCellStyleRecord = vi.spyOn(wb, 'setNamedCellStyleRecord');
      vi.spyOn(wb, 'setCellXfIndex').mockReturnValue(false);

      expect(() => syncCellFormatsToEngine(wb, store, 0, { strict: true })).toThrow();
      expect(setNamedCellStyleRecord).not.toHaveBeenCalled();
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

  it('publishes and round-trips custom and built-in 53 style records', async () => {
    const wb = await WorkbookHandle.createDefault();
    let reloaded: WorkbookHandle | null = null;
    try {
      expect(wb.isStub).toBe(false);
      const store = createSpreadsheetStore();
      mutators.upsertCustomCellStyle(store, {
        id: 'custom:Probe Custom',
        label: 'Probe Custom',
        format: { bold: true },
      });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { cellStyle: 'Probe Custom' });
      mutators.setCellFormat(store, { sheet: 0, row: 0, col: 1 }, { cellStyle: 'explanatoryText' });
      syncCellFormatsToEngine(wb, store, 0, { strict: true });

      const publishedCustom = wb
        .getNamedCellStyles()
        .find((style) => style.name === 'Probe Custom');
      expect(publishedCustom).toMatchObject({ name: 'Probe Custom', builtinId: 0xffffffff });
      const publishedBuiltin = wb.getNamedCellStyles().find((style) => style.builtinId === 53);
      expect(publishedBuiltin).toBeDefined();
      if (!publishedCustom || !publishedBuiltin) return;

      reloaded = await WorkbookHandle.loadBytes(wb.save());
      expect(reloaded.isStub).toBe(false);
      expect(
        reloaded.getNamedCellStyles().find((style) => style.name === 'Probe Custom'),
      ).toMatchObject({
        name: 'Probe Custom',
        builtinId: 0xffffffff,
      });
      expect(reloaded.getNamedCellStyles().find((style) => style.builtinId === 53)).toMatchObject({
        name: publishedBuiltin.name,
        builtinId: 53,
      });
    } finally {
      reloaded?.dispose();
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

  it('removes a newly published entry after a later catalog failure', async () => {
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
      const original = wb.setNamedCellStyleRecord.bind(wb);
      let calls = 0;
      vi.spyOn(wb, 'setNamedCellStyleRecord').mockImplementation((record) => {
        calls += 1;
        return calls < 2 ? original(record) : false;
      });

      expect(() => syncCellFormatsToEngine(wb, store, 0, { strict: true })).toThrow(AggregateError);
      expect(calls).toBe(2);
      expect(wb.getNamedCellStyles().some((style) => style.name === 'Probe One')).toBe(false);
      expect(wb.getNamedCellStyles().some((style) => style.name === 'Probe Two')).toBe(false);
    } finally {
      wb.dispose();
    }
  });

  it.each([
    { originalName: 'Hidden Probe', registeredName: 'Hidden Probe', hidden: true },
    { originalName: 'case probe', registeredName: 'Case Probe', hidden: false },
  ])(
    'restores $originalName after a later catalog failure',
    async ({ originalName, registeredName, hidden }) => {
      const wb = await WorkbookHandle.createDefault();
      try {
        const prior = {
          name: originalName,
          xfId: 0,
          builtinId: 1,
          iLevel: 4,
          hidden,
          customBuiltin: true,
        };
        expect(wb.setNamedCellStyleRecord(prior)).toBe(true);
        expect(wb.getNamedCellStyles().some((style) => style.name === prior.name)).toBe(!hidden);
        const store = createSpreadsheetStore();
        for (const [col, name] of [registeredName, 'Later Failure'].entries()) {
          mutators.upsertCustomCellStyle(store, {
            id: `custom:${name}`,
            label: name,
            format: { bold: true },
          });
          mutators.setCellFormat(store, { sheet: 0, row: 0, col }, { cellStyle: name });
        }
        const original = wb.setNamedCellStyleRecord.bind(wb);
        vi.spyOn(wb, 'setNamedCellStyleRecord').mockImplementation((record) =>
          record.name === 'Later Failure' ? false : original(record),
        );
        expect(() => syncCellFormatsToEngine(wb, store, 0, { strict: true })).toThrow(
          AggregateError,
        );
        const styles = Array.from({ length: wb.cellStyleCount() }, (_, index) =>
          wb.getCellStyle(index),
        );
        expect(styles.find((style) => style?.name === prior.name)).toMatchObject(prior);
        expect(styles.some((style) => style?.name === 'Later Failure')).toBe(false);
        if (registeredName !== originalName) {
          expect(styles.some((style) => style?.name === registeredName)).toBe(false);
        }
      } finally {
        wb.dispose();
      }
    },
  );

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

  it('removes a new style and reports a failed prior-style rollback', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      expect(wb.isStub).toBe(false);
      const normalXf = wb.getCellStyleXf(0);
      expect(normalXf).toBeDefined();
      if (!normalXf) return;
      const existingXfId = wb.addCellStyleXfRecord({ ...normalXf, fontIndex: 0 });
      expect(existingXfId).toBeGreaterThanOrEqual(0);
      expect(
        wb.setNamedCellStyleRecord({
          name: 'Probe Replace',
          xfId: existingXfId,
          builtinId: 1,
          iLevel: 4,
          hidden: false,
          customBuiltin: true,
        }),
      ).toBe(true);

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
      const originalRecord = wb.setNamedCellStyleRecord.bind(wb);
      let rollbackRecord: Parameters<typeof wb.setNamedCellStyleRecord>[0] | undefined;
      let publicationCalls = 0;
      vi.spyOn(wb, 'setNamedCellStyleRecord').mockImplementation((record) => {
        if (record.name === 'Probe Replace' && record.xfId === existingXfId) {
          rollbackRecord = record;
          return false;
        }
        publicationCalls += 1;
        if (publicationCalls === 3) return false;
        return originalRecord(record);
      });

      let error: unknown;
      try {
        syncCellFormatsToEngine(wb, store, 0, { strict: true });
      } catch (caught) {
        error = caught;
      }
      expect(error).toBeInstanceOf(AggregateError);
      const details = error instanceof AggregateError ? error.errors.map(String).join('\n') : '';
      expect(details).toContain('rollback failed: Probe Replace');
      expect(details).not.toContain('unavailable for new entries');
      expect(wb.getNamedCellStyles().some((style) => style.name === 'Probe New')).toBe(false);
      expect(rollbackRecord).toMatchObject({
        name: 'Probe Replace',
        builtinId: 1,
        iLevel: 4,
        hidden: false,
        customBuiltin: true,
      });
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
      const resolvedFont = wb.getResolvedFontRecord(themedFontIndex);
      const resolvedFill = wb.getResolvedFillRecord(indexedFillIndex);
      expect(resolvedFont).not.toBeNull();
      expect(resolvedFill).not.toBeNull();
      if (!resolvedFont || !resolvedFill) return;
      const hydratedFontColor = argbToCssColor(resolvedFont.colorArgb);
      const hydratedFillColor = argbToCssColor(resolvedFill.fgArgb);

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
          color: hydratedFontColor,
          fill: hydratedFillColor,
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
