import { afterEach, describe, expect, it } from 'vitest';

import { commitCellInput } from '../../src/commands/cell-input-commit.js';
import { MAX_COL } from '../../src/engine/address.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mutators } from '../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/index.js';

const LOCALES = [
  { locale: 'en', profile: 'windows-en_US', decimal: '1.5', formula: '=SUM(1.5,2)' },
  { locale: 'ja', profile: 'windows-ja_JP', decimal: '1.5', formula: '=SUM(1.5,2)' },
  { locale: 'de', profile: 'windows-de_DE', decimal: '1,5', formula: '=SUMME(1,5;2)' },
  { locale: 'fr', profile: 'windows-fr_FR', decimal: '1,5', formula: '=SOMME(1,5;2)' },
  { locale: 'zh', profile: 'windows-zh_CN', decimal: '1.5', formula: '=SUM(1.5,2)' },
  { locale: 'ko', profile: 'windows-ko_KR', decimal: '1.5', formula: '=SUM(1.5,2)' },
  { locale: 'th', profile: 'windows-th_TH', decimal: '1.5', formula: '=SUM(1.5,2)' },
] as const;

type LocaleFixture = (typeof LOCALES)[number];

const addr = (row: number, col: number) => ({ sheet: 0, row, col });

const expectTextAxis = (workbook: WorkbookHandle, row: number, col: number): void => {
  const effective = workbook.getEffectiveStyle(0, row, col);
  expect(effective).not.toBeNull();
  expect(effective?.source).toBe(row === 1 ? 1 : 2);
  expect(effective?.numFmtCode).toBe('@');
  expect(workbook.getCellXfIndex(0, row, col)).toBe(0);
};

const installTextAxes = (workbook: WorkbookHandle): void => {
  const base = workbook.getCellXf(0);
  expect(base).not.toBeNull();
  if (!base) throw new Error('missing native Normal XF');
  const textFormatId = workbook.addNumFmtCode('@');
  expect(textFormatId).toBeGreaterThanOrEqual(0);
  const textXf = workbook.addXfRecord({
    ...base,
    numFmtId: textFormatId,
    applyNumberFormat: true,
  });
  expect(textXf).toBeGreaterThanOrEqual(0);
  expect(workbook.setColumnStyle(0, 0, MAX_COL, textXf)).toBe(true);
  expect(workbook.setRowStyle(0, 1, textXf)).toBe(true);
};

const commit = (mounted: MountedStubSheet, row: number, col: number, raw: string): void => {
  const target = addr(row, col);
  const result = commitCellInput({
    store: mounted.instance.store,
    wb: mounted.workbook,
    addr: target,
    raw,
    origin: 'editor',
  });
  expect(result).toEqual({ status: 'applied', notice: null });
};

const disposeMounted = (mounted: MountedStubSheet | undefined): void => {
  mounted?.dispose();
  mounted?.workbook.dispose();
};

describe('native locale input and XLSX round-trip boundaries', () => {
  let mounted: MountedStubSheet | undefined;

  afterEach(() => {
    disposeMounted(mounted);
    mounted = undefined;
  });

  it.each(LOCALES)(
    'preserves typed text through blank native column and row Text XFs in $locale',
    async (fixture: LocaleFixture) => {
      const workbook = await WorkbookHandle.createDefault({ locale: fixture.locale });
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId(fixture.profile)).toBe(true);
      installTextAxes(workbook);

      // These cells must remain physically blank while the inherited styles
      // are checked. A pre-existing Normal XF would hide the row/column XF.
      expect(workbook.getCellXfIndex(0, 0, 0)).toBe(0);
      expect(workbook.getCellXfIndex(0, 1, 0)).toBe(0);
      expectTextAxis(workbook, 0, 0);
      expectTextAxis(workbook, 1, 0);

      mounted = await mountStubSheet({ workbook, locale: fixture.locale });
      expect(mounted.workbook.spreadsheetProfileId()).toBe(fixture.profile);
      expectTextAxis(mounted.workbook, 0, 0);
      expectTextAxis(mounted.workbook, 1, 0);

      commit(mounted, 0, 0, '00123');
      commit(mounted, 0, 1, fixture.decimal);
      commit(mounted, 1, 0, '=1+1');

      expect(mounted.workbook.getValue(addr(0, 0))).toEqual({ kind: 'text', value: '00123' });
      expect(mounted.workbook.getValue(addr(0, 1))).toEqual({
        kind: 'text',
        value: fixture.decimal,
      });
      expect(mounted.workbook.getValue(addr(1, 0))).toEqual({ kind: 'text', value: '=1+1' });
      expect(mounted.workbook.cellFormula(addr(0, 0))).toBeNull();
      expect(mounted.workbook.cellFormula(addr(0, 1))).toBeNull();
      expect(mounted.workbook.cellFormula(addr(1, 0))).toBeNull();

      // A direct General format is a cell XF and therefore outranks the
      // inherited text axis. The input must parse as a number in that cell.
      const explicitGeneral = addr(0, 2);
      mutators.setCellFormat(mounted.instance.store, explicitGeneral, {
        numFmt: { kind: 'general' },
      });
      const directStyle = mounted.workbook.getEffectiveStyle(
        explicitGeneral.sheet,
        explicitGeneral.row,
        explicitGeneral.col,
      );
      expect(directStyle?.source).toBe(0);
      expect(directStyle?.numFmtCode).toBe('General');
      commit(mounted, explicitGeneral.row, explicitGeneral.col, '00123');
      expect(mounted.workbook.getValue(explicitGeneral)).toEqual({ kind: 'number', value: 123 });
      expect(mounted.instance.store.getState().data.cells.get('0:0:2')?.value).toEqual({
        kind: 'number',
        value: 123,
      });
    },
  );

  it.each(LOCALES)(
    'saves canonical formula, localized value, and text then reloads native XLSX in $locale',
    async (fixture: LocaleFixture) => {
      const workbook = await WorkbookHandle.createDefault({ locale: fixture.locale });
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId(fixture.profile)).toBe(true);
      mounted = await mountStubSheet({ workbook, locale: fixture.locale });

      commit(mounted, 0, 0, fixture.formula);
      commit(mounted, 0, 1, fixture.decimal);
      const textCell = addr(0, 2);
      mutators.setCellFormat(mounted.instance.store, textCell, { numFmt: { kind: 'text' } });
      commit(mounted, textCell.row, textCell.col, '00123');

      expect(workbook.cellFormula(addr(0, 0))).toBe('=SUM(1.5,2)');
      expect(workbook.getValue(addr(0, 0))).toEqual({ kind: 'number', value: 3.5 });
      expect(workbook.getValue(addr(0, 1))).toEqual({ kind: 'number', value: 1.5 });
      expect(workbook.getValue(textCell)).toEqual({ kind: 'text', value: '00123' });

      const bytes = workbook.save();
      const reloaded = await WorkbookHandle.loadBytes(bytes);
      try {
        expect(reloaded.isStub).toBe(false);
        expect(reloaded.setSpreadsheetProfileId(fixture.profile)).toBe(true);
        expect(reloaded.cellFormula(addr(0, 0))).toBe('=SUM(1.5,2)');
        expect(reloaded.getValue(addr(0, 0))).toEqual({ kind: 'number', value: 3.5 });
        expect(reloaded.getValue(addr(0, 1))).toEqual({ kind: 'number', value: 1.5 });
        expect(reloaded.getValue(textCell)).toEqual({ kind: 'text', value: '00123' });
        expect(reloaded.getEffectiveStyle(0, 0, 2)?.numFmtCode).toBe('@');
        expect(reloaded.getFormulaForDisplay(reloaded.cellFormula(addr(0, 0)) ?? '')).toBe(
          fixture.formula,
        );
      } finally {
        reloaded.dispose();
      }
    },
  );

  it.each(LOCALES)(
    'keeps mounted history undo/redo data intact after switching from $locale',
    async (fixture: LocaleFixture) => {
      const workbook = await WorkbookHandle.createDefault({ locale: fixture.locale });
      expect(workbook.isStub).toBe(false);
      mounted = await mountStubSheet({ workbook, locale: fixture.locale });
      expect(mounted.workbook.setSpreadsheetProfileId(fixture.profile)).toBe(true);

      const target = addr(2, 0);
      commit(mounted, target.row, target.col, fixture.decimal);
      expect(mounted.workbook.getValue(target)).toEqual({ kind: 'number', value: 1.5 });
      expect(mounted.instance.store.getState().data.cells.get('0:2:0')?.value).toEqual({
        kind: 'number',
        value: 1.5,
      });
      expect(mounted.instance.history.canUndo()).toBe(true);

      const switchedTo = fixture.locale === 'en' ? 'ja' : 'en';
      expect(mounted.instance.setLocale(switchedTo)).toBe(true);
      expect(mounted.workbook.getValue(target)).toEqual({ kind: 'number', value: 1.5 });
      expect(mounted.instance.history.undo()).toBe(true);
      expect(mounted.workbook.getValue(target)).toEqual({ kind: 'blank' });
      expect(mounted.instance.store.getState().data.cells.get('0:2:0')).toMatchObject({
        value: { kind: 'blank' },
        formula: null,
      });
      expect(mounted.instance.history.redo()).toBe(true);
      expect(mounted.workbook.getValue(target)).toEqual({ kind: 'number', value: 1.5 });
      expect(mounted.instance.store.getState().data.cells.get('0:2:0')?.value).toEqual({
        kind: 'number',
        value: 1.5,
      });
    },
  );
});
