import { describe, expect, it } from 'vitest';

import { WorkbookHandle } from '../../src/engine/workbook-handle.js';

type ProfileId = Parameters<WorkbookHandle['setSpreadsheetProfileId']>[0];
type Host = 'mac' | 'windows';
type Scalar = string | number;
type CellAddress = { sheet: number; row: number; col: number };
type AutoFilter = Parameters<WorkbookHandle['setAutoFilter']>[1];

const FORMULA_CELL: CellAddress = { sheet: 0, row: 0, col: 3 };

/**
 * These are measured expectations, kept beside the cases they cover:
 * - mac-365-* `collation.golden.json` and `collation_functions.golden.json`
 *   (`col_kana_hira_kata`, `col_fullwidth_latin_upper`, `col_accent_a_acute`,
 *   `colfn_sort_hw_voiced`, `colfn_unique_hw_voiced`, and the matching COUNTIF).
 * - mac-365-* `locale_tokens.golden.json` (`weekdays_aaaa`).
 * - win-365-ja_JP `collation.golden.json` and `collation_functions.golden.json`
 *   (same case ids) for every Windows profile, and
 *   `countif_kana_folding_probes.golden.json` for the legacy COUNTIF byte behaviour.
 */
const profiles = [
  {
    id: 'mac-en_US',
    host: 'mac',
    weekday: 'Tuesday',
    collation: { kana: 'EQ', width: 'EQ', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 2,
    filterMatch: [true, true, false],
  },
  {
    id: 'mac-ja_JP',
    host: 'mac',
    weekday: '火曜日',
    collation: { kana: 'EQ', width: 'EQ', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 2,
    filterMatch: [true, true, false],
  },
  {
    id: 'mac-de_DE',
    host: 'mac',
    weekday: 'Dienstag',
    collation: { kana: 'EQ', width: 'EQ', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 2,
    filterMatch: [true, true, false],
  },
  {
    id: 'mac-fr_FR',
    host: 'mac',
    weekday: '2024',
    collation: { kana: 'EQ', width: 'EQ', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 2,
    filterMatch: [true, true, false],
  },
  {
    id: 'mac-zh_CN',
    host: 'mac',
    weekday: '星期二',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
  {
    id: 'mac-ko_KR',
    host: 'mac',
    weekday: '화요일',
    collation: { kana: 'EQ', width: 'EQ', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 2,
    filterMatch: [true, true, false],
  },
  {
    id: 'mac-th_TH',
    host: 'mac',
    weekday: 'วันอังคาร',
    collation: { kana: 'EQ', width: 'EQ', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 2,
    filterMatch: [true, true, false],
  },
  {
    id: 'windows-en_US',
    host: 'windows',
    weekday: 'Tuesday',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
  {
    id: 'windows-ja_JP',
    host: 'windows',
    weekday: '火曜日',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
  {
    id: 'windows-de_DE',
    host: 'windows',
    weekday: 'Dienstag',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
  {
    id: 'windows-fr_FR',
    host: 'windows',
    weekday: '2024',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
  {
    id: 'windows-zh_CN',
    host: 'windows',
    weekday: '星期二',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
  {
    id: 'windows-ko_KR',
    host: 'windows',
    weekday: '화요일',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
  {
    id: 'windows-th_TH',
    host: 'windows',
    weekday: 'วันอังคาร',
    collation: { kana: 'GT', width: 'GT', accent: 'LT', sort: 'ｶﾞ', unique: 1 },
    countifKana: 1,
    filterMatch: [true, false, false],
  },
] as const satisfies ReadonlyArray<{
  id: ProfileId;
  host: Host;
  weekday: string;
  collation: {
    kana: string;
    width: string;
    accent: string;
    sort: string;
    unique: number;
  };
  countifKana: number;
  filterMatch: readonly boolean[];
}>;

const scalarValue = (workbook: WorkbookHandle, formula: string): Scalar => {
  workbook.setFormula(FORMULA_CELL, formula);
  workbook.recalc();
  const value = workbook.getValue(FORMULA_CELL);
  if (value.kind === 'text' || value.kind === 'number') return value.value;
  throw new Error(`Expected a scalar text/number result, got ${value.kind}`);
};

const seedKanaCountif = (workbook: WorkbookHandle): void => {
  workbook.setText({ sheet: 0, row: 0, col: 0 }, 'あ');
  workbook.setText({ sheet: 0, row: 1, col: 0 }, 'ア');
  workbook.setText({ sheet: 0, row: 2, col: 0 }, 'い');
};

const valueFilter = (value: string): AutoFilter =>
  ({
    range: { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 0 },
    columns: [
      {
        colId: 0,
        hiddenButton: false,
        showButton: true,
        kind: 1,
        filterBlank: false,
        values: [value],
        dateGroups: [],
        customAnd: false,
        customCount: 0,
        op1: 0,
        val1: '',
        op2: 0,
        val2: '',
        top: false,
        percent: false,
        hasFilterVal: false,
        topVal: 0,
        filterVal: 0,
        dynamicType: 0,
        hasDynVal: false,
        hasDynMaxVal: false,
        dynVal: 0,
        dynMaxVal: 0,
        valIso: '',
        maxValIso: '',
        dxfId: 0xffffffff,
        cellColor: false,
        iconSet: 0,
        iconId: 0,
        hasIconId: false,
      },
    ],
    sort: null,
  }) as AutoFilter;

describe('integration: formulon 0.14 locale-aware engine refresh', () => {
  it.each(profiles)('%s keeps measured collation and scalar array results', async (profile) => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId(profile.id)).toBe(true);
      expect(scalarValue(workbook, '=IF("あ"<"ア","LT",IF("あ"="ア","EQ","GT"))')).toBe(
        profile.collation.kana,
      );
      expect(scalarValue(workbook, '=IF("Ａ"<"A","LT",IF("Ａ"="A","EQ","GT"))')).toBe(
        profile.collation.width,
      );
      expect(scalarValue(workbook, '=IF("a"<"á","LT",IF("a"="á","EQ","GT"))')).toBe(
        profile.collation.accent,
      );
      expect(scalarValue(workbook, '=INDEX(SORT({"ｶﾞ";"ガ"}),1)')).toBe(profile.collation.sort);
      expect(scalarValue(workbook, '=ROWS(UNIQUE({"ｶﾞ";"ガ"}))')).toBe(profile.collation.unique);
    } finally {
      workbook.dispose();
    }
  });

  it.each(profiles)('%s applies locale-aware COUNTIF equality', async (profile) => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId(profile.id)).toBe(true);
      seedKanaCountif(workbook);
      expect(scalarValue(workbook, '=COUNTIF(A1:A3,"あ")')).toBe(profile.countifKana);
    } finally {
      workbook.dispose();
    }
  });

  it.each(profiles)('%s evaluates typed AutoFilter kana criteria', async (profile) => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.capabilities.typedAutoFilter).toBe(true);
      expect(workbook.setSpreadsheetProfileId(profile.id)).toBe(true);
      seedKanaCountif(workbook);
      workbook.setText({ sheet: 0, row: 0, col: 0 }, '見出し');
      workbook.setText({ sheet: 0, row: 1, col: 0 }, 'あ');
      workbook.setText({ sheet: 0, row: 2, col: 0 }, 'ア');
      workbook.setText({ sheet: 0, row: 3, col: 0 }, 'い');
      expect(workbook.setAutoFilter(0, valueFilter('あ'))).toBe(true);
      expect(workbook.applyAutoFilter(0)).toBe(true);
      expect(workbook.evaluateAutoFilter(0)).toMatchObject({
        firstRow: 1,
        match: profile.filterMatch,
      });
    } finally {
      workbook.dispose();
    }
  });

  it.each(profiles)('%s uses the measured active-profile weekday token', async (profile) => {
    const workbook = await WorkbookHandle.createDefault();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.setSpreadsheetProfileId(profile.id)).toBe(true);
      expect(scalarValue(workbook, '=TEXT(DATE(2024,3,5),"aaaa")')).toBe(profile.weekday);
    } finally {
      workbook.dispose();
    }
  });

  it.each(profiles)(
    '%s formats an explicit de-DE weekday tag through both native paths',
    async (profile) => {
      const workbook = await WorkbookHandle.createDefault();
      try {
        expect(workbook.isStub).toBe(false);
        expect(workbook.capabilities.displayText).toBe(true);
        expect(workbook.setSpreadsheetProfileId(profile.id)).toBe(true);
        expect(
          workbook.formatValueForDisplay({ kind: 'number', value: 45356 }, '[$-de-DE]dddd'),
        ).toBe('Dienstag');
        const evaluated = workbook.evaluateFormulaText(
          FORMULA_CELL,
          '=TEXT(DATE(2024,3,5),"[$-de-DE]dddd")',
        );
        expect(evaluated.status.ok).toBe(true);
        // The measured locale_tokens lcid_names_407 golden keeps the token
        // literal under the German/French active format dialect. Native display
        // formatting above already proves the explicit de-DE override itself.
        const formulaExpected =
          profile.id.endsWith('-de_DE') || profile.id.endsWith('-fr_FR') ? 'dddd' : 'Dienstag';
        expect(evaluated.value.text).toBe(formulaExpected);
      } finally {
        workbook.dispose();
      }
    },
  );

  it.each([
    ['mac-en_US', 'mac-de_DE', 'Tuesday', 'Dienstag'],
    ['windows-en_US', 'windows-de_DE', 'Tuesday', 'Dienstag'],
  ] as const)(
    'recalculates cached TEXT after %s → %s',
    async (before, after, beforeText, afterText) => {
      const workbook = await WorkbookHandle.createDefault();
      try {
        expect(workbook.isStub).toBe(false);
        expect(workbook.setSpreadsheetProfileId(before)).toBe(true);
        workbook.setFormula(FORMULA_CELL, '=TEXT(DATE(2024,3,5),"aaaa")');
        expect(workbook.getValue(FORMULA_CELL)).toEqual({ kind: 'text', value: beforeText });
        const storedFormula = workbook.cellFormula(FORMULA_CELL);
        expect(workbook.setSpreadsheetProfileId(after)).toBe(true);
        workbook.recalc();
        expect(workbook.getValue(FORMULA_CELL)).toEqual({ kind: 'text', value: afterText });
        expect(workbook.cellFormula(FORMULA_CELL)).toBe(storedFormula);
      } finally {
        workbook.dispose();
      }
    },
  );
});
