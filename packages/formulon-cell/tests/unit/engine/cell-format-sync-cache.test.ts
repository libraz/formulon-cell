import { describe, expect, it } from 'vitest';
import { syncCellFormatsToEngine } from '../../../src/engine/cell-format-sync.js';
import type {
  Addr,
  BorderRecord,
  CellXf,
  ColorSpec,
  FillRecord,
  FontRecord,
  PhoneticProperties,
  PhoneticRun,
} from '../../../src/engine/types.js';
import { addrKey, type WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { type CellFormat, createSpreadsheetStore } from '../../../src/store/store.js';

const color = (rgb = 0xff000000): ColorSpec => ({
  kind: 1,
  rgb,
  theme: 0,
  tint: 0,
  indexed: 0,
});

const font = (patch: Partial<FontRecord> = {}): FontRecord => ({
  name: 'Calibri',
  size: 11,
  bold: false,
  italic: false,
  strike: false,
  hasBold: false,
  hasItalic: false,
  hasStrike: false,
  outline: false,
  shadow: false,
  condense: false,
  extend: false,
  hasOutline: false,
  hasShadow: false,
  hasCondense: false,
  hasExtend: false,
  underline: 0,
  vertAlign: 0,
  hasFamily: false,
  family: 0,
  hasCharset: false,
  charset: 0,
  hasSize: true,
  scheme: 2,
  colorArgb: 0xff000000,
  color: color(),
  ...patch,
});

const fill = (): FillRecord => ({
  pattern: 0,
  fgArgb: 0,
  bgArgb: 0,
  fg: color(0),
  bg: color(0),
});

const side = (): BorderRecord['left'] => ({ style: 0, colorArgb: 0, color: color(0) });

const border = (): BorderRecord => ({
  left: side(),
  right: side(),
  top: side(),
  bottom: side(),
  diagonal: side(),
  diagonalUp: false,
  diagonalDown: false,
});

const sourceXf = (patch: Partial<CellXf> = {}): CellXf => ({
  fontIndex: 1,
  fillIndex: 1,
  borderIndex: 1,
  numFmtId: 0,
  horizontalAlign: 0,
  verticalAlign: 2,
  wrapText: false,
  ...patch,
});

interface HarnessOptions {
  readonly indices?: ReadonlyMap<string, number | null>;
  readonly nullIndexes?: ReadonlySet<string>;
  readonly xfs?: ReadonlyMap<number, CellXf>;
  readonly missingXfIndexes?: ReadonlySet<number>;
  readonly addFontResult?: (record: FontRecord, call: number, allocated: number) => number;
  readonly getBorderRecordResult?: (
    index: number,
    call: number,
    stored: BorderRecord | null,
  ) => BorderRecord | null;
  readonly addXfResult?: (record: CellXf, call: number) => number;
  readonly setXfResult?: (addr: Addr, call: number) => boolean;
  readonly phonetic?: boolean;
  readonly effectiveStyle?: boolean;
  readonly effectiveStyles?: ReadonlyMap<
    string,
    { readonly source: number; readonly xfIndex: number }
  >;
}

interface Harness {
  readonly wb: WorkbookHandle;
  readonly counts: {
    getIndex: number;
    getXf: number;
    getEffectiveStyle: number;
    addFont: number;
    addFill: number;
    addBorder: number;
    addNumFmt: number;
    addXf: number;
    setXf: number;
    getBorder: number;
    setRuns: number;
    setPhoneticProperties: number;
    setShowPhonetic: number;
  };
  readonly addedXfs: CellXf[];
  readonly assignedXfs: number[];
  readonly setPhoneticRuns: Array<{ addr: Addr; runs: readonly PhoneticRun[] }>;
  readonly setProperties: Array<{ addr: Addr; properties: PhoneticProperties }>;
  readonly defaultFont: { current: Pick<FontRecord, 'name' | 'size' | 'scheme'> | null };
  readonly xfs: Map<number, CellXf>;
  readonly fonts: Map<number, FontRecord>;
  readonly borders: Map<number, BorderRecord>;
  readonly indices: Map<string, number>;
}

const makeHarness = (options: HarnessOptions = {}): Harness => {
  const counts = {
    getIndex: 0,
    getXf: 0,
    getEffectiveStyle: 0,
    addFont: 0,
    addFill: 0,
    addBorder: 0,
    addNumFmt: 0,
    addXf: 0,
    setXf: 0,
    getBorder: 0,
    setRuns: 0,
    setPhoneticProperties: 0,
    setShowPhonetic: 0,
  };
  const addedXfs: CellXf[] = [];
  const assignedXfs: number[] = [];
  const setPhoneticRuns: Array<{ addr: Addr; runs: readonly PhoneticRun[] }> = [];
  const setProperties: Array<{ addr: Addr; properties: PhoneticProperties }> = [];
  const defaultFont: {
    current: Pick<FontRecord, 'name' | 'size' | 'scheme'> | null;
  } = { current: { name: 'Calibri', size: 11, scheme: 2 } };
  const indices = new Map<string, number>();
  for (const [key, index] of options.indices ?? []) {
    if (index !== null) indices.set(key, index);
  }
  const xfs = new Map<number, CellXf>([[1, sourceXf()], ...(options.xfs ?? [])]);
  const fonts = new Map<number, FontRecord>([
    [1, font()],
    [2, font()],
    [3, font()],
  ]);
  const fills = new Map<number, FillRecord>([[1, fill()]]);
  const borders = new Map<number, BorderRecord>([[1, border()]]);
  const numFmts = new Map<number, string>();
  let nextFont = 100;
  let nextFill = 100;
  let nextBorder = 100;
  let nextNumFmt = 200;
  let nextXf = 100;
  const showPhonetic = new Map<string, boolean>();

  const wb = {
    capabilities: {
      cellFormatting: true,
      ...(options.effectiveStyle ? { effectiveStyle: true } : {}),
      ...(options.phonetic
        ? { phoneticRuns: true, phoneticProperties: true, showPhonetic: true }
        : {}),
    },
    get workbookDefaultFont() {
      return defaultFont.current;
    },
    getCellXfIndex(sheet: number, row: number, col: number): number | null {
      counts.getIndex += 1;
      const key = [sheet, row, col].join(':');
      return options.nullIndexes?.has(key) ? null : (indices.get(key) ?? 1);
    },
    setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): boolean {
      counts.setXf += 1;
      assignedXfs.push(xfIndex);
      const addr = { sheet, row, col };
      const succeeded = options.setXfResult?.(addr, counts.setXf) ?? true;
      if (succeeded) indices.set(addrKey(addr), xfIndex);
      return succeeded;
    },
    getCellXf(xfIndex: number): CellXf | null {
      counts.getXf += 1;
      if (options.missingXfIndexes?.has(xfIndex)) return null;
      return xfs.get(xfIndex) ?? null;
    },
    getEffectiveStyle(sheet: number, row: number, col: number) {
      counts.getEffectiveStyle += 1;
      return options.effectiveStyles?.get([sheet, row, col].join(':')) ?? null;
    },
    getFontRecord(index: number): FontRecord | null {
      return fonts.get(index) ?? null;
    },
    getFillRecord(index: number): FillRecord | null {
      return fills.get(index) ?? null;
    },
    getBorderRecord(index: number): BorderRecord | null {
      counts.getBorder += 1;
      const stored = borders.get(index) ?? null;
      return options.getBorderRecordResult
        ? options.getBorderRecordResult(index, counts.getBorder, stored)
        : stored;
    },
    getNumFmtCode(index: number): string | null {
      return numFmts.get(index) ?? null;
    },
    addFontRecord(record: FontRecord): number {
      counts.addFont += 1;
      const allocated = nextFont++;
      const result = options.addFontResult?.(record, counts.addFont, allocated) ?? allocated;
      if (result >= 0) fonts.set(result, record);
      return result;
    },
    addFillRecord(record: FillRecord): number {
      counts.addFill += 1;
      const index = nextFill++;
      fills.set(index, record);
      return index;
    },
    addBorderRecord(record: BorderRecord): number {
      counts.addBorder += 1;
      const index = nextBorder++;
      borders.set(index, record);
      return index;
    },
    addNumFmtCode(code: string): number {
      counts.addNumFmt += 1;
      const index = nextNumFmt++;
      numFmts.set(index, code);
      return index;
    },
    addXfRecord(record: CellXf): number {
      counts.addXf += 1;
      const result = options.addXfResult?.(record, counts.addXf) ?? nextXf++;
      if (result >= 0) {
        xfs.set(result, record);
        addedXfs.push(record);
      }
      return result;
    },
    getValue() {
      return { kind: 'text' as const, value: 'かな' };
    },
    setCellPhoneticRuns(sheet: number, row: number, col: number, runs: readonly PhoneticRun[]) {
      counts.setRuns += 1;
      setPhoneticRuns.push({ addr: { sheet, row, col }, runs });
      return true;
    },
    setCellPhoneticProperties(
      sheet: number,
      row: number,
      col: number,
      properties: PhoneticProperties,
    ) {
      counts.setPhoneticProperties += 1;
      setProperties.push({ addr: { sheet, row, col }, properties });
      return true;
    },
    getCellShowPhonetic(sheet: number, row: number, col: number) {
      return showPhonetic.get([sheet, row, col].join(':')) ?? false;
    },
    setCellShowPhonetic(sheet: number, row: number, col: number, value: boolean) {
      counts.setShowPhonetic += 1;
      showPhonetic.set([sheet, row, col].join(':'), value);
      return true;
    },
  };

  return {
    wb: wb as unknown as WorkbookHandle,
    counts,
    addedXfs,
    assignedXfs,
    setPhoneticRuns,
    setProperties,
    defaultFont,
    xfs,
    fonts,
    borders,
    indices,
  };
};

const storeWithFormats = (formats: Iterable<readonly [string, CellFormat]>) => {
  const store = createSpreadsheetStore();
  store.setState((state) => ({
    ...state,
    format: { ...state.format, formats: new Map(formats) },
  }));
  return store;
};

describe('syncCellFormatsToEngine resolution cache', () => {
  it('resolves 1,000 distinct format objects once per structural format and source XF', () => {
    const entries: Array<readonly [string, CellFormat]> = [];
    const colors = ['#ff0000', '#00ff00', '#0000ff', '#ffff00'];
    for (let row = 0; row < 1_000; row += 1) {
      entries.push([
        addrKey({ sheet: 0, row, col: 0 }),
        { bold: row % 2 === 0, fill: colors[row % colors.length] } as CellFormat,
      ]);
    }
    const harness = makeHarness();

    syncCellFormatsToEngine(harness.wb, storeWithFormats(entries), 0);

    expect(harness.counts.addFont).toBe(4);
    expect(harness.counts.addFill).toBe(4);
    expect(harness.counts.addBorder).toBe(4);
    expect(harness.counts.addXf).toBe(4);
    expect(harness.counts.getIndex).toBe(1_000);
    expect(harness.counts.setXf).toBe(1_000);
  });

  it('keeps source XF metadata separate when the same format targets different raw XFs', () => {
    const indexA = addrKey({ sheet: 0, row: 0, col: 0 });
    const indexB = addrKey({ sheet: 0, row: 1, col: 0 });
    const rawA = sourceXf({
      fontIndex: 2,
      quotePrefix: true,
      applyAlignment: true,
      hasAlignment: true,
      horizontalAlign: 2,
      verticalAlign: 1,
      indent: 4,
      relativeIndent: 2,
      textRotation: 35,
      shrinkToFit: true,
      readingOrder: 2,
      hasProtection: true,
      locked: false,
      hidden: true,
      applyProtection: true,
    });
    const rawB = sourceXf({
      fontIndex: 3,
      quotePrefix: false,
      applyAlignment: false,
      hasAlignment: false,
      horizontalAlign: 0,
      verticalAlign: 2,
      indent: 1,
      textRotation: 0,
      readingOrder: 1,
      hasProtection: true,
      locked: true,
      hidden: false,
      applyProtection: false,
    });
    const harness = makeHarness({
      indices: new Map([
        [indexA, 2],
        [indexB, 3],
      ]),
      xfs: new Map([
        [2, rawA],
        [3, rawB],
      ]),
    });
    harness.fonts.set(
      2,
      font({ hasCharset: true, charset: 128, color: { ...color(), kind: 2, theme: 4 } }),
    );
    harness.fonts.set(
      3,
      font({ hasCharset: true, charset: 134, color: { ...color(), kind: 2, theme: 5 } }),
    );
    const format = { fontSize: 16 } as CellFormat;

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([
        [indexA, format],
        [indexB, { ...format }],
      ]),
      0,
    );

    expect(harness.counts.addXf).toBe(2);
    expect(harness.addedXfs[0]).toMatchObject({
      quotePrefix: true,
      applyAlignment: true,
      hasAlignment: true,
      indent: 4,
      relativeIndent: 2,
      textRotation: 35,
      shrinkToFit: true,
      readingOrder: 2,
      hasProtection: true,
      locked: false,
      hidden: true,
      applyProtection: true,
    });
    expect(harness.addedXfs[1]).toMatchObject({
      quotePrefix: false,
      applyAlignment: false,
      indent: 1,
      textRotation: 0,
      readingOrder: 1,
      locked: true,
      hidden: false,
      applyProtection: false,
    });
    expect(harness.fonts.get(harness.addedXfs[0]?.fontIndex ?? -1)).toMatchObject({
      hasCharset: true,
      charset: 128,
      color: { kind: 2, theme: 4 },
    });
    expect(harness.fonts.get(harness.addedXfs[1]?.fontIndex ?? -1)).toMatchObject({
      hasCharset: true,
      charset: 134,
      color: { kind: 2, theme: 5 },
    });
  });

  it('does not merge absent fields with own undefined fields', () => {
    const keyA = addrKey({ sheet: 0, row: 0, col: 0 });
    const keyB = addrKey({ sheet: 0, row: 1, col: 0 });
    const keyC = addrKey({ sheet: 0, row: 2, col: 0 });
    const harness = makeHarness({
      xfs: new Map([[1, sourceXf({ hasProtection: true, locked: false })]]),
    });

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([
        [keyA, {} as CellFormat],
        [keyB, { align: undefined } as CellFormat],
        [keyC, { locked: undefined } as CellFormat],
      ]),
      0,
    );

    expect(harness.counts.addXf).toBe(3);
  });

  it('does not cache formats whose resolved fields are inherited or non-enumerable', () => {
    const inheritedA = Object.create({ fontFamily: 'Inherited A' }) as CellFormat;
    const inheritedB = Object.create({ fontFamily: 'Inherited B' }) as CellFormat;
    const inheritedHarness = makeHarness();
    syncCellFormatsToEngine(
      inheritedHarness.wb,
      storeWithFormats([
        [addrKey({ sheet: 0, row: 0, col: 0 }), inheritedA],
        [addrKey({ sheet: 0, row: 1, col: 0 }), inheritedB],
      ]),
      0,
    );

    expect(inheritedHarness.counts.addXf).toBe(2);
    expect(inheritedHarness.counts.addFont).toBe(2);

    const nonEnumerableA = {};
    const nonEnumerableB = {};
    Object.defineProperty(nonEnumerableA, 'fontFamily', { value: 'Hidden A' });
    Object.defineProperty(nonEnumerableB, 'fontFamily', { value: 'Hidden B' });
    const nonEnumerableHarness = makeHarness();
    syncCellFormatsToEngine(
      nonEnumerableHarness.wb,
      storeWithFormats([
        [addrKey({ sheet: 0, row: 0, col: 0 }), nonEnumerableA as CellFormat],
        [addrKey({ sheet: 0, row: 1, col: 0 }), nonEnumerableB as CellFormat],
      ]),
      0,
    );

    expect(nonEnumerableHarness.counts.addXf).toBe(2);
    expect(nonEnumerableHarness.counts.addFont).toBe(4);

    const numFmtA = Object.create({ kind: 'fixed', decimals: 1, thousands: false });
    const numFmtB = Object.create({ kind: 'fixed', decimals: 2, thousands: false });
    const nestedHarness = makeHarness();
    syncCellFormatsToEngine(
      nestedHarness.wb,
      storeWithFormats([
        [addrKey({ sheet: 0, row: 0, col: 0 }), { numFmt: numFmtA } as CellFormat],
        [addrKey({ sheet: 0, row: 1, col: 0 }), { numFmt: numFmtB } as CellFormat],
      ]),
      0,
    );
    expect(nestedHarness.counts.addNumFmt).toBe(2);
    expect(nestedHarness.addedXfs[0]?.numFmtId).not.toBe(nestedHarness.addedXfs[1]?.numFmtId);
  });

  it('does not invoke getters while deciding cache eligibility', () => {
    let reads = 0;
    const format = {};
    Object.defineProperty(format, 'fontFamily', {
      enumerable: true,
      get() {
        reads += 1;
        return 'Getter family';
      },
    });
    const harness = makeHarness();

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([
        [addrKey({ sheet: 0, row: 0, col: 0 }), format as CellFormat],
        [addrKey({ sheet: 0, row: 1, col: 0 }), format as CellFormat],
      ]),
      0,
    );

    expect(reads).toBe(18);
    expect(harness.counts.addXf).toBe(2);
  });

  it('still writes phonetic state for every cell after an XF cache hit', () => {
    const format: CellFormat = {
      bold: true,
      phonetic: [{ start: 0, end: 2, text: 'かな' }],
      phoneticProperties: { fontId: 2, type: 2, alignment: 1 },
      showPhonetic: true,
    };
    const entries = [0, 1, 2].map((row) => [addrKey({ sheet: 0, row, col: 0 }), format] as const);
    const harness = makeHarness({ phonetic: true });

    syncCellFormatsToEngine(harness.wb, storeWithFormats(entries), 0);

    expect(harness.counts.addXf).toBe(1);
    expect(harness.counts.setRuns).toBe(3);
    expect(harness.counts.setPhoneticProperties).toBe(3);
    expect(harness.counts.setShowPhonetic).toBe(3);
    expect(harness.setPhoneticRuns.map(({ addr }) => addr.row)).toEqual([0, 1, 2]);
    expect(harness.setProperties.map(({ addr }) => addr.row)).toEqual([0, 1, 2]);
  });

  it('does not reuse a missing raw XF or a failed XF resolution', () => {
    const keyA = addrKey({ sheet: 0, row: 0, col: 0 });
    const keyB = addrKey({ sheet: 0, row: 1, col: 0 });
    const format = { bold: true } as CellFormat;
    const missing = makeHarness({
      indices: new Map([
        [keyA, 42],
        [keyB, 42],
      ]),
      missingXfIndexes: new Set([42]),
    });

    syncCellFormatsToEngine(
      missing.wb,
      storeWithFormats([
        [keyA, format],
        [keyB, { ...format }],
      ]),
      0,
    );

    expect(missing.counts.getXf).toBe(2);
    expect(missing.counts.addXf).toBe(2);

    const failed = makeHarness({ addXfResult: () => -1 });
    syncCellFormatsToEngine(
      failed.wb,
      storeWithFormats([
        [keyA, format],
        [keyB, { ...format }],
      ]),
      0,
    );
    expect(failed.counts.addXf).toBe(2);
  });

  it('does not cache null or negative source XF indexes', () => {
    const keys = [
      addrKey({ sheet: 0, row: 0, col: 0 }),
      addrKey({ sheet: 0, row: 1, col: 0 }),
      addrKey({ sheet: 0, row: 2, col: 0 }),
      addrKey({ sheet: 0, row: 3, col: 0 }),
    ] as const;
    const invalid = makeHarness({
      indices: new Map([
        [keys[0], -1],
        [keys[1], -1],
      ]),
      nullIndexes: new Set([keys[2], keys[3]]),
    });

    syncCellFormatsToEngine(
      invalid.wb,
      storeWithFormats(keys.map((key) => [key, { bold: true } as CellFormat] as const)),
      0,
    );

    expect(invalid.counts.getIndex).toBe(4);
    expect(invalid.counts.getXf).toBe(2);
    expect(invalid.counts.addXf).toBe(4);
  });

  it('retries when preserving the source font fails for one XF resolution', () => {
    const keyA = addrKey({ sheet: 0, row: 0, col: 0 });
    const keyB = addrKey({ sheet: 0, row: 1, col: 0 });
    const harness = makeHarness({
      xfs: new Map([[1, sourceXf({ fontIndex: 2 })]]),
      addFontResult: (_record, call, allocated) => (call === 2 ? -1 : allocated),
    });
    harness.fonts.set(
      2,
      font({ hasCharset: true, charset: 128, color: { ...color(), kind: 2, theme: 4 } }),
    );

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([
        [keyA, { fontSize: 16 } as CellFormat],
        [keyB, { fontSize: 16 } as CellFormat],
      ]),
      0,
    );

    expect(harness.counts.addFont).toBe(4);
    expect(harness.counts.addXf).toBe(2);
    expect(harness.fonts.get(harness.addedXfs[0]?.fontIndex ?? -1)?.hasCharset).not.toBe(true);
    expect(harness.fonts.get(harness.addedXfs[1]?.fontIndex ?? -1)).toMatchObject({
      hasCharset: true,
      charset: 128,
      color: { kind: 2, theme: 4 },
    });
  });

  it('retries when a source XF projection is incomplete', () => {
    const keyA = addrKey({ sheet: 0, row: 0, col: 0 });
    const keyB = addrKey({ sheet: 0, row: 1, col: 0 });
    const harness = makeHarness({
      xfs: new Map([[1, sourceXf({ borderIndex: 2 })]]),
      getBorderRecordResult: (index, call, stored) => (index === 2 && call === 1 ? null : stored),
    });
    const rawBorder = border();
    rawBorder.top = { style: 1, colorArgb: 0xff000000, color: color() };
    harness.borders.set(2, rawBorder);

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([
        [keyA, { borders: { top: { style: 'thin' } } } as CellFormat],
        [keyB, { borders: { top: { style: 'thin' } } } as CellFormat],
      ]),
      0,
    );

    expect(harness.counts.addXf).toBe(2);
    expect(harness.addedXfs[0]?.borderIndex).not.toBe(2);
    expect(harness.addedXfs[1]?.borderIndex).toBe(2);
  });

  it('keeps the cache local to one flush and retries after strict setter failure', () => {
    const keyA = addrKey({ sheet: 0, row: 0, col: 0 });
    const keyB = addrKey({ sheet: 0, row: 1, col: 0 });
    const format = { bold: true } as CellFormat;
    const store = storeWithFormats([
      [keyA, format],
      [keyB, { ...format }],
    ]);
    let acceptSets = false;
    const harness = makeHarness({ setXfResult: () => acceptSets });

    expect(() => syncCellFormatsToEngine(harness.wb, store, 0, { strict: true })).toThrow(
      /cell-format setCellXfIndex at 0:0:0/,
    );
    expect(harness.counts.setXf).toBe(1);
    acceptSets = true;
    syncCellFormatsToEngine(harness.wb, store, 0, { strict: true });
    expect(harness.counts.addXf).toBe(2);
    expect(harness.counts.getIndex).toBe(3);
    expect(harness.counts.setXf).toBe(3);
  });

  it('refreshes a repeated format after workbook defaults change between flushes', () => {
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    const harness = makeHarness();
    const store = storeWithFormats([[key, { bold: true } as CellFormat]]);
    syncCellFormatsToEngine(harness.wb, store, 0);
    const firstFontCount = harness.counts.addFont;

    harness.defaultFont.current = { name: 'Aptos', size: 12, scheme: 1 };
    harness.indices.set(key, 1);
    harness.xfs.set(1, sourceXf({ quotePrefix: true }));
    syncCellFormatsToEngine(harness.wb, store, 0);

    expect(harness.counts.addFont).toBe(firstFontCount + 1);
    expect([...harness.fonts.values()].some((record) => record.name === 'Aptos')).toBe(true);
    expect(harness.addedXfs.at(-1)?.quotePrefix).toBe(true);
  });

  it('does not cache when the workbook has no readable default font', () => {
    const keyA = addrKey({ sheet: 0, row: 0, col: 0 });
    const keyB = addrKey({ sheet: 0, row: 1, col: 0 });
    const harness = makeHarness();
    harness.defaultFont.current = null;

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([
        [keyA, { bold: true } as CellFormat],
        [keyB, { bold: true } as CellFormat],
      ]),
      0,
    );

    expect(harness.counts.addXf).toBe(2);
  });

  it('uses one default-font snapshot throughout an XF resolution', () => {
    const harness = makeHarness();
    let reads = 0;
    Object.defineProperty(harness.wb, 'workbookDefaultFont', {
      get: () => (++reads === 1 ? { name: 'Aptos', size: 11, scheme: 2 } : null),
    });

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([[addrKey({ sheet: 0, row: 0, col: 0 }), { fontFamily: 'Aptos' }]]),
      0,
    );

    const resolved = harness.addedXfs.at(-1);
    expect(resolved).toBeDefined();
    expect(harness.fonts.get(resolved?.fontIndex ?? -1)?.scheme).toBe(2);
    expect(reads).toBe(1);
  });

  it('does not cache accounting formats that depend on native format templates', () => {
    const harness = makeHarness();
    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([
        [
          addrKey({ sheet: 0, row: 0, col: 0 }),
          { numFmt: { kind: 'accounting', decimals: 2 } } as CellFormat,
        ],
        [
          addrKey({ sheet: 0, row: 1, col: 0 }),
          { numFmt: { kind: 'accounting', decimals: 2 } } as CellFormat,
        ],
      ]),
      0,
    );

    expect(harness.counts.addNumFmt).toBe(2);
    expect(harness.counts.addXf).toBe(2);
  });

  it('bounds retained resolutions to 256 entries', () => {
    const formats: Array<readonly [string, CellFormat]> = [];
    for (let row = 0; row < 257; row += 1) {
      formats.push([addrKey({ sheet: 0, row, col: 0 }), { fontSize: row + 8 } as CellFormat]);
    }
    formats.push([addrKey({ sheet: 0, row: 257, col: 0 }), { fontSize: 8 } as CellFormat]);
    const harness = makeHarness();

    syncCellFormatsToEngine(harness.wb, storeWithFormats(formats), 0);

    expect(harness.counts.addXf).toBe(258);
  });

  it('preserves raw font metadata from an inherited row XF', () => {
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    const harness = makeHarness({
      effectiveStyle: true,
      effectiveStyles: new Map([[key, { source: 1, xfIndex: 2 }]]),
      indices: new Map([[key, 0]]),
      xfs: new Map([[2, sourceXf({ fontIndex: 2, applyFont: false, quotePrefix: true })]]),
    });
    harness.fonts.set(
      2,
      font({ hasCharset: true, charset: 128, color: { ...color(), kind: 2, theme: 4 } }),
    );

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([[key, { fill: '#fff2cc' } as CellFormat]]),
      0,
    );

    expect(harness.counts.getEffectiveStyle).toBe(1);
    expect(harness.addedXfs[0]).toMatchObject({
      fontIndex: 2,
      applyFont: false,
      quotePrefix: true,
    });
    expect(harness.fonts.get(harness.addedXfs[0]?.fontIndex ?? -1)).toMatchObject({
      hasCharset: true,
      charset: 128,
      color: { kind: 2, theme: 4 },
    });
  });

  it('preserves raw font metadata from an inherited column XF', () => {
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    const harness = makeHarness({
      effectiveStyle: true,
      effectiveStyles: new Map([[key, { source: 2, xfIndex: 3 }]]),
      indices: new Map([[key, 0]]),
      xfs: new Map([[3, sourceXf({ fontIndex: 3, applyFont: true, quotePrefix: false })]]),
    });
    harness.fonts.set(
      3,
      font({ hasCharset: true, charset: 134, color: { ...color(), kind: 2, theme: 5 } }),
    );

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([[key, { fill: '#fff2cc' } as CellFormat]]),
      0,
    );

    expect(harness.counts.getEffectiveStyle).toBe(1);
    expect(harness.addedXfs[0]).toMatchObject({
      fontIndex: 3,
      applyFont: true,
      quotePrefix: false,
    });
    expect(harness.fonts.get(harness.addedXfs[0]?.fontIndex ?? -1)).toMatchObject({
      hasCharset: true,
      charset: 134,
      color: { kind: 2, theme: 5 },
    });
  });

  it('keys repeated format resolutions by the inherited XF', () => {
    const keys = [0, 1, 2].map((row) => addrKey({ sheet: 0, row, col: 0 }));
    const harness = makeHarness({
      effectiveStyle: true,
      effectiveStyles: new Map(
        keys.map((key, index) => [
          key,
          index === 1 ? { source: 2, xfIndex: 3 } : { source: 1, xfIndex: 2 },
        ]),
      ),
      indices: new Map(keys.map((key) => [key, 0] as const)),
      xfs: new Map([
        [2, sourceXf({ fontIndex: 2 })],
        [3, sourceXf({ fontIndex: 3 })],
      ]),
    });
    const format = { fill: '#fff2cc' } as CellFormat;

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats(keys.map((key) => [key, { ...format }] as const)),
      0,
    );

    expect(harness.counts.getEffectiveStyle).toBe(3);
    expect(harness.counts.addXf).toBe(2);
    expect(harness.assignedXfs).toEqual([100, 101, 100]);
    expect(harness.addedXfs.map((xf) => xf.fontIndex)).toEqual([2, 3]);
  });

  it('uses direct XFs for cell/default sources, null indexes, and unsupported effective style', () => {
    const key0 = addrKey({ sheet: 0, row: 0, col: 0 });
    const key1 = addrKey({ sheet: 0, row: 1, col: 0 });
    const key2 = addrKey({ sheet: 0, row: 2, col: 0 });
    const key3 = addrKey({ sheet: 0, row: 3, col: 0 });
    const keys = [key0, key1, key2, key3];
    const harness = makeHarness({
      effectiveStyle: true,
      effectiveStyles: new Map([
        [key0, { source: 0, xfIndex: 2 }],
        [key1, { source: 3, xfIndex: 3 }],
        [key2, { source: 1, xfIndex: 3 }],
        [key3, { source: 1, xfIndex: 3 }],
      ]),
      indices: new Map([
        [key0, 0],
        [key1, 0],
        [key3, 2],
      ]),
      nullIndexes: new Set([key2]),
      xfs: new Map([
        [0, sourceXf({ fontIndex: 1, quotePrefix: true })],
        [2, sourceXf({ fontIndex: 2 })],
        [3, sourceXf({ fontIndex: 3 })],
      ]),
    });

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats(keys.map((key) => [key, { fill: '#fff2cc' } as CellFormat] as const)),
      0,
    );

    expect(harness.counts.getEffectiveStyle).toBe(2);
    expect(harness.counts.addXf).toBe(3);
    expect(harness.assignedXfs).toEqual([100, 100, 101, 102]);
    expect(harness.addedXfs[0]).toMatchObject({ fontIndex: 1, quotePrefix: true });
    expect(harness.addedXfs[1]?.fontIndex).not.toBe(3);
    expect(harness.addedXfs[2]).toMatchObject({ fontIndex: 2 });

    const unsupported = makeHarness({
      effectiveStyles: new Map([[key0, { source: 1, xfIndex: 3 }]]),
      indices: new Map([[key0, 0]]),
      xfs: new Map([
        [0, sourceXf({ fontIndex: 1, quotePrefix: true })],
        [3, sourceXf({ fontIndex: 3 })],
      ]),
    });
    syncCellFormatsToEngine(
      unsupported.wb,
      storeWithFormats([[key0, { fill: '#fff2cc' } as CellFormat]]),
      0,
    );

    expect(unsupported.counts.getEffectiveStyle).toBe(0);
    expect(unsupported.addedXfs[0]).toMatchObject({ fontIndex: 1, quotePrefix: true });
  });

  it('keeps an explicit font size ahead of an inherited axis font', () => {
    const key = addrKey({ sheet: 0, row: 0, col: 0 });
    const harness = makeHarness({
      effectiveStyle: true,
      effectiveStyles: new Map([[key, { source: 1, xfIndex: 2 }]]),
      indices: new Map([[key, 0]]),
      xfs: new Map([[2, sourceXf({ fontIndex: 2 })]]),
    });
    harness.fonts.set(2, font({ size: 24 }));

    syncCellFormatsToEngine(
      harness.wb,
      storeWithFormats([[key, { fontSize: (11 * 96) / 72 } as CellFormat]]),
      0,
    );

    expect(harness.counts.getEffectiveStyle).toBe(1);
    const resolved = harness.addedXfs[0];
    expect(resolved).toBeDefined();
    expect(harness.fonts.get(resolved?.fontIndex ?? -1)?.size).toBeCloseTo(11);
  });
});
