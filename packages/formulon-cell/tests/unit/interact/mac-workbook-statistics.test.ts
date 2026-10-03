import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en, ja } from '../../../src/i18n/strings.js';
import {
  attachMacWorkbookStatistics,
  collectWorkbookStatistics,
} from '../../../src/interact/mac-workbook-statistics.js';
import {
  type CellFormat,
  createSpreadsheetStore,
  mutators,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const makeWorkbook = (): WorkbookHandle =>
  ({
    sheetCount: 2,
    capabilities: { commentsEnumerable: true, hyperlinks: true },
    physicalCells: (sheet: number) => {
      const cells =
        sheet === 0
          ? [
              {
                addr: { sheet, row: 0, col: 0 },
                value: { kind: 'number', value: 1 },
                formula: null,
              },
              {
                addr: { sheet, row: 2, col: 1 },
                value: { kind: 'text', value: 'text' },
                formula: null,
              },
              { addr: { sheet, row: 3, col: 1 }, value: { kind: 'blank' }, formula: '=A1+1' },
            ]
          : [
              {
                addr: { sheet, row: 1, col: 2 },
                value: { kind: 'bool', value: true },
                formula: null,
              },
            ];
      return cells[Symbol.iterator]();
    },
    getComments: (sheet: number) => (sheet === 0 ? [{}, {}] : [{}]),
    getHyperlinks: (sheet: number) => (sheet === 0 ? [{}] : []),
    getTables: () => [
      { name: 'Sales', displayName: 'Sales', ref: 'A1:B3', sheetIndex: 0, columns: ['A', 'B'] },
    ],
  }) as unknown as WorkbookHandle;

const makeEmptyWorkbook = (capabilities = { commentsEnumerable: false, hyperlinks: false }) =>
  ({
    sheetCount: 2,
    capabilities,
    physicalCells: () => [][Symbol.iterator](),
    cells: () => [][Symbol.iterator](),
    getTables: () => [],
  }) as unknown as WorkbookHandle;

const setFormats = (store: SpreadsheetStore, formats: ReadonlyMap<string, CellFormat>): void => {
  store.setState((state) => ({
    ...state,
    format: { ...state.format, formats: new Map(formats) },
  }));
};

describe('collectWorkbookStatistics', () => {
  it('counts workbook values and metadata across every sheet', () => {
    expect(collectWorkbookStatistics(makeWorkbook())).toEqual({
      sheets: 2,
      populatedCells: 4,
      formulas: 1,
      numbers: 1,
      text: 1,
      booleans: 1,
      errors: 0,
      tables: 1,
      comments: 3,
      hyperlinks: 1,
      usedRows: 4,
      usedColumns: 3,
    });
  });

  it('uses the legacy cells iterator when physicalCells is unavailable', () => {
    const workbook = {
      sheetCount: 1,
      capabilities: { commentsEnumerable: false, hyperlinks: false },
      cells: () =>
        [
          {
            addr: { sheet: 0, row: 4, col: 5 },
            value: { kind: 'error', code: 7, text: '#N/A' },
            formula: null,
          },
        ][Symbol.iterator](),
      getTables: () => [],
    } as unknown as WorkbookHandle;

    expect(collectWorkbookStatistics(workbook)).toMatchObject({
      populatedCells: 1,
      errors: 1,
      usedRows: 1,
      usedColumns: 1,
    });
  });

  it('uses native annotation anchors without expanding hyperlink rectangles', () => {
    const workbook = {
      sheetCount: 2,
      capabilities: { commentsEnumerable: true, hyperlinks: true },
      physicalCells: () => [][Symbol.iterator](),
      getComments: (sheet: number) =>
        sheet === 1
          ? [
              { row: 12, col: 13, author: 'A', text: 'note' },
              { row: 1_048_575, col: 16_383, author: 'A', text: 'last valid' },
              { row: 1_048_576, col: 1, author: 'A', text: 'overflow row' },
            ]
          : [],
      getHyperlinks: (sheet: number) =>
        sheet === 0
          ? [
              { row: 20, col: 15, lastRow: 22, lastCol: 18, target: 'https://example.test' },
              {
                row: 1_048_575,
                col: 16_383,
                lastRow: 1_048_575,
                lastCol: 16_383,
                target: 'https://last.example',
              },
              {
                row: 2,
                col: 16_384,
                lastRow: 2,
                lastCol: 16_384,
                target: 'https://overflow.example',
              },
            ]
          : [],
      getTables: () => [],
    } as unknown as WorkbookHandle;

    expect(collectWorkbookStatistics(workbook)).toMatchObject({
      comments: 3,
      hyperlinks: 3,
      usedRows: 4,
      usedColumns: 4,
    });
  });

  it('uses valid format entries for bounds and metadata fallback only when native enumeration is absent', () => {
    const store = createSpreadsheetStore();
    setFormats(
      store,
      new Map([
        ['0:3:4', { bold: true }],
        ['1:8:9', { comment: 'note' }],
        ['1:10:11', { hyperlink: 'https://example.test' }],
        ['1:12:13', { fill: '#ffeeaa' }],
        ['0:1048575:16383', { bold: true }],
        ['0:1048576:1', { bold: true }],
        ['0:1:16384', { bold: true }],
        ['2:20:20', { bold: true }],
        ['-1:14:14', { bold: true }],
        ['0:-2:3', { bold: true }],
        ['0:4:-5', { bold: true }],
        ['0::1', { bold: true }],
      ]),
    );

    expect(collectWorkbookStatistics(makeEmptyWorkbook(), store)).toMatchObject({
      comments: 1,
      hyperlinks: 1,
      usedRows: 5,
      usedColumns: 5,
    });
  });

  it('counts live workbook data and store-only cells across two non-stub sheets', async () => {
    const workbook = await WorkbookHandle.createDefault();
    const store = createSpreadsheetStore();
    try {
      expect(workbook.isStub).toBe(false);
      expect(workbook.addSheet('Second')).toBe(1);
      expect(workbook.capabilities.commentsEnumerable).toBe(true);
      expect(workbook.capabilities.hyperlinks).toBe(true);

      workbook.setNumber({ sheet: 0, row: 0, col: 0 }, 10);
      workbook.setFormula({ sheet: 0, row: 1, col: 0 }, '=A1*2');
      workbook.setText({ sheet: 1, row: 2, col: 3 }, 'second');
      expect(workbook.setCommentEntry(0, 5, 4, 'Test', 'first blank note')).toBe(true);
      expect(workbook.setCommentEntry(1, 6, 5, 'Test', 'second blank note')).toBe(true);
      expect(workbook.addHyperlink(0, 7, 6, 'https://one.example', 'One')).toBe(true);
      expect(workbook.addHyperlink(1, 8, 7, 'https://two.example', 'Two')).toBe(true);
      mutators.setCellFormat(store, { sheet: 0, row: 9, col: 8 }, { bold: true });
      mutators.setCellFormat(store, { sheet: 1, row: 10, col: 9 }, { fill: '#ffeeaa' });

      expect(collectWorkbookStatistics(workbook, store)).toEqual({
        sheets: 2,
        populatedCells: 3,
        formulas: 1,
        numbers: 2,
        text: 1,
        booleans: 0,
        errors: 0,
        tables: 0,
        comments: 2,
        hyperlinks: 2,
        usedRows: 9,
        usedColumns: 8,
      });
    } finally {
      workbook.dispose();
    }
  });
});

describe('attachMacWorkbookStatistics', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('initializes the localized title, aria label, and close label', () => {
    for (const [strings, title, cancel] of [
      [ja, 'ブックの統計情報', '閉じる'],
      [en, 'Workbook Statistics', 'Close'],
    ] as const) {
      const handle = attachMacWorkbookStatistics({ host, getWb: makeWorkbook, strings });
      const dialog = document.querySelector<HTMLElement>('.fc-macstatsdlg');
      expect(dialog?.getAttribute('aria-label')).toBe(title);
      expect(dialog?.querySelector('.fc-fmtdlg__header')?.textContent).toBe(title);
      expect(dialog?.querySelector('.fc-fmtdlg__footer button')?.textContent).toBe(cancel);
      handle.detach();
    }
  });

  it('closes from the footer close button', () => {
    const handle = attachMacWorkbookStatistics({ host, getWb: makeWorkbook });
    handle.open();
    const dialog = document.querySelector<HTMLElement>('.fc-macstatsdlg');
    const close = dialog?.querySelector<HTMLButtonElement>('.fc-fmtdlg__footer button');
    expect(dialog?.hidden).toBe(false);
    close?.click();
    expect(dialog?.hidden).toBe(true);
    handle.detach();
  });

  it('keeps setStrings locale through refresh', () => {
    const handle = attachMacWorkbookStatistics({ host, getWb: makeWorkbook, strings: en });
    handle.open();
    handle.setStrings(ja);
    handle.refresh();

    const dialog = document.querySelector<HTMLElement>('.fc-macstatsdlg');
    expect(dialog?.getAttribute('aria-label')).toBe('ブックの統計情報');
    expect(dialog?.querySelector('.fc-fmtdlg__header')?.textContent).toBe('ブックの統計情報');
    expect(dialog?.querySelector('.fc-fmtdlg__footer button')?.textContent).toBe('閉じる');
    expect(dialog?.querySelector('[data-mac-workbook-stat="sheets"]')?.textContent).toBe(
      'シート数',
    );
    handle.detach();
  });

  it('re-reads the live workbook on reopen after replacement and detaches cleanly', async () => {
    const first = await WorkbookHandle.createDefault();
    const second = await WorkbookHandle.createDefault();
    try {
      expect(first.isStub).toBe(false);
      expect(second.isStub).toBe(false);
      first.setNumber({ sheet: 0, row: 0, col: 0 }, 1);
      second.setText({ sheet: 0, row: 0, col: 0 }, 'replacement');
      let current = first;
      const handle = attachMacWorkbookStatistics({ host, getWb: () => current, strings: en });

      handle.open();
      let dialog = document.querySelector<HTMLElement>('.fc-macstatsdlg');
      expect(dialog?.querySelector('[data-mac-workbook-stat-value="numbers"]')?.textContent).toBe(
        '1',
      );
      expect(dialog?.querySelector('[data-mac-workbook-stat-value="text"]')?.textContent).toBe('0');
      handle.close();

      current = second;
      handle.open();
      dialog = document.querySelector<HTMLElement>('.fc-macstatsdlg');
      expect(dialog?.querySelector('[data-mac-workbook-stat-value="numbers"]')?.textContent).toBe(
        '0',
      );
      expect(dialog?.querySelector('[data-mac-workbook-stat-value="text"]')?.textContent).toBe('1');

      handle.detach();
      expect(document.querySelector('.fc-macstatsdlg')).toBeNull();
    } finally {
      first.dispose();
      second.dispose();
    }
  });
});
