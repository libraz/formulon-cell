import { describe, expect, it } from 'vitest';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';
import {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  buildTranslationReviewItems,
  formatRibbonReport,
  parseScriptCommand,
  type ReviewCell,
  reviewCellsFromState,
} from '../../../src/toolbar/review-tools.js';

describe('toolbar/review-tools', () => {
  it('detects accessibility warnings and informational review items', () => {
    const cells: ReviewCell[] = [
      { label: 'A1', value: { kind: 'error', text: '#REF!' }, formula: '=Missing!A1' },
      { label: 'A2', value: { kind: 'text', value: 'https://example.com/report' } },
      { label: 'A3', value: { kind: 'text', value: 'LOUD STATUS TEXT THAT KEEPS GOING' } },
      { label: 'A4', value: { kind: 'text', value: '   ' } },
    ];

    const items = analyzeAccessibilityCells(cells);

    expect(items).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ severity: 'warning', label: 'A1' }),
        expect.objectContaining({ label: 'A2', detail: expect.stringContaining('URL') }),
        expect.objectContaining({ label: 'A3', detail: expect.stringContaining('All-caps') }),
        expect.objectContaining({ label: 'A4', detail: 'Cell contains only whitespace.' }),
      ]),
    );
  });

  it('reports an empty sheet clearly', () => {
    expect(analyzeAccessibilityCells([])).toEqual([
      {
        severity: 'info',
        label: 'Empty sheet',
        detail: 'The current sheet has no populated cells to review.',
      },
    ]);
    expect(analyzeAccessibilityCells([], 'ja')).toEqual([
      {
        severity: 'info',
        label: '空のシート',
        detail: '現在のシートにはレビュー対象の入力済みセルがありません。',
      },
    ]);
  });

  it('collects review cells from the active sheet store state', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(
      store,
      { sheet: 0, row: 1, col: 1 },
      { kind: 'text', value: 'teh report' },
      null,
    );
    mutators.setCell(
      store,
      { sheet: 1, row: 0, col: 0 },
      { kind: 'text', value: 'other sheet' },
      null,
    );

    expect(reviewCellsFromState(store.getState())).toEqual([
      { label: 'B2', value: { kind: 'text', value: 'teh report' }, formula: null, source: 'cell' },
    ]);
    expect(
      reviewCellsFromState(store.getState(), 0, { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 }),
    ).toEqual([]);
  });

  it('includes comments as reviewable text entries', () => {
    const store = createSpreadsheetStore();
    mutators.setCell(store, { sheet: 0, row: 0, col: 0 }, { kind: 'text', value: 'ok' }, null);
    mutators.setCellFormat(store, { sheet: 0, row: 0, col: 0 }, { comment: 'teh note note' });

    const cells = reviewCellsFromState(store.getState());
    expect(cells).toEqual(
      expect.arrayContaining([
        {
          label: 'A1 comment',
          value: { kind: 'text', value: 'teh note note' },
          source: 'comment',
        },
      ]),
    );
    expect(analyzeSpellingCells(cells)).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ label: 'A1 comment', detail: 'Repeated word: "note note".' }),
        expect.objectContaining({
          label: 'A1 comment',
          detail: 'Possible typo: "teh". Suggested spelling: "the".',
        }),
      ]),
    );
  });

  it('formats a compact review report for built-in ribbon actions', () => {
    expect(formatRibbonReport('Spelling', [])).toBe('Spelling\nNo issues found.');
    expect(
      formatRibbonReport('Spelling', [
        { severity: 'warning', label: 'A1', detail: 'Possible typo.' },
      ]),
    ).toBe('Spelling\nWarning - A1: Possible typo.');
    expect(formatRibbonReport('スペル チェック', [], 'ja')).toBe(
      'スペル チェック\n問題は見つかりませんでした。',
    );
  });

  it('builds a localized translation review payload from text cells', () => {
    const items = buildTranslationReviewItems(
      [
        { label: 'A1', value: { kind: 'text', value: '  翻訳  する テキスト  ' } },
        { label: 'A2', value: { kind: 'number' } },
      ],
      'ja',
    );

    expect(items).toEqual([
      {
        severity: 'info',
        label: 'A1',
        detail: '翻訳対象テキスト: "翻訳 する テキスト"',
      },
    ]);
    expect(buildTranslationReviewItems([], 'ja')).toEqual([
      {
        severity: 'info',
        label: '翻訳',
        detail: '翻訳対象のテキストセルが見つかりません。',
      },
    ]);
  });

  it('detects spelling and prose issues', () => {
    const items = analyzeSpellingCells([
      { label: 'B2', value: { kind: 'text', value: 'teh  report report . next sentence' } },
    ]);

    expect(items).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ detail: 'Repeated word: "report report".' }),
        expect.objectContaining({ detail: 'Contains repeated spaces.' }),
        expect.objectContaining({ detail: 'Contains a space before punctuation.' }),
        expect.objectContaining({
          detail: 'A sentence appears to start with a lowercase letter.',
        }),
        expect.objectContaining({
          detail: 'Possible typo: "teh". Suggested spelling: "the".',
        }),
      ]),
    );
  });

  it('localizes built-in review findings for Japanese ribbon actions', () => {
    expect(
      analyzeSpellingCells(
        [{ label: 'B2', value: { kind: 'text', value: 'teh  report report .' } }],
        'ja',
      ),
    ).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ detail: '同じ語が繰り返されています: "report report"。' }),
        expect.objectContaining({ detail: '連続した空白が含まれています。' }),
        expect.objectContaining({
          detail: 'スペルミスの可能性: "teh"。候補: "the"。',
        }),
      ]),
    );
    expect(
      analyzeAccessibilityCells([{ label: 'A4', value: { kind: 'text', value: '   ' } }], 'ja'),
    ).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ detail: 'セルには空白文字だけが含まれています。' }),
      ]),
    );
  });

  it('parses and applies supported script commands', () => {
    expect(parseScriptCommand(' UPPERCASE ')).toBe('uppercase');
    expect(parseScriptCommand('unknown')).toBeNull();
    expect(applyTextScript(' Mixed Value ', 'uppercase')).toBe(' MIXED VALUE ');
    expect(applyTextScript(' Mixed Value ', 'lowercase')).toBe(' mixed value ');
    expect(applyTextScript('  Mixed   Value  ', 'trim')).toBe('Mixed Value');
  });
});
