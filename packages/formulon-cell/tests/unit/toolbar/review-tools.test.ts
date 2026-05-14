import { describe, expect, it } from 'vitest';

import {
  analyzeAccessibilityCells,
  analyzeSpellingCells,
  applyTextScript,
  parseScriptCommand,
  type ReviewCell,
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

  it('parses and applies supported script commands', () => {
    expect(parseScriptCommand(' UPPERCASE ')).toBe('uppercase');
    expect(parseScriptCommand('unknown')).toBeNull();
    expect(applyTextScript(' Mixed Value ', 'uppercase')).toBe(' MIXED VALUE ');
    expect(applyTextScript(' Mixed Value ', 'lowercase')).toBe(' mixed value ');
    expect(applyTextScript('  Mixed   Value  ', 'trim')).toBe('Mixed Value');
  });
});
