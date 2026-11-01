import { describe, expect, it } from 'vitest';
import { paintCellText } from '../../../src/render/painters.js';
import type { ResolvedTheme } from '../../../src/theme/resolve.js';
import oracle from '../../fixtures/excel-general-number-fit.json' with { type: 'json' };

// Every glyph advances one Arial 11 pt digit; the spreadsheet charges a
// General number one digit per character, so this reproduces its fit test.
const DIGIT = oracle.digitWidthPt;

const paintedText = (value: number, widthChars: number): string => {
  const texts: string[] = [];
  const ctx = {
    font: '',
    fillStyle: '',
    direction: 'inherit',
    textAlign: 'left',
    textBaseline: 'alphabetic',
    save(): void {},
    restore(): void {},
    beginPath(): void {},
    rect(): void {},
    clip(): void {},
    measureText: (t: string) => ({ width: [...t].length * DIGIT }) as TextMetrics,
    fillText: (t: string) => texts.push(t),
  } as unknown as CanvasRenderingContext2D;
  // Column width in points is chars × the Normal font's max digit width; the
  // painter's own insets at zoom 1 stand in for the spreadsheet's.
  const available = widthChars * oracle.maxDigitWidthPt - oracle.paddingPt;
  paintCellText({
    ctx,
    bounds: { x: 0, y: 0, w: available + 6.5, h: 20 },
    theme: {
      cellAutoFg: '#000',
      cellNumFg: '#000',
      fontUi: 'Arial',
      fontMono: 'monospace',
      textCell: 11,
    } as ResolvedTheme,
    value: { kind: 'number', value },
    formula: null,
    isActive: false,
    isInRange: false,
    locale: 'en-US',
  });
  return texts.join('');
};

describe('General number fit matches the spreadsheet', () => {
  for (const [i, value] of oracle.values.entries()) {
    it(`fits ${value} into each column width`, () => {
      const got = oracle.widthsChars.map((w) => paintedText(value, w));
      expect(got).toEqual(oracle.text[i]);
    });
  }
});
