import { describe, expect, it } from 'vitest';
import { boundedAxisLabel } from '../../../src/charts/axes.js';
import { textWidth, truncateLabel } from '../../../src/charts/scene-primitives.js';

/** Quadratic reference: re-measures the whole candidate each step. */
const reference = (label: string, maxWidth: number, fontSize = 10, maxChars = Infinity): string => {
  let result = '';
  let count = 0;
  for (const character of Array.from(label)) {
    if (count >= maxChars) break;
    if (textWidth(`${result}${character}…`, fontSize) > maxWidth) break;
    result += character;
    count += 1;
  }
  return `${result}…`;
};

describe('label truncation', () => {
  const labels = [
    'Quarterly revenue by region',
    '売上高の推移（四半期ごと）',
    'a b c d e f g',
    '■■■ 2024 …',
  ];

  it('matches the re-measuring reference for latin, CJK and symbol text', () => {
    for (const label of labels) {
      for (const width of [4, 12, 30, 61, 100]) {
        for (const size of [8, 10, 14]) {
          const full = textWidth(label, size) <= width;
          expect(truncateLabel(label, width, size)).toBe(
            full ? label : reference(label, width, size),
          );
        }
      }
    }
  });

  it('bounds axis labels by character count and width', () => {
    const long = 'x'.repeat(1000);
    expect(boundedAxisLabel(long, 1e9)).toBe(reference(long, 1e9, 10, 256));
    expect(boundedAxisLabel(long, 1e9)).toHaveLength(257);
    expect(boundedAxisLabel('short')).toBe('short');
  });
});
