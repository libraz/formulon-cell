import { describe, expect, it } from 'vitest';

import { dictionaries } from '../../../src/i18n/strings.js';

type Leaf = string | readonly string[];
type RecordValue = Record<string, unknown>;

const isRecord = (value: unknown): value is RecordValue =>
  value !== null && typeof value === 'object' && !Array.isArray(value);

const collectLeaves = (value: unknown, prefix = ''): [string, Leaf][] => {
  if (!isRecord(value)) return [[prefix, value as Leaf]];
  return Object.entries(value).flatMap(([key, child]) =>
    collectLeaves(child, prefix ? `${prefix}.${key}` : key),
  );
};

const interpolationVariables = (value: Leaf): string[] => {
  const text = typeof value === 'string' ? value : value.join('\u0000');
  return [...text.matchAll(/\{([A-Za-z]\w*)\}/g)].map((match) => match[1] ?? '').sort();
};

describe('seven-language translation dictionary contract', () => {
  it('distinguishes Korean threaded comments from legacy notes across editing surfaces', () => {
    const ko = dictionaries.ko;
    expect(ko.findReplace.comments).toBe('메모');
    expect(ko.findReplace.notes).toBe('노트');
    expect(ko.ribbon.newComment).toBe('새 노트');
    expect(ko.ribbon.editComment).toBe('노트 편집');
    expect(ko.contextMenu.insertComment).toBe('노트 편집…');
    expect(ko.contextMenu.deleteComment).toBe('노트 삭제');
    expect(ko.formatDialog.comment).toBe('노트');
  });
  const entries = Object.fromEntries(
    Object.entries(dictionaries).map(([locale, dictionary]) => [locale, collectLeaves(dictionary)]),
  );
  const english = new Map(entries.en);

  it('keeps every locale on the same leaf set with non-empty values', () => {
    const englishPaths = [...english.keys()].sort();
    for (const [locale, localeEntries] of Object.entries(entries)) {
      const values = new Map(localeEntries);
      expect([...values.keys()].sort(), locale).toEqual(englishPaths);
      for (const [path, value] of localeEntries) {
        if (typeof value !== 'string') {
          expect(value.length, `${locale}.${path}`).toBeGreaterThan(0);
          expect(
            value.every((item) => item.trim() !== ''),
            `${locale}.${path}`,
          ).toBe(true);
        } else {
          expect(value.trim(), `${locale}.${path}`).not.toBe('');
        }
      }
    }
  });

  it('preserves interpolation variables in every locale', () => {
    for (const [locale, localeEntries] of Object.entries(entries)) {
      for (const [path, value] of localeEntries) {
        expect(interpolationVariables(value), `${locale}.${path}`).toEqual(
          interpolationVariables(english.get(path) ?? ''),
        );
      }
    }
  });

  it('uses the grammatically correct French diagonal-border labels', () => {
    expect(dictionaries.fr.formatDialog.borderDiagonalDown).toBe('Diagonale ↘');
    expect(dictionaries.fr.formatDialog.borderDiagonalUp).toBe('Diagonale ↗');
  });

  it('quotes localized conditional-formatting option names in status messages', () => {
    expect(dictionaries.de.cfRulesDialog.stopIfTrueUnavailable).toContain('„Anhalten, wenn wahr“');
    expect(dictionaries.th.cfRulesDialog.stopIfTrueUnavailable).toContain('"หยุดถ้าเป็นจริง"');
  });

  it('keeps the Thai fill command directional', () => {
    expect(dictionaries.th.ribbonMenu.fillGroup).toBe('ข้ามเวิร์กชีต…');
  });

  it('uses Microsoft Excel French terminology for date and color rules', () => {
    expect(dictionaries.fr.conditionalMenu.dateOccurring).toBe('Date se produisant…');
    expect(dictionaries.fr.conditionalDialog.kindDateOccurring).toBe('Date se produisant');
    expect(dictionaries.fr.conditionalMenu.colorScaleGreenYellowRed).toBe(
      'Nuances de couleurs vert-jaune-rouge',
    );
  });
});
