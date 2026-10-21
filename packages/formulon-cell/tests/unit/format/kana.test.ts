import { describe, expect, it } from 'vitest';
import { convertKana } from '../../../src/format/kana.js';

describe('format/kana — convertKana', () => {
  it('turns hiragana into full-width katakana', () => {
    expect(convertKana('とうきょう', 'katakana')).toBe('トウキョウ');
  });

  it('turns katakana, half-width included, into hiragana', () => {
    expect(convertKana('トウキョウ', 'hiragana')).toBe('とうきょう');
    expect(convertKana('ｶﾞｷﾞｸﾞ', 'hiragana')).toBe('がぎぐ');
  });

  it('turns kana into half-width katakana with separate voicing marks', () => {
    expect(convertKana('がっこう', 'halfKatakana')).toBe('ｶﾞｯｺｳ');
    expect(convertKana('パーティー', 'halfKatakana')).toBe('ﾊﾟｰﾃｨｰ');
  });

  it('joins half-width voicing marks when widening', () => {
    expect(convertKana('ｶﾞｯｺｳ', 'katakana')).toBe('ガッコウ');
    expect(convertKana('ﾊﾟｰ', 'katakana')).toBe('パー');
  });

  it('leaves non-kana text and the as-stored form alone', () => {
    expect(convertKana('ABC 東京', 'katakana')).toBe('ABC 東京');
    expect(convertKana('とうきょう', 'asStored')).toBe('とうきょう');
  });
});
