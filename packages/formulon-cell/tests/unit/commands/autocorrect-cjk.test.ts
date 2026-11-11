import { describe, expect, it } from 'vitest';
import {
  AutoCorrectStore,
  autoCorrectText,
  builtInAutoCorrectEntries,
  correctTypedEntry,
  registerAutoCorrect,
} from '../../../src/commands/autocorrect.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const ja = builtInAutoCorrectEntries('ja');

/** Types `text` a character at a time, correcting at each delimiter, then commits it. */
const typeThrough = (text: string): string => {
  const autoCorrect = new AutoCorrectStore(
    () => undefined,
    () => 'ja',
    () => {},
  );
  let value = '';
  for (const ch of text) {
    value += ch;
    const live = autoCorrect.live(value, value.length);
    if (live) value = live.value;
  }
  const store = createSpreadsheetStore();
  registerAutoCorrect(store, autoCorrect);
  return correctTypedEntry(store, value);
};

describe('AutoCorrect in text written without spaces', () => {
  it('replaces a CJK entry inside a run of CJK text', () => {
    expect(autoCorrectText('今日はこんにちわ', ja)).toBe('今日はこんにちは');
    expect(autoCorrectText('こんにちわ世界', ja)).toBe('こんにちは世界');
    expect(autoCorrectText('皆さんこんばんわ、また明日', ja)).toBe('皆さんこんばんは、また明日');
    expect(autoCorrectText('御社様', [{ from: '御社', to: '貴社' }])).toBe('貴社様');
    expect(autoCorrectText('ｺﾝﾆﾁﾜ!', [{ from: 'ｺﾝﾆﾁﾜ', to: 'ｺﾝﾆﾁﾊ' }])).toBe('ｺﾝﾆﾁﾊ!');
  });

  it('corrects mid-text CJK entries while typing and on commit', () => {
    expect(typeThrough('今日はこんにちわ。')).toBe('今日はこんにちは。');
    expect(typeThrough('こんにちわ世界')).toBe('こんにちは世界');
  });

  it('keeps the word boundary for a Latin entry', () => {
    const teh = [{ from: 'teh', to: 'the' }];
    expect(autoCorrectText('teh cat', teh)).toBe('the cat');
    expect(autoCorrectText('xteh cat', teh)).toBe('xteh cat');
    expect(autoCorrectText('tehx', teh)).toBe('tehx');
    expect(autoCorrectText('日本teh', teh)).toBe('日本teh');
  });
});
