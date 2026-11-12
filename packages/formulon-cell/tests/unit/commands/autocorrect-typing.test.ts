import { describe, expect, it } from 'vitest';
import {
  type AutoCorrectOptions,
  AutoCorrectStore,
  correctTypedEntry,
  registerAutoCorrect,
} from '../../../src/commands/autocorrect.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const storeFor = (locale: string, options: AutoCorrectOptions = {}) => {
  let current = options;
  return new AutoCorrectStore(
    () => current,
    () => locale,
    (next) => {
      current = next;
    },
  );
};

/** Types `text` a character at a time, correcting at each delimiter, then commits it. */
const typeThrough = (autoCorrect: AutoCorrectStore, text: string): string => {
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

const teh = { entries: { en: [{ from: 'teh', to: 'the' }] } };

// AutoCorrect as typed on the desktop app: each word is corrected when a delimiter follows it,
// and the last word when the entry is committed.
describe('AutoCorrect while typing', () => {
  it('replaces a word as soon as the delimiter after it is typed', () => {
    const autoCorrect = storeFor('en', teh);
    expect(autoCorrect.live('teh ', 4)).toEqual({ value: 'the ', caret: 4 });
    expect(autoCorrect.live('teh', 3)).toBeNull();
    expect(typeThrough(autoCorrect, 'teh')).toBe('the');
  });

  it.each([
    ['teh', 'the'],
    ['Teh', 'The'],
    ['TEH', 'THE'],
  ])('carries the case of %s', (typed, expected) => {
    expect(typeThrough(storeFor('en', teh), typed)).toBe(expected);
  });

  it.each([
    ['THe xx', 'The xx'],
    ['xx THe xx', 'xx The xx'],
    ['ABc', 'Abc'],
    ['ABc xx', 'Abc xx'],
    ['IDs xx', 'IDs xx'],
    ['PCs IDs xx', 'PCs IDs xx'],
    ['xx IDs', 'xx Ids'],
    ['IDs', 'Ids'],
  ])('corrects two initial capitals in %s', (typed, expected) => {
    expect(typeThrough(storeFor('en'), typed)).toBe(expected);
  });

  it('leaves plural acronyms alone with the two-capitals switch off', () => {
    expect(typeThrough(storeFor('en', { correctTwoInitialCapitals: false }), 'IDs PCs')).toBe(
      'IDs PCs',
    );
  });

  it.each([
    ['hello world', 'hello world'],
    ['hello. world', 'hello. World'],
    ['see approx. five', 'see approx. Five'],
    ['monday', 'Monday'],
    ['mon tue', 'mon tue'],
  ])('capitalises %s as typed', (typed, expected) => {
    expect(typeThrough(storeFor('ja'), typed)).toBe(expected);
  });

  it('corrects words inside a formula string and entries outside it', () => {
    expect(typeThrough(storeFor('en', teh), '="teh"')).toBe('="the"');
    expect(typeThrough(storeFor('en', teh), '=teh+"teh x"')).toBe('=the+"the x"');
  });

  it('starts a Japanese UI from the Japanese list only', () => {
    expect(typeThrough(storeFor('ja'), '(c) こんにちわ')).toBe('(c) こんにちは');
  });
});
