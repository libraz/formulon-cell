import { describe, expect, it } from 'vitest';
import { type AutoCorrectOptions, AutoCorrectStore } from '../../../src/commands/autocorrect.js';

const storeFor = (locale: string, options: AutoCorrectOptions) => {
  let current = options;
  const store = new AutoCorrectStore(
    () => current,
    () => locale,
    (next) => {
      current = next;
    },
  );
  return { store, options: () => current };
};

describe('AutoCorrect locale lookup', () => {
  it('reads a language key for a regional locale', () => {
    const { store } = storeFor('ja-JP', {
      entries: { ja: [{ from: 'a', to: 'b' }] },
      exceptions: { ja: { firstLetter: ['x'], initialCaps: [] } },
    });
    expect(store.state().entries).toEqual([{ from: 'a', to: 'b' }]);
    expect(store.state().exceptions.firstLetter).toEqual(['x']);
  });

  it('lets the full locale key win over the language key', () => {
    const { store } = storeFor('ja-JP', {
      entries: { ja: [{ from: 'a', to: 'b' }], 'ja-JP': [{ from: 'c', to: 'd' }] },
    });
    expect(store.state().entries).toEqual([{ from: 'c', to: 'd' }]);
  });

  it('writes a user edit under the full locale key', () => {
    const { store, options } = storeFor('ja-JP', { entries: { ja: [{ from: 'a', to: 'b' }] } });
    store.add('e', 'f');
    expect(options().entries?.['ja-JP']).toEqual([
      { from: 'a', to: 'b' },
      { from: 'e', to: 'f' },
    ]);
    expect(store.state().entries).toHaveLength(2);
  });
});
