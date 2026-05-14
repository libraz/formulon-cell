import { describe, expect, it } from 'vitest';
import { normalizeFormatLocale as fromShared } from '../../../src/format/locale.js';
import { normalizeFormatLocale as fromFormatDialog } from '../../../src/interact/format-dialog-model.js';
import { normalizeFormatLocale as fromHitState } from '../../../src/render/grid/hit-state.js';

describe('format/locale.ts — shared normalizer', () => {
  it('canonicalises short locales to BCP-47', () => {
    expect(fromShared('ja')).toBe('ja-JP');
    expect(fromShared('en')).toBe('en-US');
    expect(fromShared('en-GB')).toBe('en-GB');
    expect(fromShared('')).toBe('en-US');
  });

  it('hit-state re-exports the same implementation', () => {
    expect(fromHitState).toBe(fromShared);
  });

  it('format-dialog-model re-exports the same implementation', () => {
    expect(fromFormatDialog).toBe(fromShared);
  });
});
