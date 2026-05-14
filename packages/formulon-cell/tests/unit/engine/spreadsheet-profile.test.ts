import { describe, expect, it } from 'vitest';

import {
  engineProfileToPublic,
  publicProfileToEngine,
} from '../../../src/engine/spreadsheet-profile.js';

describe('engine/spreadsheet-profile', () => {
  it('round-trips windows-ja_JP through publicProfileToEngine → engineProfileToPublic', () => {
    const engine = publicProfileToEngine('windows-ja_JP');
    expect(engineProfileToPublic(engine)).toBe('windows-ja_JP');
  });

  it('round-trips mac-ja_JP', () => {
    const engine = publicProfileToEngine('mac-ja_JP');
    expect(engineProfileToPublic(engine)).toBe('mac-ja_JP');
  });

  it('returns null for an unknown engine profile id', () => {
    expect(engineProfileToPublic('zorblax-12-fr_FR')).toBeNull();
  });

  it('uses different engine ids for windows vs mac so they cannot alias', () => {
    expect(publicProfileToEngine('windows-ja_JP')).not.toBe(publicProfileToEngine('mac-ja_JP'));
  });
});
