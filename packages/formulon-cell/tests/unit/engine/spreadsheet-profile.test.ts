import { describe, expect, it } from 'vitest';

import {
  engineProfileToPublic,
  publicProfileToEngine,
} from '../../../src/engine/spreadsheet-profile.js';

describe('engine/spreadsheet-profile', () => {
  it.each([
    ['windows-ja_JP', 'win-365-ja_JP'],
    ['windows-en_US', 'win-365-en_US'],
    ['windows-de_DE', 'win-365-de_DE'],
    ['windows-fr_FR', 'win-365-fr_FR'],
    ['windows-zh_CN', 'win-365-zh_CN'],
    ['windows-ko_KR', 'win-365-ko_KR'],
    ['windows-th_TH', 'win-365-th_TH'],
    ['mac-ja_JP', 'mac-365-ja_JP'],
    ['mac-en_US', 'mac-365-en_US'],
    ['mac-de_DE', 'mac-365-de_DE'],
    ['mac-fr_FR', 'mac-365-fr_FR'],
    ['mac-zh_CN', 'mac-365-zh_CN'],
    ['mac-ko_KR', 'mac-365-ko_KR'],
    ['mac-th_TH', 'mac-365-th_TH'],
  ] as const)('maps %s to the exact engine id %s and back', (publicId, engineId) => {
    expect(publicProfileToEngine(publicId)).toBe(engineId);
    expect(engineProfileToPublic(engineId)).toBe(publicId);
  });

  it('returns null for an unknown engine profile id', () => {
    expect(engineProfileToPublic('zorblax-12-fr_FR')).toBeNull();
  });

  it('uses different engine ids for windows vs mac so they cannot alias', () => {
    expect(publicProfileToEngine('windows-ja_JP')).not.toBe(publicProfileToEngine('mac-ja_JP'));
  });
});
