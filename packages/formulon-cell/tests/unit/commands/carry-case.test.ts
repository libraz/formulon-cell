import { describe, expect, it } from 'vitest';
import { carryCase, defaultSpellingService } from '../../../src/commands/spelling.js';

describe('carryCase', () => {
  it('carries all caps, a leading capital, or nothing', () => {
    expect(carryCase('the', 'TEH')).toBe('THE');
    expect(carryCase('the', 'Teh')).toBe('The');
    expect(carryCase('the', 'teh')).toBe('the');
    expect(carryCase('the', 'T')).toBe('The');
  });

  it('treats a typed text without cased letters as lower case', () => {
    expect(carryCase('…', '...')).toBe('…');
    expect(carryCase('abc', '12')).toBe('abc');
  });

  it('is the one rule behind spelling suggestions', () => {
    expect(defaultSpellingService.suggest('TEH', 'en')).toEqual(['THE']);
    expect(defaultSpellingService.suggest('Teh', 'en')).toEqual(['The']);
  });
});
