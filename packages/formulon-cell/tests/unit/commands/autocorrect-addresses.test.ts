import { describe, expect, it } from 'vitest';
import {
  capitalizeDays,
  capitalizeSentences,
  correctTwoInitialCapitals,
} from '../../../src/commands/autocorrect.js';

describe('AutoCorrect capitalisation and addresses', () => {
  it('leaves URLs, e-mail addresses and network paths as typed', () => {
    expect(capitalizeDays('see https://x.test/monday and monday')).toBe(
      'see https://x.test/monday and Monday',
    );
    expect(capitalizeDays('monday@x.test')).toBe('monday@x.test');
    expect(correctTwoInitialCapitals('see WWw.x.test and THe')).toBe('see WWw.x.test and The');
    expect(correctTwoInitialCapitals('HTtp://x.test/ABc')).toBe('HTtp://x.test/ABc');
    expect(correctTwoInitialCapitals('JOhn@x.test')).toBe('JOhn@x.test');
  });

  it('does not capitalise an address after a sentence end', () => {
    expect(capitalizeSentences('Done. https://x.test and done. next')).toBe(
      'Done. https://x.test and done. Next',
    );
    expect(capitalizeSentences('Done. www.x.test')).toBe('Done. www.x.test');
    expect(capitalizeSentences('Done. a@x.test')).toBe('Done. a@x.test');
  });
});
