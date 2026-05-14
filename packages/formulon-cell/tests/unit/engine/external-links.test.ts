import { describe, expect, it } from 'vitest';

import { externalLinkKindLabel } from '../../../src/engine/external-links.js';

describe('engine/external-links', () => {
  it('maps each well-known kind code to its label', () => {
    expect(externalLinkKindLabel(1)).toBe('externalBook');
    expect(externalLinkKindLabel(2)).toBe('ole');
    expect(externalLinkKindLabel(3)).toBe('dde');
  });

  it('returns "unknown" for codes outside the known set', () => {
    expect(externalLinkKindLabel(0)).toBe('unknown');
    expect(externalLinkKindLabel(99)).toBe('unknown');
    expect(externalLinkKindLabel(-1)).toBe('unknown');
  });
});
