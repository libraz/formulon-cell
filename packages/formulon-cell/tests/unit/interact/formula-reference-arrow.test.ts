import { describe, expect, it } from 'vitest';
import { classifyFormulaReferenceArrow } from '../../../src/interact/formula-reference-arrow.js';

const key = (init: Partial<KeyboardEventInit>): KeyboardEvent =>
  new KeyboardEvent('keydown', { key: 'ArrowRight', ...init });

describe('formula reference arrow classifier', () => {
  it.each([
    { init: {}, mac: false, expected: { direction: 'right', extend: false, jump: false } },
    {
      init: { shiftKey: true },
      mac: false,
      expected: { direction: 'right', extend: true, jump: false },
    },
    {
      init: { ctrlKey: true },
      mac: false,
      expected: { direction: 'right', extend: false, jump: true },
    },
    {
      init: { ctrlKey: true, shiftKey: true },
      mac: true,
      expected: { direction: 'right', extend: true, jump: true },
    },
    {
      init: { metaKey: true },
      mac: true,
      expected: { direction: 'right', extend: false, jump: true },
    },
    {
      init: { metaKey: true, shiftKey: true },
      mac: true,
      expected: { direction: 'right', extend: true, jump: true },
    },
  ])('accepts the supported modifiers: $init on Mac=$mac', ({ init, mac, expected }) => {
    expect(classifyFormulaReferenceArrow(key(init), mac)).toEqual(expected);
  });

  it.each([
    { init: { altKey: true }, mac: true },
    { init: { ctrlKey: true, metaKey: true }, mac: true },
    { init: { metaKey: true }, mac: false },
    { init: { key: 'Home' }, mac: true },
    { init: { isComposing: true }, mac: true },
    { init: { key: 'Process' }, mac: true },
  ])('leaves unsupported input native: $init on Mac=$mac', ({ init, mac }) => {
    expect(classifyFormulaReferenceArrow(key(init), mac)).toBeNull();
  });
});
