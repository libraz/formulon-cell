import { afterEach, describe, expect, it } from 'vitest';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import {
  createSpreadsheetStore,
  mutators,
  type SessionTextVert,
} from '../../../src/store/store.js';

const cleanup: (() => void)[] = [];
afterEach(() => {
  for (const dispose of cleanup.splice(0).reverse()) dispose();
});

/** A shape holding two lines of text in direction `textVert`, drawn into a host. */
const drawn = (textVert?: SessionTextVert) => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  cleanup.push(() => host.remove());
  const store = createSpreadsheetStore();
  mutators.upsertIllustration(store, {
    id: 'box',
    kind: 'shape',
    shape: 'rectangle',
    sheet: 0,
    x: 40,
    y: 40,
    w: 120,
    h: 60,
    text: 'ab\ncd',
    ...(textVert ? { textVert } : {}),
  });
  const handle = attachSessionIllustrations({ host, store });
  cleanup.push(() => handle.detach());
  const text = host.querySelector('svg text');
  if (!text) throw new Error('no text');
  return { text, turned: host.querySelector('svg > g[transform]:not(:has(rect))') };
};

describe('the drawn text direction', () => {
  it('draws no direction, and horz, horizontally', () => {
    for (const vert of [undefined, 'horz' as const]) {
      const { text, turned } = drawn(vert);
      expect(turned).toBeNull();
      expect(text.getAttribute('writing-mode')).toBeNull();
    }
  });

  it('turns the text a quarter turn clockwise for vert', () => {
    const { text, turned } = drawn('vert');
    expect(turned?.getAttribute('transform')).toBe('translate(120 0) rotate(90)');
    expect(text.getAttribute('writing-mode')).toBeNull();
  });

  it('turns the text a quarter turn counter-clockwise for vert270', () => {
    const { turned } = drawn('vert270');
    expect(turned?.getAttribute('transform')).toBe('translate(0 60) rotate(-90)');
  });

  it('stacks East Asian text in columns running right to left', () => {
    const { text } = drawn('eaVert');
    expect(text.getAttribute('writing-mode')).toBe('vertical-rl');
    expect(text.getAttribute('text-orientation')).toBe('mixed');
  });

  it('stacks Mongolian text in columns running left to right', () => {
    expect(drawn('mongolianVert').text.getAttribute('writing-mode')).toBe('vertical-lr');
  });

  it('stands word art letters upright, both ways alike', () => {
    for (const vert of ['wordArtVert', 'wordArtVertRtl'] as const) {
      const { text } = drawn(vert);
      expect(text.getAttribute('writing-mode')).toBe('vertical-rl');
      expect(text.getAttribute('text-orientation')).toBe('upright');
    }
  });
});
