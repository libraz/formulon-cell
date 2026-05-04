import { describe, expect, it } from 'vitest';
import {
  applyFlashFill,
  applyFlashFillPattern,
  type FlashFillExample,
  inferFlashFillPattern,
} from '../../../src/commands/flash-fill.js';

const ex = (input: string, output: string): FlashFillExample => ({ input, output });

describe('inferFlashFillPattern', () => {
  it('returns null on empty input set', () => {
    expect(inferFlashFillPattern([])).toBeNull();
  });

  it('detects identity', () => {
    const p = inferFlashFillPattern([ex('foo', 'foo'), ex('bar', 'bar')]);
    expect(p).toEqual({ kind: 'identity' });
  });

  it('detects constant suffix', () => {
    const p = inferFlashFillPattern([ex('John', 'John Smith'), ex('Jane', 'Jane Smith')]);
    expect(p).toEqual({ kind: 'constant-suffix', suffix: ' Smith' });
  });

  it('detects constant prefix', () => {
    const p = inferFlashFillPattern([ex('apple', 'fruit-apple'), ex('peach', 'fruit-peach')]);
    expect(p).toEqual({ kind: 'constant-prefix', prefix: 'fruit-' });
  });

  it('detects affix wrap', () => {
    const p = inferFlashFillPattern([ex('foo', '[foo]'), ex('bar', '[bar]')]);
    expect(p).toEqual({ kind: 'affix', prefix: '[', suffix: ']' });
  });

  it('detects substring slice (first three chars)', () => {
    const p = inferFlashFillPattern([ex('FOOBAR', 'FOO'), ex('HELLO!', 'HEL')]);
    expect(p).toEqual({ kind: 'substring', start: 0, length: 3 });
  });

  it('detects delimited token (first name from "First Last")', () => {
    const p = inferFlashFillPattern([ex('John Smith', 'John'), ex('Jane Doe', 'Jane')]);
    expect(p).toEqual({ kind: 'token', delimiter: ' ', index: 0 });
  });

  it('detects email local-part via @ delimiter', () => {
    const p = inferFlashFillPattern([ex('alice@example.com', 'alice'), ex('bob@acme.org', 'bob')]);
    expect(p).toEqual({ kind: 'token', delimiter: '@', index: 0 });
  });

  it('detects upper-case transform', () => {
    const p = inferFlashFillPattern([ex('hi', 'HI'), ex('there', 'THERE')]);
    expect(p).toEqual({ kind: 'case', mode: 'upper' });
  });

  it('detects lower-case transform', () => {
    const p = inferFlashFillPattern([ex('HI', 'hi'), ex('There', 'there')]);
    expect(p).toEqual({ kind: 'case', mode: 'lower' });
  });

  it('detects title-case transform', () => {
    const p = inferFlashFillPattern([ex('john smith', 'John Smith'), ex('jane doe', 'Jane Doe')]);
    expect(p).toEqual({ kind: 'case', mode: 'title' });
  });

  it('returns null when no single pattern explains every example', () => {
    expect(
      inferFlashFillPattern([
        ex('foo', 'foobar'),
        ex('baz', 'qux'), // requires a different rule than the first.
      ]),
    ).toBeNull();
  });
});

describe('applyFlashFillPattern', () => {
  it('identity returns input verbatim', () => {
    expect(applyFlashFillPattern({ kind: 'identity' }, 'abc')).toBe('abc');
  });

  it('constant-suffix appends the suffix', () => {
    expect(applyFlashFillPattern({ kind: 'constant-suffix', suffix: '!' }, 'go')).toBe('go!');
  });

  it('constant-prefix prepends the prefix', () => {
    expect(applyFlashFillPattern({ kind: 'constant-prefix', prefix: 'Mr. ' }, 'Holmes')).toBe(
      'Mr. Holmes',
    );
  });

  it('substring honours start + length', () => {
    expect(applyFlashFillPattern({ kind: 'substring', start: 1, length: 3 }, 'abcdef')).toBe('bcd');
  });

  it('substring returns null when start exceeds input length', () => {
    expect(applyFlashFillPattern({ kind: 'substring', start: 10, length: 1 }, 'abc')).toBeNull();
  });

  it('token returns null when index past tokens', () => {
    expect(applyFlashFillPattern({ kind: 'token', delimiter: ',', index: 5 }, 'a,b,c')).toBeNull();
  });

  it('case modes', () => {
    expect(applyFlashFillPattern({ kind: 'case', mode: 'upper' }, 'hi')).toBe('HI');
    expect(applyFlashFillPattern({ kind: 'case', mode: 'lower' }, 'HI')).toBe('hi');
    expect(applyFlashFillPattern({ kind: 'case', mode: 'title' }, 'hello world')).toBe(
      'Hello World',
    );
  });
});

describe('applyFlashFill', () => {
  it('extends a constant-suffix pattern across a column', () => {
    const examples = [ex('A', 'A!')];
    const p = inferFlashFillPattern(examples);
    expect(p).not.toBeNull();
    if (!p) return;
    expect(applyFlashFill(p, ['B', 'C', 'D'])).toEqual(['B!', 'C!', 'D!']);
  });

  it('produces null entries when individual rows fail the pattern', () => {
    const p = { kind: 'token' as const, delimiter: ' ', index: 1 };
    expect(applyFlashFill(p, ['only-one'])).toEqual([null]);
  });
});
