import { describe, expect, it } from 'vitest';
import {
  type KeyTipBinding,
  type KeyTipState,
  reduceKeyTip,
  visibleBindings,
} from '../../../src/commands/key-tips.js';

const bindings: KeyTipBinding[] = [
  { chord: 'F', commandId: 'open-file' },
  { chord: 'H', commandId: 'home-tab' },
  { chord: 'HC', commandId: 'copy' },
  { chord: 'HV', commandId: 'paste' },
];

const idle: KeyTipState = { kind: 'idle' };

describe('reduceKeyTip', () => {
  it('idle + altDown → showing with empty chord', () => {
    const r = reduceKeyTip(idle, { kind: 'altDown' }, bindings);
    expect(r.state).toEqual({ kind: 'showing', chord: '' });
  });

  it('showing + altUp → idle', () => {
    const r = reduceKeyTip({ kind: 'showing', chord: 'H' }, { kind: 'altUp' }, bindings);
    expect(r.state).toEqual(idle);
  });

  it('showing + escape → idle', () => {
    const r = reduceKeyTip({ kind: 'showing', chord: 'H' }, { kind: 'escape' }, bindings);
    expect(r.state).toEqual(idle);
  });

  it('letter that exactly matches a binding invokes it', () => {
    const r = reduceKeyTip({ kind: 'showing', chord: '' }, { kind: 'letter', key: 'F' }, bindings);
    expect(r.state).toEqual({ kind: 'invoked', commandId: 'open-file' });
    expect(r.invoked).toBe('open-file');
  });

  it('letter that partially matches grows the chord buffer', () => {
    const r = reduceKeyTip(
      { kind: 'showing', chord: '' },
      { kind: 'letter', key: 'h' }, // case-insensitive
      bindings,
    );
    expect(r.state).toEqual({ kind: 'showing', chord: 'H' });
  });

  it('chord progression H → C invokes the HC binding', () => {
    let r = reduceKeyTip({ kind: 'showing', chord: '' }, { kind: 'letter', key: 'H' }, bindings);
    r = reduceKeyTip(r.state, { kind: 'letter', key: 'C' }, bindings);
    expect(r.state).toEqual({ kind: 'invoked', commandId: 'copy' });
  });

  it('unknown letter resets chord buffer but stays showing', () => {
    const r = reduceKeyTip({ kind: 'showing', chord: '' }, { kind: 'letter', key: 'Z' }, bindings);
    expect(r.state).toEqual({ kind: 'showing', chord: '' });
  });

  it('letter received in idle is a no-op', () => {
    const r = reduceKeyTip(idle, { kind: 'letter', key: 'F' }, bindings);
    expect(r.state).toEqual(idle);
  });
});

describe('visibleBindings', () => {
  it('returns all bindings when state is showing with empty chord', () => {
    expect(visibleBindings({ kind: 'showing', chord: '' }, bindings)).toEqual(bindings);
  });

  it('filters by the buffered prefix', () => {
    const visible = visibleBindings({ kind: 'showing', chord: 'H' }, bindings);
    expect(visible.map((b) => b.commandId)).toEqual(['home-tab', 'copy', 'paste']);
  });

  it('returns [] when state is idle', () => {
    expect(visibleBindings(idle, bindings)).toEqual([]);
  });

  it('returns [] when state is invoked', () => {
    expect(visibleBindings({ kind: 'invoked', commandId: 'open-file' }, bindings)).toEqual([]);
  });
});
