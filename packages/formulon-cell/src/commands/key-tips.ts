/**
 * KeyTips — Excel's Alt-driven "press a letter to invoke a ribbon item"
 * affordance. The state machine here drives the *behavior*; the visual
 * overlay (badge labels appended to toolbar items) is wired by the host.
 *
 * Lifecycle:
 *   idle  ─Alt held─▶  showing  ─letter pressed─▶  invoked
 *      ▲                  │
 *      └── Alt released ──┘
 *
 * Each step is pure: callers feed input events through the reducer and
 * read back the next state plus an optional command id to dispatch. No
 * DOM, no timers, fully unit-testable.
 */

export type KeyTipState =
  | { kind: 'idle' }
  | { kind: 'showing'; chord: string }
  | { kind: 'invoked'; commandId: string };

export interface KeyTipBinding {
  /** Single-letter (`F`, `H`, etc.) or chord (`F1`, `HC`). Compared
   *  case-insensitively. */
  chord: string;
  /** Stable id the host dispatches. */
  commandId: string;
}

export type KeyTipEvent =
  | { kind: 'altDown' }
  | { kind: 'altUp' }
  | { kind: 'letter'; key: string }
  | { kind: 'escape' };

export interface KeyTipReducerOutput {
  state: KeyTipState;
  /** Set when the latest event invoked a binding — host should dispatch
   *  the command immediately and call `reduce({ kind: 'altUp' })` to
   *  return to idle. */
  invoked?: string;
}

const initialState: KeyTipState = { kind: 'idle' };

/** Step the state machine once. The reducer is a pure function. */
export function reduceKeyTip(
  state: KeyTipState,
  event: KeyTipEvent,
  bindings: readonly KeyTipBinding[],
): KeyTipReducerOutput {
  if (event.kind === 'altUp' || event.kind === 'escape') {
    return { state: initialState };
  }
  if (event.kind === 'altDown') {
    if (state.kind === 'idle') return { state: { kind: 'showing', chord: '' } };
    return { state };
  }
  // Letter event — only meaningful while showing.
  if (state.kind !== 'showing') return { state };

  const nextChord = (state.chord + event.key).toUpperCase();
  const direct = bindings.find((b) => b.chord.toUpperCase() === nextChord);
  // A prefix-extension exists when at least one OTHER binding starts with
  // `nextChord` but is longer. Excel's KeyTips: pressing `H` while both
  // `H` and `HC` are bound holds the chord open instead of firing `H` —
  // the user has to press a second letter (or release Alt) to commit.
  const extension = bindings.some((b) => {
    const c = b.chord.toUpperCase();
    return c.length > nextChord.length && c.startsWith(nextChord);
  });
  if (direct && !extension) {
    return {
      state: { kind: 'invoked', commandId: direct.commandId },
      invoked: direct.commandId,
    };
  }
  if (direct || extension) {
    return { state: { kind: 'showing', chord: nextChord } };
  }
  // Dead end — unknown letter. Reset chord buffer to drop the bad input
  // but stay in showing mode; Excel parity.
  return { state: { kind: 'showing', chord: '' } };
}

/** Filter bindings by the current chord prefix — the caller renders these
 *  as the visible KeyTip badges. */
export function visibleBindings(
  state: KeyTipState,
  bindings: readonly KeyTipBinding[],
): KeyTipBinding[] {
  if (state.kind !== 'showing') return [];
  if (state.chord.length === 0) return [...bindings];
  const prefix = state.chord.toUpperCase();
  return bindings.filter((b) => b.chord.toUpperCase().startsWith(prefix));
}
