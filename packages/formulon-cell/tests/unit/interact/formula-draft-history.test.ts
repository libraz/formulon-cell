import { describe, expect, it } from 'vitest';
import {
  FormulaDraftHistory,
  type FormulaDraftSnapshot,
} from '../../../src/interact/formula-draft-history.js';

const snapshot = (
  value: string,
  caret = value.length,
  mode: 'edit' | 'point' = 'edit',
): FormulaDraftSnapshot => ({
  value,
  start: caret,
  end: caret,
  direction: 'none',
  mode,
});

describe('FormulaDraftHistory', () => {
  it('records one browser input transaction and restores its caret and mode', () => {
    const history = new FormulaDraftHistory(snapshot('=1+'));
    history.beforeInput(snapshot('=1+'));
    history.input(snapshot('=1+2', 4, 'point'));

    expect(history.undo()).toEqual(snapshot('=1+'));
    expect(history.redo()).toEqual(snapshot('=1+2', 4, 'point'));
  });

  it('uses the last observed snapshot when beforeinput is missing', () => {
    const history = new FormulaDraftHistory(snapshot('=1+'));
    history.input(snapshot('=1+2'));

    expect(history.undo()).toEqual(snapshot('=1+'));
  });

  it('groups pointer and keyboard reference updates into one transaction', () => {
    const history = new FormulaDraftHistory(snapshot('=1+'));
    history.beginGroup(snapshot('=1+'), 'reference');
    history.updateGroup(snapshot('=1+C4', 5, 'point'), 'reference');
    history.updateGroup(snapshot('=1+C4:D5', 10, 'point'), 'reference');
    history.updateGroup(snapshot('=1+A4:C5', 10, 'point'), 'reference');
    history.endGroup();

    expect(history.undo()).toEqual(snapshot('=1+'));
    expect(history.redo()).toEqual(snapshot('=1+A4:C5', 10, 'point'));
  });

  it('records an IME composition once and folds the final composition input into it', () => {
    const history = new FormulaDraftHistory(snapshot(''));
    history.startComposition(snapshot(''));
    history.input(snapshot('あ', 1), { composing: true });
    history.endComposition(snapshot('あ', 1));
    history.beforeInput(snapshot('あ', 1));
    history.input(snapshot('あい', 2), { inputType: 'insertFromComposition' });

    expect(history.undo()).toEqual(snapshot(''));
    expect(history.redo()).toEqual(snapshot('あい', 2));
    expect(history.undo()).toEqual(snapshot(''));
  });

  it('folds a same-text final composition input without creating a second edit', () => {
    const history = new FormulaDraftHistory(snapshot(''));
    history.startComposition(snapshot(''));
    history.input(snapshot('あ', 1), { composing: true });
    history.endComposition(snapshot('あ', 1));
    history.beforeInput(snapshot('あ', 1));
    history.input(
      { ...snapshot('あ', 0), end: 0, direction: 'backward' },
      { inputType: 'insertFromComposition' },
    );

    expect(history.undo()).toEqual(snapshot(''));
    expect(history.undo()).toBeNull();
  });

  it('drops the redo branch after a new edit and bounds the undo stack', () => {
    const history = new FormulaDraftHistory(snapshot(''));
    history.recordCommand(snapshot(''), snapshot('a'), 'command');
    history.recordCommand(snapshot('a'), snapshot(''), 'command');
    expect(history.undo()).toEqual(snapshot('a'));
    history.recordCommand(snapshot('a'), snapshot('b'), 'command');
    expect(history.redo()).toBeNull();

    const bounded = new FormulaDraftHistory(snapshot(''));
    for (let index = 0; index < 105; index += 1) {
      const before = snapshot(String(index));
      bounded.recordCommand(before, snapshot(String(index + 1)), 'command');
    }
    for (let index = 0; index < 100; index += 1) expect(bounded.undo()).not.toBeNull();
    expect(bounded.undo()).toBeNull();
  });

  it('treats same-value commands as no-ops without dropping redo', () => {
    const history = new FormulaDraftHistory(snapshot('abc'));
    history.recordCommand(snapshot('abc'), snapshot('abcd'), 'command');
    expect(history.undo()).toEqual(snapshot('abc'));
    history.recordCommand(snapshot('abc'), snapshot('abc', 1, 'point'), 'noop');
    expect(history.redo()).toEqual(snapshot('abcd'));
  });

  it('updates selection baselines without adding entries and restores private checkpoints', () => {
    const history = new FormulaDraftHistory(snapshot('=A1'));
    history.beforeInput(snapshot('=A1'));
    history.input(snapshot('=A1+B1', 6));
    history.observeSelection({ ...snapshot('=A1+B1', 3), end: 5, direction: 'backward' });
    const checkpoint = history.captureState();
    history.recordCommand(snapshot('=A1+B1', 3), snapshot('=A1+C4', 6, 'point'), 'reference');
    history.restoreState(checkpoint);

    expect(history.undo()).toEqual(snapshot('=A1'));
    expect(history.redo()).toEqual(snapshot('=A1+B1', 6));
  });
});
