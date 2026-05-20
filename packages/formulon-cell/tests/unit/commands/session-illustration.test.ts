import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import { setProtectedSheet } from '../../../src/commands/protection.js';
import {
  arrangeSessionIllustration,
  clearSessionIllustration,
  createRibbonImageFromSelection,
  createRibbonShapeFromSelection,
  createSessionImage,
  createSessionShape,
  listSessionIllustrations,
  sessionIllustrationById,
  updateSessionIllustration,
} from '../../../src/commands/session-illustration.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const range = (r0: number, c0: number, r1: number, c1: number) =>
  ({ sheet: 0, r0, c0, r1, c1 }) as const;

describe('session illustration commands', () => {
  it('creates reusable session shape and image overlays with stable defaults', () => {
    const store = createSpreadsheetStore();
    const r = range(2, 1, 2, 1);

    const shape = createSessionShape(store, r, { shape: 'oval' });
    const image = createSessionImage(store, r, { src: 'data:image/png;base64,a', alt: 'a.png' });

    expect(shape).toEqual({
      id: 'shape-0-2-1-oval',
      kind: 'shape',
      shape: 'oval',
      sheet: 0,
      x: undefined,
      y: undefined,
      w: undefined,
      h: undefined,
      color: undefined,
    });
    expect(image).toEqual({
      id: 'image-0-2-1',
      kind: 'image',
      src: 'data:image/png;base64,a',
      alt: 'a.png',
      sheet: 0,
      x: undefined,
      y: undefined,
      w: undefined,
      h: undefined,
    });
    expect(listSessionIllustrations(store.getState())).toEqual([shape, image]);
  });

  it('creates ribbon images and shapes with Excel-like overlay placement defaults', () => {
    const store = createSpreadsheetStore();
    const r = range(0, 0, 0, 0);

    const shape = createRibbonShapeFromSelection(store, r, 'arrow');
    const image = createRibbonImageFromSelection(
      store,
      r,
      'data:image/png;base64,b',
      null,
      'b.png',
    );

    expect(shape).toMatchObject({
      id: 'ribbon-shape-0-0-0-arrow-0',
      kind: 'shape',
      shape: 'arrow',
      sheet: 0,
      x: 300,
      y: 88,
      w: 180,
      h: 80,
      color: '#0f6cbd',
    });
    expect(image).toMatchObject({
      id: 'ribbon-image-0-0-0-1',
      kind: 'image',
      src: 'data:image/png;base64,b',
      alt: 'b.png',
      sheet: 0,
      x: 324,
      y: 112,
      w: 240,
      h: 160,
    });
  });

  it('lists, finds, updates, and clears session illustrations for host object panes', () => {
    const store = createSpreadsheetStore();
    const shape = createSessionShape(store, range(0, 0, 0, 0), { id: 'shape-a', shape: 'line' });
    createSessionImage(store, range(1, 0, 1, 0), { id: 'image-a', src: 'https://example.test/a' });

    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'shape-a',
      'image-a',
    ]);
    expect(sessionIllustrationById(store.getState(), 'shape-a')).toEqual(shape);
    expect(sessionIllustrationById(store.getState(), 'missing')).toBeNull();
    expect(updateSessionIllustration(store, 'missing', { x: 10 })).toBeNull();

    expect(updateSessionIllustration(store, 'shape-a', { x: 24, y: 40, w: 160 })).toMatchObject({
      id: 'shape-a',
      x: 24,
      y: 40,
      w: 160,
    });
    expect(clearSessionIllustration(store, 'missing')).toBe(false);
    expect(clearSessionIllustration(store, 'image-a')).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);
  });

  it('rejects create, update, and clear on protected sheets', () => {
    const store = createSpreadsheetStore();
    const shape = createSessionShape(store, range(0, 0, 0, 0), { id: 'shape-a', shape: 'line' });
    expect(shape).not.toBeNull();
    setProtectedSheet(store, 0, true);

    expect(
      createSessionShape(store, range(1, 0, 1, 0), { id: 'shape-b', shape: 'oval' }),
    ).toBeNull();
    expect(createSessionImage(store, range(1, 0, 1, 0), { id: 'image-a', src: 'x' })).toBeNull();
    expect(updateSessionIllustration(store, 'shape-a', { x: 32 })).toBeNull();
    expect(clearSessionIllustration(store, 'shape-a')).toBe(false);
    expect(listSessionIllustrations(store.getState())).toEqual([shape]);
  });

  it('records create, update, and clear in history', () => {
    const store = createSpreadsheetStore();
    const history = new History();

    const shape = createSessionShape(
      store,
      range(0, 0, 0, 0),
      { id: 'shape-a', shape: 'line' },
      history,
    );
    expect(shape).not.toBeNull();
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);
    expect(history.undo()).toBe(true);
    expect(listSessionIllustrations(store.getState())).toEqual([]);
    expect(history.redo()).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);

    expect(updateSessionIllustration(store, 'shape-a', { x: 10 }, history)).toMatchObject({
      x: 10,
    });
    expect(history.undo()).toBe(true);
    expect(sessionIllustrationById(store.getState(), 'shape-a')?.x).toBeUndefined();
    expect(history.redo()).toBe(true);
    expect(sessionIllustrationById(store.getState(), 'shape-a')?.x).toBe(10);

    expect(clearSessionIllustration(store, 'shape-a', history)).toBe(true);
    expect(listSessionIllustrations(store.getState())).toEqual([]);
    expect(history.undo()).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual(['shape-a']);
  });

  it('arranges session illustrations within their sheet and records history', () => {
    const store = createSpreadsheetStore();
    const history = new History();
    createSessionShape(store, range(0, 0, 0, 0), {
      id: 'back',
      shape: 'rectangle',
    });
    createSessionShape(store, range(0, 1, 0, 1), {
      id: 'middle',
      shape: 'oval',
    });
    createSessionShape(
      store,
      { sheet: 1, r0: 0, c0: 0, r1: 0, c1: 0 },
      {
        id: 'other-sheet',
        shape: 'line',
      },
    );
    createSessionImage(store, range(0, 2, 0, 2), { id: 'front', src: 'x' });

    expect(arrangeSessionIllustration(store, 'middle', 'bring-forward', history)).toMatchObject({
      id: 'middle',
    });
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'back',
      'front',
      'other-sheet',
      'middle',
    ]);

    expect(arrangeSessionIllustration(store, 'middle', 'send-back', history)).toMatchObject({
      id: 'middle',
    });
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'middle',
      'back',
      'other-sheet',
      'front',
    ]);

    expect(history.undo()).toBe(true);
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'back',
      'front',
      'other-sheet',
      'middle',
    ]);
  });

  it('ignores arrange requests that cannot change order or target protected sheets', () => {
    const store = createSpreadsheetStore();
    createSessionShape(store, range(0, 0, 0, 0), {
      id: 'shape-a',
      shape: 'rectangle',
    });
    createSessionShape(store, range(0, 1, 0, 1), {
      id: 'shape-b',
      shape: 'oval',
    });

    expect(arrangeSessionIllustration(store, 'missing', 'bring-front')).toBeNull();
    expect(arrangeSessionIllustration(store, 'shape-a', 'send-back')).toBeNull();
    setProtectedSheet(store, 0, true);
    expect(arrangeSessionIllustration(store, 'shape-b', 'send-backward')).toBeNull();
    expect(listSessionIllustrations(store.getState()).map((item) => item.id)).toEqual([
      'shape-a',
      'shape-b',
    ]);
  });
});
