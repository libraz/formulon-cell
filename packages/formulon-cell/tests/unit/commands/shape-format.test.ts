import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  createSessionShape,
  selectedSessionIllustration,
  sessionIllustrationById,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionDrawingFlip,
  applySessionShapeFill,
  applySessionShapeKind,
  applySessionShapeOutline,
  applySessionShapeOutlineDash,
  applySessionShapeOutlineWeight,
  applySessionShapeRotation,
  DEFAULT_SHAPE_COLOR,
  DEFAULT_SHAPE_LINE_WIDTH,
  DEFAULT_SHAPE_OUTLINE_COLOR,
  shapeDashStroke,
  shapeFillColor,
  shapeOutlineColor,
  shapeOutlineDash,
  shapeOutlineWeight,
} from '../../../src/commands/shape-format.js';
import {
  createSpreadsheetStore,
  mutators,
  type SessionIllustration,
  type SessionShapeKind,
  type SpreadsheetStore,
} from '../../../src/store/store.js';

const range = { sheet: 0, r0: 1, c0: 1, r1: 1, c1: 1 };

describe('shape format commands', () => {
  let store: SpreadsheetStore;
  let history: History;

  beforeEach(() => {
    store = createSpreadsheetStore();
    history = new History();
  });

  const shape = (kind: SessionShapeKind = 'rectangle'): SessionIllustration => {
    const created = createSessionShape(store, range, { id: `shape-${kind}`, shape: kind });
    if (!created) throw new Error('shape was not created');
    return created;
  };
  const current = (id: string): SessionIllustration => {
    const found = sessionIllustrationById(store.getState(), id);
    if (!found) throw new Error('shape is gone');
    return found;
  };

  it('finds the selected picture or shape on the active sheet only', () => {
    const item = shape();
    expect(selectedSessionIllustration(store.getState())).toBeNull();
    mutators.setSelectedObject(store, { id: item.id, name: 'Rectangle 1' });
    expect(selectedSessionIllustration(store.getState())?.id).toBe(item.id);
    mutators.setSelectedObject(store, { id: 'chart-native-0-1', name: 'Chart 1' });
    expect(selectedSessionIllustration(store.getState())).toBeNull();
  });

  it('fills a shape without moving its outline, one undo step at a time', () => {
    const item = shape();
    expect(shapeFillColor(item)).toBe(DEFAULT_SHAPE_COLOR);
    expect(shapeOutlineColor(item)).toBe(DEFAULT_SHAPE_OUTLINE_COLOR);

    applySessionShapeFill(store, item, '#ff0000', history);
    expect(shapeFillColor(current(item.id))).toBe('#ff0000');
    expect(current(item.id).opacity).toBe(1);
    expect(shapeOutlineColor(current(item.id))).toBe(DEFAULT_SHAPE_OUTLINE_COLOR);

    applySessionShapeFill(store, current(item.id), null, history);
    expect(shapeFillColor(current(item.id))).toBeNull();
    expect(shapeOutlineColor(current(item.id))).toBe(DEFAULT_SHAPE_OUTLINE_COLOR);

    history.undo();
    expect(shapeFillColor(current(item.id))).toBe('#ff0000');
    history.undo();
    expect(current(item.id)).toEqual(item);
  });

  it('takes the outline away and brings it back at its weight', () => {
    const item = shape();
    applySessionShapeOutline(store, item, null, history);
    expect(shapeOutlineColor(current(item.id))).toBeNull();
    expect(shapeFillColor(current(item.id))).toBe(DEFAULT_SHAPE_COLOR);

    applySessionShapeOutline(store, current(item.id), '#00b050', history);
    expect(shapeOutlineColor(current(item.id))).toBe('#00b050');
    expect(current(item.id).lineWidth).toBeUndefined();
    expect(shapeFillColor(current(item.id))).toBe(DEFAULT_SHAPE_COLOR);
  });

  it('keeps the weight while the outline is off and restores it', () => {
    const item = shape();
    applySessionShapeOutlineWeight(store, item, 3, history);
    applySessionShapeOutline(store, current(item.id), null, history);
    expect(shapeOutlineColor(current(item.id))).toBeNull();
    expect(shapeOutlineWeight(current(item.id))).toBeNull();
    expect(current(item.id).lineWidth).toBe(3);

    applySessionShapeOutline(store, current(item.id), '#00b050', history);
    expect(shapeOutlineWeight(current(item.id))).toBe(3);
    expect(shapeOutlineColor(current(item.id))).toBe('#00b050');
  });

  it('turns the outline on at the default weight when none was ever set', () => {
    const item = shape();
    applySessionShapeOutline(store, item, null, history);
    applySessionShapeOutline(store, current(item.id), '#00b050', history);
    expect(shapeOutlineWeight(current(item.id))).toBe(DEFAULT_SHAPE_LINE_WIDTH);
  });

  it('weighs and dashes the outline, turning a missing outline back on', () => {
    const item = shape();
    expect(shapeOutlineWeight(item)).toBe(DEFAULT_SHAPE_LINE_WIDTH);
    expect(shapeOutlineDash(item)).toBe('solid');

    applySessionShapeOutline(store, item, null, history);
    applySessionShapeOutlineWeight(store, current(item.id), 4.5, history);
    expect(shapeOutlineWeight(current(item.id))).toBe(4.5);
    expect(shapeOutlineColor(current(item.id))).toBe(DEFAULT_SHAPE_OUTLINE_COLOR);

    applySessionShapeOutline(store, current(item.id), null, history);
    applySessionShapeOutlineDash(store, current(item.id), 'longDashDot', history);
    expect(shapeOutlineDash(current(item.id))).toBe('longDashDot');
    expect(shapeOutlineWeight(current(item.id))).toBe(4.5);

    applySessionShapeOutlineDash(store, current(item.id), 'solid', history);
    expect(current(item.id).lineDash).toBeUndefined();
    expect(applySessionShapeOutlineWeight(store, current(item.id), 0, history)).toBeNull();

    history.undo();
    expect(shapeOutlineDash(current(item.id))).toBe('longDashDot');
  });

  it('dashes in multiples of the line width', () => {
    expect(shapeDashStroke('solid', 2)).toEqual({ dasharray: 'none', linecap: 'butt' });
    expect(shapeDashStroke('roundDot', 2)).toEqual({ dasharray: '0 4', linecap: 'round' });
    expect(shapeDashStroke('longDashDotDot', 2).dasharray).toBe('16 6 2 6 2 6');
  });

  it('gives a line an outline and no fill', () => {
    const line = shape('line');
    expect(shapeOutlineColor(line)).toBe(DEFAULT_SHAPE_COLOR);
    expect(shapeFillColor(line)).toBeNull();
    expect(applySessionShapeFill(store, line, '#ff0000', history)).toBeNull();
    expect(history.canUndo()).toBe(false);

    applySessionShapeOutline(store, line, '#7030a0', history);
    expect(shapeOutlineColor(current(line.id))).toBe('#7030a0');
  });

  it('turns a shape a quarter turn each way, one undo step each', () => {
    const item = shape();
    applySessionShapeRotation(store, item, 90, history);
    expect(current(item.id).rotationDeg).toBe(90);
    applySessionShapeRotation(store, current(item.id), -90, history);
    applySessionShapeRotation(store, current(item.id), -90, history);
    expect(current(item.id).rotationDeg).toBe(270);

    history.undo();
    expect(current(item.id).rotationDeg).toBe(0);
    history.undo();
    history.undo();
    expect(current(item.id)).toEqual(item);
  });

  it('turns a rotation the file stored past a full turn into the range drawn', () => {
    const item = shape();
    mutators.updateIllustration(store, item.id, { rotationDeg: 400 });
    applySessionShapeRotation(store, current(item.id), 90, history);
    expect(current(item.id).rotationDeg).toBe(130);
  });

  it('flips a shape on either axis, mirroring its rotation, one undo step each', () => {
    const item = shape();
    mutators.updateIllustration(store, item.id, { rotationDeg: 30 });
    applySessionDrawingFlip(store, current(item.id), 'horizontal', history);
    expect(current(item.id)).toMatchObject({ flipH: true, rotationDeg: 330 });
    expect(current(item.id).flipV).toBeUndefined();
    applySessionDrawingFlip(store, current(item.id), 'vertical', history);
    expect(current(item.id)).toMatchObject({ flipH: true, flipV: true, rotationDeg: 30 });
    applySessionDrawingFlip(store, current(item.id), 'horizontal', history);
    expect(current(item.id).flipH).toBeUndefined();
    expect(current(item.id).flipV).toBe(true);

    history.undo();
    expect(current(item.id)).toMatchObject({ flipH: true, flipV: true, rotationDeg: 30 });
    history.undo();
    history.undo();
    expect(current(item.id).flipH).toBeUndefined();
    expect(current(item.id).rotationDeg).toBe(30);
  });

  it('keeps an unrotated flipped shape at no rotation', () => {
    const item = shape();
    applySessionDrawingFlip(store, item, 'vertical', history);
    expect(current(item.id)).toMatchObject({ flipV: true, rotationDeg: 0 });
  });

  it('changes the kind of a shape, keeping its paint, text and frame', () => {
    const created = createSessionShape(store, range, {
      id: 'box',
      shape: 'rectangle',
      x: 40,
      y: 30,
      w: 120,
      h: 60,
      color: '#ff0000',
      lineWidth: 3,
      extra: { text: 'Note', textBox: true, lineColor: '#00ff00', rotationDeg: 30 },
    });
    if (!created) throw new Error('shape was not created');

    applySessionShapeKind(store, created, 'oval', history);
    expect(current(created.id)).toEqual({ ...created, shape: 'oval' });
    history.undo();
    expect(current(created.id)).toEqual(created);
  });

  it('changes no line, and makes nothing into one', () => {
    const line = shape('line');
    expect(applySessionShapeKind(store, line, 'oval', history)).toBeNull();
    const box = shape();
    expect(applySessionShapeKind(store, box, 'line', history)).toBeNull();
    expect(applySessionShapeKind(store, box, 'arrow', history)).toBeNull();
    expect(history.canUndo()).toBe(false);
  });

  it('leaves a shape drawn only as its bounding box as it is', () => {
    const box = shape();
    mutators.updateIllustration(store, box.id, { nativeLocked: true });
    expect(applySessionShapeKind(store, current(box.id), 'oval', history)).toBeNull();
    expect(applySessionShapeRotation(store, current(box.id), 90, history)).toBeNull();
    expect(applySessionDrawingFlip(store, current(box.id), 'horizontal', history)).toBeNull();
    expect(history.canUndo()).toBe(false);
  });
});
