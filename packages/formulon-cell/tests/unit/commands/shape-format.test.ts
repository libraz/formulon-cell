import { beforeEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  createSessionShape,
  selectedSessionIllustration,
  sessionIllustrationById,
} from '../../../src/commands/session-illustration.js';
import {
  applySessionShapeFill,
  applySessionShapeOutline,
  applySessionShapeOutlineDash,
  applySessionShapeOutlineWeight,
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
    expect(shapeOutlineWeight(current(item.id))).toBe(DEFAULT_SHAPE_LINE_WIDTH);

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
});
