import { AnchorEditAs, AnchorKind, DrawingObjectKind, ShapePaint } from '@libraz/formulon';
import {
  DEFAULT_SHAPE_COLOR,
  shapeFillColor,
  shapeHasFill,
  shapeOutlineColor,
  shapeOutlineWeight,
} from '../commands/shape-format.js';
import type {
  SessionIllustration,
  SessionShapeKind,
  SpreadsheetStore,
  State,
} from '../store/store.js';
import {
  ANCHOR_KIND_ABSOLUTE,
  drawingAnchorOf,
  placementForPixels,
  sessionPosition,
  twoCellAnchorForPixels,
} from './drawing-geometry.js';
import { argbFromCssColor, cssColorFromArgb } from './theme-color.js';
import type { EngineDrawingAnchor } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { EngineShape, EngineShapeInput } from './workbook-handle-drawing-objects.js';

/** Where a shape without a frame is drawn, in grid pixels. */
export const DEFAULT_ILLUSTRATION_FRAME = { x: 300, y: 88, w: 160, h: 96 } as const;

const EMU_PER_PT = 12700;

/** Outline of a shape the overlay can only approximate. */
const LOCKED_OUTLINE_COLOR = '#7f7f7f';

/** The preset geometry each saved shape kind is written as. An arrow has no
 *  entry: the engine cannot express a line end. */
const PRST_OF_KIND: Partial<Record<SessionShapeKind, string>> = {
  rectangle: 'rect',
  'rounded-rectangle': 'roundRect',
  oval: 'ellipse',
  triangle: 'triangle',
  diamond: 'diamond',
  line: 'line',
};

const KIND_OF_PRST = new Map<string, SessionShapeKind>([
  ...(Object.entries(PRST_OF_KIND) as [SessionShapeKind, string][]).map(
    ([kind, prst]) => [prst, kind] as const,
  ),
  ['straightConnector1', 'line'],
]);

/** Whether the engine can store `item`. */
export const shapeIsSaved = (item: SessionIllustration): boolean =>
  item.kind === 'shape' && item.shape !== undefined && PRST_OF_KIND[item.shape] !== undefined;

const isLineKind = (kind: SessionShapeKind | undefined): boolean =>
  kind === 'line' || kind === 'arrow';

const rgbOf = (color: string | null | undefined): number | null => {
  const argb = color ? argbFromCssColor(color) : null;
  return argb === null ? null : argb & 0xffffff;
};

const anchorInput = (anchor: EngineDrawingAnchor): EngineShapeInput => ({
  anchorKind: anchor.anchorKind,
  editAs: anchor.editAs,
  fromRow: anchor.fromRow,
  fromCol: anchor.fromCol,
  fromRowOff: anchor.fromRowOff,
  fromColOff: anchor.fromColOff,
  toRow: anchor.toRow,
  toCol: anchor.toCol,
  toRowOff: anchor.toRowOff,
  toColOff: anchor.toColOff,
  cx: anchor.cx,
  cy: anchor.cy,
});

const sameAnchor = (a: EngineDrawingAnchor | undefined, b: EngineDrawingAnchor | undefined) =>
  a === b || (!!a && !!b && JSON.stringify(a) === JSON.stringify(b));

/** The fill and outline fields that write `item`'s paint. */
const paintInput = (item: SessionIllustration): EngineShapeInput => {
  const input: EngineShapeInput = {};
  if (shapeHasFill(item)) {
    const rgb = rgbOf(shapeFillColor(item) ?? undefined);
    if (item.opacity === 0 || shapeFillColor(item) === null) input.fillKind = ShapePaint.None;
    else {
      input.fillKind = ShapePaint.Solid;
      input.fillRgb = rgb ?? (rgbOf(DEFAULT_SHAPE_COLOR) as number);
    }
  }
  const weight = shapeOutlineWeight(item);
  const outline = rgbOf(shapeOutlineColor(item));
  if (weight === null || outline === null) input.lineKind = ShapePaint.None;
  else {
    input.lineKind = ShapePaint.Solid;
    input.lineRgb = outline;
    input.lineWidthEmu = Math.round(weight * EMU_PER_PT);
  }
  return input;
};

const pick = <K extends keyof EngineShapeInput>(
  input: EngineShapeInput,
  keys: readonly K[],
): EngineShapeInput =>
  Object.fromEntries(keys.flatMap((key) => (input[key] === undefined ? [] : [[key, input[key]]])));

/** Overlay-only settings of a saved shape, reset so the screen shows what the
 *  file holds: the file has no corner radius, dash style or partial opacity
 *  (a partly transparent fill is saved opaque). */
const unsaveable = (item: SessionIllustration): Partial<SessionIllustration> => {
  const reset: Partial<SessionIllustration> = {};
  if (item.nativeLocked) return reset;
  if (item.radius !== undefined) reset.radius = undefined;
  if (item.lineDash !== undefined) reset.lineDash = undefined;
  if (item.opacity !== undefined && item.opacity !== 0 && item.opacity !== 1) reset.opacity = 1;
  return reset;
};

/** The overlay for a shape of the file, or null for a hidden one, which stays
 *  untouched in the engine. A shape the overlay cannot draw (unknown preset,
 *  gradient or picture paint, rotation, group) becomes a locked bounding box. */
export function shapeFromNative(
  store: SpreadsheetStore,
  sheet: number,
  shape: EngineShape,
): SessionIllustration | null {
  if (shape.hidden) return null;
  const kind = KIND_OF_PRST.get(shape.prst);
  // The engine reports no flip, so a loaded line's direction is unknown.
  const locked =
    kind === undefined ||
    kind === 'line' ||
    shape.kind === DrawingObjectKind.Connector ||
    shape.kind === DrawingObjectKind.Group ||
    shape.rotation !== 0 ||
    shape.fillKind === ShapePaint.Other ||
    shape.lineKind === ShapePaint.Other;
  const anchor = drawingAnchorOf(shape);
  const item: SessionIllustration = {
    id: `shape-native-${sheet}-${shape.objectId}`,
    kind: 'shape',
    shape: locked || kind === undefined ? 'rectangle' : kind,
    sheet,
    ...sessionPosition(store, anchor),
    nativeObjectId: shape.objectId,
    nativeAnchor: anchor,
    nativeGeometryEditable: shape.anchorKind !== ANCHOR_KIND_ABSOLUTE,
  };
  if (locked) {
    return {
      ...item,
      nativeLocked: true,
      opacity: 0,
      lineColor: LOCKED_OUTLINE_COLOR,
      lineWidth: 1,
      lineDash: 'dash',
    };
  }
  if (shape.fillKind === ShapePaint.Solid) item.color = cssColorFromArgb(shape.fillRgb);
  else if (shape.fillKind === ShapePaint.None) item.opacity = 0;
  if (shape.lineKind === ShapePaint.None) item.lineWidth = 0;
  else if (shape.lineKind === ShapePaint.Solid) {
    item.lineColor = cssColorFromArgb(shape.lineRgb);
    if (shape.lineWidthEmu > 0) item.lineWidth = shape.lineWidthEmu / EMU_PER_PT;
  }
  if (shape.text) item.text = shape.text;
  return item;
}

/** A shape as an insert spec, or null when re-inserting it would lose paint
 *  or geometry the engine cannot take back. */
const shapeInputOf = (shape: EngineShape): EngineShapeInput | null => {
  if (
    shape.prst === '' ||
    shape.kind === DrawingObjectKind.Group ||
    shape.fillKind === ShapePaint.Other ||
    shape.lineKind === ShapePaint.Other
  ) {
    return null;
  }
  const input: EngineShapeInput = {
    ...anchorInput(drawingAnchorOf(shape)),
    name: shape.name,
    prst: shape.prst,
    rotation: shape.rotation,
    hidden: shape.hidden,
  };
  if (shape.kind !== DrawingObjectKind.Connector) {
    input.text = shape.text;
    input.textBox = shape.textBox;
  }
  if (shape.fillKind === ShapePaint.Solid) {
    input.fillKind = ShapePaint.Solid;
    input.fillRgb = shape.fillRgb;
  } else if (shape.fillKind === ShapePaint.None) input.fillKind = ShapePaint.None;
  if (shape.lineKind === ShapePaint.Solid) {
    input.lineKind = ShapePaint.Solid;
    input.lineRgb = shape.lineRgb;
    input.lineWidthEmu = shape.lineWidthEmu;
  } else if (shape.lineKind === ShapePaint.None) input.lineKind = ShapePaint.None;
  return input;
};

/** Capture shape `objectId` so a row or column delete that drops it can be
 *  undone. The restorer inserts it again; null when it cannot be restored
 *  without loss. The engine's `setDrawingZOrder` refuses shapes, so a restored
 *  shape goes on top of the drawing order. */
export function captureShapeRestore(
  wb: WorkbookHandle,
  sheet: number,
  objectId: number,
): (() => boolean) | null {
  const shape = wb.listShapes(sheet).find((candidate) => candidate.objectId === objectId);
  const input = shape ? shapeInputOf(shape) : null;
  return input ? () => wb.insertShape(sheet, input) !== null : null;
}

/** Capture the anchors of `sheet`'s shapes ahead of a row or column delete,
 *  which collapses the shapes it does not remove. The restorer puts the
 *  survivors back where they were once the delete is undone. */
export function captureShapeAnchors(wb: WorkbookHandle, sheet: number): () => void {
  if (!wb.capabilities.shapes) return () => undefined;
  const anchors = new Map(
    wb
      .listShapes(sheet)
      .filter((shape) => shape.anchorKind !== ANCHOR_KIND_ABSOLUTE)
      .map((shape) => [shape.objectId, drawingAnchorOf(shape)] as const),
  );
  return () => {
    for (const shape of wb.listShapes(sheet)) {
      const anchor = anchors.get(shape.objectId);
      if (anchor && !sameAnchor(anchor, drawingAnchorOf(shape))) {
        wb.updateShape(sheet, shape.objectId, anchorInput(anchor));
      }
    }
  };
}

/** Overlay shapes mirrored into the engine. Items are matched by overlay id;
 *  the engine object id is kept per item, so a history snapshot that restores
 *  an item re-creates its shape. Hydration runs with the mirror muted. */
export class ShapeMirror {
  constructor(private readonly handle: WorkbookHandle) {}

  private last: State['illustrations'] | null = null;

  /** Engine object id of each mirrored overlay item. */
  private readonly objectIds = new Map<string, number>();

  /** The engine's description of each shape removed through the mirror, so a
   *  later undo re-inserts it as it was (name, text, anchor kind, geometry).
   *  The engine cannot reorder shapes, so a restored shape goes on top. */
  private readonly removed = new Map<string, { objectId: number; input: EngineShapeInput }>();

  /** Adopt `state` as what the engine already carries. */
  rebase(state: State | null): void {
    this.last = state?.illustrations ?? null;
    this.objectIds.clear();
    for (const item of state?.illustrations.illustrations ?? []) {
      if (item.kind === 'shape' && item.nativeObjectId !== undefined) {
        this.objectIds.set(item.id, item.nativeObjectId);
      }
    }
  }

  sync(store: SpreadsheetStore): void {
    const slice = store.getState().illustrations;
    if (slice === this.last) return;
    const previous = this.last?.illustrations ?? [];
    this.last = slice;
    const wb = this.handle;
    if (!wb.capabilities.shapes) return;
    const before = new Map(previous.map((item) => [item.id, item] as const));
    const after = new Set(slice.illustrations.map((item) => item.id));
    for (const item of previous) {
      if (item.kind !== 'shape' || after.has(item.id)) continue;
      const objectId = this.objectIds.get(item.id);
      if (objectId === undefined) continue;
      this.forget(item.id, item.sheet, objectId);
    }
    // null drops the item: the engine refused it.
    const patches = new Map<string, Partial<SessionIllustration> | null>();
    for (const item of slice.illustrations) {
      const old = before.get(item.id);
      if (old === item || item.kind !== 'shape' || !shapeIsSaved(item)) continue;
      const objectId = this.objectIds.get(item.id);
      if (objectId === undefined || (old && old.sheet !== item.sheet)) {
        if (objectId !== undefined && old) this.forget(old.id, old.sheet, objectId);
        const tied = this.insert(store, item);
        patches.set(item.id, tied && { ...tied, ...unsaveable(item) });
      } else if (old?.kind === 'shape') {
        const patch = { ...this.update(store, objectId, old, item), ...unsaveable(item) };
        if (Object.keys(patch).length > 0) patches.set(item.id, patch);
      }
    }
    if (patches.size === 0) return;
    store.setState((state) => {
      const illustrations = state.illustrations.illustrations.flatMap((item) => {
        const patch = patches.get(item.id);
        if (patch === undefined) return [item];
        return patch === null ? [] : [{ ...item, ...patch }];
      });
      const next = { ...state, illustrations: { illustrations } };
      this.last = next.illustrations;
      return next;
    });
  }

  /** Remove the engine shape of an overlay item, keeping its description. */
  private forget(id: string, sheet: number, objectId: number): void {
    const shape = this.handle
      .listShapes(sheet)
      .find((candidate) => candidate.objectId === objectId);
    const description = shape ? shapeInputOf(shape) : null;
    if (description) this.removed.set(id, { objectId, input: description });
    this.handle.removeShape(sheet, objectId);
    this.objectIds.delete(id);
  }

  /** Insert `item`; returns the fields that tie it to the new shape, or null. */
  private insert(store: SpreadsheetStore, item: SessionIllustration) {
    const wb = this.handle;
    // Only the item that was removed comes back as described; a new item reusing its id does not.
    const removed = this.removed.get(item.id);
    this.removed.delete(item.id);
    const restored = removed?.objectId === item.nativeObjectId ? removed?.input : undefined;
    const objectId = wb.insertShape(item.sheet, restored ?? this.newShape(store, item));
    if (objectId === null) return null;
    this.objectIds.set(item.id, objectId);
    return this.tie(item.sheet, objectId);
  }

  /** The insert spec of an item the engine does not hold yet. An item that
   *  still carries its anchor keeps it; a new one is placed from its pixel frame. */
  private newShape(store: SpreadsheetStore, item: SessionIllustration): EngineShapeInput {
    const spec: EngineShapeInput = {
      prst: PRST_OF_KIND[item.shape as SessionShapeKind] as string,
      ...paintInput(item),
      ...(item.text && !isLineKind(item.shape) ? { text: item.text } : {}),
    };
    if (item.nativeAnchor) return { ...spec, ...anchorInput(item.nativeAnchor) };
    const frame = {
      x: item.x ?? DEFAULT_ILLUSTRATION_FRAME.x,
      y: item.y ?? DEFAULT_ILLUSTRATION_FRAME.y,
      w: item.w ?? DEFAULT_ILLUSTRATION_FRAME.w,
      h: item.h ?? DEFAULT_ILLUSTRATION_FRAME.h,
    };
    const placed = placementForPixels(
      store,
      {
        anchorKind: AnchorKind.OneCell,
        editAs: AnchorEditAs.OneCell,
        fromRow: 0,
        fromCol: 0,
        fromRowOff: 0,
        fromColOff: 0,
        toRow: 0,
        toCol: 0,
        toRowOff: 0,
        toColOff: 0,
        cx: 1,
        cy: 1,
      },
      frame,
    );
    return {
      ...spec,
      anchorKind: AnchorKind.OneCell,
      editAs: AnchorEditAs.OneCell,
      fromRow: placed.row ?? 0,
      fromCol: placed.col ?? 0,
      fromRowOff: placed.rowOffEmu ?? 0,
      fromColOff: placed.colOffEmu ?? 0,
      cx: placed.widthEmu,
      cy: placed.heightEmu,
    };
  }

  /** Write what changed between `old` and `item`; returns `old` when the
   *  engine refuses, so the screen goes back to the file's state. */
  private update(
    store: SpreadsheetStore,
    objectId: number,
    old: SessionIllustration,
    item: SessionIllustration,
  ): Partial<SessionIllustration> | null {
    const input: EngineShapeInput = {};
    const anchor = this.anchorChange(store, old, item);
    if (anchor) Object.assign(input, anchorInput(anchor));
    if (!item.nativeLocked) {
      const paint = paintInput(item);
      const oldPaint = paintInput(old);
      const same = (keys: readonly (keyof EngineShapeInput)[]): boolean =>
        keys.every((key) => paint[key] === oldPaint[key]);
      const fill = ['fillKind', 'fillRgb'] as const;
      const line = ['lineKind', 'lineRgb', 'lineWidthEmu'] as const;
      if (!same(fill)) Object.assign(input, pick(paint, fill));
      if (!same(line)) Object.assign(input, pick(paint, line));
      if (item.shape !== old.shape) input.prst = PRST_OF_KIND[item.shape as SessionShapeKind];
      if (item.text !== old.text) input.text = item.text ?? '';
    }
    if (Object.keys(input).length === 0) return null;
    if (this.handle.updateShape(item.sheet, objectId, input)) {
      const tie = this.tie(item.sheet, objectId);
      return tie.nativeAnchor && sameAnchor(tie.nativeAnchor, item.nativeAnchor) ? null : tie;
    }
    return old;
  }

  /** The anchor `item` moved or resized to, or null when its geometry is unchanged. */
  private anchorChange(
    store: SpreadsheetStore,
    old: SessionIllustration,
    item: SessionIllustration,
  ): EngineDrawingAnchor | null {
    if (item.nativeGeometryEditable === false) return null;
    if (item.nativeAnchor && !sameAnchor(old.nativeAnchor, item.nativeAnchor)) {
      return item.nativeAnchor;
    }
    const changed: Partial<Pick<SessionIllustration, 'x' | 'y' | 'w' | 'h'>> = {};
    for (const key of ['x', 'y', 'w', 'h'] as const) {
      if (item[key] !== old[key] && item[key] !== undefined) changed[key] = item[key];
    }
    const anchor = old.nativeAnchor;
    if (!anchor || Object.keys(changed).length === 0) return null;
    if (anchor.anchorKind === AnchorKind.TwoCell) {
      return twoCellAnchorForPixels(
        store,
        { ...sessionPosition(store, anchor), ...changed },
        anchor.editAs,
      );
    }
    const placed = placementForPixels(store, anchor, changed);
    return {
      ...anchor,
      fromRow: placed.row ?? anchor.fromRow,
      fromCol: placed.col ?? anchor.fromCol,
      fromRowOff: placed.rowOffEmu ?? anchor.fromRowOff,
      fromColOff: placed.colOffEmu ?? anchor.fromColOff,
      cx: placed.widthEmu ?? anchor.cx,
      cy: placed.heightEmu ?? anchor.cy,
    };
  }

  /** The fields that tie an overlay item to shape `objectId` as the engine holds it now. */
  private tie(sheet: number, objectId: number): Partial<SessionIllustration> {
    const shape = this.handle.listShapes(sheet).find((s) => s.objectId === objectId);
    return shape
      ? {
          nativeObjectId: objectId,
          nativeAnchor: drawingAnchorOf(shape),
          nativeGeometryEditable: shape.anchorKind !== ANCHOR_KIND_ABSOLUTE,
        }
      : { nativeObjectId: objectId };
  }
}
