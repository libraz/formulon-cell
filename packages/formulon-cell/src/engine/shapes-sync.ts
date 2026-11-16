import type { ChartFill, ChartLine, ChartLineDash, Color } from '@libraz/formulon';
import { AnchorEditAs, AnchorKind, ColorKind, DrawingObjectKind } from '@libraz/formulon';
import {
  DEFAULT_SHAPE_COLOR,
  SAVED_LINE_DASH,
  shapeFillColor,
  shapeHasFill,
  shapeOutlineColor,
  shapeOutlineWeight,
  shapeStoredWeight,
} from '../commands/shape-format.js';
import { writeDrawingOrder } from '../store/drawing-order.js';
import type {
  SessionIllustration,
  SessionNativeColor,
  SessionShapeKind,
  SessionShapeLineDash,
  SpreadsheetStore,
  State,
} from '../store/store.js';
import {
  ANCHOR_KIND_ABSOLUTE,
  drawingAnchorOf,
  placementForAnchor,
  placementForPixels,
  sameAnchor,
  sessionPosition,
  twoCellAnchorForPixels,
} from './drawing-geometry.js';
import { argbFromCssColor, cssColorFromArgb, renderedThemeColor } from './theme-color.js';
import type { EngineDrawingAnchor } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type {
  EngineShape,
  EngineShapeInput,
  EngineShapePatch,
} from './workbook-handle-drawing-objects.js';

/** Where a shape without a frame is drawn, in grid pixels. */
export const DEFAULT_ILLUSTRATION_FRAME = { x: 300, y: 88, w: 160, h: 96 } as const;

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
const shapeIsSaved = (item: SessionIllustration): boolean =>
  item.kind === 'shape' && item.shape !== undefined && PRST_OF_KIND[item.shape] !== undefined;

const isLineKind = (kind: SessionShapeKind | undefined): boolean =>
  kind === 'line' || kind === 'arrow';

/** `#rrggbb` of a CSS colour the engine can store; null for any other. */
const hexOf = (color: string | null | undefined): string | null => {
  const argb = color ? argbFromCssColor(color) : null;
  return argb === null ? null : cssColorFromArgb(argb);
};

/** How a shape paints its fill or outline: its style's (`unset`), a solid colour, nothing, or a paint
 *  the overlay cannot draw. */
type Paint = 'unset' | 'solid' | 'none' | 'other';

const paintOf = (paint: ChartFill | ChartLine | null): Paint => {
  if (paint === null) return 'unset';
  if (paint.other) return 'other';
  if (paint.none) return 'none';
  return paint.color !== null ? 'solid' : 'unset';
};

/** A paint's colour: as drawn, and as the file stores it. */
const nativeColor = (
  wb: Pick<WorkbookHandle, 'resolveColorArgb'>,
  spec: Color,
): SessionNativeColor => ({ spec, css: cssColorFromArgb(renderedThemeColor(wb, spec).argb) });

/** The dash style an engine dash reads as; none for one the overlay has no style for. */
const dashOfSaved = (saved: ChartLineDash): SessionShapeLineDash | undefined =>
  (Object.keys(SAVED_LINE_DASH) as SessionShapeLineDash[]).find(
    (dash) => SAVED_LINE_DASH[dash] === saved,
  );

/** The engine dash `item` is saved with; none for solid or a style the file cannot store. */
const savedDash = (item: SessionIllustration): ChartLineDash | undefined =>
  item.lineDash === undefined ? undefined : SAVED_LINE_DASH[item.lineDash];

/** The colour to write for `css`: the file's own while the shape still shows it. */
const colorInput = (css: string, native: SessionNativeColor | undefined): string | Color =>
  native !== undefined && native.css === css ? native.spec : css;

const anchorInput = (anchor: EngineDrawingAnchor) => ({
  anchor: placementForAnchor(undefined, anchor),
});

/** The fill and outline fields that write `item`'s paint. */
const paintInput = (item: SessionIllustration): Pick<EngineShapeInput, 'fill' | 'line'> => {
  const input: { -readonly [K in 'fill' | 'line']?: EngineShapeInput[K] } = {};
  // A paint the cell cannot state is never written: a colour sent for it replaces it.
  if (shapeHasFill(item) && !item.nativeFillOther) {
    const color = hexOf(shapeFillColor(item) ?? undefined);
    const transparencyPct = Math.round((1 - (item.opacity ?? 1)) * 100);
    if (item.opacity === 0 || shapeFillColor(item) === null) input.fill = { none: true };
    else {
      input.fill = {
        color: colorInput(color ?? (hexOf(DEFAULT_SHAPE_COLOR) as string), item.nativeFill),
        ...(transparencyPct > 0 ? { transparencyPct } : {}),
      };
    }
  }
  const weight = shapeOutlineWeight(item);
  const outline = hexOf(shapeOutlineColor(item));
  const dash = savedDash(item);
  const dashInput = dash === undefined ? {} : { dash };
  if (item.nativeLineOther) {
    input.line = { widthPt: shapeStoredWeight(item), ...dashInput };
  } else if (weight === null || outline === null) {
    input.line = { none: true, widthPt: shapeStoredWeight(item), ...dashInput };
  } else {
    input.line = { color: colorInput(outline, item.nativeLine), widthPt: weight, ...dashInput };
  }
  return input;
};

/** Overlay-only settings of a saved shape, reset so the screen shows what the
 *  file holds: the file has no corner radius and only some dash styles. */
const unsaveable = (item: SessionIllustration): Partial<SessionIllustration> => {
  const reset: Partial<SessionIllustration> = {};
  if (item.nativeLocked) return reset;
  if (item.radius !== undefined) reset.radius = undefined;
  if (item.lineDash !== undefined && savedDash(item) === undefined) reset.lineDash = undefined;
  return reset;
};

/** `next` with a null for each optional member `prev` wrote and `next` leaves
 *  out, so the engine's merge removes it. */
const clearDropped = <T extends object>(next: T, prev: T | null | undefined): T => {
  const cleared: Record<string, null> = {};
  for (const key of ['dash', 'transparencyPct']) {
    if (prev && key in prev && !(key in next)) cleared[key] = null;
  }
  return { ...next, ...cleared };
};

/** The overlay for a shape of the file, or null for a hidden one, which stays
 *  untouched in the engine. A shape the overlay cannot draw (unknown preset,
 *  line, connector, group) becomes a locked bounding box; a fill or outline it
 *  cannot state is kept as the file holds it. */
export function shapeFromNative(
  store: SpreadsheetStore,
  sheet: number,
  shape: EngineShape,
  wb: Pick<WorkbookHandle, 'resolveColorArgb'>,
): SessionIllustration | null {
  if (shape.hidden) return null;
  // A custom geometry has no preset name.
  const kind = shape.presetGeometry === null ? undefined : KIND_OF_PRST.get(shape.presetGeometry);
  // The engine reports no flip, so a loaded line's direction is unknown.
  const locked =
    kind === undefined ||
    kind === 'line' ||
    shape.kind === DrawingObjectKind.Connector ||
    shape.kind === DrawingObjectKind.Group;
  const anchor = drawingAnchorOf(shape);
  const item: SessionIllustration = {
    id: `shape-native-${sheet}-${shape.objectId}`,
    kind: 'shape',
    shape: locked || kind === undefined ? 'rectangle' : kind,
    sheet,
    ...sessionPosition(store, anchor),
    nativeObjectId: shape.objectId,
    nativeAnchor: anchor,
    nativeGeometryEditable: shape.anchor.kind !== ANCHOR_KIND_ABSOLUTE,
    ...(shape.rotationDeg !== 0 ? { rotationDeg: shape.rotationDeg } : {}),
    ...(shape.name ? { name: shape.name } : {}),
    ...(shape.description ? { alt: shape.description } : {}),
    ...(shape.textBox ? { textBox: true } : {}),
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
  const fill = paintOf(shape.fill);
  const line = paintOf(shape.line);
  const fillColor = shape.fill?.color;
  const lineColor = shape.line?.color;
  // A paint the cell cannot state is drawn in the colour the engine reads for it.
  if ((fill === 'solid' || fill === 'other') && fillColor) {
    const color = nativeColor(wb, fillColor);
    item.color = color.css;
    if (fill === 'solid' && color.spec.kind !== ColorKind.Rgb) item.nativeFill = color;
    const transparencyPct = shape.fill?.transparencyPct ?? 0;
    if (transparencyPct > 0) item.opacity = (100 - transparencyPct) / 100;
  } else if (fill === 'none') item.opacity = 0;
  if (fill === 'other') item.nativeFillOther = true;
  if (line === 'none') item.lineHidden = true;
  else if ((line === 'solid' || line === 'other') && lineColor) {
    const color = nativeColor(wb, lineColor);
    item.lineColor = color.css;
    if (line === 'solid' && color.spec.kind !== ColorKind.Rgb) item.nativeLine = color;
  }
  if (line === 'other') item.nativeLineOther = true;
  const width = shape.line?.widthPt ?? 0;
  if (width > 0) item.lineWidth = width;
  const dash = shape.line?.dash == null ? undefined : dashOfSaved(shape.line.dash);
  if (dash !== undefined && dash !== 'solid') item.lineDash = dash;
  if (shape.text) item.text = shape.text;
  return item;
}

/** A shape as an insert spec, or null when re-inserting it would lose paint
 *  or geometry the engine cannot take back. */
const shapeInputOf = (shape: EngineShape): EngineShapeInput | null => {
  const fill = paintOf(shape.fill);
  const line = paintOf(shape.line);
  if (
    shape.presetGeometry === null ||
    shape.kind === DrawingObjectKind.Group ||
    fill === 'other' ||
    line === 'other'
  ) {
    return null;
  }
  // Every member the outline states, so the copy keeps it.
  const lineStated = {
    ...(shape.line?.widthPt != null ? { widthPt: shape.line.widthPt } : {}),
    ...(shape.line?.dash != null ? { dash: shape.line.dash } : {}),
    ...(shape.line?.transparencyPct != null ? { transparencyPct: shape.line.transparencyPct } : {}),
  };
  return {
    ...anchorInput(drawingAnchorOf(shape)),
    name: shape.name,
    ...(shape.description ? { description: shape.description } : {}),
    presetGeometry: shape.presetGeometry,
    rotationDeg: shape.rotationDeg,
    hidden: shape.hidden,
    ...(shape.kind !== DrawingObjectKind.Connector
      ? { text: shape.text, textBox: shape.textBox }
      : {}),
    ...(fill === 'solid' && shape.fill?.color
      ? {
          fill: {
            color: shape.fill.color,
            ...(shape.fill.transparencyPct != null
              ? { transparencyPct: shape.fill.transparencyPct }
              : {}),
          },
        }
      : {}),
    ...(fill === 'none' ? { fill: { none: true } } : {}),
    ...(line === 'solid' && shape.line?.color
      ? { line: { color: shape.line.color, ...lineStated } }
      : {}),
    ...(line === 'none' ? { line: { none: true, ...lineStated } } : {}),
  };
};

/** Shape `objectId` as the engine holds it, as an insert spec; null when it
 *  cannot be inserted again without loss. */
export function captureShapeInput(
  wb: WorkbookHandle,
  sheet: number,
  objectId: number,
): EngineShapeInput | null {
  const shape = wb.listShapes(sheet).find((candidate) => candidate.objectId === objectId);
  return shape ? shapeInputOf(shape) : null;
}

/** Capture shape `objectId` so a row or column delete that drops it can be
 *  undone. The restorer inserts it again at its slot in the drawing order; null
 *  when it cannot be restored without loss. */
export function captureShapeRestore(
  wb: WorkbookHandle,
  sheet: number,
  objectId: number,
): (() => boolean) | null {
  const input = captureShapeInput(wb, sheet, objectId);
  if (!input) return null;
  const index = wb.listDrawingObjects(sheet).findIndex((object) => object.objectId === objectId);
  return () => {
    const restored = wb.insertShape(sheet, input);
    return restored !== null && (index < 0 || wb.setDrawingZOrder(sheet, restored, index));
  };
}

/** Capture the anchors of `sheet`'s shapes ahead of a row or column delete,
 *  which collapses the shapes it does not remove. The restorer puts the
 *  survivors back where they were once the delete is undone. */
export function captureShapeAnchors(wb: WorkbookHandle, sheet: number): () => void {
  if (!wb.capabilities.shapes) return () => undefined;
  const anchors = new Map(
    wb
      .listShapes(sheet)
      .filter((shape) => shape.anchor.kind !== ANCHOR_KIND_ABSOLUTE)
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

  /** Each shape removed through the mirror, so a later undo puts it back as it
   *  was: from the engine's capture, at its slot in the drawing order, or else
   *  re-inserted from its description (name, text, anchor kind, geometry). */
  private readonly removed = new Map<
    string,
    {
      sheet: number;
      objectId: number;
      snapshot: Uint8Array | null;
      input: EngineShapeInput | null;
    }
  >();

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
    // Sheets where a shape gained an engine object, which takes its slot from the store order.
    const placed = new Set<number>();
    for (const item of slice.illustrations) {
      const old = before.get(item.id);
      if (old === item || item.kind !== 'shape' || !shapeIsSaved(item)) continue;
      const objectId = this.objectIds.get(item.id);
      if (objectId === undefined || (old && old.sheet !== item.sheet)) {
        if (objectId !== undefined && old) this.forget(old.id, old.sheet, objectId);
        const tied = this.insert(store, item);
        patches.set(item.id, tied && { ...tied, ...unsaveable(item) });
        if (tied) placed.add(item.sheet);
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
    for (const sheet of placed) writeDrawingOrder(wb, store.getState(), sheet);
  }

  /** Remove the engine shape of an overlay item, keeping its description. */
  private forget(id: string, sheet: number, objectId: number): void {
    const description = captureShapeInput(this.handle, sheet, objectId);
    const snapshot = this.handle.snapshotDrawing(sheet, objectId);
    if (description || snapshot) {
      this.removed.set(id, { sheet, objectId, snapshot, input: description });
    }
    this.handle.removeShape(sheet, objectId);
    this.objectIds.delete(id);
  }

  /** Insert `item`; returns the fields that tie it to the new shape, or null. */
  private insert(store: SpreadsheetStore, item: SessionIllustration) {
    const wb = this.handle;
    // Only the item that was removed comes back as described; a new item reusing its id does not.
    const removed = this.removed.get(item.id);
    this.removed.delete(item.id);
    const restored = removed?.objectId === item.nativeObjectId ? removed : undefined;
    const objectId =
      (restored?.snapshot && restored.sheet === item.sheet
        ? wb.restoreDrawing(item.sheet, restored.snapshot)
        : null) ??
      wb.insertShape(
        item.sheet,
        restored?.input ?? item.nativeSource ?? this.newShape(store, item),
      );
    if (objectId === null) return null;
    this.objectIds.set(item.id, objectId);
    const tie = this.tie(item.sheet, objectId);
    return item.nativeSource ? { ...tie, nativeSource: undefined } : tie;
  }

  /** The insert spec of an item the engine does not hold yet. An item that
   *  still carries its anchor keeps it; a new one is placed from its pixel frame. */
  private newShape(store: SpreadsheetStore, item: SessionIllustration): EngineShapeInput {
    const spec: EngineShapeInput = {
      presetGeometry: PRST_OF_KIND[item.shape as SessionShapeKind] as string,
      ...paintInput(item),
      ...(item.text && !isLineKind(item.shape) ? { text: item.text } : {}),
      ...(item.textBox && !isLineKind(item.shape) ? { textBox: true } : {}),
      ...(item.name ? { name: item.name } : {}),
      ...(item.alt ? { description: item.alt } : {}),
      ...(item.rotationDeg !== undefined ? { rotationDeg: item.rotationDeg } : {}),
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
        kind: AnchorKind.OneCell,
        editAs: AnchorEditAs.OneCell,
        fromRow: 0,
        fromCol: 0,
        fromRowOffsetEmu: 0,
        fromColOffsetEmu: 0,
        toRow: 0,
        toCol: 0,
        toRowOffsetEmu: 0,
        toColOffsetEmu: 0,
        widthEmu: 1,
        heightEmu: 1,
      },
      frame,
    );
    return { ...spec, anchor: placed };
  }

  /** Write what changed between `old` and `item`; returns `old` when the
   *  engine refuses, so the screen goes back to the file's state. */
  private update(
    store: SpreadsheetStore,
    objectId: number,
    old: SessionIllustration,
    item: SessionIllustration,
  ): Partial<SessionIllustration> | null {
    const input: { -readonly [K in keyof EngineShapePatch]: EngineShapePatch[K] } = {};
    const anchor = this.anchorChange(store, old, item);
    if (anchor) Object.assign(input, anchorInput(anchor));
    if (!item.nativeLocked) {
      const paint = paintInput(item);
      const oldPaint = paintInput(old);
      const same = (key: 'fill' | 'line'): boolean =>
        JSON.stringify(paint[key]) === JSON.stringify(oldPaint[key]);
      if (!same('fill') && paint.fill) input.fill = clearDropped(paint.fill, oldPaint.fill);
      if (!same('line') && paint.line) input.line = clearDropped(paint.line, oldPaint.line);
      if (item.shape !== old.shape) {
        input.presetGeometry = PRST_OF_KIND[item.shape as SessionShapeKind];
      }
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
    if (anchor.kind === AnchorKind.TwoCell) {
      return twoCellAnchorForPixels(
        store,
        { ...sessionPosition(store, anchor), ...changed },
        anchor.editAs,
      );
    }
    const placed = placementForPixels(store, anchor, changed);
    return {
      ...anchor,
      fromRow: placed.fromRow ?? anchor.fromRow,
      fromCol: placed.fromCol ?? anchor.fromCol,
      fromRowOffsetEmu: placed.fromRowOffsetEmu ?? anchor.fromRowOffsetEmu,
      fromColOffsetEmu: placed.fromColOffsetEmu ?? anchor.fromColOffsetEmu,
      widthEmu: placed.widthEmu ?? anchor.widthEmu,
      heightEmu: placed.heightEmu ?? anchor.heightEmu,
    };
  }

  /** The fields that tie an overlay item to shape `objectId` as the engine holds it now. */
  private tie(sheet: number, objectId: number): Partial<SessionIllustration> {
    const shape = this.handle.listShapes(sheet).find((s) => s.objectId === objectId);
    return shape
      ? {
          nativeObjectId: objectId,
          nativeAnchor: drawingAnchorOf(shape),
          nativeGeometryEditable: shape.anchor.kind !== ANCHOR_KIND_ABSOLUTE,
        }
      : { nativeObjectId: objectId };
  }
}
