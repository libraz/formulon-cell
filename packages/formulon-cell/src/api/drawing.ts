import {
  type BytesInput,
  DrawingObjectKind,
  type DrawingPlacementInput,
  type Drawing as EngineDrawing,
  type Picture as EnginePicture,
  type Shape as EngineShape,
  type Slicer as EngineSlicer,
  type Table as EngineTable,
  type RestoreDrawingOptions,
  type ShapeInput,
  type SlicerInput,
} from '@libraz/formulon';
import type {
  FacadeClass,
  HandleContext,
  MemberSpec,
  SheetContext,
  WorkbookContext,
} from './facade-context.js';
import type { Table } from './table.js';
import type { Sheet } from './workbook.js';

/**
 * An object of a sheet's drawing: a picture, a shape, a chart, a slicer or another kind. Members
 * are the engine `Drawing` members of the same name. Once the object is removed every member
 * throws `NotFound`.
 */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class Drawing {
  /** @internal */
  constructor() {
    throw new TypeError('get a drawing object from a Sheet');
  }
}

/** A picture of a sheet's drawing. Add one with `Sheet.addPicture`. */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class Picture extends Drawing {}

/** A shape of a sheet's drawing. Add one with `Sheet.addShape`. */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class Shape extends Drawing {}

/** A slicer of a sheet's drawing. Add one with `Sheet.addSlicer`. */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class Slicer extends Drawing {}

/** The drawing members every kind shares. */
export const DRAWING_MEMBERS = {
  props: ['objectId', 'kind', 'name'],
  getters: { sheet: (ctx) => ctx.sheet(ctx.target.sheet) },
  read: ['getDrawingProperties'],
  change: { drawing: ['update'] },
} as const satisfies MemberSpec<HandleContext<EngineDrawing>>;

export const PICTURE_MEMBERS = {
  read: ['getProperties', 'snapshot'],
  change: { drawing: ['setZOrder'] },
} as const satisfies MemberSpec<HandleContext<EnginePicture>>;

export const SHAPE_MEMBERS = {
  read: ['getProperties', 'snapshot'],
  change: { drawing: ['setZOrder'] },
} as const satisfies MemberSpec<HandleContext<EngineShape>>;

export const SLICER_MEMBERS = {
  read: ['getProperties'],
  change: { drawing: ['setZOrder'] },
} as const satisfies MemberSpec<HandleContext<EngineSlicer>>;

/** Classes of the kinds with their own handle; the chart class registers itself here. */
export const DRAWING_CLASSES = new Map<number, FacadeClass<Drawing>>([
  [DrawingObjectKind.Picture, Picture],
  [DrawingObjectKind.Shape, Shape],
]);

/** The facade of any drawing object, by its kind; a slicer is told apart by the sheet's slicer list. */
export const drawingOf = (ctx: WorkbookContext, engine: EngineDrawing): Drawing => {
  const cls: FacadeClass<Drawing> =
    DRAWING_CLASSES.get(engine.kind) ??
    (engine.sheet.listSlicers().some((slicer) => slicer.objectId === engine.objectId)
      ? Slicer
      : Drawing);
  return ctx.facade(engine, cls);
};

/** A drawing object passed by its id or by its handle. */
export type DrawingRef = number | Drawing;

const engineDrawingRef = (ctx: WorkbookContext, ref: DrawingRef): number | EngineDrawing =>
  ctx.engineOf(ref) as number | EngineDrawing;

export const SHEET_DRAWING_MEMBERS = {
  methods: {
    listDrawings: (ctx: SheetContext): Drawing[] =>
      ctx.sheetEngine.listDrawings().map((engine) => drawingOf(ctx, engine)),
    drawing: (ctx: SheetContext, ref: DrawingRef): Drawing =>
      drawingOf(ctx, ctx.sheetEngine.drawing(engineDrawingRef(ctx, ref))),
    removeDrawing: (ctx: SheetContext, ref: DrawingRef): void =>
      ctx.sheetChange('drawing', () => ctx.sheetEngine.removeDrawing(engineDrawingRef(ctx, ref))),
    restoreDrawing: (ctx: SheetContext, bytes: BytesInput, opts?: RestoreDrawingOptions): Drawing =>
      ctx.sheetChange('drawing', () => drawingOf(ctx, ctx.sheetEngine.restoreDrawing(bytes, opts))),
    listPictures: (ctx: SheetContext): Picture[] =>
      ctx.sheetEngine.listPictures().map((engine) => ctx.facade(engine, Picture)),
    picture: (ctx: SheetContext, ref: DrawingRef): Picture =>
      ctx.facade(ctx.sheetEngine.picture(engineDrawingRef(ctx, ref)), Picture),
    addPicture: (ctx: SheetContext, bytes: BytesInput, input?: DrawingPlacementInput): Picture =>
      ctx.sheetChange('drawing', () =>
        ctx.facade(ctx.sheetEngine.addPicture(bytes, input), Picture),
      ),
    listShapes: (ctx: SheetContext): Shape[] =>
      ctx.sheetEngine.listShapes().map((engine) => ctx.facade(engine, Shape)),
    shape: (ctx: SheetContext, ref: DrawingRef): Shape =>
      ctx.facade(ctx.sheetEngine.shape(engineDrawingRef(ctx, ref)), Shape),
    addShape: (ctx: SheetContext, spec: ShapeInput): Shape =>
      ctx.sheetChange('drawing', () => ctx.facade(ctx.sheetEngine.addShape(spec), Shape)),
    listSlicers: (ctx: SheetContext): Slicer[] =>
      ctx.sheetEngine.listSlicers().map((engine) => ctx.facade(engine, Slicer)),
    slicer: (ctx: SheetContext, ref: DrawingRef): Slicer =>
      ctx.facade(ctx.sheetEngine.slicer(engineDrawingRef(ctx, ref)), Slicer),
    addSlicer: (
      ctx: SheetContext,
      table: Table | string,
      column: number | string,
      spec: SlicerInput,
    ): Slicer =>
      ctx.sheetChange('drawing', () =>
        ctx.facade(
          ctx.sheetEngine.addSlicer(ctx.engineOf(table) as EngineTable | string, column, spec),
          Slicer,
        ),
      ),
  },
} as const satisfies MemberSpec<SheetContext>;

type DrawingMemberName =
  | (typeof DRAWING_MEMBERS.props)[number]
  | (typeof DRAWING_MEMBERS.read)[number]
  | (typeof DRAWING_MEMBERS.change.drawing)[number];
type KindMemberName = 'getProperties' | 'setZOrder';

export interface Drawing extends Pick<EngineDrawing, DrawingMemberName> {
  /** The sheet whose drawing holds the object. */
  readonly sheet: Sheet;
}
export interface Picture extends Pick<EnginePicture, KindMemberName | 'snapshot'> {
  update: EnginePicture['update'];
}
export interface Shape extends Pick<EngineShape, KindMemberName | 'snapshot'> {
  update: EngineShape['update'];
}
export interface Slicer extends Pick<EngineSlicer, KindMemberName> {
  update: EngineSlicer['update'];
}

declare module './workbook.js' {
  interface Sheet {
    /** Every object of the sheet's drawing, in stacking order (0 is the back). */
    listDrawings(): Drawing[];
    /** The drawing object of an id; throws `NotFound` when there is none. */
    drawing(ref: DrawingRef): Drawing;
    /** Removes a drawing object. */
    removeDrawing(ref: DrawingRef): void;
    /** Puts back an object captured with `snapshot()`. */
    restoreDrawing(bytes: BytesInput, opts?: RestoreDrawingOptions): Drawing;
    listPictures(): Picture[];
    /** The picture of an id; throws `NotFound` when there is none. */
    picture(ref: DrawingRef): Picture;
    addPicture(bytes: BytesInput, input?: DrawingPlacementInput): Picture;
    listShapes(): Shape[];
    /** The shape of an id; throws `NotFound` when there is none. */
    shape(ref: DrawingRef): Shape;
    addShape(spec: ShapeInput): Shape;
    listSlicers(): Slicer[];
    /** The slicer of an id; throws `NotFound` when there is none. */
    slicer(ref: DrawingRef): Slicer;
    /** Adds a slicer over a column of a table (a `Table` or a table name). */
    addSlicer(table: Table | string, column: number | string, spec: SlicerInput): Slicer;
  }
}
