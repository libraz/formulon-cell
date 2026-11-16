import type {
  DrawingPatch,
  DrawingProperties,
  ShapeInput,
  ShapePatch,
  ShapeProperties,
  SlicerInput,
  SlicerPatch,
  SlicerProperties,
  SparklineGroup,
  SparklineGroupInput,
  SparklineGroupPatch,
  SparklineInput,
} from '@libraz/formulon';
import { attemptWrite, orFallback } from './result.js';
import type { EngineCapabilities } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { assertAlive, internals, ws } from './workbook-handle-internals.js';
import { workbookTableAt } from './workbook-handle-tables.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleDrawingObjectsMethods {}
}

/** A shape with the drawing fields (name, visibility, anchor) every object has. */
export type EngineShape = ShapeProperties & DrawingProperties;
export type EngineShapeInput = ShapeInput;
export type EngineShapePatch = ShapePatch & DrawingPatch;
/** A table slicer with the drawing fields of its frame. */
export type EngineSlicer = SlicerProperties & DrawingProperties;
export type EngineSlicerInput = SlicerInput;
export type EngineSlicerPatch = SlicerPatch & Omit<DrawingPatch, 'name'>;
export type EngineSparklineInput = SparklineInput;
type EngineSparklineGroup = SparklineGroup;
type EngineSparklineGroupInput = SparklineGroupInput;
type EngineSparklineGroupPatch = SparklineGroupPatch;

/** Shapes, sparkline groups and table slicers of a sheet. Sparkline groups and
 *  slicers are addressed by their position in the sheet's list, which a removal
 *  shifts; shapes by the object id the engine reports. */
export abstract class WorkbookHandleDrawingObjectsMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Shapes, connectors and groups of `sheet`, in drawing order. */
  listShapes(sheet: number): EngineShape[] {
    assertAlive(this);
    if (!this.capabilities.shapes) return [];
    return orFallback([], () =>
      ws(this, sheet)
        .listShapes()
        .map((shape) => ({ ...shape.getDrawingProperties(), ...shape.getProperties() })),
    );
  }

  /** Insert a shape and return its object id, or null when the engine refuses. */
  insertShape(sheet: number, spec: EngineShapeInput): number | null {
    assertAlive(this);
    if (!this.capabilities.shapes) return null;
    return orFallback(null, () => ws(this, sheet).addShape(spec).objectId);
  }

  /** Apply the fields `patch` names to shape `objectId`. */
  updateShape(sheet: number, objectId: number, patch: EngineShapePatch): boolean {
    assertAlive(this);
    if (!this.capabilities.shapes) return false;
    return attemptWrite(() => ws(this, sheet).shape(objectId).update(patch));
  }

  removeShape(sheet: number, objectId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.shapes) return false;
    return attemptWrite(() => ws(this, sheet).removeDrawing(objectId));
  }

  /** Sparkline groups of `sheet` in document order. */
  listSparklineGroups(sheet: number): EngineSparklineGroup[] {
    assertAlive(this);
    if (!this.capabilities.sparklines) return [];
    return orFallback([], () => [...ws(this, sheet).listSparklineGroups()]);
  }

  addSparklineGroup(sheet: number, spec: EngineSparklineGroupInput): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return attemptWrite(() => {
      ws(this, sheet).addSparklineGroup(spec);
    });
  }

  updateSparklineGroup(sheet: number, index: number, patch: EngineSparklineGroupPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return attemptWrite(() => ws(this, sheet).updateSparklineGroup(index, patch));
  }

  /** Replace the members of group `index`. */
  setSparklineGroupSparklines(
    sheet: number,
    index: number,
    sparklines: readonly EngineSparklineInput[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return attemptWrite(() =>
      ws(this, sheet).updateSparklineGroup(index, { sparklines: [...sparklines] }),
    );
  }

  removeSparklineGroup(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return attemptWrite(() => ws(this, sheet).removeSparklineGroup(index));
  }

  /** Table slicers hosted on `sheet`, in document order. */
  listSlicers(sheet: number): EngineSlicer[] {
    assertAlive(this);
    if (!this.capabilities.slicers) return [];
    return orFallback([], () =>
      ws(this, sheet)
        .listSlicers()
        .map((slicer) => ({ ...slicer.getDrawingProperties(), ...slicer.getProperties() })),
    );
  }

  /** Add a slicer over column `column` of table `table` (workbook-wide table
   *  index) and return its frame's object id, or null when the engine refuses. */
  addSlicer(sheet: number, table: number, column: number, spec: EngineSlicerInput): number | null {
    assertAlive(this);
    if (!this.capabilities.slicers) return null;
    const target = workbookTableAt(internals(this).wb, table);
    if (!target) return null;
    return orFallback(null, () => ws(this, sheet).addSlicer(target, column, spec).objectId);
  }

  updateSlicer(sheet: number, index: number, patch: EngineSlicerPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.slicers) return false;
    const slicer = orFallback(undefined, () => ws(this, sheet).listSlicers()[index]);
    return slicer ? attemptWrite(() => slicer.update(patch)) : false;
  }

  removeSlicer(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.slicers) return false;
    const slicer = orFallback(undefined, () => ws(this, sheet).listSlicers()[index]);
    return slicer ? attemptWrite(() => ws(this, sheet).removeDrawing(slicer)) : false;
  }
}

export function installDrawingObjectsMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleDrawingObjectsMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleDrawingObjectsMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
