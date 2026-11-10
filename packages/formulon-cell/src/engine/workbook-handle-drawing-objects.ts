import type {
  DrawingPatch,
  DrawingProperties,
  ShapeInput,
  ShapePatch,
  ShapeProperties,
  Sheet,
  SlicerInput,
  SlicerPatch,
  SlicerProperties,
  Sparkline,
  SparklineGroup,
  SparklineGroupInput,
  SparklineGroupPatch,
  SparklineInput,
} from '@libraz/formulon';
import { isFormulonError } from './result.js';
import type { EngineCapabilities, Workbook } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import { workbookTableAt } from './workbook-handle-tables.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

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
export type EngineSparkline = Sparkline;
export type EngineSparklineInput = SparklineInput;
export type EngineSparklineGroup = SparklineGroup;
export type EngineSparklineGroupInput = SparklineGroupInput;
export type EngineSparklineGroupPatch = SparklineGroupPatch;

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function ws(handle: unknown, sheet: number): Sheet {
  return internals(handle).wb.sheet(sheet);
}

/** Run an engine edit; an engine error is a `false` outcome. */
function edit(run: () => void): boolean {
  try {
    run();
    return true;
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return false;
  }
}

/** Run an engine read; an engine error reads as `fallback`. */
function read<T>(fallback: T, run: () => T): T {
  try {
    return run();
  } catch (e) {
    if (!isFormulonError(e)) throw e;
    return fallback;
  }
}

/** Shapes, sparkline groups and table slicers of a sheet. Sparkline groups and
 *  slicers are addressed by their position in the sheet's list, which a removal
 *  shifts; shapes by the object id the engine reports. */
export abstract class WorkbookHandleDrawingObjectsMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Shapes, connectors and groups of `sheet`, in drawing order. */
  listShapes(sheet: number): EngineShape[] {
    assertAlive(this);
    if (!this.capabilities.shapes) return [];
    return read([], () =>
      ws(this, sheet)
        .listShapes()
        .map((shape) => ({ ...shape.getDrawingProperties(), ...shape.getProperties() })),
    );
  }

  /** Insert a shape and return its object id, or null when the engine refuses. */
  insertShape(sheet: number, spec: EngineShapeInput): number | null {
    assertAlive(this);
    if (!this.capabilities.shapes) return null;
    return read(null, () => ws(this, sheet).addShape(spec).objectId);
  }

  /** Apply the fields `patch` names to shape `objectId`. */
  updateShape(sheet: number, objectId: number, patch: EngineShapePatch): boolean {
    assertAlive(this);
    if (!this.capabilities.shapes) return false;
    return edit(() => ws(this, sheet).shape(objectId).update(patch));
  }

  removeShape(sheet: number, objectId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.shapes) return false;
    return edit(() => ws(this, sheet).removeDrawing(objectId));
  }

  /** Sparkline groups of `sheet` in document order. */
  listSparklineGroups(sheet: number): EngineSparklineGroup[] {
    assertAlive(this);
    if (!this.capabilities.sparklines) return [];
    return read([], () => [...ws(this, sheet).listSparklineGroups()]);
  }

  addSparklineGroup(sheet: number, spec: EngineSparklineGroupInput): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return edit(() => {
      ws(this, sheet).addSparklineGroup(spec);
    });
  }

  updateSparklineGroup(sheet: number, index: number, patch: EngineSparklineGroupPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return edit(() => ws(this, sheet).updateSparklineGroup(index, patch));
  }

  /** Replace the members of group `index`. */
  setSparklineGroupSparklines(
    sheet: number,
    index: number,
    sparklines: readonly EngineSparklineInput[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return edit(() => ws(this, sheet).updateSparklineGroup(index, { sparklines: [...sparklines] }));
  }

  removeSparklineGroup(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return edit(() => ws(this, sheet).removeSparklineGroup(index));
  }

  /** Table slicers hosted on `sheet`, in document order. */
  listSlicers(sheet: number): EngineSlicer[] {
    assertAlive(this);
    if (!this.capabilities.slicers) return [];
    return read([], () =>
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
    return read(null, () => ws(this, sheet).addSlicer(target, column, spec).objectId);
  }

  updateSlicer(sheet: number, index: number, patch: EngineSlicerPatch): boolean {
    assertAlive(this);
    if (!this.capabilities.slicers) return false;
    const slicer = read(undefined, () => ws(this, sheet).listSlicers()[index]);
    return slicer ? edit(() => slicer.update(patch)) : false;
  }

  removeSlicer(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.slicers) return false;
    const slicer = read(undefined, () => ws(this, sheet).listSlicers()[index]);
    return slicer ? edit(() => ws(this, sheet).removeDrawing(slicer)) : false;
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
