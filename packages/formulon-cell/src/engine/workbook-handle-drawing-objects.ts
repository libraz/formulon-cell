import type {
  Shape,
  ShapeInput,
  Slicer,
  SlicerInput,
  Sparkline,
  SparklineGroup,
  SparklineGroupInput,
} from '@libraz/formulon';
import type { EngineCapabilities, Workbook } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleDrawingObjectsMethods {}
}

export type EngineShape = Shape;
export type EngineShapeInput = ShapeInput;
export type EngineSlicer = Slicer;
export type EngineSlicerInput = SlicerInput;
export type EngineSparkline = Sparkline;
export type EngineSparklineGroup = SparklineGroup;
export type EngineSparklineGroupInput = SparklineGroupInput;

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
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
    const result = wb(this).listShapes(sheet);
    return result.status.ok ? [...result] : [];
  }

  /** Insert a shape and return its object id, or null when the engine refuses. */
  insertShape(sheet: number, spec: EngineShapeInput): number | null {
    assertAlive(this);
    if (!this.capabilities.shapes) return null;
    const result = wb(this).insertShape(sheet, spec);
    return result.status.ok ? result.objectId : null;
  }

  /** Apply the fields `spec` names to shape `objectId`. */
  updateShape(sheet: number, objectId: number, spec: EngineShapeInput): boolean {
    assertAlive(this);
    if (!this.capabilities.shapes) return false;
    return wb(this).updateShape(sheet, objectId, spec).ok;
  }

  removeShape(sheet: number, objectId: number): boolean {
    assertAlive(this);
    if (!this.capabilities.shapes) return false;
    return wb(this).removeShape(sheet, objectId).ok;
  }

  /** Sparkline groups of `sheet` in document order. */
  listSparklineGroups(sheet: number): EngineSparklineGroup[] {
    assertAlive(this);
    if (!this.capabilities.sparklines) return [];
    const result = wb(this).listSparklineGroups(sheet);
    return result.status.ok ? [...result] : [];
  }

  addSparklineGroup(sheet: number, spec: EngineSparklineGroupInput): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return wb(this).addSparklineGroup(sheet, spec).ok;
  }

  updateSparklineGroup(sheet: number, index: number, spec: EngineSparklineGroupInput): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return wb(this).updateSparklineGroup(sheet, index, spec).ok;
  }

  /** Replace the members of group `index`. */
  setSparklineGroupSparklines(
    sheet: number,
    index: number,
    sparklines: readonly EngineSparkline[],
  ): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return wb(this).setSparklineGroupSparklines(sheet, index, [...sparklines]).ok;
  }

  removeSparklineGroup(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.sparklines) return false;
    return wb(this).removeSparklineGroup(sheet, index).ok;
  }

  /** Table slicers hosted on `sheet`, in document order. */
  listSlicers(sheet: number): EngineSlicer[] {
    assertAlive(this);
    if (!this.capabilities.slicers) return [];
    const result = wb(this).listSlicers(sheet);
    return result.status.ok ? [...result] : [];
  }

  /** Add a slicer over column `column` of table `table` (workbook-wide table
   *  index) and return its frame's object id, or null when the engine refuses. */
  addSlicer(sheet: number, table: number, column: number, spec: EngineSlicerInput): number | null {
    assertAlive(this);
    if (!this.capabilities.slicers) return null;
    const result = wb(this).addSlicer(sheet, table, column, spec);
    return result.status.ok ? result.objectId : null;
  }

  updateSlicer(sheet: number, index: number, spec: EngineSlicerInput): boolean {
    assertAlive(this);
    if (!this.capabilities.slicers) return false;
    return wb(this).updateSlicer(sheet, index, spec).ok;
  }

  removeSlicer(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.slicers) return false;
    return wb(this).removeSlicer(sheet, index).ok;
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
