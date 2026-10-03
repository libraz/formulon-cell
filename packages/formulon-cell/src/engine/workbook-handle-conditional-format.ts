import { completeDxfRecord } from './style-records.js';
import type {
  ConditionalFormatEntry,
  ConditionalFormatInput,
  DxfRecord,
  EngineCapabilities,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  assertAlive(): void;
};

declare module './workbook-handle.js' {
  interface WorkbookHandle extends WorkbookHandleConditionalFormatMethods {}
}

function internals(handle: unknown): WorkbookHandleInternals {
  return handle as WorkbookHandleInternals;
}

function assertAlive(handle: unknown): void {
  internals(handle).assertAlive();
}

function wb(handle: unknown): Workbook {
  return internals(handle).wb;
}

export abstract class WorkbookHandleConditionalFormatMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Evaluate every CF block on `sheet` against the inclusive viewport rect.
   *  Returns a sparse list — only cells with at least one match appear. The
   *  underlying embind vectors are released before this method returns, so
   *  the JS objects own no engine memory. Pass `NaN` for `todaySerial` to
   *  disable `TimePeriod` rules; defaults to `NaN`. Returns `[]` when the
   *  engine doesn't expose `evaluateCfRange`. */
  evaluateCfRange(
    sheet: number,
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    todaySerial = Number.NaN,
  ): {
    row: number;
    col: number;
    matches: {
      kind: number;
      priority: number;
      dxfIdEngaged: boolean;
      dxfId: number;
      color: { r: number; g: number; b: number; a: number };
      barLengthPct: number;
      barAxisPositionPct: number;
      barIsNegative: boolean;
      barFill: { r: number; g: number; b: number; a: number };
      barBorderEngaged: boolean;
      barBorder: { r: number; g: number; b: number; a: number };
      barGradient: boolean;
      barDirection: number;
      iconSetName: number;
      iconIndex: number;
    }[];
  }[] {
    assertAlive(this);
    if (!this.capabilities.conditionalFormat) return [];
    const r = wb(this).evaluateCfRange(sheet, firstRow, firstCol, lastRow, lastCol, todaySerial);
    if (!r.status.ok) return [];
    const out: ReturnType<WorkbookHandle['evaluateCfRange']> = [];
    for (const cell of r.cells) {
      const matches: ReturnType<WorkbookHandle['evaluateCfRange']>[number]['matches'] = [];
      for (const m of cell.matches) {
        matches.push({
          kind: m.kind as number,
          priority: m.priority,
          dxfIdEngaged: m.dxfIdEngaged !== 0,
          dxfId: m.dxfId,
          color: { r: m.color.r, g: m.color.g, b: m.color.b, a: m.color.a },
          barLengthPct: m.barLengthPct,
          barAxisPositionPct: m.barAxisPositionPct,
          barIsNegative: m.barIsNegative !== 0,
          barFill: { r: m.barFill.r, g: m.barFill.g, b: m.barFill.b, a: m.barFill.a },
          barBorderEngaged: m.barBorderEngaged !== 0,
          barBorder: {
            r: m.barBorder.r,
            g: m.barBorder.g,
            b: m.barBorder.b,
            a: m.barBorder.a,
          },
          barGradient: m.barGradient !== 0,
          barDirection: m.barDirection,
          iconSetName: m.iconSetName,
          iconIndex: m.iconIndex,
        });
      }
      out.push({ row: cell.row, col: cell.col, matches });
    }
    return out;
  }

  /** Snapshot of every CF rule on `sheet`, in flattened priority order.
   *  Returns `[]` when the engine doesn't expose `getConditionalFormats`
   *  or when there are no rules. The entries borrow rule ids from the
   *  engine's storage; treat them as immutable view objects. */
  getConditionalFormats(sheet: number): readonly ConditionalFormatEntry[] {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return [];
    const entries = wb(this).getConditionalFormats(sheet);
    return entries.status.ok ? [...entries] : [];
  }

  /** Removes the CF rule at `index` (flattened priority order). When the
   *  containing block becomes empty, the engine drops it too. Returns
   *  `false` (no-op) under stub mode and older engine package builds. */
  removeConditionalFormatAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return false;
    return wb(this).removeConditionalFormatAt(sheet, index).ok;
  }

  /** Drops every `<conditionalFormatting>` block on `sheet`. Returns
   *  `false` (no-op) under stub mode and older engine package builds. */
  clearConditionalFormats(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return false;
    return wb(this).clearConditionalFormats(sheet).ok;
  }

  /** Adds one conditional-format rule to `sheet` so it round-trips through
   *  .xlsx. `rule.type` mirrors `formulon::cf::RuleType` (0 expression,
   *  1 cellIs, 2 colorScale, 3 dataBar, 4 iconSet, 5 top10, 6 aboveAverage,
   *  7 containsText, 8 notContainsText, 9 beginsWith, 10 endsWith, …). The
   *  applied differential format is referenced by `dxfId`. Returns the
   *  flattened rule index, or -1 when the engine refuses the write. */
  addConditionalFormat(sheet: number, rule: ConditionalFormatInput): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return -1;
    const r = wb(this).addConditionalFormat(sheet, rule);
    return r.status.ok ? r.index : -1;
  }

  getDxf(index: number): DxfRecord | null {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return null;
    const r = wb(this).getDxf(index);
    if (!r.status.ok) return null;
    const { status: _status, ...record } = r;
    return record;
  }

  addDxf(record: DxfRecord): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return -1;
    const r = wb(this).addDxf(completeDxfRecord(record));
    return r.status.ok ? r.index : -1;
  }

  dxfCount(): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return 0;
    const count = wb(this).dxfCount();
    return count.status.ok ? count.value : 0;
  }
}

export function installConditionalFormatMethods(target: WorkbookHandleCtor): void {
  for (const key of Object.getOwnPropertyNames(WorkbookHandleConditionalFormatMethods.prototype)) {
    if (key === 'constructor') continue;
    const descriptor = Object.getOwnPropertyDescriptor(
      WorkbookHandleConditionalFormatMethods.prototype,
      key,
    );
    if (!descriptor) continue;
    Object.defineProperty(target.prototype, key, descriptor);
  }
}
