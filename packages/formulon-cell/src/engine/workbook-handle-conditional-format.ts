import {
  CfMatchKind,
  type CfRangeEvaluationOptions,
  type Color,
  type ConditionalFormat,
  type Formulon,
} from '@libraz/formulon';
import { orFallback } from './result.js';
import { completeDxfRecord } from './style-records.js';
import type {
  ConditionalFormatInput,
  DxfRecord,
  EngineCapabilities,
  Sheet,
  Workbook,
} from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

type WorkbookHandleCtor = { prototype: WorkbookHandle };
type WorkbookHandleInternals = {
  wb: Workbook;
  module: Formulon;
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

function ws(handle: unknown, sheet: number): Sheet {
  return internals(handle).wb.sheet(sheet);
}

type Rgba = { r: number; g: number; b: number; a: number };

const rgbaOf = (color: Color | null | undefined): Rgba => {
  const argb = color ? color.argb >>> 0 : 0;
  return {
    a: (argb >>> 24) & 0xff,
    r: (argb >>> 16) & 0xff,
    g: (argb >>> 8) & 0xff,
    b: argb & 0xff,
  };
};

const NO_RGBA: Rgba = { r: 0, g: 0, b: 0, a: 0 };

export abstract class WorkbookHandleConditionalFormatMethods {
  declare readonly capabilities: EngineCapabilities;

  /** Evaluate every CF block on `sheet` against the inclusive viewport rect.
   *  Returns a sparse list — only cells with at least one match appear. Pass
   *  `NaN` for `todaySerial` to disable `TimePeriod` rules; defaults to `NaN`. */
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
    const native = wb(this);
    const options: CfRangeEvaluationOptions = Number.isFinite(todaySerial)
      ? {
          today: internals(this).module.convertSerialToCivil(todaySerial, {
            date1904: native.getDate1904(),
          }),
        }
      : { timePeriodRules: false };
    const cells = orFallback(null, () =>
      ws(this, sheet).evaluateCfRange({ firstRow, firstCol, lastRow, lastCol }, options),
    );
    if (!cells) return [];
    const out: ReturnType<WorkbookHandle['evaluateCfRange']> = [];
    for (const cell of cells) {
      const matches: ReturnType<WorkbookHandle['evaluateCfRange']>[number]['matches'] = [];
      for (const m of cell.matches) {
        const bar = m.kind === CfMatchKind.DataBar ? m.bar : null;
        const icon = m.kind === CfMatchKind.IconSet ? m.icon : null;
        const dxfIndex = m.kind === CfMatchKind.DifferentialFormat ? m.dxfIndex : null;
        matches.push({
          kind: m.kind,
          priority: m.priority,
          dxfIdEngaged: dxfIndex !== null,
          dxfId: dxfIndex ?? 0,
          color: m.kind === CfMatchKind.ColorScale ? rgbaOf(m.color) : { ...NO_RGBA },
          barLengthPct: bar?.lengthPct ?? 0,
          barAxisPositionPct: bar?.axisPositionPct ?? 0,
          barIsNegative: bar?.isNegative === true,
          barFill: bar ? rgbaOf(bar.fill) : { ...NO_RGBA },
          barBorderEngaged: bar?.border != null,
          barBorder: bar?.border ? rgbaOf(bar.border) : { ...NO_RGBA },
          barGradient: bar?.gradient === true,
          barDirection: bar?.direction ?? 0,
          iconSetName: icon?.iconSet ?? 0,
          iconIndex: icon?.iconIndex ?? 0,
        });
      }
      out.push({ row: cell.row, col: cell.col, matches });
    }
    return out;
  }

  /** Snapshot of every CF rule on `sheet`, in flattened priority order.
   *  Returns `[]` when there are no rules. Treat the entries as immutable
   *  view objects. */
  getConditionalFormats(sheet: number): readonly ConditionalFormat[] {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return [];
    return orFallback([], () => [...ws(this, sheet).listCfRules()]);
  }

  /** Removes the CF rule at `index` (flattened priority order). When the
   *  containing block becomes empty, the engine drops it too. Returns
   *  `false` (no-op) under stub mode or when the engine refuses. */
  removeConditionalFormatAt(sheet: number, index: number): boolean {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return false;
    return orFallback(false, () => {
      ws(this, sheet).removeCfRule(index);
      return true;
    });
  }

  /** Drops every `<conditionalFormatting>` block on `sheet`. Returns
   *  `false` (no-op) under stub mode. */
  clearConditionalFormats(sheet: number): boolean {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return false;
    return orFallback(false, () => {
      ws(this, sheet).clearCfRules();
      return true;
    });
  }

  /** Adds one conditional-format rule to `sheet` so it round-trips through
   *  .xlsx. The applied differential format is referenced by `dxfIndex`.
   *  Returns the flattened rule index, or -1 when the engine refuses the
   *  write. */
  addConditionalFormat(sheet: number, rule: ConditionalFormatInput | ConditionalFormat): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatMutate) return -1;
    return orFallback(-1, () => {
      const target = ws(this, sheet);
      const id = target.addCfRule(rule as ConditionalFormatInput);
      return target.listCfRules().findIndex((entry) => entry.id === id);
    });
  }

  getDxf(index: number): DxfRecord | null {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return null;
    return orFallback(null, () => wb(this).ooxml.getDxf(index));
  }

  addDxf(record: DxfRecord): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return -1;
    return orFallback(-1, () => wb(this).ooxml.addDxf(completeDxfRecord(record)));
  }

  dxfCount(): number {
    assertAlive(this);
    if (!this.capabilities.conditionalFormatDxf) return 0;
    return orFallback(0, () => wb(this).ooxml.listDxfs().length);
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
