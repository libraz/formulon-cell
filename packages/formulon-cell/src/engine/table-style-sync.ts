import type { TableStyleElementInput, TableStyleInput } from '@libraz/formulon';
import {
  type CustomTableStyle,
  customPivotTableStyleId,
  customTableStyleId,
  type PivotTableStyleAssignment,
  type StripeElement,
} from '../commands/format-as-table.js';
import { sameJsonValue } from '../commands/prepared-change.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { nameKey } from '../store/table-names.js';
import type { CellFormat, TablesSlice } from '../store/types.js';
import { dxfToApply } from './cf-sync.js';
import { dxfRecordFromApply } from './cf-writeback.js';
import {
  ENGINE_ELEMENT_TYPE,
  elementTypeFromEngine,
  type StyleElementFormat,
  type StyleElementFormats,
  type TableStyleElementType,
} from './table-style-elements.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { CustomEngineTableStyle } from './workbook-handle-table-styles.js';

/**
 * Custom table / PivotTable styles between the store (a format per element,
 * plus stripe band sizes) and the engine's `<tableStyles>` (elements over
 * differential formats). Every element and size is written and read back, so a
 * style round-trips unchanged.
 */

const STRIPES: ReadonlySet<TableStyleElementType> = new Set<TableStyleElementType>([
  'firstRowStripe',
  'secondRowStripe',
  'firstColumnStripe',
  'secondColumnStripe',
]);

/** Pivot style toggles for a PivotTable that had no style element yet. */
const DEFAULT_PIVOT_STYLE = {
  showRowHeaders: true,
  showColHeaders: true,
  showRowStripes: false,
  showColStripes: false,
  showLastColumn: true,
};

type StyleReader = Pick<
  WorkbookHandle,
  'getCustomTableStyles' | 'getPivotStyleInfo' | 'getPivotTables' | 'capabilities'
>;

const dxfIdCache = new WeakMap<WorkbookHandle, Map<string, number>>();

/** Id of a differential format holding `apply`; addDxf appends, so an id
 *  issued for the same content is reused. */
function dxfIdFor(wb: WorkbookHandle, apply: Partial<CellFormat>): number | null {
  const dxf = dxfRecordFromApply(wb, apply);
  if (!dxf) return null;
  const key = JSON.stringify(dxf);
  let ids = dxfIdCache.get(wb);
  if (!ids) {
    ids = new Map();
    dxfIdCache.set(wb, ids);
  }
  let id = ids.get(key);
  if (id === undefined || !wb.getDxf(id)) {
    id = wb.addDxf(dxf);
    if (id < 0) return null;
    ids.set(key, id);
  }
  return id;
}

/** The engine style for a custom style; null when its differential formats cannot be authored. */
export function engineStyleFromCustom(
  wb: WorkbookHandle,
  style: CustomTableStyle,
  usage: { pivot: boolean; table: boolean },
): TableStyleInput | null {
  if (!wb.capabilities.conditionalFormatDxf) return null;
  const elements: TableStyleElementInput[] = [];
  for (const [type, apply] of Object.entries(style.elements)) {
    const engineType = ENGINE_ELEMENT_TYPE[type as TableStyleElementType];
    const size = STRIPES.has(type as TableStyleElementType)
      ? (style.stripeSizes?.[type as StripeElement] ?? 1)
      : undefined;
    const sized = size === undefined ? {} : { size };
    if (Object.keys(apply).length === 0) {
      elements.push({ type: engineType, ...sized });
      continue;
    }
    const dxfIndex = dxfIdFor(wb, apply);
    if (dxfIndex === null) return null;
    elements.push({ type: engineType, ...sized, dxfIndex });
  }
  return { name: style.label, ...usage, elements };
}

/** The store model of an engine style: every element it lists, and its stripe band sizes. */
export function customStyleFromEngine(
  wb: WorkbookHandle,
  style: CustomEngineTableStyle,
  pivot: boolean,
): CustomTableStyle {
  const elements: StyleElementFormats = {};
  const stripeSizes: Partial<Record<StripeElement, number>> = {};
  for (const element of style.elements) {
    const type = elementTypeFromEngine(element.type);
    if (!type) continue;
    elements[type] = (
      element.dxfIndex == null ? {} : dxfToApply(wb, element.dxfIndex)
    ) as StyleElementFormat;
    if (STRIPES.has(type) && (element.size ?? 1) > 1) {
      stripeSizes[type as StripeElement] = element.size;
    }
  }
  return {
    id: pivot ? customPivotTableStyleId(style.name) : customTableStyleId(style.name),
    label: style.name,
    elements,
    ...(Object.keys(stripeSizes).length > 0 ? { stripeSizes } : {}),
  };
}

/** Custom styles, and the PivotTable assignments naming them, as the engine holds them. */
function tableStylesFromEngine(
  wb: StyleReader,
): Pick<TablesSlice, 'customTableStyles' | 'customPivotTableStyles' | 'pivotTableStyles'> {
  const customTableStyles: CustomTableStyle[] = [];
  const customPivotTableStyles: CustomTableStyle[] = [];
  for (const style of wb.getCustomTableStyles()) {
    if (style.table) {
      customTableStyles.push(customStyleFromEngine(wb as WorkbookHandle, style, false));
    }
    if (style.pivot) {
      customPivotTableStyles.push(customStyleFromEngine(wb as WorkbookHandle, style, true));
    }
  }
  const pivotTableStyles: PivotTableStyleAssignment[] = [];
  for (const pivot of wb.getPivotTables()) {
    const info = wb.getPivotStyleInfo(pivot.sheetIndex, pivot.pivotIndex);
    const styleId = info ? customPivotTableStyleId(info.name) : null;
    if (styleId && customPivotTableStyles.some((s) => s.id === styleId)) {
      pivotTableStyles.push({
        sheetIndex: pivot.sheetIndex,
        pivotIndex: pivot.pivotIndex,
        styleId,
      });
    }
  }
  return { customTableStyles, customPivotTableStyles, pivotTableStyles };
}

/** Load the engine's custom styles and PivotTable assignments into the store. */
export function hydrateTableStylesFromEngine(wb: StyleReader, store: SpreadsheetStore): void {
  if (!wb.capabilities.customTableStyles) return;
  mutators.replaceTableStyles(store, tableStylesFromEngine(wb));
}

type StyleSlices = Pick<
  TablesSlice,
  'customTableStyles' | 'customPivotTableStyles' | 'pivotTableStyles'
>;

const sameStyle = (a: CustomTableStyle, b: CustomTableStyle | undefined): boolean =>
  !!b &&
  a.label === b.label &&
  sameJsonValue(a.elements, b.elements) &&
  sameJsonValue(a.stripeSizes ?? {}, b.stripeSizes ?? {});

const pivotKey = (a: PivotTableStyleAssignment): string => `${a.sheetIndex}:${a.pivotIndex}`;

/** Writes the store's custom styles and PivotTable style assignments to the engine as they change. */
export class TableStyleMirror {
  constructor(private readonly wb: WorkbookHandle) {}

  /** Style a PivotTable carried before an assignment replaced it, by pivot. */
  private readonly replaced = new Map<string, string | undefined>();

  /** True when `after` differs from `before` in a style slice. */
  static changed(before: StyleSlices | null, after: StyleSlices): boolean {
    return (
      before?.customTableStyles !== after.customTableStyles ||
      before?.customPivotTableStyles !== after.customPivotTableStyles ||
      before?.pivotTableStyles !== after.pivotTableStyles
    );
  }

  /** Write what differs between `before` and `after`. Returns the ids of custom styles the engine
   *  refused, or that clash with a same-named style of the other kind, so the caller can drop them. */
  sync(before: StyleSlices | null, after: StyleSlices): string[] {
    const wb = this.wb;
    const rejected = wb.capabilities.customTableStyles ? this.syncCustomStyles(before, after) : [];
    if (wb.capabilities.pivotStyleInfo) this.syncAssignments(before, after, new Set(rejected));
    return rejected;
  }

  private syncCustomStyles(before: StyleSlices | null, after: StyleSlices): string[] {
    const tables = after.customTableStyles ?? [];
    const pivots = after.customPivotTableStyles ?? [];
    const rejected: string[] = [];
    const sameLabel = (list: readonly CustomTableStyle[], label: string) =>
      list.find((s) => nameKey(s.label) === nameKey(label));
    const write = (
      list: readonly CustomTableStyle[],
      previous: readonly CustomTableStyle[],
      otherList: readonly CustomTableStyle[],
      isPivot: boolean,
    ): void => {
      for (const style of list) {
        if (
          sameStyle(
            style,
            previous.find((p) => p.id === style.id),
          )
        )
          continue;
        // One engine style holds both kinds, so a same-named pair must look the same.
        const twin = sameLabel(otherList, style.label);
        if (twin && !sameStyle(style, twin)) {
          rejected.push(style.id);
          continue;
        }
        const engine = engineStyleFromCustom(this.wb, style, {
          pivot: isPivot || !!twin,
          table: !isPivot || !!twin,
        });
        if (!engine || !this.wb.setCustomTableStyle(engine)) rejected.push(style.id);
      }
      for (const style of previous) {
        if (list.some((s) => s.id === style.id)) continue;
        const twin = sameLabel(otherList, style.label);
        if (!twin) {
          this.wb.removeCustomTableStyle(style.label);
          continue;
        }
        // The other kind keeps the style, now on its own.
        const engine = engineStyleFromCustom(this.wb, twin, { pivot: isPivot, table: !isPivot });
        if (engine) this.wb.setCustomTableStyle(engine);
      }
    };
    write(tables, before?.customTableStyles ?? [], pivots, false);
    write(pivots, before?.customPivotTableStyles ?? [], tables, true);
    return rejected;
  }

  private syncAssignments(
    before: StyleSlices | null,
    after: StyleSlices,
    rejected: ReadonlySet<string>,
  ): void {
    const previous = before?.pivotTableStyles ?? [];
    const next = after.pivotTableStyles ?? [];
    const labelOf = (id: string): string | undefined =>
      (after.customPivotTableStyles ?? []).find((s) => s.id === id)?.label;
    for (const gone of previous) {
      if (next.some((a) => pivotKey(a) === pivotKey(gone))) continue;
      const key = pivotKey(gone);
      if (!this.replaced.has(key)) continue;
      const info = this.wb.getPivotStyleInfo(gone.sheetIndex, gone.pivotIndex);
      const name = this.replaced.get(key);
      this.replaced.delete(key);
      if (info) {
        this.wb.setPivotStyleInfo(gone.sheetIndex, gone.pivotIndex, { ...info, name });
      }
    }
    for (const assignment of next) {
      const old = previous.find((a) => pivotKey(a) === pivotKey(assignment));
      if (old?.styleId === assignment.styleId) continue;
      const label = rejected.has(assignment.styleId) ? undefined : labelOf(assignment.styleId);
      if (label === undefined) continue;
      const key = pivotKey(assignment);
      const info = this.wb.getPivotStyleInfo(assignment.sheetIndex, assignment.pivotIndex);
      if (!this.replaced.has(key)) this.replaced.set(key, info?.name || undefined);
      this.wb.setPivotStyleInfo(assignment.sheetIndex, assignment.pivotIndex, {
        ...(info ?? DEFAULT_PIVOT_STYLE),
        name: label,
      });
    }
  }
}
