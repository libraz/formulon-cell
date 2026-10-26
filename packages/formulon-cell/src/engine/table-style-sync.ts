import type { TableStyle as EngineTableStyle, TableStyleElement } from '@libraz/formulon';
import {
  type CustomTableStyle,
  customPivotTableStyleId,
  customTableStyleId,
  DEFAULT_TABLE_COLOR,
  type PivotTableStyleAssignment,
  type TableStyle,
  tableStyleSwatch,
  tableVariantFromOptions,
  tableVariantOptions,
} from '../commands/format-as-table.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import type { CellFormat, TablesSlice } from '../store/types.js';
import { dxfToApply } from './cf-sync.js';
import { dxfRecordFromApply } from './cf-writeback.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { CustomEngineTableStyle } from './workbook-handle-table-styles.js';

/**
 * Custom table / PivotTable styles between the store's gallery model (family,
 * hue, variant) and the engine's `<tableStyles>` (elements over differential
 * formats). The gallery model is written as a fixed element set - whole-table
 * border in the hue, header fill and text, optional stripe and first-column
 * elements - and read back from the same elements, so a style authored here
 * returns unchanged. A style authored elsewhere is approximated from its
 * header fill and stripe elements.
 */

const FAMILIES: readonly TableStyle[] = ['medium', 'light', 'dark'];

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

/** The engine style for a gallery style; null when its differential formats cannot be authored. */
export function engineStyleFromCustom(
  wb: WorkbookHandle,
  style: CustomTableStyle,
  usage: { pivot: boolean; table: boolean },
): EngineTableStyle | null {
  if (!wb.capabilities.conditionalFormatDxf) return null;
  const color = style.color ?? DEFAULT_TABLE_COLOR;
  const swatch = tableStyleSwatch(style.style, color);
  const { banded, firstCol } = tableVariantOptions(style.variant);
  const side = { style: 'thin' as const, color };
  const elements: TableStyleElement[] = [];
  const add = (type: TableStyleElement['type'], apply: Partial<CellFormat>): boolean => {
    const dxfId = dxfIdFor(wb, apply);
    if (dxfId === null) return false;
    elements.push({ type, dxfId });
    return true;
  };
  const ok =
    add('wholeTable', { borders: { top: side, bottom: side, left: side, right: side } }) &&
    add('headerRow', { fill: swatch.header, color: swatch.headerText, bold: true }) &&
    add('totalRow', { bold: true }) &&
    (!banded || add('firstRowStripe', { fill: swatch.band })) &&
    (!firstCol || add('firstColumn', { bold: true }));
  return ok ? { name: style.label, ...usage, elements } : null;
}

const norm = (color: string | undefined): string | undefined => color?.toLowerCase();

const luminance = (hex: string): number => {
  const h = hex.replace('#', '');
  const channel = (i: number): number => Number.parseInt(h.slice(i, i + 2), 16) / 255;
  return 0.2126 * channel(0) + 0.7152 * channel(2) + 0.0722 * channel(4);
};

/** The gallery model of an engine style. */
export function customStyleFromEngine(
  wb: WorkbookHandle,
  style: CustomEngineTableStyle,
  pivot: boolean,
): CustomTableStyle {
  const apply = (type: TableStyleElement['type']): Partial<CellFormat> | null => {
    const element = style.elements.find((e) => e.type === type);
    return element?.dxfId === undefined ? null : dxfToApply(wb, element.dxfId);
  };
  const header = apply('headerRow');
  const stripe = apply('firstRowStripe');
  const border = apply('wholeTable')?.borders?.top;
  const borderColor = !border
    ? undefined
    : border === true
      ? '#000000'
      : (border.color ?? '#000000');
  const headerFill = norm(header?.fill);
  const color = norm(borderColor ?? headerFill ?? stripe?.fill) ?? DEFAULT_TABLE_COLOR;
  const matching = FAMILIES.find((family) => {
    const swatch = tableStyleSwatch(family, color);
    return headerFill === norm(swatch.header) && norm(header?.color) === norm(swatch.headerText);
  });
  const family: TableStyle =
    matching ??
    (headerFill === undefined || luminance(headerFill) > 0.75
      ? 'light'
      : luminance(headerFill) < luminance(color) * 0.8
        ? 'dark'
        : 'medium');
  return {
    id: pivot ? customPivotTableStyleId(style.name) : customTableStyleId(style.name),
    label: style.name,
    style: family,
    color,
    variant: tableVariantFromOptions({
      banded: style.elements.some((e) => e.type === 'firstRowStripe'),
      firstCol: style.elements.some((e) => e.type === 'firstColumn'),
    }),
  };
}

/** Custom styles, and the PivotTable assignments naming them, as the engine holds them. */
export function tableStylesFromEngine(
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
  a.style === b.style &&
  a.color === b.color &&
  a.variant === b.variant;

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
      list.find((s) => s.label.toLocaleLowerCase() === label.toLocaleLowerCase());
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
