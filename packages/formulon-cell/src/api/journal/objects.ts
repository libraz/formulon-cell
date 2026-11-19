import type {
  AutoFilter,
  CellRect,
  Sheet as EngineSheet,
  PivotOptions,
  TableProperties,
  TableStyleOptions,
} from '@libraz/formulon';
import { pivotFilterBlockRows } from '../../commands/pivot-actions.js';
import { recreatePivotTable } from '../../commands/pivot-table.js';
import { captureSlicerRestore } from '../../engine/slicer-sync.js';
import type { Addr } from '../../engine/types.js';
import type { CellSnapshot } from '../../engine/workbook-handle.js';
import { PIVOT_CACHE_CHANGES } from '../pivot-cache.js';
import type { JournalCodec, JournalFamily, JournalScope } from './index.js';

const sheetOf = (scope: JournalScope): EngineSheet => scope.engine.sheet(scope.sheet as number);

/** The rows of a range a table write may rewrite: its first and last, `grow` rows beyond. */
const edgeRows = (rect: CellRect, grow: number): CellRect[] => [
  { ...rect, firstRow: Math.max(0, rect.firstRow - grow), lastRow: rect.firstRow },
  { ...rect, firstRow: rect.lastRow, lastRow: rect.lastRow + grow },
];

/** The range a table write names in its first argument, if any. */
const writtenRange = (scope: JournalScope): CellRect | null => {
  const arg = scope.args[0] as { range?: unknown } | undefined;
  if (scope.name !== 'addTable' && scope.name !== 'update') return null;
  const range = arg?.range;
  if (typeof range === 'string') return scope.engine.parseReference(range).range;
  if (range && typeof range === 'object') {
    return 'rect' in range ? (range as { rect: CellRect }).rect : (range as CellRect);
  }
  return null;
};

/** The header and totals rows the write can rewrite: those of the sheet's tables (one row
 *  beyond for a toggled header or totals row) and of the range the write names. */
const captureTableEdges = (
  scope: JournalScope,
  tables: readonly TableProperties[],
): CellSnapshot[] => {
  const sheet = scope.sheet as number;
  const rects = tables.flatMap((table) => edgeRows(table.range.rect, 1));
  const written = writtenRange(scope);
  if (written) rects.push(...edgeRows(written, 1));
  const cells = new Map<string, CellSnapshot>();
  for (const rect of rects) {
    for (let row = rect.firstRow; row <= rect.lastRow; row += 1) {
      for (let col = rect.firstCol; col <= rect.lastCol; col += 1) {
        const key = `${row}:${col}`;
        if (!cells.has(key)) cells.set(key, scope.handle.captureCell({ sheet, row, col }));
      }
    }
  }
  return [...cells.values()];
};

/** Tables of a sheet: properties, style and filter of each, restored by name. */
const TABLES: JournalCodec = {
  id: 'tables',
  scopeKey: (scope) => `sheet:${scope.sheet}`,
  capture(scope) {
    const sheet = sheetOf(scope);
    const before = sheet.listTables().map((table) => ({
      props: table.getProperties() as TableProperties,
      style: table.getStyleOptions() as TableStyleOptions | null,
      filter: table.getAutoFilter() as AutoFilter | null,
    }));
    const cells = captureTableEdges(
      scope,
      before.map((entry) => entry.props),
    );
    return () => {
      const names = new Set(before.map((entry) => entry.props.name.toLowerCase()));
      for (const table of [...sheet.listTables()].reverse()) {
        if (!names.has(table.name.toLowerCase())) sheet.removeTable(table);
      }
      for (const { props, style, filter } of before) {
        const present = sheet
          .listTables()
          .find((table) => table.name.toLowerCase() === props.name.toLowerCase());
        if (!present) {
          sheet.addTable({
            range: props.range.rect,
            name: props.name,
            columns: props.columns.map((column) => column.name),
            headerRow: props.headerRow,
            totalsRow: props.totalsRow,
            styled: style !== null,
            ...(style
              ? {
                  styleName: style.styleName,
                  showFirstColumn: style.showFirstColumn,
                  showLastColumn: style.showLastColumn,
                  showRowStripes: style.showRowStripes,
                  showColumnStripes: style.showColumnStripes,
                }
              : {}),
          });
          continue;
        }
        present.update(props);
        if (style) present.updateStyleOptions(style);
        present.setAutoFilter(filter);
      }
      scope.handle.restoreCells(cells);
    };
  },
};

/** Custom table styles and the workbook's style defaults. */
const TABLE_STYLES: JournalCodec = {
  id: 'table-styles',
  scopeKey: () => 'workbook',
  capture(scope) {
    const { engine } = scope;
    const styles = engine.listTableStyles();
    const defaults = engine.getTableStyleDefaults();
    return () => {
      for (const style of styles) engine.setTableStyle(style);
      engine.updateTableStyleDefaults(defaults);
      const names = new Set(styles.map((style) => style.name));
      for (const style of engine.listTableStyles()) {
        if (!names.has(style.name)) engine.removeTableStyle(style.name);
      }
    };
  },
};

type DrawingBefore =
  | { readonly id: number; readonly snapshot: Uint8Array }
  | { readonly id: number; readonly restore: () => boolean };

/** Every drawing object of a sheet, restored under its id and stacking slot. */
const DRAWINGS: JournalCodec = {
  id: 'drawings',
  scopeKey: (scope) => `sheet:${scope.sheet}`,
  capture(scope) {
    const sheetIndex = scope.sheet as number;
    const sheet = sheetOf(scope);
    const before: DrawingBefore[] = [];
    for (const { objectId: id } of sheet.listDrawings()) {
      const snapshot = scope.handle.snapshotDrawing(sheetIndex, id);
      if (snapshot) {
        before.push({ id, snapshot });
        continue;
      }
      const restore = captureSlicerRestore(scope.handle, sheetIndex, id);
      if (!restore) return null;
      before.push({ id, restore });
    }
    return () => {
      for (const { objectId: id } of [...sheet.listDrawings()].reverse()) {
        if (sheet.listDrawings().some((drawing) => drawing.objectId === id))
          sheet.removeDrawing(id);
      }
      for (const entry of before) {
        if ('snapshot' in entry) sheet.restoreDrawing(entry.snapshot);
        else entry.restore();
      }
    };
  },
};

interface PivotBefore {
  readonly definition: NonNullable<ReturnType<JournalScope['handle']['getPivotDefinition']>>;
  readonly anchor: { row: number; col: number };
  readonly options: PivotOptions | null;
}

/** Pivot tables of a sheet, rebuilt from their stored definitions. */
const PIVOTS: JournalCodec = {
  id: 'pivots',
  scopeKey: (scope) => `sheet:${scope.sheet}`,
  capture(scope) {
    const { handle } = scope;
    const sheetIndex = scope.sheet as number;
    const pivots = sheetOf(scope).listPivots();
    const before: PivotBefore[] = [];
    for (let index = 0; index < pivots.length; index += 1) {
      const pivot = pivots[index];
      const definition = handle.getPivotDefinition(sheetIndex, index);
      const range = pivot?.getLayout().range;
      if (!definition || !range) return null;
      before.push({
        definition,
        anchor: { row: range.firstRow + pivotFilterBlockRows(definition), col: range.firstCol },
        options: handle.pivotOptions(sheetIndex, index),
      });
    }
    return () => {
      for (let index = sheetOf(scope).listPivots().length - 1; index >= 0; index -= 1) {
        handle.removePivotTable(sheetIndex, index);
      }
      for (const { definition, anchor, options } of before) {
        const index = recreatePivotTable(handle, sheetIndex, anchor, definition);
        if (index < 0) throw new Error(`cannot restore pivot table ${definition.name}`);
        if (options) handle.updatePivotOptions(sheetIndex, index, options);
      }
    };
  },
};

/** The ids of the workbook's pivot caches. */
const CACHE_SET: JournalCodec = {
  id: 'pivot-cache-set',
  scopeKey: () => 'workbook',
  capture(scope) {
    const { engine } = scope;
    const ids = engine.listPivotCaches().map((cache) => cache.cacheId);
    return () => {
      const keep = new Set(ids);
      for (const cache of engine.listPivotCaches()) {
        if (!keep.has(cache.cacheId)) engine.removePivotCache(cache);
      }
      const present = new Set(engine.listPivotCaches().map((cache) => cache.cacheId));
      for (const id of ids) if (!present.has(id)) engine.addPivotCache(id);
    };
  },
};

/** The data table whose result block starts at the write's range, and the cells it overwrites. */
const DATA_TABLE: JournalCodec = {
  id: 'data-table',
  scopeKey: (scope) => `${scope.sheet}:${scope.range?.firstRow}:${scope.range?.firstCol}`,
  capture(scope) {
    const range = scope.range as CellRect | undefined;
    if (!range) return null;
    const sheet = sheetOf(scope);
    const sheetIndex = scope.sheet as number;
    const before = sheet.getDataTable(range.firstRow, range.firstCol);
    const cells: CellSnapshot[] = [];
    if (!before) {
      for (let row = range.firstRow; row <= range.lastRow; row += 1) {
        for (let col = range.firstCol; col <= range.lastCol; col += 1) {
          cells.push(scope.handle.captureCell({ sheet: sheetIndex, row, col }));
        }
      }
    }
    const addrs: Addr[] = [];
    for (let row = range.firstRow; row <= range.lastRow; row += 1) {
      for (let col = range.firstCol; col <= range.lastCol; col += 1) {
        addrs.push({ sheet: sheetIndex, row, col });
      }
    }
    // The result cells change with the record, so each one is announced.
    return () =>
      scope.handle.writeCellsUnjournaled(addrs, () => {
        if (sheet.getDataTable(range.firstRow, range.firstCol)) {
          sheet.removeDataTable(range.firstRow, range.firstCol);
        }
        if (before) sheet.addDataTable(before);
        else scope.handle.restoreCells(cells);
      });
  },
};

export const OBJECTS_JOURNAL: JournalFamily = {
  kinds: {
    'table/sheet': TABLES,
    'table/workbook': TABLE_STYLES,
    'drawing/sheet': DRAWINGS,
    'pivot/sheet': PIVOTS,
    'pivot/workbook': CACHE_SET,
  },
  members: {
    pivotCache: Object.fromEntries(PIVOT_CACHE_CHANGES.pivot.map((name) => [name, 'reset'])),
    sheet: { addDataTable: DATA_TABLE, removeDataTable: DATA_TABLE },
  },
};
