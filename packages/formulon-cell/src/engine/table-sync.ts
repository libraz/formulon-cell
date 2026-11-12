import { TableTotalsRowFunction } from '@libraz/formulon';
import {
  DEFAULT_TABLE_COLOR,
  overlayStyleName,
  parseTableStyleName,
  type TableColumnTotals,
  type TableOverlay,
  type TableTotalsFunction,
} from '../commands/format-as-table.js';
import { mutators, type SpreadsheetStore } from '../store/store.js';
import { parseRangeRef } from './range-resolver.js';
import { hydrateTableStylesFromEngine } from './table-style-sync.js';
import type { Addr, Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { TableStyleDetails } from './workbook-handle-table-styles.js';

type TableReader = Pick<WorkbookHandle, 'getTables'> &
  Partial<
    Pick<
      WorkbookHandle,
      | 'withEngineSyncMuted'
      | 'getTableStyleDetails'
      | 'getCustomTableStyles'
      | 'getPivotStyleInfo'
      | 'getPivotTables'
      | 'getTableAutoFilter'
      | 'capabilities'
    >
  >;

function tableId(name: string, sheet: number, ref: string): string {
  const safe = `${name || 'Table'}-${sheet}-${ref}`.replace(/[^A-Za-z0-9_-]+/g, '-');
  return `engine-table-${safe}`;
}

function tableRange(sheet: number, ref: string): Range | null {
  const parsed = parseRangeRef(ref);
  if (!parsed) return null;
  return {
    sheet,
    r0: parsed.r0,
    c0: parsed.c0,
    r1: parsed.r1,
    c1: parsed.c1,
  };
}

const TOTALS_TO_ENGINE: Readonly<Record<TableTotalsFunction, TableTotalsRowFunction>> = {
  none: TableTotalsRowFunction.None,
  sum: TableTotalsRowFunction.Sum,
  min: TableTotalsRowFunction.Min,
  max: TableTotalsRowFunction.Max,
  average: TableTotalsRowFunction.Average,
  count: TableTotalsRowFunction.Count,
  countNums: TableTotalsRowFunction.CountNumbers,
  stdDev: TableTotalsRowFunction.StdDev,
  var: TableTotalsRowFunction.Var,
  custom: TableTotalsRowFunction.Custom,
};

/** The engine's totals-row function for the store's model name. */
export function totalsFunctionToEngine(fn: TableTotalsFunction): TableTotalsRowFunction {
  return TOTALS_TO_ENGINE[fn];
}

/** The store's model name for an engine totals-row function. */
export function totalsFunctionFromEngine(fn: TableTotalsRowFunction): TableTotalsFunction {
  for (const [name, value] of Object.entries(TOTALS_TO_ENGINE)) {
    if (value === fn) return name as TableTotalsFunction;
  }
  return 'none';
}

/** A totals-row formula as the store keeps it: without the leading `=`. */
export function totalsFormulaText(formula: string | null): string {
  return (formula ?? '').replace(/^=/, '');
}

/** Per-column totals-row settings as the engine reports them. */
export function totalsFromDetails(details: TableStyleDetails): TableColumnTotals[] {
  return details.columns.map((c) => ({
    function: totalsFunctionFromEngine(c.totalsRowFunction),
    label: c.totalsRowLabel ?? '',
    formula: totalsFormulaText(c.totalsRowFormula),
  }));
}

/** SUBTOTAL function codes the engine writes for the built-in totals functions. */
const SUBTOTAL_FUNCTIONS: Readonly<Record<number, TableColumnTotals['function']>> = {
  101: 'average',
  102: 'countNums',
  103: 'count',
  104: 'max',
  105: 'min',
  107: 'stdDev',
  109: 'sum',
  110: 'var',
};

const SUBTOTAL_PATTERN = /^SUBTOTAL\(\s*(\d+)\s*,\s*[^[\]]*\[[^\]]*\]\s*\)$/i;

/** Totals settings that reproduce what a totals-row cell holds: the engine's SUBTOTAL formula is
 *  its built-in function, any other formula a custom one, and a blank cell none. Only a formula
 *  takes the formula path; any other entry, a typed number included, is the column label. */
export function totalsFromCell(wb: Pick<WorkbookHandle, 'cellFormula' | 'getValue'>, addr: Addr) {
  const formula = wb.cellFormula(addr)?.replace(/^=/, '');
  if (formula) {
    const code = SUBTOTAL_PATTERN.exec(formula)?.[1];
    const builtIn = code === undefined ? undefined : SUBTOTAL_FUNCTIONS[Number(code)];
    return builtIn
      ? { function: builtIn, label: '', formula: '' }
      : { function: 'custom' as const, label: '', formula };
  }
  const value = wb.getValue(addr);
  if (value.kind === 'text') return { function: 'none' as const, label: value.value, formula: '' };
  if (value.kind === 'number')
    return { function: 'none' as const, label: String(value.value), formula: '' };
  if (value.kind === 'bool') {
    return { function: 'none' as const, label: value.value ? 'TRUE' : 'FALSE', formula: '' };
  }
  return { function: 'none' as const, label: '', formula: '' };
}

/** True when two totals models describe the same settings; absent equals empty. */
export function sameTotals(
  a: readonly TableColumnTotals[] | undefined,
  b: readonly TableColumnTotals[] | undefined,
): boolean {
  if (!a || !b) return a === b;
  return (
    a.length === b.length &&
    a.every(
      (t, i) =>
        t.function === b[i]?.function &&
        (t.label ?? '') === (b[i]?.label ?? '') &&
        (t.formula ?? '') === (b[i]?.formula ?? ''),
    )
  );
}

/** Look and flags of an engine table from its style details; a built-in style
 *  supplies its family and hue, any other the defaults. */
function overlayLook(
  details: TableStyleDetails,
): Pick<
  TableOverlay,
  | 'style'
  | 'color'
  | 'showHeader'
  | 'showTotal'
  | 'banded'
  | 'firstCol'
  | 'lastCol'
  | 'colStripes'
  | 'styleName'
> {
  const look = parseTableStyleName(details.styleName);
  return {
    style: look?.style ?? 'medium',
    color: look?.color ?? DEFAULT_TABLE_COLOR,
    showHeader: details.headerRowCount > 0,
    showTotal: details.totalsRowCount > 0,
    banded: details.showRowStripes,
    firstCol: details.showFirstColumn,
    lastCol: details.showLastColumn,
    colStripes: details.showColumnStripes,
    styleName: details.styleName,
  };
}

/** Convert the engine's worksheet tables into engine-backed overlays, with
 *  the style, flags and header / totals rows the engine reports. */
export function tableOverlaysFromEngine(wb: TableReader): TableOverlay[] {
  const out: TableOverlay[] = [];
  for (const table of wb.getTables()) {
    const range = tableRange(table.sheetIndex, table.ref);
    if (!range) continue;
    const details = wb.getTableStyleDetails?.(table.index);
    out.push({
      id: tableId(table.displayName || table.name, table.sheetIndex, table.ref),
      source: 'engine',
      name: table.name,
      range,
      ...(details
        ? { ...overlayLook(details), totals: totalsFromDetails(details) }
        : { style: 'medium', showHeader: true, showTotal: false, banded: true }),
      // A table without its AutoFilter shows no header filter buttons.
      ...(wb.capabilities?.typedTableAutoFilter && wb.getTableAutoFilter?.(table.index) === null
        ? { filterButton: false }
        : {}),
    });
  }
  return out;
}

/** Load the engine's tables and custom styles into the store without echoing
 *  them back. A known table whose engine style and flags are unchanged keeps
 *  the exact look the user gave it. */
export function hydrateTableOverlaysFromEngine(wb: TableReader, store: SpreadsheetStore): void {
  const hydrate = (): void => {
    if (wb.capabilities?.customTableStyles) {
      hydrateTableStylesFromEngine(wb as Parameters<typeof hydrateTableStylesFromEngine>[0], store);
    }
    const known = new Map<string, TableOverlay>();
    for (const t of store.getState().tables.tables) {
      if (t.source === 'engine' && t.name) known.set(t.name.toLocaleLowerCase(), t);
    }
    const overlays = tableOverlaysFromEngine(wb).map((t) => {
      const previous = t.name ? known.get(t.name.toLocaleLowerCase()) : undefined;
      return previous && t.styleName !== undefined && sameEngineLook(previous, t)
        ? { ...t, style: previous.style, color: previous.color, styleName: previous.styleName }
        : t;
    });
    mutators.replaceEngineTableOverlays(store, overlays);
  };
  // A reader without the mute (a partial view of the handle) hydrates directly.
  if (wb.withEngineSyncMuted) wb.withEngineSyncMuted(hydrate);
  else hydrate();
}

/** True when two overlays are written to the engine as the same style and flags. */
function sameEngineLook(a: TableOverlay, b: TableOverlay): boolean {
  return (
    overlayStyleName(a) === overlayStyleName(b) &&
    a.showHeader === b.showHeader &&
    a.showTotal === b.showTotal &&
    a.banded === b.banded &&
    !!a.firstCol === !!b.firstCol &&
    !!a.lastCol === !!b.lastCol &&
    !!a.colStripes === !!b.colStripes &&
    (a.filterButton !== false) === (b.filterButton !== false)
  );
}
