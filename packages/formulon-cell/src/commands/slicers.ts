import { rangeKey } from '../engine/address.js';
import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SlicerSpec, type SpreadsheetStore } from '../store/store.js';
import { authorizedWrite, authorizeWrite } from './authorized-write.js';
import { drawingLocked } from './drawing-lock.js';
import {
  applyValueFilter,
  distinctValues,
  recordFilterChange,
  resetFilterCriteria,
} from './filter.js';
import { History } from './history.js';
import type { OperationIntent } from './interaction-policy.js';
import { type ObjectActionContext, objectChangeAllowed, objectIntents } from './object-actions.js';
import { alertFilterRefused } from './protection-alert.js';
import { recordSlicersChange } from './slice-history.js';

export interface CreateSlicerOptions {
  id?: string;
  tableName: string;
  column: string;
  selected?: readonly string[];
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  /** Sheet hosting the slicer; defaults to the sheet of its table. */
  sheet?: number;
  caption?: string;
  style?: string;
  columnCount?: number;
}

export type CreateSlicerResult =
  | { ok: true; spec: SlicerSpec }
  | { ok: false; reason: 'table-not-found' | 'column-not-found' };

interface SlicerTable {
  name: string;
  displayName: string;
  ref: string;
  sheetIndex: number;
  columns: string[];
}

export function listSlicers(store: SpreadsheetStore): readonly SlicerSpec[] {
  return store.getState().slicers.slicers;
}

export function findSlicerTable(workbook: WorkbookHandle, tableName: string): SlicerTable | null {
  const index = workbook.tableIndexByName(tableName);
  return index === null ? null : (workbook.getTables()[index] ?? null);
}

export function resolveSlicerSpec(
  workbook: WorkbookHandle,
  spec: SlicerSpec,
): { range: Range; byCol: number } | null {
  const table = findSlicerTable(workbook, spec.tableName);
  if (!table) return null;
  const parsed = parseRangeRef(table.ref);
  if (!parsed) return null;
  const colIdx = table.columns.indexOf(spec.column);
  if (colIdx < 0) return null;
  return {
    range: {
      sheet: table.sheetIndex,
      r0: parsed.r0,
      c0: parsed.c0,
      r1: parsed.r1,
      c1: parsed.c1,
    },
    byCol: parsed.c0 + colIdx,
  };
}

export function listSlicerValues(
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  spec: SlicerSpec,
): string[] {
  const resolved = resolveSlicerSpec(workbook, spec);
  if (!resolved) return [];
  return distinctValues(store.getState(), resolved.range, resolved.byCol);
}

export function createSlicer(
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  options: CreateSlicerOptions,
): CreateSlicerResult {
  const table = findSlicerTable(workbook, options.tableName);
  if (!table) return { ok: false, reason: 'table-not-found' };
  if (!table.columns.includes(options.column)) return { ok: false, reason: 'column-not-found' };
  const spec: SlicerSpec = {
    id: options.id ?? nextSlicerId(store, table.name, options.column),
    tableName: table.name,
    column: options.column,
    selected: options.selected ? [...options.selected] : [],
    x: options.x,
    y: options.y,
    ...(options.sheet !== undefined ? { sheet: options.sheet } : {}),
    ...(options.w !== undefined ? { w: options.w } : {}),
    ...(options.h !== undefined ? { h: options.h } : {}),
    ...(options.caption !== undefined ? { caption: options.caption } : {}),
    ...(options.style !== undefined ? { style: options.style } : {}),
    ...(options.columnCount !== undefined ? { columnCount: options.columnCount } : {}),
  };
  mutators.addSlicer(store, spec);
  return { ok: true, spec };
}

export function updateSlicer(
  store: SpreadsheetStore,
  id: string,
  patch: Partial<Omit<SlicerSpec, 'id'>>,
): void {
  mutators.updateSlicer(store, id, patch);
}

export function removeSlicer(store: SpreadsheetStore, id: string): void {
  mutators.removeSlicer(store, id);
}

export function setSlicerSelected(
  store: SpreadsheetStore,
  id: string,
  values: readonly string[],
): void {
  mutators.setSlicerSelected(store, id, values);
}

export function clearSlicerSelection(store: SpreadsheetStore, id: string): void {
  setSlicerSelected(store, id, []);
}

export function recomputeSlicerFilters(store: SpreadsheetStore, workbook: WorkbookHandle): number {
  const resolved = listSlicers(store)
    .map((spec) => {
      const r = resolveSlicerSpec(workbook, spec);
      return r ? { ...r, selected: spec.selected } : null;
    })
    .filter(
      (entry): entry is { range: Range; byCol: number; selected: readonly string[] } =>
        entry !== null,
    );

  // Group every slicer by its range so each range's columns are ANDed in a
  // single pass. A range with no active selection is fully revealed. Each
  // selection becomes a value criterion on the range's filter, so a table
  // slicer filters through the table's own criteria.
  const groups = new Map<
    string,
    { range: Range; cols: Array<{ byCol: number; wanted: Set<string> }> }
  >();
  const order: string[] = [];
  for (const entry of resolved) {
    const key = rangeKey(entry.range);
    let group = groups.get(key);
    if (!group) {
      group = { range: entry.range, cols: [] };
      groups.set(key, group);
      order.push(key);
    }
    if (entry.selected.length === 0) continue;
    group.cols.push({ byCol: entry.byCol, wanted: new Set(entry.selected) });
  }

  let hiddenCount = 0;
  for (const key of order) {
    const group = groups.get(key);
    if (!group) continue;
    // Reset rather than clear, so the range's own filter (a table's) survives.
    resetFilterCriteria(store, group.range);
    // Each call reports the range's total under the criteria applied so far.
    let groupHidden = 0;
    for (const { byCol, wanted } of group.cols) {
      const state = store.getState();
      const hiddenValues = distinctValues(state, group.range, byCol).filter((v) => !wanted.has(v));
      groupHidden = applyValueFilter(state, store, group.range, byCol, hiddenValues);
    }
    hiddenCount += groupHidden;
  }
  return hiddenCount;
}

const nextSlicerId = (store: SpreadsheetStore, tableName: string, column: string): string => {
  const base = `slicer-${slug(tableName)}-${slug(column)}`;
  const existing = new Set(listSlicers(store).map((s) => s.id));
  if (!existing.has(base)) return base;
  let i = 2;
  while (existing.has(`${base}-${i}`)) i += 1;
  return `${base}-${i}`;
};

const slug = (value: string): string =>
  value
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'item';

/** A slicer change and the entry point that asked for it. */
export interface SlicerActionContext extends ObjectActionContext {
  readonly workbook: WorkbookHandle;
}

/** Whether slicer `spec` is locked (see {@link drawingLocked}). */
const slicerLocked = (_spec: SlicerSpec): boolean => drawingLocked('slicer');

const slicerSheet = (ctx: SlicerActionContext, spec: SlicerSpec): number =>
  spec.sheet ?? ctx.store.getState().data.sheetIndex;

/** Whether slicer `spec` responds to clicks, as a drawing: unlocked, or Edit objects allowed. A
 *  slicer that does not is inert: its items and Clear do nothing and show no alert. */
export const slicerInteractive = (ctx: SlicerActionContext, spec: SlicerSpec): boolean =>
  !slicerLocked(spec) || objectChangeAllowed(ctx, slicerSheet(ctx, spec));

/**
 * Change a slicer and recompute the rows its filter hides, as one authorized step: an object
 * change of the slicer's sheet (a locked slicer, or deleting one) and, for a change of what it
 * selects (`filters`), a filter use of the table it filters (Use AutoFilter on a protected sheet).
 * A usable slicer whose filter the protected sheet refuses says so. The slicer and the filter it
 * drives both undo together.
 */
const slicerWrite = (
  ctx: SlicerActionContext,
  spec: SlicerSpec,
  mutate: () => void,
  filters = true,
): boolean => {
  const sheet = slicerSheet(ctx, spec);
  const range = filters ? resolveSlicerSpec(ctx.workbook, spec)?.range : undefined;
  const asObject = !filters || slicerLocked(spec);
  if (asObject && !objectChangeAllowed(ctx, sheet)) return false;
  const filterIntent = (origin: OperationIntent['origin']): OperationIntent[] =>
    range
      ? [
          {
            operation: 'filter' as const,
            origin,
            commandId: ctx.commandId,
            effects: [{ kind: 'range' as const, range }],
          },
        ]
      : [];
  const filterDecision = authorizeWrite(ctx.store, filterIntent, ctx.origin);
  if (!filterDecision.allowed) {
    if (filterDecision.code === 'protected') alertFilterRefused(ctx.workbook);
    return false;
  }
  const intents = (origin: OperationIntent['origin']): OperationIntent[] => [
    ...(asObject ? objectIntents(ctx, sheet)(origin) : []),
    ...filterIntent(origin),
  ];
  const history = ctx.history ?? new History();
  return authorizedWrite({
    store: ctx.store,
    history,
    intents,
    origin: ctx.origin,
    write: () => {
      recordSlicersChange(history, ctx.store, mutate);
      recordFilterChange(history, ctx.store, () => recomputeSlicerFilters(ctx.store, ctx.workbook));
    },
  }).ok;
};

/** Select `values` in slicer `id` (none selects every item), filtering the rows it drives. */
export function setSlicerSelectionAction(
  ctx: SlicerActionContext,
  id: string,
  values: readonly string[],
): boolean {
  const spec = listSlicers(ctx.store).find((entry) => entry.id === id);
  return spec ? slicerWrite(ctx, spec, () => setSlicerSelected(ctx.store, id, values)) : false;
}

/** Delete slicer `id`, showing again the rows only it hid. The slicer is a drawing, so Edit
 *  objects decides it on a protected sheet; Use AutoFilter does not. */
export function removeSlicerAction(ctx: SlicerActionContext, id: string): boolean {
  const spec = listSlicers(ctx.store).find((entry) => entry.id === id);
  return spec ? slicerWrite(ctx, spec, () => removeSlicer(ctx.store, id), false) : false;
}
