import { parseRangeRef } from '../engine/range-resolver.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { mutators, type SlicerSpec, type SpreadsheetStore } from '../store/store.js';
import { applyFilter, clearFilter, distinctValues } from './filter.js';

export interface CreateSlicerOptions {
  id?: string;
  tableName: string;
  column: string;
  selected?: readonly string[];
  x?: number;
  y?: number;
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
  const target = tableName.toLowerCase();
  for (const table of workbook.getTables()) {
    if (table.name.toLowerCase() === target || table.displayName.toLowerCase() === target) {
      return table;
    }
  }
  return null;
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

  const cleared = new Set<string>();
  for (const entry of resolved) {
    const key = rangeKey(entry.range);
    if (cleared.has(key)) continue;
    clearFilter(store.getState(), store, entry.range);
    cleared.add(key);
  }

  let hiddenCount = 0;
  for (const entry of resolved) {
    if (entry.selected.length === 0) continue;
    const wanted = new Set(entry.selected);
    hiddenCount += applyFilter(store.getState(), store, entry.range, entry.byCol, (cell) =>
      wanted.has(cellToKey(cell?.value)),
    );
  }
  return hiddenCount;
}

const rangeKey = (r: Range): string => `${r.sheet}:${r.r0}:${r.c0}:${r.r1}:${r.c1}`;

const cellToKey = (v: unknown): string => {
  if (!v || typeof v !== 'object') return '';
  const cv = v as { kind: string; value?: unknown };
  if (cv.kind === 'number') return String(cv.value);
  if (cv.kind === 'text') return String(cv.value ?? '');
  if (cv.kind === 'bool') return cv.value ? 'TRUE' : 'FALSE';
  return '';
};

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
