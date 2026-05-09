import type { WorkbookHandle } from './workbook-handle.js';

export interface PassthroughSummary {
  /** Total preserved OOXML object parts surfaced for host chrome. */
  count: number;
  /** Path roots seen — `xl/charts`, `xl/drawings`, `xl/pivotTables`, etc.
   *  Lets the UI label the badge ("3 charts, 1 pivot"). */
  byCategory: Record<string, number>;
  /** Top-N raw paths for diagnostics / tooltips. */
  paths: string[];
}

export type WorkbookObjectKind =
  | 'charts'
  | 'drawings'
  | 'pivotTables'
  | 'pivotCaches'
  | 'media'
  | 'embeddings'
  | 'comments'
  | 'threadedComments'
  | 'queryTables'
  | 'slicers'
  | 'timelines'
  | 'connections'
  | 'externalLinks'
  | 'vbaProject'
  | 'controls'
  | 'printerSettings'
  | 'customXml'
  | 'other';

export const WORKBOOK_OBJECT_KINDS: readonly WorkbookObjectKind[] = [
  'charts',
  'drawings',
  'media',
  'embeddings',
  'comments',
  'threadedComments',
  'pivotTables',
  'pivotCaches',
  'queryTables',
  'slicers',
  'timelines',
  'connections',
  'externalLinks',
  'controls',
  'printerSettings',
  'customXml',
  'vbaProject',
  'other',
];

const KIND_LABELS: Record<WorkbookObjectKind, string> = {
  charts: 'Charts',
  drawings: 'Drawings',
  media: 'Media',
  embeddings: 'Embedded objects',
  comments: 'Comments',
  threadedComments: 'Threaded comments',
  pivotTables: 'PivotTables',
  pivotCaches: 'Pivot caches',
  queryTables: 'Query tables',
  slicers: 'Slicers',
  timelines: 'Timelines',
  connections: 'Connections',
  externalLinks: 'External links',
  controls: 'Controls',
  printerSettings: 'Printer settings',
  customXml: 'Custom XML',
  vbaProject: 'Macro project',
  other: 'Other',
};

export interface WorkbookObjectRecord {
  kind: WorkbookObjectKind;
  path: string;
  name: string;
  extension: string;
}

export interface TableSummary {
  /** Total Excel Tables across all sheets. */
  count: number;
  /** Per-sheet count keyed by sheet index. */
  bySheet: Record<number, number>;
  /** Display-friendly names — useful for a tooltip. */
  names: string[];
  /** Detailed table metadata for object inspectors and host chrome. */
  items: {
    name: string;
    displayName: string;
    ref: string;
    sheetIndex: number;
    columns: string[];
  }[];
}

export interface PivotTableSummary {
  /** Total projected PivotTables across all sheets. */
  count: number;
  /** Per-sheet count keyed by sheet index. */
  bySheet: Record<number, number>;
  /** Detailed read-only layout metadata. */
  items: {
    sheetIndex: number;
    pivotIndex: number;
    top: number;
    left: number;
    rows: number;
    cols: number;
    cells: number;
    fields: string[];
  }[];
}

/** Categorize OOXML passthrough paths into rough buckets. Path root after
 *  the `xl/` prefix drives the bucket; unknown paths bucket as `other`. */
const CATEGORIES: { match: RegExp; bucket: WorkbookObjectKind }[] = [
  { match: /^xl\/charts\//, bucket: 'charts' },
  { match: /^xl\/drawings\//, bucket: 'drawings' },
  { match: /^xl\/pivotTables\//, bucket: 'pivotTables' },
  { match: /^xl\/pivotCache\//, bucket: 'pivotCaches' },
  { match: /^xl\/media\//, bucket: 'media' },
  { match: /^xl\/embeddings\//, bucket: 'embeddings' },
  { match: /^xl\/comments/, bucket: 'comments' },
  { match: /^xl\/threadedComments\//, bucket: 'threadedComments' },
  { match: /^xl\/queryTables\//, bucket: 'queryTables' },
  { match: /^xl\/slicer/, bucket: 'slicers' },
  { match: /^xl\/timeline/, bucket: 'timelines' },
  { match: /^xl\/connections\.xml$/, bucket: 'connections' },
  { match: /^xl\/externalLinks\//, bucket: 'externalLinks' },
  { match: /^xl\/vbaProject\.bin$/, bucket: 'vbaProject' },
  { match: /^xl\/(?:ctrlProps|activeX)\//, bucket: 'controls' },
  { match: /^xl\/printerSettings\//, bucket: 'printerSettings' },
  { match: /^customXml\//, bucket: 'customXml' },
];

export function classifyWorkbookObjectPath(path: string): WorkbookObjectKind {
  return CATEGORIES.find((c) => c.match.test(path))?.bucket ?? 'other';
}

export function workbookObjectKindLabel(kind: WorkbookObjectKind): string {
  return KIND_LABELS[kind];
}

export function workbookObjectName(path: string): string {
  const slash = path.lastIndexOf('/');
  return slash >= 0 ? path.slice(slash + 1) : path;
}

export function workbookObjectExtension(path: string): string {
  const name = workbookObjectName(path);
  const dot = name.lastIndexOf('.');
  return dot > 0 && dot < name.length - 1 ? name.slice(dot + 1).toLowerCase() : '';
}

export function listWorkbookObjects(wb: WorkbookHandle): readonly WorkbookObjectRecord[] {
  return wb
    .getPassthroughs()
    .map((it) => ({
      kind: classifyWorkbookObjectPath(it.path),
      path: it.path,
      name: workbookObjectName(it.path),
      extension: workbookObjectExtension(it.path),
    }))
    .sort((a, b) => a.kind.localeCompare(b.kind) || a.path.localeCompare(b.path));
}

export function workbookObjectsByKind(
  objects: readonly WorkbookObjectRecord[],
): Record<WorkbookObjectKind, WorkbookObjectRecord[]> {
  const out = emptyObjectBuckets();
  for (const object of objects) out[object.kind].push(object);
  return out;
}

export function workbookObjectKindCounts(
  objects: readonly WorkbookObjectRecord[],
): Record<WorkbookObjectKind, number> {
  const byKind = workbookObjectsByKind(objects);
  return Object.fromEntries(
    WORKBOOK_OBJECT_KINDS.map((kind) => [kind, byKind[kind].length]),
  ) as Record<WorkbookObjectKind, number>;
}

/**
 * Build a passthrough-badge summary off the workbook handle. Pure read —
 * does not mutate engine state. Empty on the stub or when the engine doesn't
 * expose `passthroughCount`.
 */
export function summarizePassthroughs(wb: WorkbookHandle): PassthroughSummary {
  const items = listWorkbookObjects(wb);
  const counts = workbookObjectKindCounts(items);
  const byCategory = Object.fromEntries(
    WORKBOOK_OBJECT_KINDS.filter((kind) => counts[kind] > 0).map((kind) => [kind, counts[kind]]),
  );
  return {
    count: items.length,
    byCategory,
    paths: items.slice(0, 32).map((it) => it.path),
  };
}

const emptyObjectBuckets = (): Record<WorkbookObjectKind, WorkbookObjectRecord[]> =>
  WORKBOOK_OBJECT_KINDS.reduce(
    (out, kind) => {
      out[kind] = [];
      return out;
    },
    {} as Record<WorkbookObjectKind, WorkbookObjectRecord[]>,
  );

/** Build an Excel-Tables summary off the workbook handle. */
export function summarizeTables(wb: WorkbookHandle): TableSummary {
  const items = wb.getTables();
  const bySheet: Record<number, number> = {};
  for (const t of items) {
    bySheet[t.sheetIndex] = (bySheet[t.sheetIndex] ?? 0) + 1;
  }
  return {
    count: items.length,
    bySheet,
    names: items.map((t) => t.displayName || t.name).filter((n) => n.length > 0),
    items,
  };
}

/** Build a read-only PivotTable projection summary off the workbook handle. */
export function summarizePivotTables(wb: WorkbookHandle): PivotTableSummary {
  const items = wb.getPivotTables();
  const bySheet: Record<number, number> = {};
  for (const p of items) {
    bySheet[p.sheetIndex] = (bySheet[p.sheetIndex] ?? 0) + 1;
  }
  return {
    count: items.length,
    bySheet,
    items,
  };
}
