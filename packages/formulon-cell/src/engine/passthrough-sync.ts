import type { WorkbookHandle } from './workbook-handle.js';

export interface PassthroughSummary {
  /** Total non-rendered OOXML parts preserved by the engine. */
  count: number;
  /** Path roots seen — `xl/charts`, `xl/drawings`, `xl/pivotTables`, etc.
   *  Lets the UI label the badge ("3 charts, 1 pivot"). */
  byCategory: Record<string, number>;
  /** Top-N raw paths for diagnostics / tooltips. */
  paths: string[];
}

export interface TableSummary {
  /** Total Excel Tables across all sheets. */
  count: number;
  /** Per-sheet count keyed by sheet index. */
  bySheet: Record<number, number>;
  /** Display-friendly names — useful for a tooltip. */
  names: string[];
}

/** Categorize OOXML passthrough paths into rough buckets. Path root after
 *  the `xl/` prefix drives the bucket; unknown paths bucket as `other`. */
const CATEGORIES: { match: RegExp; bucket: string }[] = [
  { match: /^xl\/charts\//, bucket: 'charts' },
  { match: /^xl\/drawings\//, bucket: 'drawings' },
  { match: /^xl\/pivotTables\//, bucket: 'pivotTables' },
  { match: /^xl\/pivotCache\//, bucket: 'pivotCaches' },
  { match: /^xl\/media\//, bucket: 'media' },
  { match: /^xl\/embeddings\//, bucket: 'embeddings' },
  { match: /^xl\/threadedComments\//, bucket: 'threadedComments' },
  { match: /^xl\/queryTables\//, bucket: 'queryTables' },
];

/**
 * Build a passthrough-badge summary off the workbook handle. Pure read —
 * does not mutate engine state. Empty on the stub or when the engine doesn't
 * expose `passthroughCount`.
 */
export function summarizePassthroughs(wb: WorkbookHandle): PassthroughSummary {
  const items = wb.getPassthroughs();
  const byCategory: Record<string, number> = {};
  for (const it of items) {
    const bucket = CATEGORIES.find((c) => c.match.test(it.path))?.bucket ?? 'other';
    byCategory[bucket] = (byCategory[bucket] ?? 0) + 1;
  }
  return {
    count: items.length,
    byCategory,
    paths: items.slice(0, 32).map((it) => it.path),
  };
}

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
  };
}
