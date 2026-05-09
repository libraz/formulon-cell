import {
  summarizePassthroughs,
  summarizePivotTables,
  summarizeTables,
} from './passthrough-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

export type SpreadsheetCompatibilityStatus = 'writable' | 'read-only' | 'session' | 'unsupported';

export type SpreadsheetCompatibilityId =
  | 'cell-formatting'
  | 'conditional-formatting'
  | 'data-validation'
  | 'hyperlinks'
  | 'comments'
  | 'defined-names'
  | 'sheet-protection'
  | 'sheet-views'
  | 'loaded-tables'
  | 'format-as-table'
  | 'pivot-layouts'
  | 'pivot-authoring'
  | 'session-charts'
  | 'charts-drawings'
  | 'chart-authoring'
  | 'external-links';

export interface SpreadsheetCompatibilityItem {
  id: SpreadsheetCompatibilityId;
  label: string;
  status: SpreadsheetCompatibilityStatus;
  count?: number;
  reason: string;
}

export interface SpreadsheetCompatibilitySummary {
  items: SpreadsheetCompatibilityItem[];
  byStatus: Record<SpreadsheetCompatibilityStatus, number>;
  byId: Record<SpreadsheetCompatibilityId, SpreadsheetCompatibilityItem>;
}

function byStatus(
  items: readonly SpreadsheetCompatibilityItem[],
): Record<SpreadsheetCompatibilityStatus, number> {
  return {
    writable: items.filter((i) => i.status === 'writable').length,
    'read-only': items.filter((i) => i.status === 'read-only').length,
    session: items.filter((i) => i.status === 'session').length,
    unsupported: items.filter((i) => i.status === 'unsupported').length,
  };
}

function byId(
  items: readonly SpreadsheetCompatibilityItem[],
): Record<SpreadsheetCompatibilityId, SpreadsheetCompatibilityItem> {
  return Object.fromEntries(items.map((item) => [item.id, item])) as Record<
    SpreadsheetCompatibilityId,
    SpreadsheetCompatibilityItem
  >;
}

export function spreadsheetCompatibilityItem(
  summary: SpreadsheetCompatibilitySummary,
  id: SpreadsheetCompatibilityId,
): SpreadsheetCompatibilityItem {
  return summary.byId[id];
}

export function spreadsheetCompatibilityStatus(
  summary: SpreadsheetCompatibilitySummary,
  id: SpreadsheetCompatibilityId,
): SpreadsheetCompatibilityStatus {
  return spreadsheetCompatibilityItem(summary, id).status;
}

export function isSpreadsheetFeatureWritable(
  summary: SpreadsheetCompatibilitySummary,
  id: SpreadsheetCompatibilityId,
): boolean {
  return spreadsheetCompatibilityStatus(summary, id) === 'writable';
}

export function isSpreadsheetFeatureAvailable(
  summary: SpreadsheetCompatibilitySummary,
  id: SpreadsheetCompatibilityId,
): boolean {
  return spreadsheetCompatibilityStatus(summary, id) !== 'unsupported';
}

/** Workbook-level spreadsheet compatibility summary for host chrome. This is a
 *  conservative contract: a feature is marked `writable` only when the engine
 *  has the APIs needed to round-trip it; visual-only UI affordances are
 *  marked `session`; preserved OOXML without authoring APIs is `read-only`. */
export function summarizeSpreadsheetCompatibility(
  wb: WorkbookHandle,
): SpreadsheetCompatibilitySummary {
  const c = wb.capabilities;
  const passthroughs = summarizePassthroughs(wb);
  const tables = summarizeTables(wb);
  const pivots = summarizePivotTables(wb);
  const chartDrawingCount =
    (passthroughs.byCategory.charts ?? 0) +
    (passthroughs.byCategory.drawings ?? 0) +
    (passthroughs.byCategory.media ?? 0);

  const items: SpreadsheetCompatibilityItem[] = [
    {
      id: 'cell-formatting',
      label: 'Cell formatting',
      status: c.cellFormatting ? 'writable' : 'session',
      reason: c.cellFormatting
        ? 'XF/font/fill/border/number-format APIs are available.'
        : 'Formatting can be shown in the UI but cannot be written through this engine.',
    },
    {
      id: 'conditional-formatting',
      label: 'Conditional formatting',
      status: c.conditionalFormatMutate
        ? 'writable'
        : c.conditionalFormat
          ? 'read-only'
          : 'session',
      reason: c.conditionalFormatMutate
        ? 'Conditional-format rules can be enumerated, added, removed, and cleared.'
        : c.conditionalFormat
          ? 'Rules can be evaluated but visual authoring/writeback is limited.'
          : 'Session rules can be painted, but the engine has no CF surface.',
    },
    {
      id: 'data-validation',
      label: 'Data validation',
      status: c.dataValidation ? 'writable' : 'session',
      reason: c.dataValidation
        ? 'Validation ranges can be read and written.'
        : 'Validation dropdown UI can be hosted, but engine writeback is unavailable.',
    },
    {
      id: 'hyperlinks',
      label: 'Hyperlinks',
      status: c.hyperlinks ? 'writable' : 'session',
      reason: c.hyperlinks
        ? 'Hyperlinks can be enumerated, added, and cleared.'
        : 'Hyperlink UI can be hosted for the session, but engine writeback is unavailable.',
    },
    {
      id: 'comments',
      label: 'Comments',
      status: c.comments ? 'writable' : 'session',
      reason: c.comments
        ? 'Cell comments can be read and written.'
        : 'Comment UI can be hosted for the session, but engine writeback is unavailable.',
    },
    {
      id: 'defined-names',
      label: 'Defined names',
      status: c.definedNameMutate ? 'writable' : 'read-only',
      reason: c.definedNameMutate
        ? 'Defined names can be listed and updated.'
        : 'Defined names can be listed when present, but mutation is unavailable.',
    },
    {
      id: 'sheet-protection',
      label: 'Sheet protection',
      status: c.sheetProtectionRoundtrip ? 'writable' : 'session',
      reason: c.sheetProtectionRoundtrip
        ? 'Sheet protection metadata can round-trip through the engine.'
        : 'Protection UI state can be hosted for the session, but engine writeback is unavailable.',
    },
    {
      id: 'sheet-views',
      label: 'Sheet views',
      status: 'session',
      reason:
        c.freeze || c.sheetZoom || c.hiddenRowsCols || c.outlines
          ? 'Sheet views can be captured and restored in the UI; individual view settings may still persist through engine view APIs.'
          : 'Sheet views can be captured and restored in the UI for the current session.',
    },
    {
      id: 'loaded-tables',
      label: 'Loaded tables',
      status: tables.count > 0 ? 'read-only' : 'unsupported',
      count: tables.count,
      reason:
        tables.count > 0
          ? 'Loaded ListObjects are visible and preserved, but not authorable.'
          : 'No loaded ListObjects were reported by the engine.',
    },
    {
      id: 'format-as-table',
      label: 'Format as Table',
      status: 'session',
      reason: 'The UI can create session table overlays; engine ListObject authoring is absent.',
    },
    {
      id: 'pivot-layouts',
      label: 'PivotTable layouts',
      status: pivots.count > 0 ? 'read-only' : c.pivotTables ? 'read-only' : 'unsupported',
      count: pivots.count,
      reason: c.pivotTables
        ? 'Loaded PivotTables can be projected into grid cells.'
        : 'The engine has no PivotTable projection API.',
    },
    {
      id: 'pivot-authoring',
      label: 'PivotTable authoring',
      status: c.pivotTableMutate ? 'writable' : 'unsupported',
      reason: c.pivotTableMutate
        ? 'PivotCache and PivotTable mutation APIs are available.'
        : 'Creating or editing PivotTable definitions needs new engine APIs.',
    },
    {
      id: 'session-charts',
      label: 'Session chart previews',
      status: 'session',
      reason:
        'Column/line chart overlays can be created in the UI; persisted chart writeback is gated on engine chart APIs.',
    },
    {
      id: 'charts-drawings',
      label: 'Charts, drawings, and images',
      status: chartDrawingCount > 0 ? 'read-only' : 'unsupported',
      count: chartDrawingCount,
      reason:
        chartDrawingCount > 0
          ? 'OOXML parts are preserved and inventoried, but anchors/content are not exposed.'
          : 'No preserved chart/drawing/media parts were reported by the engine.',
    },
    {
      id: 'chart-authoring',
      label: 'Chart authoring',
      status: 'unsupported',
      reason: 'Chart creation/editing requires a chart model in the engine.',
    },
    {
      id: 'external-links',
      label: 'External links',
      status: c.externalLinks ? 'read-only' : 'unsupported',
      reason: c.externalLinks
        ? 'External link records can be inventoried.'
        : 'External link enumeration is unavailable.',
    },
  ];

  return { items, byStatus: byStatus(items), byId: byId(items) };
}
