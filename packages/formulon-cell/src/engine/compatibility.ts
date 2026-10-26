import {
  summarizePassthroughs,
  summarizePivotTables,
  summarizeTables,
} from './passthrough-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

export type SpreadsheetCompatibilityStatus = 'writable' | 'read-only' | 'session' | 'unsupported';

export type SpreadsheetCompatibilityId =
  | 'cell-formatting'
  | 'font-effects'
  | 'gradient-fills'
  | 'xlsb-save'
  | 'conditional-formatting'
  | 'data-validation'
  | 'hyperlinks'
  | 'comments'
  | 'threaded-comments'
  | 'images'
  | 'shapes'
  | 'sparklines'
  | 'slicers'
  | 'workbook-theme'
  | 'phonetic'
  | 'defined-names'
  | 'named-cell-styles'
  | 'auto-filter'
  | 'sheet-protection'
  | 'workbook-protection'
  | 'protected-ranges'
  | 'sheet-views'
  | 'show-formulas'
  | 'outline-symbols'
  | 'outline-collapse'
  | 'split-panes'
  | 'hidden-default-rows'
  | 'page-setup'
  | 'loaded-tables'
  | 'format-as-table'
  | 'pivot-layouts'
  | 'pivot-authoring'
  | 'charts-drawings'
  | 'chart-authoring'
  | 'external-links'
  | 'calc-settings'
  | 'data-tables';

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
      id: 'threaded-comments',
      label: 'Threaded comments',
      status: c.threadedComments ? 'writable' : 'unsupported',
      reason: c.threadedComments
        ? 'Threaded comments, replies, people and resolution can be read and written.'
        : 'Threaded comment editing requires the native comment and person APIs.',
    },
    {
      id: 'images',
      label: 'Images',
      status: c.images ? 'writable' : 'session',
      reason: c.images
        ? 'Picture bytes and anchors can be loaded, inserted and removed.'
        : 'Pictures are available as session overlays.',
    },
    {
      id: 'shapes',
      label: 'Shapes',
      status: c.shapes ? 'writable' : 'session',
      reason: c.shapes
        ? 'Rectangles, rounded rectangles, ovals, triangles, diamonds and lines are saved with the workbook; arrows, dash styles and partial transparency cannot be saved.'
        : 'Shapes are available as session overlays.',
    },
    {
      id: 'sparklines',
      label: 'Sparklines',
      status: c.sparklines ? 'writable' : 'session',
      reason: c.sparklines
        ? 'Sparkline groups are loaded, edited and saved with the workbook.'
        : 'Sparklines are available as session overlays.',
    },
    {
      id: 'slicers',
      label: 'Table slicers',
      status: c.slicers ? 'writable' : 'session',
      reason: c.slicers
        ? 'Table slicers are loaded, edited and saved with the workbook; the selection lives in the table filter.'
        : 'Slicers are available as session overlays.',
    },
    {
      id: 'workbook-theme',
      label: 'Workbook theme',
      status: c.workbookTheme ? 'writable' : 'unsupported',
      reason: c.workbookTheme
        ? 'The workbook colour scheme and major/minor typefaces can be read and written.'
        : 'Workbook theme editing requires the native theme APIs.',
    },
    {
      id: 'cell-formatting',
      label: 'Cell formatting',
      status: c.cellFormatting ? 'writable' : 'session',
      reason: c.cellFormatting
        ? 'XF/font/fill/border/number-format APIs are available.'
        : 'Formatting can be shown in the UI but cannot be written through this engine.',
    },
    {
      id: 'font-effects',
      label: 'Font effects',
      status: c.cellFormatting ? 'writable' : 'session',
      reason: c.cellFormatting
        ? 'Outline, shadow, condense and extend font flags are read and written under every profile. Outline and Shadow are set in the Mac Format Cells dialog and drawn as hollow and shadowed text only under the Mac profile; condense and extend are kept but not drawn.'
        : 'Font effects are shown for the session but cannot be written through this engine.',
    },
    {
      id: 'gradient-fills',
      label: 'Gradient fills',
      status: c.cellFormatting ? 'writable' : 'session',
      reason: c.cellFormatting
        ? 'Linear and path gradient fills are read, written and set through Fill Effects. Path gradients are painted as radial approximations.'
        : 'Gradient fills are shown for the session but cannot be written through this engine.',
    },
    {
      id: 'xlsb-save',
      label: 'Binary workbook (.xlsb)',
      status: c.xlsbSave ? 'writable' : 'unsupported',
      reason: c.xlsbSave
        ? 'Workbooks can be saved as .xlsb. Notes, threaded comments and filters are not written to .xlsb. Pictures, shapes and charts, PivotTables and slicers read from .xlsx, page setup (margins, header and footer, page breaks), protected ranges, some sheet views and workbook protection settings, custom table styles, and tables or sparklines whose references cannot be encoded may also be left out. The engine reports one total for these, so saving warns with the categories present before the file is delivered.'
        : 'Saving as .xlsb requires the engine saveWithDiagnostics API.',
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
        ? 'Validation ranges, messages, and in-cell dropdown visibility can be read and written.'
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
        ? `${
            c.commentsEnumerable
              ? 'Cell comments can be enumerated, read, and written.'
              : 'Cell comments can be read and written for known cells; blank-cell comment hydration needs sheet-wide enumeration.'
          }${
            c.commentVisibility
              ? ' Shown notes stay shown; a note moved by sort, insert/delete, cut or paste takes the default box size and position.'
              : ' Note visibility is session-only.'
          }`
        : 'Comment UI can be hosted for the session, but engine writeback is unavailable.',
    },
    {
      id: 'phonetic',
      label: 'Phonetic guides',
      status: c.phonetic ? 'writable' : 'unsupported',
      reason: c.phonetic
        ? c.phoneticRuns
          ? 'Cell-level OOXML rPh guides round-trip span by span, so a partially annotated cell keeps its spans.'
          : 'Guides can be read and written, but only as one reading for the whole cell — editing a partially annotated cell collapses its spans.'
        : 'Reading or writing OOXML rPh/phoneticPr records requires engine APIs.',
    },
    {
      id: 'defined-names',
      label: 'Defined names',
      status: c.definedNameMutate ? 'writable' : 'read-only',
      reason: c.definedNameMutate
        ? c.definedNameScopes
          ? 'Workbook- and sheet-scoped defined names can be listed, updated, and round-tripped.'
          : 'Workbook-scoped defined names can be listed and updated; sheet-scoped names need engine scope support.'
        : 'Defined names can be listed when present, but mutation is unavailable.',
    },
    {
      id: 'named-cell-styles',
      label: 'Named cell styles',
      status: c.cellStyleMutate ? 'writable' : c.cellStyles ? 'read-only' : 'unsupported',
      reason: c.cellStyleMutate
        ? 'A style applied from the gallery is written as a named cellStyle/cellStyleXf pair, so editing the style reaches every cell that uses it.'
        : c.cellStyles
          ? 'Named cell styles can be inspected, but creating or updating cellStyle/cellStyleXf records needs engine writer APIs.'
          : 'The engine cannot enumerate named cell styles.',
    },
    {
      id: 'auto-filter',
      label: 'AutoFilter',
      status: c.autoFilter ? 'writable' : 'unsupported',
      reason: c.autoFilter
        ? 'Worksheet AutoFilter definitions can be preserved and written from the filter UI.'
        : 'Reading or writing worksheet autoFilter definitions requires engine APIs.',
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
      id: 'workbook-protection',
      label: 'Workbook protection',
      status: c.workbookProtectionRoundtrip ? 'writable' : 'session',
      reason: c.workbookProtectionRoundtrip
        ? 'Workbook structure protection and its password hash can round-trip through the engine.'
        : 'Workbook structure protection can be hosted for the session, but engine writeback is unavailable.',
    },
    {
      id: 'protected-ranges',
      label: 'Protected ranges',
      status: c.protectedRanges ? 'writable' : 'session',
      reason: c.protectedRanges
        ? 'Allowed edit ranges, their password hashes and security descriptors can round-trip through the engine.'
        : 'Allowed edit ranges can be hosted for the session, but engine writeback is unavailable.',
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
      id: 'show-formulas',
      label: 'Show Formulas',
      status: c.sheetShowFormulas ? 'writable' : 'session',
      reason: c.sheetShowFormulas
        ? 'Show Formulas is saved per sheet.'
        : 'Show Formulas applies for the session; the engine cannot save the sheet flag.',
    },
    {
      id: 'outline-symbols',
      label: 'Outline symbols',
      status: c.outlineSymbols ? 'writable' : 'session',
      reason: c.outlineSymbols
        ? 'Hiding the outline symbols is saved per sheet.'
        : 'Outline symbols can be hidden for the session; the engine cannot save the sheet flag.',
    },
    {
      id: 'outline-collapse',
      label: 'Collapsed outline groups',
      status: c.outlines && c.outlineCollapse ? 'writable' : 'session',
      reason:
        c.outlines && c.outlineCollapse
          ? 'Collapsed groups are saved with their summary row or column marked collapsed.'
          : 'Groups collapse for the session; the engine cannot save outline levels or the collapsed flag.',
    },
    {
      id: 'split-panes',
      label: 'Split panes',
      status: c.sheetSplit && c.sheetView ? 'writable' : 'session',
      reason:
        c.sheetSplit && c.sheetView
          ? 'Split panes are saved with the rows and columns each pane shows. The top and left panes scroll on their own but end where the other pane begins.'
          : 'Split panes apply for the session; the engine cannot save a split.',
    },
    {
      id: 'hidden-default-rows',
      label: 'Rows hidden by default',
      status: c.geometry && c.sheetFormatDefaults ? 'writable' : 'read-only',
      reason:
        c.geometry && c.sheetFormatDefaults
          ? 'Sheets that hide rows by default keep that setting. A row the file keeps visible only through an empty row record opens hidden.'
          : 'Rows hidden by default are shown hidden, but the setting cannot be changed.',
    },
    {
      id: 'page-setup',
      label: 'Page setup',
      status: c.printSettings ? 'writable' : 'session',
      reason: c.printSettings
        ? c.pageBreaks
          ? 'Orientation, paper, scaling, margins, print options, header/footer, print area, print titles and manual breaks round-trip through the engine. Printer minimum margins stay session state — they describe a device, not the workbook.'
          : 'Print settings round-trip through the engine; manual page breaks stay session state until the engine exposes the break table.'
        : 'The Page Setup dialog drives the print document for the session, but the engine has no print-settings surface to write it back.',
    },
    {
      id: 'loaded-tables',
      label: 'Loaded tables',
      status: c.tableMutate ? 'writable' : tables.count > 0 ? 'read-only' : 'unsupported',
      count: tables.count,
      reason: c.tableMutate
        ? 'Loaded ListObjects are visible, editable, and written back on save.'
        : tables.count > 0
          ? 'Loaded ListObjects are visible and preserved, but not authorable.'
          : 'No loaded ListObjects were reported by the engine.',
    },
    {
      id: 'format-as-table',
      label: 'Format as Table',
      status: c.tableMutate ? 'writable' : 'session',
      reason: c.tableMutate
        ? 'Format as Table creates a real ListObject alongside the UI overlay, so the saved workbook opens as a table.'
        : 'The UI can create session table overlays; engine ListObject authoring is absent.',
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
      id: 'charts-drawings',
      label: 'Charts, drawings, and images',
      status:
        c.charts && c.images ? 'writable' : chartDrawingCount > 0 ? 'read-only' : 'unsupported',
      count: chartDrawingCount,
      reason:
        c.charts && c.images
          ? 'Charts and pictures are anchored to cells, edited in the workbook and saved with it; other drawing and media parts are preserved.'
          : chartDrawingCount > 0
            ? 'OOXML parts are preserved and inventoried, but anchors/content are not exposed.'
            : 'No preserved chart/drawing/media parts were reported by the engine.',
    },
    {
      id: 'chart-authoring',
      label: 'Chart authoring',
      status: c.charts ? 'writable' : 'unsupported',
      reason: c.charts
        ? 'Column, bar, line, area, pie, doughnut and scatter charts are created and edited in the workbook; other chart types can be moved, resized and deleted.'
        : 'Chart creation/editing requires the engine chart APIs.',
    },
    {
      id: 'external-links',
      label: 'External links',
      status: c.externalLinks ? 'read-only' : 'unsupported',
      reason: c.externalLinks
        ? 'External link records can be inventoried.'
        : 'External link enumeration is unavailable.',
    },
    {
      id: 'calc-settings',
      label: 'Calculation options',
      status: c.calcMode && c.calcSettings ? 'writable' : 'unsupported',
      reason:
        c.calcMode && c.calcSettings
          ? 'Calculation mode, recalculation before saving, iteration, R1C1 reference style, precision as displayed and the 1904 date system are saved with the workbook. Precision as displayed rounds stored constants and new entries to their format; formula results keep full precision because the engine stores the option as metadata only.'
          : 'Calculation options cannot be saved without the engine calcPr APIs.',
    },
    {
      id: 'data-tables',
      label: 'Data tables',
      status: c.dataTables ? 'writable' : 'unsupported',
      reason: c.dataTables
        ? 'What-If data tables are created, computed and saved with the workbook; their result cells cannot be edited.'
        : 'Creating or computing data tables requires the engine data-table APIs.',
    },
  ];

  return { items, byStatus: byStatus(items), byId: byId(items) };
}
