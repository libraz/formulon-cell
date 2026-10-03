import {
  type SpreadsheetCompatibilityId,
  type SpreadsheetCompatibilityStatus,
  summarizeSpreadsheetCompatibility,
} from '../engine/compatibility.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { Strings } from '../i18n/strings.js';

type WorkbookObjectsStrings = Strings['workbookObjects'];

export interface SpreadsheetCompatibilityReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

const COMPATIBILITY_LABEL_KEYS: Record<
  SpreadsheetCompatibilityId,
  keyof WorkbookObjectsStrings['compatibilityLabels']
> = {
  'cell-formatting': 'cellFormatting',
  'conditional-formatting': 'conditionalFormatting',
  'data-validation': 'dataValidation',
  hyperlinks: 'hyperlinks',
  comments: 'comments',
  phonetic: 'phonetic',
  'defined-names': 'definedNames',
  'named-cell-styles': 'namedCellStyles',
  'auto-filter': 'autoFilter',
  'sheet-protection': 'sheetProtection',
  'sheet-views': 'sheetViews',
  'page-setup': 'pageSetup',
  'loaded-tables': 'loadedTables',
  'format-as-table': 'formatAsTable',
  'pivot-layouts': 'pivotLayouts',
  'pivot-authoring': 'pivotAuthoring',
  'session-charts': 'sessionCharts',
  'charts-drawings': 'chartsDrawings',
  'chart-authoring': 'chartAuthoring',
  'external-links': 'externalLinks',
};

const STATUS_LABEL_KEYS: Record<
  SpreadsheetCompatibilityStatus,
  keyof Pick<WorkbookObjectsStrings, 'writable' | 'readOnly' | 'sessionOnly' | 'unsupported'>
> = {
  writable: 'writable',
  'read-only': 'readOnly',
  session: 'sessionOnly',
  unsupported: 'unsupported',
};

export const spreadsheetCompatibilityLabel = (
  id: SpreadsheetCompatibilityId,
  strings: WorkbookObjectsStrings,
): string => strings.compatibilityLabels[COMPATIBILITY_LABEL_KEYS[id]];

export const spreadsheetCompatibilityDetail = (
  id: SpreadsheetCompatibilityId,
  strings: WorkbookObjectsStrings,
): string => strings.compatibilityDetails[COMPATIBILITY_LABEL_KEYS[id]];

export const spreadsheetCompatibilityStatusLabel = (
  status: SpreadsheetCompatibilityStatus,
  strings: WorkbookObjectsStrings,
): string => strings[STATUS_LABEL_KEYS[status]];

/** Build the flat severity/label/detail list rendered by the React, Vue, and
 *  playground "inspect workbook" backstage actions. Three call sites used to
 *  carry ~100 lines of switch-case duplication each; this helper is the single
 *  source of truth for that mapping. */
export const buildSpreadsheetCompatibilityReport = (
  wb: WorkbookHandle,
  strings: WorkbookObjectsStrings,
): SpreadsheetCompatibilityReportItem[] => {
  const summary = summarizeSpreadsheetCompatibility(wb);
  const banner: SpreadsheetCompatibilityReportItem = {
    severity: 'info',
    label: strings.compatibility,
    detail:
      `${strings.writable} ${summary.byStatus.writable}, ` +
      `${strings.readOnly} ${summary.byStatus['read-only']}, ` +
      `${strings.sessionOnly} ${summary.byStatus.session}, ` +
      `${strings.unsupported} ${summary.byStatus.unsupported}`,
  };
  const rows = summary.items.map<SpreadsheetCompatibilityReportItem>((entry) => {
    const detail = spreadsheetCompatibilityDetail(entry.id, strings);
    return {
      severity: entry.status === 'unsupported' || entry.status === 'read-only' ? 'warning' : 'info',
      label: `${spreadsheetCompatibilityLabel(entry.id, strings)} · ${spreadsheetCompatibilityStatusLabel(entry.status, strings)}`,
      detail: entry.count === undefined ? detail : `${detail} (${entry.count})`,
    };
  });
  return [banner, ...rows];
};

export const compatibilityLabelKey = (
  id: SpreadsheetCompatibilityId,
): keyof WorkbookObjectsStrings['compatibilityLabels'] => COMPATIBILITY_LABEL_KEYS[id];
