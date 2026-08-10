import { distinctValues } from '../commands/filter.js';
import type { SpreadsheetStore, State, ValueFilterCriteria } from '../store/store.js';
import { parseRangeRef } from './range-resolver.js';
import type { Range } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

const escapeXml = (value: string): string =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const columnLabel = (col: number): string => {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
};

export const autoFilterRangeRef = (range: Range): string =>
  `${columnLabel(range.c0)}${range.r0 + 1}:${columnLabel(range.c1)}${range.r1 + 1}`;

const sameRange = (left: Range, right: Range): boolean =>
  left.sheet === right.sheet &&
  left.r0 === right.r0 &&
  left.r1 === right.r1 &&
  left.c0 === right.c0 &&
  left.c1 === right.c1;

const conditionOperator: Record<NonNullable<ValueFilterCriteria['condition']>['op'], string> = {
  equals: 'equal',
  notEquals: 'notEqual',
  contains: 'equal',
  notContains: 'notEqual',
  greaterThan: 'greaterThan',
  greaterThanOrEqual: 'greaterThanOrEqual',
  lessThan: 'lessThan',
  lessThanOrEqual: 'lessThanOrEqual',
};

/** Serialises the application-owned subset of AutoFilter: its range, value
 * checklist and one custom comparison per column. Color criteria cannot be
 * represented without a differential-style id, so the caller deliberately
 * leaves them out rather than emitting a misleading OOXML definition. */
export function autoFilterXmlFromState(state: State, sheet: number): string {
  const range = state.ui.filterRange;
  if (!range || range.sheet !== sheet) return '';
  const criteria = state.ui.filterCriteria
    .filter((entry) => sameRange(entry.range, range))
    .sort((a, b) => a.byCol - b.byCol);
  if (criteria.length === 0) return `<autoFilter ref="${autoFilterRangeRef(range)}"/>`;

  const columns = criteria.flatMap((entry) => {
    const colId = entry.byCol - range.c0;
    if (colId < 0 || colId > range.c1 - range.c0 || entry.color) return [];
    if (entry.condition) {
      const wildcard = entry.condition.op === 'contains' || entry.condition.op === 'notContains';
      const value = wildcard ? `*${entry.condition.value}*` : entry.condition.value;
      return [
        `<filterColumn colId="${colId}"><customFilters><customFilter operator="${conditionOperator[entry.condition.op]}" val="${escapeXml(value)}"/></customFilters></filterColumn>`,
      ];
    }
    const hidden = new Set(entry.hiddenValues);
    const visible = distinctValues(state, range, entry.byCol).filter((value) => !hidden.has(value));
    const includeBlank = visible.includes('');
    const filters = visible
      .filter((value) => value !== '')
      .map((value) => `<filter val="${escapeXml(value)}"/>`)
      .join('');
    const blankAttr = includeBlank ? ' blank="1"' : '';
    return [
      `<filterColumn colId="${colId}"><filters${blankAttr}>${filters}</filters></filterColumn>`,
    ];
  });
  return `<autoFilter ref="${autoFilterRangeRef(range)}">${columns.join('')}</autoFilter>`;
}

/** Parses only the range, which is enough to restore header affordances while
 * leaving every imported criterion in the engine's verbatim XML until the
 * user changes the filter. */
export function autoFilterRangeFromXml(xml: string, sheet: number): Range | null {
  const ref = /<autoFilter\b[^>]*\bref="([^"]+)"/i.exec(xml)?.[1];
  const parsed = ref ? parseRangeRef(ref) : null;
  return parsed ? { sheet, r0: parsed.r0, c0: parsed.c0, r1: parsed.r1, c1: parsed.c1 } : null;
}

/** Writes the current sheet's user-owned AutoFilter state through the optional
 * engine seam. This is intentionally invoked only after a UI filter mutation;
 * untouched imported XML is preserved verbatim by the engine. */
export function syncAutoFilterToEngine(wb: WorkbookHandle, state: State, sheet: number): boolean {
  if (!wb.capabilities.autoFilter) return false;
  return wb.setSheetAutoFilterXml(sheet, autoFilterXmlFromState(state, sheet));
}

export function hydrateAutoFilterFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet: number,
): void {
  if (!wb.capabilities.autoFilter || typeof wb.getSheetAutoFilterXml !== 'function') return;
  const xml = wb.getSheetAutoFilterXml(sheet);
  if (xml === null) return;
  const range = autoFilterRangeFromXml(xml, sheet);
  wb.withAutoFilterSyncMuted(() => {
    store.setState((state) => ({
      ...state,
      ui: {
        ...state.ui,
        filterRange: range,
        // Criteria remain engine-owned until a UI interaction replaces them.
        filterCriteria: [],
      },
    }));
  });
}
