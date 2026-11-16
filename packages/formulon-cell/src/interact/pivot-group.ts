import { PivotDatePeriod } from '@libraz/formulon';
import { PIVOT_CELL_KIND } from '../engine/pivot-cell-kind.js';
import { type Addr, PivotCalendar } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { dateGroupingOf } from '../engine/workbook-handle-pivot.js';
import { commitPivotChange, type PivotCommitContext } from './pivot-field-settings-dialog-apply.js';

/** The By list of a date field, finest first. */
export const PIVOT_GROUP_PERIODS = [
  PivotDatePeriod.Seconds,
  PivotDatePeriod.Minutes,
  PivotDatePeriod.Hours,
  PivotDatePeriod.Days,
  PivotDatePeriod.Months,
  PivotDatePeriod.Quarters,
  PivotDatePeriod.Years,
] as const;

/** The field a report label groups: the label's own field, or the field its group field groups. */
export interface PivotGroupTarget {
  sheet: number;
  pivotIndex: number;
  /** Pivot field index of the field to group. */
  fieldIdx: number;
  /** Dates only, numbers only, or anything else (text, blanks, a mix). */
  values: 'date' | 'number' | 'other';
}

/** The field the row or column item label at `addr` groups; null for any other cell. */
export function pivotGroupTargetAt(wb: WorkbookHandle, addr: Addr): PivotGroupTarget | null {
  const info = wb.pivotCellInfo(addr);
  if (!info || info.fieldIndex < 0) return null;
  if (info.kind !== PIVOT_CELL_KIND.RowLabel && info.kind !== PIVOT_CELL_KIND.ColLabel) {
    return null;
  }
  const definition = wb.getPivotDefinition(addr.sheet, info.pivotIndex);
  const label = definition?.fields.find((f) => f.cacheFieldIndex === info.fieldIndex);
  if (!definition || !label) return null;
  const field = definition.fields[label.groupBaseIndex ?? label.index] ?? label;
  const cache = wb.pivotCacheFields(definition.cacheId)[field.cacheFieldIndex];
  const values =
    !cache || cache.containsBlank || cache.containsString
      ? 'other'
      : cache.containsDate && !cache.containsNonDate
        ? 'date'
        : cache.containsNumber && !cache.containsDate
          ? 'number'
          : 'other';
  return { sheet: addr.sheet, pivotIndex: info.pivotIndex, fieldIdx: field.index, values };
}

/** What the Grouping dialog shows and OK writes. Bounds are serials for a date field; a ticked
 *  bound follows the data and its value is only shown. */
export type PivotGroupSettings =
  | (PivotGroupBounds & { kind: 'date'; periods: number; intervalDays: number })
  | (PivotGroupBounds & { kind: 'number'; interval: number });

export interface PivotGroupBounds {
  autoStart: boolean;
  autoEnd: boolean;
  start: number;
  end: number;
}

// A span equal to 3 × 10^k takes 10^k; the slack keeps that edge when the span carries float error.
const SPAN_EDGE_SLACK = 1e-9;

/** The interval the desktop proposes for a span of values: 10^floor(log10(span / 3)), the edge
 *  `span === 3 × 10^k` included. */
export function pivotAutoGroupInterval(span: number): number {
  if (!(span > 0) || !Number.isFinite(span)) return 1;
  let k = Math.floor(Math.log10(span / 3));
  if (span >= 3 * 10 ** (k + 1) * (1 - SPAN_EDGE_SLACK)) k += 1;
  else if (span < 3 * 10 ** k * (1 - SPAN_EDGE_SLACK)) k -= 1;
  return Number((10 ** k).toPrecision(1));
}

/** The field's current grouping, or the defaults for an ungrouped field: a date field by Months
 *  from its first date to the midnight after its last, a number field from its smallest value to
 *  its largest by the proposed interval. Null when the target has nothing to group. */
export function pivotGroupSettings(
  wb: WorkbookHandle,
  target: PivotGroupTarget,
): PivotGroupSettings | null {
  if (target.values === 'other') return null;
  const definition = wb.getPivotDefinition(target.sheet, target.pivotIndex);
  const field = definition?.fields[target.fieldIdx];
  if (!definition || !field) return null;
  const values = wb
    .pivotCacheSharedItems(definition.cacheId, field.cacheFieldIndex)
    .flatMap((value) => (value.kind === 'number' ? [value.value] : []));
  if (values.length === 0) return null;
  const min = Math.min(...values);
  const max = Math.max(...values);
  if (target.values === 'date') {
    const dataEnd = Number.isInteger(max) ? max + 1 : max;
    const group = field.dateGroup;
    return group
      ? {
          kind: 'date',
          periods: group.periods,
          intervalDays: group.intervalDays,
          autoStart: group.startSerial < 0,
          autoEnd: group.endSerial < 0,
          start: group.startSerial < 0 ? min : group.startSerial,
          end: group.endSerial < 0 ? dataEnd : group.endSerial,
        }
      : {
          kind: 'date',
          periods: PivotDatePeriod.Months,
          intervalDays: 1,
          autoStart: true,
          autoEnd: true,
          start: min,
          end: dataEnd,
        };
  }
  const group = field.numericGroup;
  return group
    ? {
        kind: 'number',
        interval: group.interval,
        autoStart: group.start === null,
        autoEnd: group.end === null,
        start: group.start ?? min,
        end: group.end ?? max,
      }
    : {
        kind: 'number',
        interval: pivotAutoGroupInterval(max - min),
        autoStart: true,
        autoEnd: true,
        start: min,
        end: max,
      };
}

/** Groups the target's field as the Grouping dialog's OK does; one undo step. */
export function groupPivotField(
  ctx: PivotCommitContext,
  target: PivotGroupTarget,
  settings: PivotGroupSettings,
): boolean {
  const { wb } = ctx;
  const { sheet, pivotIndex, fieldIdx } = target;
  let write: () => boolean;
  if (settings.kind === 'date') {
    const granularity = dateGroupingOf(settings.periods);
    if (granularity === null) return false;
    const calendar =
      wb.getPivotDefinition(sheet, pivotIndex)?.fields[fieldIdx]?.dateGroup?.calendar ??
      PivotCalendar.Gregorian;
    write = () =>
      wb.setPivotFieldDateGroup(sheet, pivotIndex, fieldIdx, granularity, calendar, {
        periods: settings.periods,
        ...(settings.periods === PivotDatePeriod.Days
          ? { intervalDays: settings.intervalDays }
          : {}),
        ...(settings.autoStart ? {} : { startSerial: settings.start }),
        ...(settings.autoEnd ? {} : { endSerial: settings.end }),
      });
  } else {
    write = () =>
      wb.setPivotFieldNumericGroup(sheet, pivotIndex, fieldIdx, {
        interval: settings.interval,
        start: settings.autoStart ? null : settings.start,
        end: settings.autoEnd ? null : settings.end,
      });
  }
  return commitPivotChange(ctx, sheet, pivotIndex, write);
}
