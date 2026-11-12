import { type AutoFilterDateGroupInput, DateTimeGrouping } from '@libraz/formulon';
import { addrKey } from '../engine/address.js';
import type { Range } from '../engine/types.js';
import { excelSerialToDateParts } from '../format/date-serial.js';
import type { State } from '../store/types.js';
import { numFmtShowsDate } from './date-periods.js';

/** The checklist keys of a column's dated numbers, each with its serial. A number counts as a
 *  date when its cell shows a calendar date. */
export function columnDateSerials(state: State, range: Range, byCol: number): Map<string, number> {
  const dates = new Map<string, number>();
  for (let row = range.r0 + 1; row <= range.r1; row += 1) {
    const key = addrKey({ sheet: range.sheet, row, col: byCol });
    const value = state.data.cells.get(key)?.value;
    if (value?.kind !== 'number') continue;
    if (numFmtShowsDate(state.format.formats.get(key)?.numFmt)) {
      dates.set(String(value.value), value.value);
    }
  }
  return dates;
}

interface DayParts {
  year: number;
  month: number;
  day: number;
}

const partsOf = (serial: number, date1904: boolean): DayParts | null => {
  const parts = excelSerialToDateParts(serial, date1904);
  return parts ? { year: parts.year, month: parts.month, day: parts.day } : null;
};

/** The fewest date-group items that keep exactly the `visible` dates of `all`: a whole year when
 *  every date of that year is kept, else a whole month, else each kept day. */
export function dateGroupsForVisible(
  visible: Iterable<number>,
  all: Iterable<number>,
  date1904: boolean,
): AutoFilterDateGroupInput[] {
  const kept = new Set<string>();
  for (const serial of visible) {
    const p = partsOf(serial, date1904);
    if (p) kept.add(`${p.year}-${p.month}-${p.day}`);
  }
  // year -> month -> day -> kept
  const tree = new Map<number, Map<number, Map<number, boolean>>>();
  for (const serial of all) {
    const p = partsOf(serial, date1904);
    if (!p) continue;
    const months = tree.get(p.year) ?? new Map<number, Map<number, boolean>>();
    tree.set(p.year, months);
    const days = months.get(p.month) ?? new Map<number, boolean>();
    months.set(p.month, days);
    days.set(p.day, kept.has(`${p.year}-${p.month}-${p.day}`));
  }
  const groups: AutoFilterDateGroupInput[] = [];
  for (const [year, months] of [...tree].sort(([a], [b]) => a - b)) {
    const allDays = [...months.values()].flatMap((days) => [...days.values()]);
    if (allDays.every(Boolean)) {
      groups.push({ year, grouping: DateTimeGrouping.Year });
      continue;
    }
    for (const [month, days] of [...months].sort(([a], [b]) => a - b)) {
      if ([...days.values()].every(Boolean)) {
        groups.push({ year, month, grouping: DateTimeGrouping.Month });
        continue;
      }
      for (const [day, on] of [...days].sort(([a], [b]) => a - b)) {
        if (on) groups.push({ year, month, day, grouping: DateTimeGrouping.Day });
      }
    }
  }
  return groups;
}

/** Whether a date falls in any of `groups`; a group finer than a day matches its whole day. */
export function serialInDateGroups(
  serial: number,
  groups: readonly AutoFilterDateGroupInput[],
  date1904: boolean,
): boolean {
  const p = partsOf(serial, date1904);
  if (!p) return false;
  return groups.some((group) => {
    if (group.year !== p.year) return false;
    const grouping = group.grouping ?? DateTimeGrouping.Year;
    if (grouping === DateTimeGrouping.Year) return true;
    if (group.month !== p.month) return false;
    return grouping === DateTimeGrouping.Month || group.day === p.day;
  });
}

/** `groups` as `<dateGroupItem>` elements. */
export function dateGroupItemsXml(groups: readonly AutoFilterDateGroupInput[]): string {
  return groups
    .map((group) => {
      const grouping = group.grouping ?? DateTimeGrouping.Year;
      const name =
        grouping === DateTimeGrouping.Year
          ? 'year'
          : grouping === DateTimeGrouping.Month
            ? 'month'
            : 'day';
      const month = grouping === DateTimeGrouping.Year ? '' : ` month="${group.month}"`;
      const day = grouping === DateTimeGrouping.Day ? ` day="${group.day}"` : '';
      return `<dateGroupItem year="${group.year}"${month}${day} dateTimeGrouping="${name}"/>`;
    })
    .join('');
}
