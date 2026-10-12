import { addrKey, MAX_COL, MAX_ROW, parseAddrKey } from './address.js';
import type {
  Addr,
  SheetAxisDirectXf,
  SheetAxisFormatOverride,
  SheetFontAxisSnapshot,
  SheetFontOverride,
} from './types.js';

/** The one dimensional coordinate change made by a row/column structure edit. */
export interface AxisStructureShift {
  readonly axis: 'row' | 'col';
  readonly split: number;
  /** Positive for insertion and negative for deletion. */
  readonly delta: number;
}

/** Map an axis coordinate without enumerating the worksheet. Coordinates in a
 * deleted band are removed; coordinates after the band move by `delta`. */
export const mapAxisIndex = (index: number, shift: AxisStructureShift): number | null => {
  if (index < shift.split) return index;
  if (shift.delta < 0 && index < shift.split - shift.delta) return null;
  const max = shift.axis === 'row' ? MAX_ROW : MAX_COL;
  const mapped = index + shift.delta;
  return mapped < 0 || mapped > max ? null : mapped;
};

/** Map a cell address through a structure edit. Other sheet coordinates are
 * deliberately left untouched. */
export const mapAxisAddress = (addr: Addr, shift: AxisStructureShift): Addr | null => {
  const index = shift.axis === 'row' ? addr.row : addr.col;
  const mapped = mapAxisIndex(index, shift);
  if (mapped === null) return null;
  return shift.axis === 'row' ? { ...addr, row: mapped } : { ...addr, col: mapped };
};

const mapInterval = (
  first: number,
  last: number,
  shift: AxisStructureShift,
): [number, number] | null => {
  const max = shift.axis === 'row' ? MAX_ROW : MAX_COL;
  if (shift.delta > 0) {
    const mappedFirst = first >= shift.split ? first + shift.delta : first;
    const mappedLast = last >= shift.split ? last + shift.delta : last;
    if (mappedFirst > max || mappedFirst > mappedLast) return null;
    return [Math.max(0, mappedFirst), Math.min(max, mappedLast)];
  }
  const count = -shift.delta;
  const bandLast = shift.split + count - 1;
  const mappedFirst = first < shift.split ? first : first > bandLast ? first - count : shift.split;
  const mappedLast = last < shift.split ? last : last > bandLast ? last - count : shift.split - 1;
  if (mappedFirst > mappedLast || mappedFirst > max || mappedLast < 0) return null;
  return [Math.max(0, mappedFirst), Math.min(max, mappedLast)];
};

const mergeColumns = (
  entries: readonly { first: number; last: number; styleXf: number }[],
): SheetFontAxisSnapshot['columns'] => {
  const sorted = [...entries]
    .filter((entry) => entry.first <= entry.last)
    .sort((a, b) => a.first - b.first || a.last - b.last);
  const out: { first: number; last: number; styleXf: number }[] = [];
  for (const entry of sorted) {
    const previous = out.at(-1);
    if (previous && previous.last + 1 >= entry.first && previous.styleXf === entry.styleXf) {
      previous.last = Math.max(previous.last, entry.last);
    } else out.push({ ...entry });
  }
  return out;
};

/** Remap a compact native axis snapshot. Explicit records are preserved even
 * when their XF index is zero. A new gap stays implicit unless a source span
 * crosses it, in which case that span widens naturally. Column spans are
 * clipped to the worksheet boundary. */
export const remapAxisSnapshot = (
  snapshot: SheetFontAxisSnapshot,
  shift: AxisStructureShift,
): SheetFontAxisSnapshot => {
  if (shift.axis === 'row') {
    return {
      columns: snapshot.columns.map((entry) => ({ ...entry })),
      rows: snapshot.rows.flatMap((entry) => {
        const row = mapAxisIndex(entry.row, shift);
        return row === null ? [] : [{ row, styleXf: entry.styleXf }];
      }),
    };
  }
  return {
    columns: mergeColumns(
      snapshot.columns.flatMap((entry) => {
        const mapped = mapInterval(entry.first, entry.last, shift);
        return mapped ? [{ first: mapped[0], last: mapped[1], styleXf: entry.styleXf }] : [];
      }),
    ),
    rows: snapshot.rows.map((entry) => ({ ...entry })),
  };
};

const remapDirectXfs = (
  entries: readonly SheetAxisDirectXf[] | undefined,
  shift: AxisStructureShift,
): readonly SheetAxisDirectXf[] | undefined => {
  if (!entries) return undefined;
  const mapped = new Map<string, SheetAxisDirectXf>();
  for (const entry of entries) {
    const addr = parseAddrKey(entry.key);
    if (!addr) {
      mapped.set(entry.key, entry);
      continue;
    }
    const next = mapAxisAddress(addr, shift);
    if (!next) continue;
    mapped.set(addrKey(next), { ...entry, key: addrKey(next) });
  }
  return [...mapped.values()];
};

/** Remap a row/column axis override, retaining its immutable before XF data. */
export const remapAxisFormatOverride = (
  override: SheetAxisFormatOverride,
  shift: AxisStructureShift,
): SheetAxisFormatOverride => ({
  ...override,
  baseline: remapAxisSnapshot(override.baseline, shift),
  desired: remapAxisSnapshot(override.desired, shift),
  ...(override.directXfs ? { directXfs: remapDirectXfs(override.directXfs, shift) } : {}),
});

/** Remap a legacy whole-sheet font baseline. */
export const remapSheetFontOverride = (
  override: SheetFontOverride,
  shift: AxisStructureShift,
): SheetFontOverride => ({
  ...override,
  ...(override.baseline ? { baseline: remapAxisSnapshot(override.baseline, shift) } : {}),
});

/** Return the structural address keys that must be included in a native
 * checkpoint. The output contains both pre-edit keys and their mapped keys;
 * deleted-band keys remain in the pre-edit half for Undo restoration. */
export const structuralKeySet = (
  keys: Iterable<string>,
  sheet: number,
  shift: AxisStructureShift,
): Set<string> => {
  const out = new Set<string>();
  for (const key of keys) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    out.add(key);
    const mapped = mapAxisAddress(addr, shift);
    if (mapped) out.add(addrKey(mapped));
  }
  return out;
};
