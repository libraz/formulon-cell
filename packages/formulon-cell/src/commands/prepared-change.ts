import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { CellPatch, CellSnapshot } from '../engine/workbook-handle.js';
import type { CellFormat } from '../store/types.js';
import type { CoercedInput } from './coerce-input.js';

export type PreparedChange = {
  readonly patch: CellPatch;
  readonly implicitFormat?: CellFormat['numFmt'];
  readonly input?: CoercedInput;
  /** Font axes inherited from a worksheet-wide override when a newly written
   *  cell would otherwise be flushed with the native default XF. */
  readonly inheritedFormat?: Partial<CellFormat>;
  /** Raw row/column XF used by a loaded workbook's worksheet font. Pinning
   *  this XF before a format sync preserves its font metadata (theme link,
   *  charset, colour and apply flags) instead of rebuilding it from paint
   *  values. */
  readonly inheritedAxisXf?: {
    readonly xfIndex: number;
    readonly projected: CellFormat;
  };
};

export const snapshotToPatch = (snapshot: CellSnapshot): CellPatch => ({
  addr: snapshot.addr,
  value: snapshot.value,
  formula: snapshot.formula,
});

const allKeys = (record: Record<string, unknown>): string[] => Object.keys(record).sort();

const definedKeys = (record: Record<string, unknown>): string[] =>
  Object.keys(record)
    .filter((key) => record[key] !== undefined)
    .sort();

const structurallyEqual = (
  left: unknown,
  right: unknown,
  keysOf: (record: Record<string, unknown>) => string[],
): boolean => {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object')
    return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((value, index) => structurallyEqual(value, right[index], keysOf));
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = keysOf(leftRecord);
  const rightKeys = keysOf(rightRecord);
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every(
    (key, index) =>
      key === rightKeys[index] && structurallyEqual(leftRecord[key], rightRecord[key], keysOf),
  );
};

/** Structural equality: key order is ignored, array order is not, and an own key holding
 *  `undefined` differs from a missing key. */
export const sameStructuredValue = (left: unknown, right: unknown): boolean =>
  structurallyEqual(left, right, allKeys);

/** Structural equality with the key semantics of comparing JSON forms: a key holding
 *  `undefined` equals a missing key. */
export const sameJsonValue = (left: unknown, right: unknown): boolean =>
  structurallyEqual(left, right, definedKeys);

const mergeInheritedFormat = (
  current: CellFormat | undefined,
  inherited: Partial<CellFormat> | undefined,
): CellFormat | undefined => {
  if (!inherited) return current;
  // Inherited axes are an implicit base. A direct format already present in
  // the map remains authoritative even if a caller supplied both fields in
  // the prepared change.
  return { ...inherited, ...(current ?? {}) };
};

const projectPreparedFormat = (
  current: CellFormat | undefined,
  change: PreparedChange,
): CellFormat | undefined => {
  let next = current;
  const needsStoredFormat =
    current !== undefined ||
    change.implicitFormat !== undefined ||
    (change.inheritedAxisXf === undefined && change.inheritedFormat !== undefined);
  if (needsStoredFormat) {
    const inheritedBase = {
      ...(change.inheritedAxisXf?.projected ?? {}),
      ...(change.inheritedFormat ?? {}),
    };
    next = mergeInheritedFormat(next, inheritedBase);
  }
  if (change.implicitFormat && (next?.numFmt === undefined || next.numFmt.kind === 'general')) {
    next = { ...(next ?? {}), numFmt: change.implicitFormat };
  }
  return next;
};

export const projectPreparedFormats = (
  before: ReadonlyMap<string, CellFormat>,
  prepared: readonly PreparedChange[],
): Map<string, CellFormat> => {
  const next = new Map(before);
  for (const change of prepared) {
    const key = addrKey(change.patch.addr);
    const projected = projectPreparedFormat(next.get(key), change);
    if (projected === undefined) next.delete(key);
    else next.set(key, projected);
  }
  return next;
};

export const changedFormatAddresses = (
  before: ReadonlyMap<string, CellFormat>,
  after: ReadonlyMap<string, CellFormat>,
  prepared: readonly PreparedChange[],
): readonly Addr[] => {
  const changed: Addr[] = [];
  const seen = new Set<string>();
  for (const change of prepared) {
    const key = addrKey(change.patch.addr);
    if (seen.has(key)) continue;
    seen.add(key);
    if (!sameStructuredValue(before.get(key), after.get(key))) changed.push(change.patch.addr);
  }
  return changed;
};

export const unionAddresses = (...groups: readonly (readonly Addr[])[]): readonly Addr[] => {
  const result: Addr[] = [];
  const seen = new Set<string>();
  for (const group of groups) {
    for (const addr of group) {
      const key = addrKey(addr);
      if (seen.has(key)) continue;
      seen.add(key);
      result.push(addr);
    }
  }
  return result;
};
