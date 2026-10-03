import { addrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import type { CellPatch, CellSnapshot } from '../engine/workbook-handle.js';
import type { CellFormat } from '../store/types.js';
import type { CoercedInput } from './coerce-input.js';

export type PreparedChange = {
  readonly patch: CellPatch;
  readonly implicitFormat?: CellFormat['numFmt'];
  readonly input?: CoercedInput;
  /** A format staged by the editor/formula bar for this resolved anchor. */
  readonly pendingFormat?: Partial<CellFormat>;
};

export const snapshotToPatch = (snapshot: CellSnapshot): CellPatch => ({
  addr: snapshot.addr,
  value: snapshot.value,
  formula: snapshot.formula,
});

export const sameStructuredValue = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object')
    return false;
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) return false;
    return left.every((value, index) => sameStructuredValue(value, right[index]));
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const leftKeys = Object.keys(leftRecord).sort();
  const rightKeys = Object.keys(rightRecord).sort();
  if (leftKeys.length !== rightKeys.length) return false;
  return leftKeys.every(
    (key, index) =>
      key === rightKeys[index] && sameStructuredValue(leftRecord[key], rightRecord[key]),
  );
};

const mergeFormatPatch = (
  current: CellFormat | undefined,
  patch: Partial<CellFormat> | undefined,
): CellFormat | undefined => {
  if (!patch) return current;
  const next: CellFormat = { ...(current ?? {}), ...patch };
  if (patch.borders) next.borders = { ...(current?.borders ?? {}), ...patch.borders };
  return next;
};

const projectPreparedFormat = (
  current: CellFormat | undefined,
  change: PreparedChange,
): CellFormat | undefined => {
  let next = current;
  if (change.implicitFormat && (next?.numFmt === undefined || next.numFmt.kind === 'general')) {
    next = { ...(next ?? {}), numFmt: change.implicitFormat };
  }
  return mergeFormatPatch(next, change.pendingFormat);
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

/** Return only addresses whose matching pending patch changes the format after
 * implicit input coercion has been projected. Implicit formats alone retain
 * the value-edit authorization path used before pending-format integration. */
export const effectivePendingFormatAddresses = (
  before: ReadonlyMap<string, CellFormat>,
  prepared: readonly PreparedChange[],
): readonly Addr[] => {
  const implicitBase = projectPreparedFormats(
    before,
    prepared.map((change) => ({ ...change, pendingFormat: undefined })),
  );
  const pendingPrepared = prepared
    .filter((change) => change.pendingFormat !== undefined)
    .map((change) => ({ ...change, implicitFormat: undefined }));
  const pendingApplied = projectPreparedFormats(implicitBase, pendingPrepared);
  return changedFormatAddresses(implicitBase, pendingApplied, pendingPrepared);
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
