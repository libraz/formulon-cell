import { addrKey } from '../engine/address.js';
import { flushProtectionToEngine } from '../engine/protection-sync.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  type AllowedEditRange,
  type CellFormat,
  mutators,
  type SheetProtectionPasswordHash,
  type SheetProtectionPermissions,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import type { History } from './history.js';

export interface SheetProtectionOptions {
  password?: string;
  passwordHash?: SheetProtectionPasswordHash;
  permissions?: SheetProtectionPermissions;
  workbook?: WorkbookHandle;
}

export interface WorkbookStructureProtectionOptions {
  password?: string;
}

export interface AllowedEditRangeOptions {
  id?: string;
  title?: string;
  password?: string;
}

interface ProtectionSnapshot {
  protectedSheets: Map<
    number,
    {
      password?: string;
      passwordHash?: SheetProtectionPasswordHash;
      permissions?: SheetProtectionPermissions;
    }
  >;
  workbookStructure?: { password?: string };
  allowedEditRanges: AllowedEditRange[];
}

const rangeContainsAddr = (range: Range, addr: Addr): boolean =>
  range.sheet === addr.sheet &&
  addr.row >= range.r0 &&
  addr.row <= range.r1 &&
  addr.col >= range.c0 &&
  addr.col <= range.c1;

/** Whether `sheet` is currently flagged protected on the workbook. Mirrors
 *  the protection slice as a pure helper so call sites don't reach into the
 *  Map shape directly. */
export function isSheetProtected(state: State, sheet: number): boolean {
  return state.protection.protectedSheets.has(sheet);
}

export function setProtectedSheet(
  store: SpreadsheetStore,
  sheet: number,
  on: boolean,
  options: SheetProtectionOptions = {},
): void {
  mutators.setSheetProtected(
    store,
    sheet,
    on,
    options.password !== undefined ||
      options.passwordHash !== undefined ||
      options.permissions !== undefined
      ? {
          password: options.password,
          passwordHash: options.passwordHash,
          permissions: options.permissions,
        }
      : undefined,
  );
  if (options.workbook)
    flushProtectionToEngine(
      options.workbook,
      sheet,
      on,
      options.password,
      options.permissions,
      options.passwordHash,
    );
}

export function toggleProtectedSheet(
  store: SpreadsheetStore,
  sheet: number,
  options: SheetProtectionOptions = {},
): boolean {
  const on = !isSheetProtected(store.getState(), sheet);
  setProtectedSheet(store, sheet, on, options);
  return on;
}

export function protectedSheetPassword(state: State, sheet: number): string | undefined {
  return state.protection.protectedSheets.get(sheet)?.password;
}

export function protectedSheetPasswordHash(
  state: State,
  sheet: number,
): SheetProtectionPasswordHash | undefined {
  const passwordHash = state.protection.protectedSheets.get(sheet)?.passwordHash;
  return passwordHash ? { ...passwordHash } : undefined;
}

const OOXML_HASH_ALGORITHMS: Readonly<Record<string, AlgorithmIdentifier>> = {
  sha1: 'SHA-1',
  'sha-1': 'SHA-1',
  sha256: 'SHA-256',
  'sha-256': 'SHA-256',
  sha384: 'SHA-384',
  'sha-384': 'SHA-384',
  sha512: 'SHA-512',
  'sha-512': 'SHA-512',
};

const normalizeOoxmlHashAlgorithm = (name: string): AlgorithmIdentifier | null =>
  OOXML_HASH_ALGORITHMS[name.trim().toLowerCase()] ?? null;

const base64ToBytes = (value: string): Uint8Array | null => {
  try {
    const binary = globalThis.atob(value);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  } catch {
    return null;
  }
};

const concatBytes = (left: Uint8Array, right: Uint8Array): Uint8Array => {
  const out = new Uint8Array(left.length + right.length);
  out.set(left, 0);
  out.set(right, left.length);
  return out;
};

const utf16LeBytes = (value: string): Uint8Array => {
  const out = new Uint8Array(value.length * 2);
  for (let i = 0; i < value.length; i += 1) {
    const code = value.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = (code >>> 8) & 0xff;
  }
  return out;
};

const iteratorBytes = (value: number): Uint8Array => {
  const out = new Uint8Array(4);
  out[0] = value & 0xff;
  out[1] = (value >>> 8) & 0xff;
  out[2] = (value >>> 16) & 0xff;
  out[3] = (value >>> 24) & 0xff;
  return out;
};

const sameBytes = (left: Uint8Array, right: Uint8Array): boolean => {
  if (left.length !== right.length) return false;
  let diff = 0;
  for (let i = 0; i < left.length; i += 1) diff |= (left.at(i) ?? 0) ^ (right.at(i) ?? 0);
  return diff === 0;
};

const nodeHashAlgorithm = (algorithm: AlgorithmIdentifier): string | null => {
  if (algorithm === 'SHA-1') return 'sha1';
  if (algorithm === 'SHA-256') return 'sha256';
  if (algorithm === 'SHA-384') return 'sha384';
  if (algorithm === 'SHA-512') return 'sha512';
  return null;
};

type NodeProcessLike = {
  versions?: { node?: unknown };
  getBuiltinModule?: (specifier: string) => unknown;
};

const nodeProcess = (): NodeProcessLike | undefined =>
  (globalThis as { process?: NodeProcessLike }).process;

const isNodeRuntime = (): boolean => typeof nodeProcess()?.versions?.node === 'string';

type NodeHashModule = {
  createHash: (algorithm: string) => {
    update: (data: Uint8Array) => { digest: () => Uint8Array };
  };
};

const getNodeHashModule = (): NodeHashModule | null => {
  if (!isNodeRuntime()) return null;
  try {
    const getBuiltinModule = nodeProcess()?.getBuiltinModule;
    const mod = getBuiltinModule?.('node:crypto') as Partial<NodeHashModule> | undefined;
    return mod && typeof mod.createHash === 'function' ? (mod as NodeHashModule) : null;
  } catch {
    return null;
  }
};

const digestBytes = async (
  subtle: SubtleCrypto,
  algorithm: AlgorithmIdentifier,
  bytes: Uint8Array,
): Promise<Uint8Array> => {
  const input = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(input).set(bytes);
  return new Uint8Array(await subtle.digest(algorithm, input));
};

export async function verifySheetProtectionPasswordHash(
  password: string,
  passwordHash: SheetProtectionPasswordHash,
): Promise<boolean> {
  const algorithm = normalizeOoxmlHashAlgorithm(passwordHash.algorithmName);
  const salt = base64ToBytes(passwordHash.saltValue);
  const expected = base64ToBytes(passwordHash.hashValue);
  if (!algorithm || !salt || !expected || passwordHash.spinCount < 0) return false;

  const nodeCrypto = getNodeHashModule();
  const nodeAlgorithm = nodeHashAlgorithm(algorithm);
  if (nodeCrypto && nodeAlgorithm) {
    let hash = nodeCrypto
      .createHash(nodeAlgorithm)
      .update(concatBytes(salt, utf16LeBytes(password)))
      .digest();
    const block = new Uint8Array(hash.length + 4);
    for (let iterator = 0; iterator < passwordHash.spinCount; iterator += 1) {
      block.set(hash, 0);
      block.set(iteratorBytes(iterator), hash.length);
      hash = nodeCrypto.createHash(nodeAlgorithm).update(block).digest();
    }
    return sameBytes(hash, expected);
  }

  const subtle = globalThis.crypto?.subtle;
  if (!subtle) return false;

  let hash = await digestBytes(subtle, algorithm, concatBytes(salt, utf16LeBytes(password)));
  const block = new Uint8Array(hash.length + 4);
  for (let iterator = 0; iterator < passwordHash.spinCount; iterator += 1) {
    block.set(hash, 0);
    block.set(iteratorBytes(iterator), hash.length);
    hash = await digestBytes(subtle, algorithm, block);
  }
  return sameBytes(hash, expected);
}

export function protectedSheetPermissions(
  state: State,
  sheet: number,
): SheetProtectionPermissions | undefined {
  const permissions = state.protection.protectedSheets.get(sheet)?.permissions;
  return permissions ? { ...permissions } : undefined;
}

export function isWorkbookStructureProtected(state: State): boolean {
  return !!state.protection.workbookStructure;
}

export function workbookStructurePassword(state: State): string | undefined {
  return state.protection.workbookStructure?.password;
}

export function setWorkbookStructureProtected(
  store: SpreadsheetStore,
  on: boolean,
  options: WorkbookStructureProtectionOptions = {},
): void {
  mutators.setWorkbookStructureProtected(
    store,
    on,
    options.password !== undefined ? { password: options.password } : undefined,
  );
}

export function toggleWorkbookStructureProtected(
  store: SpreadsheetStore,
  options: WorkbookStructureProtectionOptions = {},
): boolean {
  const on = !isWorkbookStructureProtected(store.getState());
  setWorkbookStructureProtected(store, on, options);
  return on;
}

export function allowedEditRangesForSheet(state: State, sheet: number): AllowedEditRange[] {
  return state.protection.allowedEditRanges.filter((entry) => entry.range.sheet === sheet);
}

export function isAddrInAllowedEditRange(state: State, addr: Addr): boolean {
  return state.protection.allowedEditRanges.some((entry) => rangeContainsAddr(entry.range, addr));
}

export function addAllowedEditRange(
  store: SpreadsheetStore,
  range: Range,
  options: AllowedEditRangeOptions = {},
): string {
  return mutators.addAllowedEditRange(store, {
    id: options.id,
    title: options.title ?? `Range ${store.getState().protection.allowedEditRanges.length + 1}`,
    range,
    ...(options.password !== undefined ? { password: options.password } : {}),
  });
}

export function clearAllowedEditRanges(store: SpreadsheetStore, sheet?: number): void {
  mutators.clearAllowedEditRanges(store, sheet);
}

const cloneRange = (range: Range): Range => ({ ...range });

const cloneAllowedEditRange = (entry: AllowedEditRange): AllowedEditRange => ({
  ...entry,
  range: cloneRange(entry.range),
});

const cloneSheetProtectionEntry = (entry: {
  password?: string;
  passwordHash?: SheetProtectionPasswordHash;
  permissions?: SheetProtectionPermissions;
}): {
  password?: string;
  passwordHash?: SheetProtectionPasswordHash;
  permissions?: SheetProtectionPermissions;
} => ({
  ...(entry.password !== undefined ? { password: entry.password } : {}),
  ...(entry.passwordHash ? { passwordHash: { ...entry.passwordHash } } : {}),
  ...(entry.permissions ? { permissions: { ...entry.permissions } } : {}),
});

const captureProtectionSnapshot = (state: State): ProtectionSnapshot => ({
  protectedSheets: new Map(
    [...state.protection.protectedSheets.entries()].map(([sheet, entry]) => [
      sheet,
      cloneSheetProtectionEntry(entry),
    ]),
  ),
  ...(state.protection.workbookStructure
    ? { workbookStructure: { ...state.protection.workbookStructure } }
    : {}),
  allowedEditRanges: state.protection.allowedEditRanges.map(cloneAllowedEditRange),
});

const sameProtectionSnapshot = (a: ProtectionSnapshot, b: ProtectionSnapshot): boolean =>
  JSON.stringify({
    protectedSheets: [...a.protectedSheets.entries()].sort(([x], [y]) => x - y),
    workbookStructure: a.workbookStructure ?? null,
    allowedEditRanges: a.allowedEditRanges,
  }) ===
  JSON.stringify({
    protectedSheets: [...b.protectedSheets.entries()].sort(([x], [y]) => x - y),
    workbookStructure: b.workbookStructure ?? null,
    allowedEditRanges: b.allowedEditRanges,
  });

const applyProtectionSnapshot = (
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  snap: ProtectionSnapshot,
): void => {
  const previousSheets = new Set(store.getState().protection.protectedSheets.keys());
  store.setState((s) => ({
    ...s,
    protection: {
      protectedSheets: new Map(
        [...snap.protectedSheets.entries()].map(([sheet, entry]) => [
          sheet,
          cloneSheetProtectionEntry(entry),
        ]),
      ),
      ...(snap.workbookStructure ? { workbookStructure: { ...snap.workbookStructure } } : {}),
      allowedEditRanges: snap.allowedEditRanges.map(cloneAllowedEditRange),
    },
  }));
  if (!wb) return;
  const sheets = new Set([...previousSheets, ...snap.protectedSheets.keys()]);
  for (const sheet of sheets) {
    const entry = snap.protectedSheets.get(sheet);
    flushProtectionToEngine(
      wb,
      sheet,
      !!entry,
      entry?.password,
      entry?.permissions,
      entry?.passwordHash,
    );
  }
};

export function recordProtectionChange<T>(
  history: History | null,
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  mutate: () => T,
): T {
  if (!history || history.isReplaying()) return mutate();
  const before = captureProtectionSnapshot(store.getState());
  const result = mutate();
  const after = captureProtectionSnapshot(store.getState());
  if (!sameProtectionSnapshot(before, after)) {
    history.push({
      undo: () => applyProtectionSnapshot(store, wb, before),
      redo: () => applyProtectionSnapshot(store, wb, after),
    });
  }
  return result;
}

/** Whether the cell at `addr` is locked. desktop default is locked, so a
 *  missing format entry (or `locked === undefined`) returns `true`. Only
 *  an explicit `locked: false` opts the cell out. */
export function isCellLocked(state: State, addr: Addr): boolean {
  const fmt = state.format.formats.get(addrKey(addr));
  return fmt?.locked !== false;
}

/** Combined gate. A cell is writable when EITHER the sheet is unprotected
 *  OR the cell is explicitly unlocked. Used by the command layer at write
 *  time so locked + protected → no-op rather than throw. */
export function isCellWritable(state: State, addr: Addr): boolean {
  if (!isSheetProtected(state, addr.sheet)) return true;
  if (isAddrInAllowedEditRange(state, addr)) return true;
  return !isCellLocked(state, addr);
}

/** Soft warn-and-return helper used by writeInput / writeCoerced wrappers.
 *  Centralizes the console message so the test suite can stub `console.warn`
 *  in one place. */
export function warnProtected(addr: Addr): void {
  // eslint-disable-next-line no-console
  console.warn(
    `formulon-cell: cell at sheet=${addr.sheet} row=${addr.row} col=${addr.col} is locked on a protected sheet; write skipped`,
  );
}

/** Determine whether the protection gate would let a write through anywhere
 *  in `range`. Returns `range` unchanged when the sheet is unprotected, or
 *  when at least one cell inside is explicitly unlocked. Returns `null`
 *  when every cell in the range is gated (sheet protected + no unlock
 *  flag) — format mutators short-circuit on `null` and emit a single
 *  console warning rather than enumerating cells.
 *
 *  Note: when the returned value is non-null, individual locked cells
 *  inside the range are still gated by the lower-level
 *  `gateProtectionAddr` check applied per-cell by the mutator. */
export function gateProtection(state: State, range: Range): Range | null {
  if (!isSheetProtected(state, range.sheet)) return range;
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const fmt = state.format.formats.get(addrKey({ sheet: range.sheet, row: r, col: c }));
      if (
        fmt?.locked === false ||
        isAddrInAllowedEditRange(state, { sheet: range.sheet, row: r, col: c })
      ) {
        return range;
      }
    }
  }
  return null;
}

/** Subset of `range` that survives the per-cell protection gate. Yields
 *  every writable address; an empty iterator means the entire range is
 *  blocked. Used by format / paste paths that need to skip locked cells
 *  while still touching unlocked ones. */
export function* writableAddrs(state: State, range: Range): IterableIterator<Addr> {
  const protectedSheet = isSheetProtected(state, range.sheet);
  for (let r = range.r0; r <= range.r1; r += 1) {
    for (let c = range.c0; c <= range.c1; c += 1) {
      const addr: Addr = { sheet: range.sheet, row: r, col: c };
      if (!protectedSheet || !isCellLocked(state, addr) || isAddrInAllowedEditRange(state, addr)) {
        yield addr;
      }
    }
  }
}

/** Set the `locked` flag across a range. `locked === undefined` is the
 *  desktop default (treated as locked); pass `false` to opt cells out, or
 *  `true` to make the lock explicit. NOT gated by sheet protection — the
 *  whole point of this mutator is to configure protection up front. */
export function setCellLocked(store: SpreadsheetStore, range: Range, locked: boolean): void {
  const patch: Partial<CellFormat> = { locked };
  mutators.setRangeFormat(store, range, patch);
}
