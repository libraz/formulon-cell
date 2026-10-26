import { addrKey, parseAddrKey } from '../engine/address.js';
import {
  hashSheetProtectionPassword,
  legacySheetPasswordHash,
  verifySheetProtectionPasswordHash,
} from '../engine/protection-password.js';
import {
  allowedEditRangeAreas,
  flushProtectionToEngine,
  flushWorkbookProtectionToEngine,
  syncAllowedEditRangesToEngine,
} from '../engine/protection-sync.js';
import type { Addr, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import { rangeArea, rangeContainsAddr, rangeContainsRange } from '../store/selection-geometry.js';
import {
  type AllowedEditRange,
  type CellFormat,
  mutators,
  type SheetProtectionPasswordHash,
  type SheetProtectionPermissions,
  type SpreadsheetStore,
  type State,
  type WorkbookStructureProtectionState,
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
  passwordHash?: SheetProtectionPasswordHash;
  workbook?: WorkbookHandle;
}

export interface AllowedEditRangeOptions {
  id?: string;
  title?: string;
  password?: string;
  extraAreas?: readonly Range[];
  workbook?: WorkbookHandle;
}

export interface AllowedEditRangePatch {
  title?: string;
  range?: Range;
  extraAreas?: readonly Range[];
  /** A new password, or `null` to remove the current one. */
  password?: string | null;
}

interface ProtectionSnapshot {
  protectedSheets: Map<
    number,
    {
      password?: string;
      legacyPasswordHash?: string;
      passwordHash?: SheetProtectionPasswordHash;
      permissions?: SheetProtectionPermissions;
    }
  >;
  workbookStructure?: WorkbookStructureProtectionState;
  allowedEditRanges: AllowedEditRange[];
}

const MAX_MATERIALIZED_LOCK_CELLS = 100_000;

/** Whether `sheet` is currently flagged protected on the workbook. Mirrors
 *  the protection slice as a pure helper so call sites don't reach into the
 *  Map shape directly. */
export function isSheetProtected(state: State, sheet: number): boolean {
  return state.protection.protectedSheets.has(sheet);
}

/** Spreadsheet-parity gate for row/col structure changes. When `sheet` is
 *  protected the operation is rejected (no-op + warning) regardless of
 *  per-cell locks — spreadsheets disable the insert/delete row/col commands
 *  wholesale on protected sheets. */
export function blockedByProtection(store: SpreadsheetStore, sheet: number, op: string): boolean {
  if (!isSheetProtected(store.getState(), sheet)) return false;
  // eslint-disable-next-line no-console
  console.warn(`formulon-cell: ${op} blocked — sheet ${sheet} is protected`);
  return true;
}

export { verifySheetProtectionPasswordHash };

/** Sheets whose typed password is still being hashed. */
const pendingPasswordHashes = new Set<Promise<void>>();

/** Resolves once every typed password (sheet, workbook structure and allowed
 *  edit range) has been hashed and written; await it before saving to be sure
 *  the file carries the modern hash rather than the 16-bit verifier written
 *  immediately. */
export const whenSheetPasswordsHashed = async (): Promise<void> => {
  while (pendingPasswordHashes.size > 0) await Promise.all([...pendingPasswordHashes]);
};

/** Replace the 16-bit verifier written for a freshly typed password with the
 *  modern salted hash, unless the sheet was changed in the meantime. */
const upgradeSheetPasswordHash = (
  store: SpreadsheetStore,
  workbook: WorkbookHandle,
  sheet: number,
  password: string,
): void => {
  upgradeStoredPassword(
    password,
    () => store.getState().protection.protectedSheets.get(sheet),
    (entry, passwordHash) => {
      mutators.setSheetProtected(store, sheet, true, { ...entry, passwordHash });
      flushProtectionToEngine(workbook, sheet, true, undefined, entry.permissions, passwordHash);
    },
  );
};

interface PasswordFields {
  password?: string;
  legacyPasswordHash?: string;
  passwordHash?: SheetProtectionPasswordHash;
}

const hasPasswordFields = (entry: PasswordFields): boolean =>
  !!entry.password || !!entry.passwordHash || !!entry.legacyPasswordHash;

/** Whether `password` matches typed text of this session, a modern salted
 *  hash, or the 16-bit verifier of a loaded file; true when none is set. */
const verifyPasswordFields = async (entry: PasswordFields, password: string): Promise<boolean> => {
  if (entry.password !== undefined) return password === entry.password;
  if (entry.passwordHash) return verifySheetProtectionPasswordHash(password, entry.passwordHash);
  if (entry.legacyPasswordHash !== undefined) {
    return legacySheetPasswordHash(password) === entry.legacyPasswordHash.toUpperCase();
  }
  return true;
};

/** Whether `password` unlocks the protected `sheet`. */
export async function verifySheetProtectionPassword(
  state: State,
  sheet: number,
  password: string,
): Promise<boolean> {
  const entry = state.protection.protectedSheets.get(sheet);
  return entry ? verifyPasswordFields(entry, password) : true;
}

/** Whether unprotecting `sheet` requires a password. */
export function sheetHasPassword(state: State, sheet: number): boolean {
  const entry = state.protection.protectedSheets.get(sheet);
  return !!entry && hasPasswordFields(entry);
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
  if (on && options.workbook && options.password && !options.passwordHash) {
    upgradeSheetPasswordHash(store, options.workbook, sheet, options.password);
  }
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

const rangeOverlaps = (left: Range, right: Range): boolean =>
  left.sheet === right.sheet &&
  left.r0 <= right.r1 &&
  left.r1 >= right.r0 &&
  left.c0 <= right.c1 &&
  left.c1 >= right.c0;

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

/** The workbook structure password typed this session.
 *  @deprecated It is undefined once the password is hashed (right after
 *  protecting) and for loaded files; use `verifyWorkbookStructurePassword`. */
export function workbookStructurePassword(state: State): string | undefined {
  return state.protection.workbookStructure?.password;
}

/** Whether unprotecting the workbook structure requires a password. */
export function workbookStructureHasPassword(state: State): boolean {
  const entry = state.protection.workbookStructure;
  return !!entry && hasPasswordFields(entry);
}

/** Whether `password` unlocks the protected workbook structure. */
export async function verifyWorkbookStructurePassword(
  state: State,
  password: string,
): Promise<boolean> {
  const entry = state.protection.workbookStructure;
  return entry ? verifyPasswordFields(entry, password) : true;
}

/** Swap a freshly typed password for its salted hash once computed, unless
 *  `current` no longer holds the typed text. The typed text is dropped. */
const upgradeStoredPassword = <T extends PasswordFields>(
  password: string,
  current: () => T | undefined,
  apply: (entry: T, passwordHash: SheetProtectionPasswordHash) => void,
): void => {
  const task = hashSheetProtectionPassword(password)
    .then((passwordHash) => {
      const entry = current();
      if (!passwordHash || entry?.password !== password || entry.passwordHash) return;
      apply(entry, passwordHash);
    })
    .catch(() => {})
    .finally(() => {
      pendingPasswordHashes.delete(task);
    });
  pendingPasswordHashes.add(task);
};

const upgradeWorkbookPasswordHash = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  password: string,
): void => {
  upgradeStoredPassword(
    password,
    () => store.getState().protection.workbookStructure,
    (_entry, passwordHash) => {
      mutators.setWorkbookStructureProtected(store, true, { passwordHash });
      flushWorkbookProtectionToEngine(wb, { passwordHash });
    },
  );
};

export function setWorkbookStructureProtected(
  store: SpreadsheetStore,
  on: boolean,
  options: WorkbookStructureProtectionOptions = {},
): void {
  const state: WorkbookStructureProtectionState = {
    ...(options.password ? { password: options.password } : {}),
    ...(options.passwordHash ? { passwordHash: options.passwordHash } : {}),
  };
  mutators.setWorkbookStructureProtected(store, on, on ? state : undefined);
  const wb = options.workbook;
  if (!wb) return;
  flushWorkbookProtectionToEngine(wb, on ? state : null);
  if (on && state.password && !state.passwordHash) {
    upgradeWorkbookPasswordHash(store, wb, state.password);
  }
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

/** Whether an allowed edit range admits edits now: it has no password, or
 *  its password was entered this session. */
export const isAllowedEditRangeOpen = (state: State, entry: AllowedEditRange): boolean =>
  !hasPasswordFields(entry) || !!state.protection.unlockedAllowedRanges?.has(entry.id);

export function isAddrInAllowedEditRange(state: State, addr: Addr): boolean {
  return state.protection.allowedEditRanges.some(
    (entry) =>
      isAllowedEditRangeOpen(state, entry) &&
      allowedEditRangeAreas(entry).some((area) => rangeContainsAddr(area, addr)),
  );
}

/** Whether `range` lies inside one area of an allowed edit range. */
export function isRangeInAllowedEditRange(state: State, range: Range): boolean {
  return state.protection.allowedEditRanges.some(
    (entry) =>
      isAllowedEditRangeOpen(state, entry) &&
      allowedEditRangeAreas(entry).some((area) => rangeContainsRange(area, range)),
  );
}

const MAX_SCANNED_EDIT_CELLS = 100_000;

interface AllowedRangeEditInspection {
  /** Guarded ranges whose password the edit needs. */
  pending: AllowedEditRange[];
  /** The edit also reaches a locked cell no allowed range covers. */
  blocked: boolean;
}

/** Which password-guarded ranges hold back an edit of `ranges` on a protected
 *  sheet, and whether the edit reaches a locked cell outside every allowed
 *  range (desktop refuses such a mixed edit as a whole). Unlocked cells never
 *  need a password. Targets above the scan limit report their overlapping
 *  guarded ranges and are never marked blocked. */
export function inspectAllowedRangeEdit(
  state: State,
  ranges: readonly Range[],
): AllowedRangeEditInspection {
  const pending = new Set<AllowedEditRange>();
  let blocked = false;
  for (const range of ranges) {
    if (!isSheetProtected(state, range.sheet)) continue;
    const onSheet = state.protection.allowedEditRanges.filter((e) => e.range.sheet === range.sheet);
    const guarded = onSheet.filter(
      (e) =>
        !isAllowedEditRangeOpen(state, e) &&
        allowedEditRangeAreas(e).some((area) => rangeOverlaps(area, range)),
    );
    if (guarded.length === 0) continue;
    if (rangeArea(range) > MAX_SCANNED_EDIT_CELLS) {
      for (const entry of guarded) pending.add(entry);
      continue;
    }
    for (let row = range.r0; row <= range.r1; row += 1) {
      for (let col = range.c0; col <= range.c1; col += 1) {
        const addr = { sheet: range.sheet, row, col };
        if (!isCellLocked(state, addr)) continue;
        const covering = onSheet.filter((e) =>
          allowedEditRangeAreas(e).some((area) => rangeContainsAddr(area, addr)),
        );
        if (covering.length === 0) blocked = true;
        else if (!covering.some((e) => isAllowedEditRangeOpen(state, e))) {
          const first = covering[0];
          if (first) pending.add(first);
        }
      }
    }
  }
  return { pending: [...pending], blocked };
}

/** Password-guarded ranges that still hold back an edit of `ranges` on a
 *  protected sheet. */
export function allowedRangesPendingUnlock(
  state: State,
  ranges: readonly Range[],
): AllowedEditRange[] {
  return inspectAllowedRangeEdit(state, ranges).pending;
}

/** The dialogs a host offers for unlocking a password-guarded range. */
export interface AllowedRangeUnlockUi {
  /** The typed password, or null when the user cancels. */
  askPassword(title: string): Promise<string | null>;
  reportIncorrect(title: string): Promise<void>;
}

const unlockUis = new WeakMap<SpreadsheetStore, AllowedRangeUnlockUi>();

/** Register how `store` asks for range passwords; returns the unregister call.
 *  Without one, edits in a guarded range stay refused. */
export function registerAllowedRangeUnlockUi(
  store: SpreadsheetStore,
  ui: AllowedRangeUnlockUi,
): () => void {
  unlockUis.set(store, ui);
  return () => {
    if (unlockUis.get(store) === ui) unlockUis.delete(store);
  };
}

/** Ask for the password of every guarded range an edit of `ranges` touches and
 *  unlock them for the session. False when none is registered, the user
 *  cancels, or a password is wrong. */
export function requestAllowedRangeUnlock(
  store: SpreadsheetStore,
  ranges: readonly Range[],
): Promise<boolean> {
  // One prompt at a time per store: a later request waits, then re-checks what is still locked.
  const previous = inflightUnlocks.get(store);
  const task = (async () => {
    if (previous) await previous.catch(() => {});
    return unlockPending(store, ranges);
  })();
  inflightUnlocks.set(store, task);
  const clear = (): void => {
    if (inflightUnlocks.get(store) === task) inflightUnlocks.delete(store);
  };
  task.then(clear, clear);
  return task;
}

const inflightUnlocks = new WeakMap<SpreadsheetStore, Promise<boolean>>();

async function unlockPending(store: SpreadsheetStore, ranges: readonly Range[]): Promise<boolean> {
  const { pending: needed, blocked } = inspectAllowedRangeEdit(store.getState(), ranges);
  if (blocked) return false;
  for (const pending of needed) {
    if (isAllowedEditRangeOpen(store.getState(), pending)) continue;
    const ui = unlockUis.get(store);
    if (!ui) return false;
    const password = await ui.askPassword(pending.title);
    if (password === null) return false;
    if (!(await verifyAllowedEditRangePassword(store.getState(), pending.id, password))) {
      await ui.reportIncorrect(pending.title);
      return false;
    }
    mutators.setAllowedEditRangeUnlocked(store, pending.id, true);
  }
  return true;
}

/** When a guarded range holds back an edit of `ranges`, ask for its password and
 *  call `run` once unlocked (a wrong or cancelled password drops it). Returns
 *  true when the edit was deferred, false when it can go ahead now.
 *  `onSettled` fires first, whichever way the prompt ends. */
export function deferForAllowedRangeUnlock(
  store: SpreadsheetStore,
  ranges: readonly Range[],
  run: () => void,
  onSettled?: () => void,
): boolean {
  const { pending, blocked } = inspectAllowedRangeEdit(store.getState(), ranges);
  if (pending.length === 0) return false;
  if (blocked) {
    // eslint-disable-next-line no-console
    console.warn(
      'formulon-cell: edit blocked — it reaches locked cells outside every allowed range',
    );
    return true;
  }
  void requestAllowedRangeUnlock(store, ranges).then((ok) => {
    onSettled?.();
    if (ok) run();
  });
  return true;
}

/** The ranges a selection-based edit writes: the selection and its extra areas. */
export const selectionEditRanges = (state: State): Range[] => [
  state.selection.range,
  ...(state.selection.extraRanges ?? []),
];

/** Run `run` now, or once the guarded ranges `ranges` touches are unlocked. */
export function runAfterAllowedRangeUnlock(
  store: SpreadsheetStore,
  ranges: readonly Range[],
  run: () => void,
): void {
  if (!deferForAllowedRangeUnlock(store, ranges, run)) run();
}

/** Whether unlocking the allowed edit range `id` requires a password. */
export function allowedEditRangeHasPassword(state: State, id: string): boolean {
  const entry = state.protection.allowedEditRanges.find((e) => e.id === id);
  return !!entry && hasPasswordFields(entry);
}

/** Whether `password` unlocks the allowed edit range `id`. */
export async function verifyAllowedEditRangePassword(
  state: State,
  id: string,
  password: string,
): Promise<boolean> {
  const entry = state.protection.allowedEditRanges.find((e) => e.id === id);
  return entry ? verifyPasswordFields(entry, password) : true;
}

/** Run `mutate` and write the resulting change in allowed edit ranges to `wb`. */
const syncRangeChange = (
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  mutate: () => void,
): void => {
  const before = store.getState().protection.allowedEditRanges;
  mutate();
  if (wb) syncAllowedEditRangesToEngine(wb, before, store.getState().protection.allowedEditRanges);
};

/** `title`, made unique among the ranges of `sheet` (the engine requires it). */
const uniqueRangeTitle = (state: State, sheet: number, title: string, selfId?: string): string => {
  const taken = new Set(
    state.protection.allowedEditRanges
      .filter((e) => e.range.sheet === sheet && e.id !== selfId)
      .map((e) => e.title),
  );
  const base = title.trim() || `Range ${taken.size + 1}`;
  let candidate = base;
  for (let n = 2; taken.has(candidate); n += 1) candidate = `${base} (${n})`;
  return candidate;
};

const upgradeRangePasswordHash = (
  store: SpreadsheetStore,
  wb: WorkbookHandle | undefined,
  id: string,
  password: string,
): void => {
  upgradeStoredPassword(
    password,
    () => store.getState().protection.allowedEditRanges.find((e) => e.id === id),
    (entry, passwordHash) => {
      const { password: _typed, ...rest } = entry;
      syncRangeChange(store, wb, () =>
        mutators.replaceAllowedEditRange(store, { ...rest, passwordHash }),
      );
    },
  );
};

/** Allowed edit ranges cannot be edited while their sheet is protected. */
const rangesLocked = (state: State, sheet: number): boolean => {
  if (!isSheetProtected(state, sheet)) return false;
  // eslint-disable-next-line no-console
  console.warn(`formulon-cell: allowed edit ranges are locked — sheet ${sheet} is protected`);
  return true;
};

/** Add an allowed edit range. Null when the sheet is protected. */
export function addAllowedEditRange(
  store: SpreadsheetStore,
  range: Range,
  options: AllowedEditRangeOptions = {},
): string | null {
  const state = store.getState();
  if (rangesLocked(state, range.sheet)) return null;
  const title = uniqueRangeTitle(
    state,
    range.sheet,
    options.title ?? `Range ${state.protection.allowedEditRanges.length + 1}`,
    options.id,
  );
  let id = '';
  syncRangeChange(store, options.workbook, () => {
    id = mutators.addAllowedEditRange(store, {
      id: options.id,
      title,
      range,
      ...(options.extraAreas?.length ? { extraAreas: options.extraAreas } : {}),
      ...(options.password ? { password: options.password } : {}),
    });
  });
  if (options.password) upgradeRangePasswordHash(store, options.workbook, id, options.password);
  return id;
}

/** Change an allowed edit range in place. Returns false for an unknown id. */
export function updateAllowedEditRange(
  store: SpreadsheetStore,
  id: string,
  patch: AllowedEditRangePatch,
  options: { workbook?: WorkbookHandle } = {},
): boolean {
  const current = store.getState().protection.allowedEditRanges.find((e) => e.id === id);
  if (!current || rangesLocked(store.getState(), current.range.sheet)) return false;
  const { password, legacyPasswordHash, passwordHash, extraAreas: currentAreas, ...base } = current;
  const range = patch.range ?? current.range;
  const extraAreas = patch.range ? patch.extraAreas : (patch.extraAreas ?? currentAreas);
  const next: AllowedEditRange = {
    ...base,
    title: uniqueRangeTitle(store.getState(), range.sheet, patch.title ?? current.title, id),
    range,
    ...(extraAreas?.length ? { extraAreas } : {}),
    ...(patch.password === undefined
      ? {
          ...(password !== undefined ? { password } : {}),
          ...(legacyPasswordHash !== undefined ? { legacyPasswordHash } : {}),
          ...(passwordHash ? { passwordHash } : {}),
        }
      : patch.password
        ? { password: patch.password }
        : {}),
  };
  syncRangeChange(store, options.workbook, () => mutators.replaceAllowedEditRange(store, next));
  // A new or removed password starts locked again.
  if (patch.password !== undefined) mutators.setAllowedEditRangeUnlocked(store, id, false);
  if (patch.password) upgradeRangePasswordHash(store, options.workbook, id, patch.password);
  return true;
}

export function removeAllowedEditRange(
  store: SpreadsheetStore,
  id: string,
  options: { workbook?: WorkbookHandle } = {},
): void {
  const current = store.getState().protection.allowedEditRanges.find((e) => e.id === id);
  if (!current || rangesLocked(store.getState(), current.range.sheet)) return;
  syncRangeChange(store, options.workbook, () => mutators.removeAllowedEditRange(store, id));
}

export function clearAllowedEditRanges(
  store: SpreadsheetStore,
  sheet?: number,
  options: { workbook?: WorkbookHandle } = {},
): void {
  if (sheet !== undefined && rangesLocked(store.getState(), sheet)) return;
  syncRangeChange(store, options.workbook, () => {
    if (sheet !== undefined) {
      mutators.clearAllowedEditRanges(store, sheet);
      return;
    }
    // Every sheet: the protected ones keep their ranges.
    for (const entry of store.getState().protection.allowedEditRanges) {
      if (!isSheetProtected(store.getState(), entry.range.sheet)) {
        mutators.removeAllowedEditRange(store, entry.id);
      }
    }
  });
}

const cloneRange = (range: Range): Range => ({ ...range });

const cloneAllowedEditRange = (entry: AllowedEditRange): AllowedEditRange => ({
  ...entry,
  range: cloneRange(entry.range),
  ...(entry.extraAreas ? { extraAreas: entry.extraAreas.map(cloneRange) } : {}),
  ...(entry.passwordHash ? { passwordHash: { ...entry.passwordHash } } : {}),
});

const cloneSheetProtectionEntry = (entry: {
  password?: string;
  legacyPasswordHash?: string;
  passwordHash?: SheetProtectionPasswordHash;
  permissions?: SheetProtectionPermissions;
}): {
  password?: string;
  legacyPasswordHash?: string;
  passwordHash?: SheetProtectionPasswordHash;
  permissions?: SheetProtectionPermissions;
} => ({
  ...(entry.password !== undefined ? { password: entry.password } : {}),
  ...(entry.legacyPasswordHash !== undefined
    ? { legacyPasswordHash: entry.legacyPasswordHash }
    : {}),
  ...(entry.passwordHash ? { passwordHash: { ...entry.passwordHash } } : {}),
  ...(entry.permissions ? { permissions: { ...entry.permissions } } : {}),
});

const cloneWorkbookStructure = (
  entry: WorkbookStructureProtectionState,
): WorkbookStructureProtectionState => ({
  ...entry,
  ...(entry.passwordHash ? { passwordHash: { ...entry.passwordHash } } : {}),
});

const captureProtectionSnapshot = (state: State): ProtectionSnapshot => ({
  protectedSheets: new Map(
    [...state.protection.protectedSheets.entries()].map(([sheet, entry]) => [
      sheet,
      cloneSheetProtectionEntry(entry),
    ]),
  ),
  ...(state.protection.workbookStructure
    ? { workbookStructure: cloneWorkbookStructure(state.protection.workbookStructure) }
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
  const previousRanges = store.getState().protection.allowedEditRanges;
  // Sheets this snapshot unprotects end their session unlocks, as the mutator does.
  const unprotected = new Set(
    [...previousSheets].filter((sheet) => !snap.protectedSheets.has(sheet)),
  );
  store.setState((s) => ({
    ...s,
    protection: {
      protectedSheets: new Map(
        [...snap.protectedSheets.entries()].map(([sheet, entry]) => [
          sheet,
          cloneSheetProtectionEntry(entry),
        ]),
      ),
      ...(snap.workbookStructure
        ? { workbookStructure: cloneWorkbookStructure(snap.workbookStructure) }
        : {}),
      allowedEditRanges: snap.allowedEditRanges.map(cloneAllowedEditRange),
      ...(s.protection.unlockedAllowedRanges
        ? {
            unlockedAllowedRanges: new Set(
              [...s.protection.unlockedAllowedRanges].filter(
                (id) =>
                  !s.protection.allowedEditRanges.some(
                    (e) => e.id === id && unprotected.has(e.range.sheet),
                  ),
              ),
            ),
          }
        : {}),
    },
  }));
  // A snapshot taken before hashing finished still holds typed text: hash it again.
  for (const entry of snap.allowedEditRanges) {
    if (entry.password && !entry.passwordHash) {
      upgradeRangePasswordHash(store, wb, entry.id, entry.password);
    }
  }
  const structure = snap.workbookStructure;
  if (wb && structure?.password && !structure.passwordHash) {
    upgradeWorkbookPasswordHash(store, wb, structure.password);
  }
  if (wb) {
    for (const [sheet, entry] of snap.protectedSheets) {
      if (entry.password && !entry.passwordHash) {
        upgradeSheetPasswordHash(store, wb, sheet, entry.password);
      }
    }
  }
  if (!wb) return;
  flushWorkbookProtectionToEngine(wb, snap.workbookStructure ?? null);
  syncAllowedEditRangesToEngine(wb, previousRanges, snap.allowedEditRanges);
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
      entry?.legacyPasswordHash,
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
  for (const [key, fmt] of state.format.formats) {
    if (fmt.locked !== false) continue;
    const addr = parseAddrKey(key);
    if (
      addr &&
      addr.sheet === range.sheet &&
      addr.row >= range.r0 &&
      addr.row <= range.r1 &&
      addr.col >= range.c0 &&
      addr.col <= range.c1
    ) {
      return range;
    }
  }
  return state.protection.allowedEditRanges.some(
    (entry) =>
      isAllowedEditRangeOpen(state, entry) &&
      allowedEditRangeAreas(entry).some((area) => rangeOverlaps(area, range)),
  )
    ? range
    : null;
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
  if (rangeArea(range) > MAX_MATERIALIZED_LOCK_CELLS) {
    if (!locked) return;
    store.setState((s) => {
      const formats = new Map(s.format.formats);
      for (const [key, current] of s.format.formats) {
        const addr = parseAddrKey(key);
        if (!addr || !rangeContainsAddr(range, addr) || current.locked !== false) continue;
        const next: CellFormat = { ...current };
        delete next.locked;
        if (Object.keys(next).length > 0) formats.set(key, next);
        else formats.delete(key);
      }
      return { ...s, format: { ...s.format, formats } };
    });
    return;
  }
  const patch: Partial<CellFormat> = { locked };
  mutators.setRangeFormat(store, range, patch);
}
