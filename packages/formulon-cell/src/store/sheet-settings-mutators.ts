import { DEFAULT_OUTLINE_SUMMARY } from './outline-summary.js';
import type { SpreadsheetStore } from './store.js';
import type {
  AllowedEditRange,
  OutlineSummary,
  PageSetup,
  ProtectionSlice,
  SheetProtectionPasswordHash,
  SheetProtectionPermissions,
  SheetProtectionState,
  WorkbookStructureProtectionState,
} from './types.js';
import { defaultPageSetup } from './types.js';

const withoutSheetUnlocks = (
  protection: ProtectionSlice,
  sheet: number,
): ReadonlySet<string> | undefined => {
  const unlocked = protection.unlockedAllowedRanges;
  if (!unlocked?.size) return unlocked;
  const sheetIds = new Set(
    protection.allowedEditRanges.filter((e) => e.range.sheet === sheet).map((e) => e.id),
  );
  return new Set([...unlocked].filter((id) => !sheetIds.has(id)));
};

const setRangeUnlocked = (store: SpreadsheetStore, id: string, unlocked: boolean): void => {
  store.setState((s) => {
    const current = s.protection.unlockedAllowedRanges ?? new Set<string>();
    if (current.has(id) === unlocked) return s;
    const next = new Set(current);
    if (unlocked) next.add(id);
    else next.delete(id);
    return { ...s, protection: { ...s.protection, unlockedAllowedRanges: next } };
  });
};

export const newAllowedEditRangeId = (): string =>
  `allowed-edit-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;

/** Per-sheet and workbook settings: protection, allowed-edit ranges, page
 *  setup and sheet tab colors. */
export const sheetSettingsMutators = {
  /** Toggle sheet-level protection for `sheet`. When `on` is `true` the sheet
   *  enters protected mode and the command layer gates writes against
   *  per-cell `locked` flags. The optional password and permission flags
   *  round-trip through the slice so callers can persist them. */
  setSheetProtected(
    store: SpreadsheetStore,
    sheet: number,
    on: boolean,
    options?: {
      password?: string;
      legacyPasswordHash?: string;
      passwordHash?: SheetProtectionPasswordHash;
      permissions?: SheetProtectionPermissions;
    },
  ): void {
    store.setState((s) => {
      const next = new Map(s.protection.protectedSheets);
      if (on) {
        const entry: SheetProtectionState = {};
        if (options?.password !== undefined) entry.password = options.password;
        if (options?.legacyPasswordHash !== undefined) {
          entry.legacyPasswordHash = options.legacyPasswordHash;
        }
        if (options?.passwordHash !== undefined) entry.passwordHash = { ...options.passwordHash };
        if (options?.permissions !== undefined) entry.permissions = { ...options.permissions };
        next.set(sheet, entry);
      } else if (next.has(sheet)) {
        next.delete(sheet);
      } else {
        return s;
      }
      // Unprotecting ends the session unlocks of that sheet's ranges.
      const unlockedAllowedRanges = on
        ? s.protection.unlockedAllowedRanges
        : withoutSheetUnlocks(s.protection, sheet);
      return {
        ...s,
        protection: {
          ...s.protection,
          protectedSheets: next,
          ...(unlockedAllowedRanges ? { unlockedAllowedRanges } : {}),
        },
      };
    });
  },

  /** Toggle workbook-structure protection. When enabled, sheet-structure
   *  commands reject adding, deleting, moving, renaming, hiding, and unhiding
   *  sheets. This is intentionally not history-tracked, matching spreadsheet
   *  protection settings. */
  setWorkbookStructureProtected(
    store: SpreadsheetStore,
    on: boolean,
    options?: WorkbookStructureProtectionState,
  ): void {
    store.setState((s) => {
      const current = s.protection.workbookStructure;
      if (on) {
        const next: WorkbookStructureProtectionState = {};
        if (options?.password !== undefined) next.password = options.password;
        if (options?.legacyPasswordHash !== undefined) {
          next.legacyPasswordHash = options.legacyPasswordHash;
        }
        if (options?.passwordHash !== undefined) next.passwordHash = { ...options.passwordHash };
        if (current && JSON.stringify(current) === JSON.stringify(next)) return s;
        return { ...s, protection: { ...s.protection, workbookStructure: next } };
      }
      if (!current) return s;
      const { workbookStructure: _discard, ...rest } = s.protection;
      return { ...s, protection: rest };
    });
  },

  addAllowedEditRange(
    store: SpreadsheetStore,
    entry: Omit<AllowedEditRange, 'id'> & { id?: string },
  ): string {
    const id = entry.id ?? newAllowedEditRangeId();
    store.setState((s) => {
      const nextEntry: AllowedEditRange = { ...entry, id };
      const next = s.protection.allowedEditRanges.filter((r) => r.id !== id);
      return {
        ...s,
        protection: { ...s.protection, allowedEditRanges: [...next, nextEntry] },
      };
    });
    return id;
  },

  /** Replace one entry in place, keeping its position. A no-op for an unknown id. */
  replaceAllowedEditRange(store: SpreadsheetStore, entry: AllowedEditRange): void {
    store.setState((s) => {
      const at = s.protection.allowedEditRanges.findIndex((r) => r.id === entry.id);
      if (at < 0) return s;
      const next = [...s.protection.allowedEditRanges];
      next[at] = entry;
      return { ...s, protection: { ...s.protection, allowedEditRanges: next } };
    });
  },

  removeAllowedEditRange(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.protection.allowedEditRanges.filter((r) => r.id !== id);
      if (next.length === s.protection.allowedEditRanges.length) return s;
      return { ...s, protection: { ...s.protection, allowedEditRanges: next } };
    });
    setRangeUnlocked(store, id, false);
  },

  /** Mark a password-guarded range unlocked (or locked again) for this session. */
  setAllowedEditRangeUnlocked: setRangeUnlocked,

  /** Replace every entry of `sheet` with `entries`, leaving other sheets alone. */
  setAllowedEditRangesForSheet(
    store: SpreadsheetStore,
    sheet: number,
    entries: readonly AllowedEditRange[],
  ): void {
    store.setState((s) => {
      const others = s.protection.allowedEditRanges.filter((r) => r.range.sheet !== sheet);
      return {
        ...s,
        protection: { ...s.protection, allowedEditRanges: [...others, ...entries] },
      };
    });
  },

  clearAllowedEditRanges(store: SpreadsheetStore, sheet?: number): void {
    store.setState((s) => {
      const next =
        sheet === undefined
          ? []
          : s.protection.allowedEditRanges.filter((entry) => entry.range.sheet !== sheet);
      if (next.length === s.protection.allowedEditRanges.length) return s;
      return { ...s, protection: { ...s.protection, allowedEditRanges: next } };
    });
  },

  /** Merge a partial patch into the page-setup for `sheet`. Pass `null` to
   *  reset that sheet back to defaults. The merge is shallow except for
   *  `margins` and `printableBounds`, which are deep-merged — so a patch like
   *  `{ margins: { top: 1 } }` preserves the other three sides. */
  setPageSetup(store: SpreadsheetStore, sheet: number, patch: Partial<PageSetup> | null): void {
    store.setState((s) => {
      const setupBySheet = new Map(s.pageSetup.setupBySheet);
      if (patch === null) {
        setupBySheet.delete(sheet);
      } else {
        const prev = setupBySheet.get(sheet) ?? defaultPageSetup();
        const next: PageSetup = { ...prev, ...patch };
        if (patch.margins) {
          next.margins = { ...prev.margins, ...patch.margins };
        }
        if (patch.printableBounds) {
          next.printableBounds = {
            ...(prev.printableBounds ?? patch.printableBounds),
            ...patch.printableBounds,
          };
        }
        setupBySheet.set(sheet, next);
      }
      return { ...s, pageSetup: { setupBySheet } };
    });
  },

  setSheetTabColor(store: SpreadsheetStore, sheet: number, color: string | null): void {
    store.setState((s) => {
      const sheetTabColors = new Map(s.layout.sheetTabColors);
      if (color) {
        sheetTabColors.set(sheet, color);
      } else {
        sheetTabColors.delete(sheet);
      }
      return { ...s, layout: { ...s.layout, sheetTabColors } };
    });
  },

  /** Set where a sheet's outline summary rows / columns sit. A sheet back at
   *  the defaults (below, right) drops its entry. */
  setOutlineSummary(store: SpreadsheetStore, sheet: number, patch: Partial<OutlineSummary>): void {
    store.setState((s) => {
      const outlineSummary = new Map(s.layout.outlineSummary);
      const next = { ...DEFAULT_OUTLINE_SUMMARY, ...outlineSummary.get(sheet), ...patch };
      if (next.below && next.right) outlineSummary.delete(sheet);
      else outlineSummary.set(sheet, next);
      return { ...s, layout: { ...s.layout, outlineSummary } };
    });
  },
};
