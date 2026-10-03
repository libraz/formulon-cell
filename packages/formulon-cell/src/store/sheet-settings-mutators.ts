import type { Range } from '../engine/types.js';
import type { SpreadsheetStore } from './store.js';
import type {
  PageSetup,
  SheetProtectionPasswordHash,
  SheetProtectionPermissions,
  SheetProtectionState,
} from './types.js';
import { defaultPageSetup } from './types.js';

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
      passwordHash?: SheetProtectionPasswordHash;
      permissions?: SheetProtectionPermissions;
    },
  ): void {
    store.setState((s) => {
      const next = new Map(s.protection.protectedSheets);
      if (on) {
        const entry: SheetProtectionState = {};
        if (options?.password !== undefined) entry.password = options.password;
        if (options?.passwordHash !== undefined) entry.passwordHash = { ...options.passwordHash };
        if (options?.permissions !== undefined) entry.permissions = { ...options.permissions };
        next.set(sheet, entry);
      } else if (next.has(sheet)) {
        next.delete(sheet);
      } else {
        return s;
      }
      return { ...s, protection: { ...s.protection, protectedSheets: next } };
    });
  },

  /** Toggle workbook-structure protection. When enabled, sheet-structure
   *  commands reject adding, deleting, moving, renaming, hiding, and unhiding
   *  sheets. This is intentionally not history-tracked, matching spreadsheet
   *  protection settings. */
  setWorkbookStructureProtected(
    store: SpreadsheetStore,
    on: boolean,
    options?: { password?: string },
  ): void {
    store.setState((s) => {
      const current = s.protection.workbookStructure;
      if (on) {
        const next: { password?: string } = {};
        if (options?.password !== undefined) next.password = options.password;
        if (current && current.password === next.password) return s;
        return { ...s, protection: { ...s.protection, workbookStructure: next } };
      }
      if (!current) return s;
      const { workbookStructure: _discard, ...rest } = s.protection;
      return { ...s, protection: rest };
    });
  },

  addAllowedEditRange(
    store: SpreadsheetStore,
    entry: { id?: string; title: string; range: Range; password?: string },
  ): string {
    const id =
      entry.id ??
      `allowed-edit-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    store.setState((s) => {
      const nextEntry = {
        id,
        title: entry.title,
        range: entry.range,
        ...(entry.password !== undefined ? { password: entry.password } : {}),
      };
      const next = s.protection.allowedEditRanges.filter((r) => r.id !== id);
      return {
        ...s,
        protection: { ...s.protection, allowedEditRanges: [...next, nextEntry] },
      };
    });
    return id;
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
};
