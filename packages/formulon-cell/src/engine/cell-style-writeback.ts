import {
  CELL_STYLES,
  type CellStyleFallbackProfile,
  type CellStyleFormatGroup,
  type CellStyleId,
  cellStyleFallbackFormat,
  cellStyleGroups,
  customCellStyleId,
  formatForCellStyleGroups,
  getCellStyle,
  stylePayloadMatches,
  stylePayloadMatchesNative,
} from '../commands/cell-styles.js';
import type { CellFormat, CustomCellStyle, State } from '../store/store.js';
import { cellFormatFromXf } from './cell-format-sync.js';
import { localeForWorkbook } from './profile-text-semantics.js';
import type { CellXf } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';

const OOXML_STYLE_NAMES: Partial<Record<CellStyleId, string>> = {
  warning: 'Warning Text',
};

/** The `<cellStyleXfs>` row every unstyled cell inherits from. OOXML reserves
 *  index 0 for it, and a cell XF that names no style leaves `xfId` at 0. */
const NORMAL_STYLE_XF_ID = 0;

/** A named style as it is about to be written: the store-side key that cells
 *  carry in `CellFormat.cellStyle`, the OOXML name, and the format the style
 *  itself defines. */
export interface NamedStyleRegistration {
  key: string;
  name: string;
  builtinId: number | null;
  format: Partial<CellFormat>;
  groups: readonly CellStyleFormatGroup[];
}

export interface NamedStyleSyncPlan {
  readonly styleXfIds: ReadonlyMap<string, number>;
  commit(): void;
}

interface PendingNamedStyleRegistration {
  readonly record: NamedCellStyleRecord;
  readonly previous?: NamedCellStyleRecord;
}

type NamedCellStyleRecord = Parameters<WorkbookHandle['setNamedCellStyleRecord']>[0];
type NamedStyleCatalogEntry = NonNullable<ReturnType<WorkbookHandle['getCellStyle']>> & {
  readonly index: number;
};

export interface NamedStyleXfResolution {
  readonly xfId: number;
  readonly raw: CellXf;
  readonly groups: readonly CellStyleFormatGroup[];
  readonly projected: Partial<CellFormat>;
}

const ooxmlNameFor = (id: CellStyleId, label: string): string => OOXML_STYLE_NAMES[id] ?? label;

/** Reverse of `ooxmlNameFor` + the gallery labels: maps an OOXML style name
 *  back onto the gallery id a hydrated cell should carry, so a style survives
 *  a save/load without turning into a look-alike custom entry. */
export function cellStyleKeyForOoxmlName(name: string): string {
  const needle = name.trim().toLowerCase();
  for (const def of CELL_STYLES) {
    if (ooxmlNameFor(def.id, def.label).toLowerCase() === needle) return def.id;
  }
  return name;
}

/** Every named style referenced by at least one cell, plus every custom style
 *  the user defined, in a stable order. Unused built-ins are not written —
 *  the desktop spreadsheet keeps the full gallery in every file, but emitting
 *  each unused entry would bloat a workbook the user never styled. */
export function collectNamedStyles(state: State, locale = 'en-US'): NamedStyleRegistration[] {
  const keys = new Set<string>();
  const materialized = new Map<string, Partial<CellFormat>[]>();
  for (const fmt of state.format.formats.values()) {
    if (!fmt.cellStyle) continue;
    keys.add(fmt.cellStyle);
    const values = materialized.get(fmt.cellStyle) ?? [];
    values.push(fmt);
    materialized.set(fmt.cellStyle, values);
  }
  const out: NamedStyleRegistration[] = [];
  for (const key of keys) {
    const builtin = getCellStyle(key as CellStyleId);
    if (builtin) {
      const groups = cellStyleGroups(builtin);
      const candidates = materialized.get(key) ?? [];
      const projected = candidates.map((candidate) => formatForCellStyleGroups(candidate, groups));
      const profiles: readonly CellStyleFallbackProfile[] = ['default', 'excel365Mac'];
      const profile = profiles.find((candidateProfile) => {
        const fallback = cellStyleFallbackFormat(builtin.id, candidateProfile, locale);
        return (
          projected.length > 0 &&
          projected.every((candidate) => stylePayloadMatches(candidate, fallback, groups))
        );
      });
      out.push({
        key,
        name: ooxmlNameFor(builtin.id, builtin.label),
        builtinId: builtin.builtinId ?? null,
        format: profile
          ? cellStyleFallbackFormat(builtin.id, profile, locale)
          : cellStyleFallbackFormat(builtin.id, 'default', locale),
        groups,
      });
      continue;
    }
    const custom = state.format.customCellStyles?.find((s) => s.label === key);
    if (custom) {
      out.push({
        key,
        name: custom.label,
        builtinId: null,
        format: custom.format,
        groups: cellStyleGroups(custom),
      });
    }
  }
  for (const custom of state.format.customCellStyles ?? []) {
    if (keys.has(custom.label) || getCellStyle(custom.label as CellStyleId)) continue;
    out.push({
      key: custom.label,
      name: custom.label,
      builtinId: null,
      format: custom.format,
      groups: cellStyleGroups(custom),
    });
  }
  return out;
}

/** Resolve the raw catalog XFs used by named styles. Component selectors and
 * presence bits are retained alongside the UI projection so cell writeback
 * can reuse them without flattening theme/indexed colors into RGB. */
export function resolveNamedStyleXfs(
  wb: WorkbookHandle,
  state: State,
  ids: ReadonlyMap<string, number>,
): ReadonlyMap<string, NamedStyleXfResolution> {
  const registrations = new Map(
    collectNamedStyles(state, localeForWorkbook(wb)).map((registration) => [
      registration.key,
      registration,
    ]),
  );
  const resolved = new Map<string, NamedStyleXfResolution>();
  if (!wb.capabilities.cellStyles || typeof wb.getNamedCellStyles !== 'function') return resolved;
  const add = (key: string, xfId: number, groups: readonly CellStyleFormatGroup[]): void => {
    const raw = wb.getCellStyleXf(xfId);
    if (!raw) return;
    resolved.set(key, {
      xfId,
      raw,
      groups,
      projected: cellFormatFromXf(wb, raw, wb.workbookDefaultFont),
    });
  };
  for (const [key, xfId] of ids) {
    const registration = registrations.get(key);
    const style = registration ? undefined : getCellStyle(key as CellStyleId);
    const groups = registration?.groups ?? (style ? cellStyleGroups(style) : []);
    add(key, xfId, groups);
  }
  const normal = wb.getNamedCellStyles().find((style) => style.builtinId === 0);
  if (normal && !resolved.has('normal')) {
    const style = getCellStyle('normal');
    add('normal', normal.xfId, style ? cellStyleGroups(style) : []);
  }
  return resolved;
}

const sameNamedStyleFormat = (
  wb: WorkbookHandle,
  format: Partial<CellFormat>,
  xfId: number,
): boolean => {
  const xf = wb.getCellStyleXf(xfId);
  if (!xf) return false;
  const native = cellFormatFromXf(wb, xf, wb.workbookDefaultFont);
  return stylePayloadMatchesNative(wb, xf, native, format, cellStyleGroups({ format }));
};

const isUnrepresentedNativeBuiltin = (entry: NamedStyleCatalogEntry): boolean =>
  entry.builtinId >= 0 &&
  entry.builtinId <= 53 &&
  !CELL_STYLES.some((style) => style.builtinId === entry.builtinId);

const isImportedNativeBuiltin = (
  registration: NamedStyleRegistration,
  entry: NamedStyleCatalogEntry,
): boolean =>
  registration.builtinId === null &&
  registration.name === entry.name &&
  isUnrepresentedNativeBuiltin(entry);

const sameImportedNativeBuiltinFormat = (
  wb: WorkbookHandle,
  registration: NamedStyleRegistration,
  entry: NamedStyleCatalogEntry,
): boolean => {
  if (!isImportedNativeBuiltin(registration, entry)) return false;
  const xf = wb.getCellStyleXf(entry.xfId);
  if (!xf) return false;
  const native = cellFormatFromXf(wb, xf, wb.workbookDefaultFont);
  return stylePayloadMatches(native, registration.format, registration.groups);
};

const namedStyleCatalog = (
  wb: WorkbookHandle,
): {
  readonly entries: readonly NamedStyleCatalogEntry[];
  readonly byBuiltin: ReadonlyMap<number, NamedStyleCatalogEntry>;
  readonly byName: ReadonlyMap<string, NamedStyleCatalogEntry>;
} => {
  const entries: NamedStyleCatalogEntry[] =
    typeof wb.cellStyleCount === 'function' && typeof wb.getCellStyle === 'function'
      ? Array.from({ length: wb.cellStyleCount() }, (_, index) => {
          const style = wb.getCellStyle(index);
          return style ? [{ ...style, index }] : [];
        }).flat()
      : wb.getNamedCellStyles().map((style) => ({ ...style, hidden: false }));
  const byBuiltin = new Map<number, NamedStyleCatalogEntry>();
  const byName = new Map<string, NamedStyleCatalogEntry>();
  for (const entry of entries) {
    byName.set(entry.name.trim().toLowerCase(), entry);
    if (entry.builtinId >= 0) byBuiltin.set(entry.builtinId, entry);
  }
  return { entries, byBuiltin, byName };
};

const existingNamedStyle = (
  catalog: ReturnType<typeof namedStyleCatalog>,
  registration: NamedStyleRegistration,
): NamedStyleCatalogEntry | undefined => {
  const builtin =
    registration.builtinId === null ? undefined : catalog.byBuiltin.get(registration.builtinId);
  const exact = catalog.byName.get(registration.name.trim().toLowerCase());
  // Built-in identity comes from its ordinal, never from a same-name custom row.
  return registration.builtinId === null ? exact : builtin;
};

const canReuseNamedStyle = (
  wb: WorkbookHandle,
  registration: NamedStyleRegistration,
  existing: NamedStyleCatalogEntry | undefined,
): existing is NamedStyleCatalogEntry =>
  existing !== undefined &&
  (registration.builtinId !== null ||
    sameNamedStyleFormat(wb, registration.format, existing.xfId) ||
    sameImportedNativeBuiltinFormat(wb, registration, existing));

const namedStyleRecordFor = (
  registration: NamedStyleRegistration,
  xfId: number,
  existing: NamedStyleCatalogEntry | undefined,
): NamedCellStyleRecord => {
  if (existing && isImportedNativeBuiltin(registration, existing)) {
    return {
      name: existing.name,
      xfId,
      builtinId: existing.builtinId,
      iLevel: existing.iLevel,
      hidden: existing.hidden,
      customBuiltin: true,
    };
  }
  return {
    name: registration.name,
    xfId,
    builtinId: writableBuiltinId(registration.builtinId),
  };
};

/** Built-in ordinals in the OOXML catalog are accepted by formulon 0.13. */
const writableBuiltinId = (builtinId: number | null): number | null =>
  builtinId !== null && builtinId >= 0 && builtinId <= 53 ? builtinId : null;

/** Stage named-style catalog changes until all cell XF and metadata writes have
 *  succeeded. A later failed registration rolls back replacements and removes
 *  newly registered rows when the engine exposes that operation. */
export function planNamedCellStylesToEngine(
  wb: WorkbookHandle,
  state: State,
  resolveStyleXf: (format: Partial<CellFormat>) => number,
): NamedStyleSyncPlan {
  const styleXfIds = new Map<string, number>();
  const pending: PendingNamedStyleRegistration[] = [];
  if (!wb.capabilities.cellStyleMutate) {
    return { styleXfIds, commit: () => undefined };
  }

  const registrations = collectNamedStyles(state, localeForWorkbook(wb));
  // Keep hidden gallery entries too, since a same-name write can replace them.
  const catalog = namedStyleCatalog(wb);

  const existingNormal = catalog.byBuiltin.get(getCellStyle('normal')?.builtinId ?? 0);
  if (existingNormal) {
    styleXfIds.set('normal', existingNormal.xfId);
  } else if (wb.cellStyleXfCount() === 0) {
    const normalXf = resolveStyleXf({});
    if (normalXf < 0) {
      throw new Error('Strict engine sync failed: resolve Normal cell-style XF');
    }
    if (normalXf !== NORMAL_STYLE_XF_ID) {
      throw new Error('Strict engine sync failed: Normal cell-style XF is not index 0');
    }
    styleXfIds.set('normal', NORMAL_STYLE_XF_ID);
    pending.push({
      record: {
        name: 'Normal',
        xfId: NORMAL_STYLE_XF_ID,
        builtinId: writableBuiltinId(getCellStyle('normal')?.builtinId ?? null),
      },
    });
  }

  if (registrations.length === 0 && pending.length === 0) {
    return { styleXfIds, commit: () => undefined };
  }

  for (const reg of registrations) {
    const existing = existingNamedStyle(catalog, reg);
    // Native writes replace only an exact name (kept so a failed strict publication can
    // restore it); a case-insensitive match may append a row that rollback must remove.
    const replacement = catalog.entries.find((entry) => entry.name === reg.name);
    if (canReuseNamedStyle(wb, reg, existing)) {
      styleXfIds.set(reg.key, existing.xfId);
      continue;
    }
    const xfId = resolveStyleXf(reg.format);
    if (xfId < 0) {
      throw new Error(`Strict engine sync failed: resolve named style XF ${reg.name}`);
    }
    styleXfIds.set(reg.key, xfId);
    pending.push({
      record: namedStyleRecordFor(reg, xfId, replacement),
      ...(replacement
        ? {
            previous: {
              name: replacement.name,
              xfId: replacement.xfId,
              builtinId: replacement.builtinId,
              iLevel: replacement.iLevel,
              hidden: replacement.hidden,
              customBuiltin: replacement.customBuiltin,
            },
          }
        : {}),
    });
  }

  return {
    styleXfIds,
    commit(): void {
      let committed = 0;
      const applied: PendingNamedStyleRegistration[] = [];
      for (const registration of pending) {
        if (wb.setNamedCellStyleRecord(registration.record)) {
          committed += 1;
          applied.push(registration);
          continue;
        }
        const failure = new Error(
          `Strict engine sync failed: setNamedCellStyle ${registration.record.name}`,
        );
        if (committed > 0) {
          const rollbackErrors: unknown[] = [];
          for (const appliedRegistration of applied.reverse()) {
            if (!appliedRegistration.previous) {
              if (!wb.removeNamedCellStyle(appliedRegistration.record.name)) {
                rollbackErrors.push(
                  new Error(
                    `Named-style catalog rollback failed: ${appliedRegistration.record.name}`,
                  ),
                );
              }
              continue;
            }
            if (!wb.setNamedCellStyleRecord(appliedRegistration.previous)) {
              rollbackErrors.push(
                new Error(
                  `Named-style catalog rollback failed: ${appliedRegistration.previous.name}`,
                ),
              );
            }
          }
          throw new AggregateError(
            [failure, ...rollbackErrors],
            'Strict named-style publication failed after a partial catalog commit',
            { cause: failure },
          );
        }
        throw failure;
      }
    },
  };
}

/**
 * Register every named style the sheet uses as a real `<cellStyle>` +
 * `<cellStyleXfs>` pair and return the style-xf index each store-side key
 * resolved to. Callers hand that map to the cell XF writeback, which points
 * each styled cell's `xfId` at its style — the two-layer model Excel needs for
 * "edit the style, every cell using it follows".
 *
 * Returns an empty map when the engine cannot author named styles; cells then
 * keep their direct formatting and simply inherit the default style.
 */
export function syncNamedCellStylesToEngine(
  wb: WorkbookHandle,
  state: State,
  resolveStyleXf: (format: Partial<CellFormat>) => number,
): Map<string, number> {
  const resolved = new Map<string, number>();
  if (!wb.capabilities.cellStyleMutate) return resolved;
  const registrations = collectNamedStyles(state, localeForWorkbook(wb));
  const catalog = namedStyleCatalog(wb);

  // A workbook that has never carried a named style has an empty
  // `<cellStyleXfs>` table, so the first row added would land at index 0 —
  // the row every unstyled cell inherits. Seed Normal there first.
  const existingNormal = catalog.byBuiltin.get(getCellStyle('normal')?.builtinId ?? 0);
  if (existingNormal) {
    resolved.set('normal', existingNormal.xfId);
  } else if (wb.cellStyleXfCount() === 0) {
    const normalXf = resolveStyleXf({});
    if (normalXf !== NORMAL_STYLE_XF_ID) return resolved;
    if (
      wb.setNamedCellStyle('Normal', NORMAL_STYLE_XF_ID, getCellStyle('normal')?.builtinId ?? null)
    ) {
      resolved.set('normal', NORMAL_STYLE_XF_ID);
    }
  }

  for (const reg of registrations) {
    const existing = existingNamedStyle(catalog, reg);
    if (canReuseNamedStyle(wb, reg, existing)) {
      resolved.set(reg.key, existing.xfId);
      continue;
    }
    const xfId = resolveStyleXf(reg.format);
    if (xfId < 0) continue;
    if (!wb.setNamedCellStyleRecord(namedStyleRecordFor(reg, xfId, existing))) continue;
    resolved.set(reg.key, xfId);
  }
  return resolved;
}

/** Map each `<cellStyleXfs>` row back to the gallery key a hydrated cell
 *  should carry. Built-in ordinal 0 ("Normal") is left out: an unstyled cell
 *  points at it and must not come back tagged with a style. */
export function cellStyleKeysByXfId(wb: WorkbookHandle): Map<number, string> {
  const out = new Map<number, string>();
  if (!wb.capabilities.cellStyles) return out;
  for (const style of wb.getNamedCellStyles()) {
    const builtin = CELL_STYLES.find((entry) => entry.builtinId === style.builtinId);
    if (
      style.xfId === NORMAL_STYLE_XF_ID ||
      style.builtinId === getCellStyle('normal')?.builtinId
    ) {
      continue;
    }
    if (!out.has(style.xfId)) {
      out.set(style.xfId, builtin?.id ?? cellStyleKeyForOoxmlName(style.name));
    }
  }
  return out;
}

/** Named styles the workbook carries that the built-in gallery does not know,
 *  as custom styles. Cells hydrate against these by name, so the store must
 *  hold them for a later flush to keep each cell's `xfId`. */
export function importedCustomCellStyles(wb: WorkbookHandle): CustomCellStyle[] {
  const out: CustomCellStyle[] = [];
  if (!wb.capabilities.cellStyles) return out;
  const normalId = getCellStyle('normal')?.builtinId;
  for (const style of wb.getNamedCellStyles()) {
    if (style.xfId === NORMAL_STYLE_XF_ID || style.builtinId === normalId) continue;
    if (CELL_STYLES.some((entry) => entry.builtinId === style.builtinId)) continue;
    if (cellStyleKeyForOoxmlName(style.name) !== style.name) continue;
    const xf = wb.getCellStyleXf(style.xfId);
    if (!xf) continue;
    out.push({
      id: customCellStyleId(style.name),
      label: style.name,
      format: cellFormatFromXf(wb, xf, wb.workbookDefaultFont),
    });
  }
  return out;
}
