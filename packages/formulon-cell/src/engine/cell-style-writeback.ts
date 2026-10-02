import {
  CELL_STYLES,
  type CellStyleFallbackProfile,
  type CellStyleFormatGroup,
  type CellStyleId,
  cellStyleFallbackFormat,
  cellStyleGroups,
  formatForCellStyleGroups,
  getCellStyle,
  stylePayloadMatches,
  stylePayloadMatchesNative,
} from '../commands/cell-styles.js';
import type { CellFormat, State } from '../store/store.js';
import { cellFormatFromXf } from './cell-format-sync.js';
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
  readonly name: string;
  readonly xfId: number;
  readonly builtinId: number | null;
  readonly previous?: {
    readonly name: string;
    readonly xfId: number;
    readonly builtinId: number | null;
  };
}

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

/** Every named style referenced by at least one cell on the sheet set, in a
 *  stable order. Styles nothing references are not written — Excel keeps the
 *  full built-in gallery in every file, but emitting every unused entry would
 *  bloat a workbook the user never styled. */
export function collectNamedStyles(state: State): NamedStyleRegistration[] {
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
        const fallback = cellStyleFallbackFormat(builtin.id, candidateProfile);
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
          ? cellStyleFallbackFormat(builtin.id, profile)
          : cellStyleFallbackFormat(builtin.id, 'default'),
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
    collectNamedStyles(state).map((registration) => [registration.key, registration]),
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

/** Built-in ordinals above 47 are useful when reading an OOXML catalog, but
 *  the current engine rejects them when authoring a cellStyle record. */
const writableBuiltinId = (builtinId: number | null): number | null =>
  builtinId !== null && builtinId >= 0 && builtinId <= 47 ? builtinId : null;

/** Stage named-style catalog changes until all cell XF and metadata writes have
 *  succeeded. The plan intentionally has no rollback claim for catalog rows:
 *  the current engine can replace a row but cannot delete a newly registered
 *  one. A later failed registration is surfaced as an AggregateError. */
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

  const registrations = collectNamedStyles(state);
  const named = wb.getNamedCellStyles();
  const byBuiltin = new Map<number, (typeof named)[number]>();
  const byName = new Map<string, (typeof named)[number]>();
  for (const entry of named) {
    byName.set(entry.name.trim().toLowerCase(), entry);
    if (entry.builtinId >= 0) byBuiltin.set(entry.builtinId, entry);
  }

  const existingNormal = byBuiltin.get(getCellStyle('normal')?.builtinId ?? 0);
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
      name: 'Normal',
      xfId: NORMAL_STYLE_XF_ID,
      builtinId: writableBuiltinId(getCellStyle('normal')?.builtinId ?? null),
    });
  }

  if (registrations.length === 0) {
    return { styleXfIds, commit: () => undefined };
  }

  for (const reg of registrations) {
    const builtin = reg.builtinId === null ? undefined : byBuiltin.get(reg.builtinId);
    const exact = byName.get(reg.name.trim().toLowerCase());
    // Built-in identity comes from its ordinal, never from a same-name
    // custom/imported row. Keep an exact row only as the replacement target so
    // a failed strict publication can restore it.
    const existing = reg.builtinId === null ? exact : builtin;
    const replacement = builtin ?? exact;
    if (
      existing &&
      (reg.builtinId !== null || sameNamedStyleFormat(wb, reg.format, existing.xfId))
    ) {
      styleXfIds.set(reg.key, existing.xfId);
      continue;
    }
    const xfId = resolveStyleXf(reg.format);
    if (xfId < 0) {
      throw new Error(`Strict engine sync failed: resolve named style XF ${reg.name}`);
    }
    styleXfIds.set(reg.key, xfId);
    pending.push({
      name: reg.name,
      xfId,
      builtinId: writableBuiltinId(reg.builtinId),
      ...(replacement
        ? {
            previous: {
              name: replacement.name,
              xfId: replacement.xfId,
              builtinId: writableBuiltinId(replacement.builtinId),
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
        if (wb.setNamedCellStyle(registration.name, registration.xfId, registration.builtinId)) {
          committed += 1;
          applied.push(registration);
          continue;
        }
        const failure = new Error(
          `Strict engine sync failed: setNamedCellStyle ${registration.name}`,
        );
        if (committed > 0) {
          const rollbackErrors: unknown[] = [];
          let unrestorable = false;
          for (const appliedRegistration of applied.reverse()) {
            if (!appliedRegistration.previous) {
              unrestorable = true;
              continue;
            }
            if (
              !wb.setNamedCellStyle(
                appliedRegistration.previous.name,
                appliedRegistration.previous.xfId,
                appliedRegistration.previous.builtinId,
              )
            ) {
              rollbackErrors.push(
                new Error(
                  `Named-style catalog rollback failed: ${appliedRegistration.previous.name}`,
                ),
              );
            }
          }
          throw new AggregateError(
            [
              failure,
              ...rollbackErrors,
              ...(unrestorable
                ? [new Error('Named-style catalog rollback is unavailable for new entries')]
                : []),
            ],
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
  const registrations = collectNamedStyles(state);
  const named = wb.getNamedCellStyles();
  const byBuiltin = new Map<number, (typeof named)[number]>();
  const byName = new Map<string, (typeof named)[number]>();
  for (const entry of named) {
    byName.set(entry.name.trim().toLowerCase(), entry);
    if (entry.builtinId >= 0) byBuiltin.set(entry.builtinId, entry);
  }

  // A workbook that has never carried a named style has an empty
  // `<cellStyleXfs>` table, so the first row added would land at index 0 —
  // the row every unstyled cell inherits. Seed Normal there first.
  const existingNormal = byBuiltin.get(getCellStyle('normal')?.builtinId ?? 0);
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
    const builtin = reg.builtinId === null ? undefined : byBuiltin.get(reg.builtinId);
    const exact = byName.get(reg.name.trim().toLowerCase());
    // A same-name custom row is not the built-in row when the ordinal is
    // absent; publish the desired built-in identity instead of reusing it.
    const existing = reg.builtinId === null ? exact : builtin;
    if (
      existing &&
      (reg.builtinId !== null || sameNamedStyleFormat(wb, reg.format, existing.xfId))
    ) {
      resolved.set(reg.key, existing.xfId);
      continue;
    }
    const xfId = resolveStyleXf(reg.format);
    if (xfId < 0) continue;
    if (!wb.setNamedCellStyle(reg.name, xfId, writableBuiltinId(reg.builtinId))) continue;
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
