import { addrKey } from '../engine/address.js';
import { cellFormatFromXf } from '../engine/cell-format-sync.js';
import type { CellXf, Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import {
  type CellFormat,
  type CustomCellStyle,
  mutators,
  type SpreadsheetStore,
  type State,
} from '../store/store.js';
import {
  CELL_STYLES,
  type CellStyleDef,
  type CellStyleFormatGroup,
  type CellStyleId,
  EXCEL365_MAC_STYLE_OVERRIDES,
} from './cell-style-catalog.js';
import { recordDialogFormatChange } from './dialog-format-history.js';
import { applyFormatPatch, applySelectionFormatAction, planSelectionFormat } from './format.js';
import type { History } from './history.js';
import { recordFormatChange, recordFormatChangeWithRepeat } from './history.js';
import type { InteractionOrigin } from './interaction-policy.js';
import { mergeAnchorOf } from './merge.js';

export type {
  CellStyleDef,
  CellStyleFormatGroup,
  CellStyleGroupDef,
  CellStyleGroupId,
  CellStyleId,
} from './cell-style-catalog.js';
export { CELL_STYLE_GROUPS, CELL_STYLES } from './cell-style-catalog.js';

export interface MergeCellStylesResult {
  imported: number;
  skipped: number;
}

export interface CellStyleIncludeOptions {
  number?: boolean;
  alignment?: boolean;
  font?: boolean;
  border?: boolean;
  fill?: boolean;
  protection?: boolean;
}

export interface CreateCellStyleOptions {
  include?: CellStyleIncludeOptions;
}

const STYLE_BY_ID = new Map<CellStyleId, CellStyleDef>(CELL_STYLES.map((s) => [s.id, s]));
const CUSTOM_STYLE_PREFIX = 'custom:';
const BUILT_IN_STYLE_NAMES = new Set(
  CELL_STYLES.flatMap((style) => [style.id.toLowerCase(), style.label.toLowerCase()]),
);

export type CellStyleFallbackProfile = 'default' | 'excel365Mac';

export const cellStyleFallbackProfileForPlatform = (
  platform: string | undefined,
): CellStyleFallbackProfile => (platform?.toLowerCase() === 'mac' ? 'excel365Mac' : 'default');

/** Return a fresh portable/profile-aware format for a built-in style. */
export function cellStyleFallbackFormat(
  id: CellStyleId,
  profile: CellStyleFallbackProfile = 'default',
): Partial<CellFormat> {
  const style = STYLE_BY_ID.get(id);
  const format = { ...(style?.format ?? {}) };
  if (profile === 'excel365Mac') Object.assign(format, EXCEL365_MAC_STYLE_OVERRIDES[id] ?? {});
  return format;
}

const CELL_STYLE_GROUP_FIELDS: Record<CellStyleFormatGroup, readonly (keyof CellFormat)[]> = {
  number: ['numFmt'],
  alignment: [
    'align',
    'vAlign',
    'wrap',
    'justifyLastLine',
    'shrinkToFit',
    'indent',
    'rotation',
    'textDirection',
  ],
  font: [
    'bold',
    'italic',
    'underline',
    'strike',
    'fontVertAlign',
    'color',
    'fontFamily',
    'fontSize',
  ],
  border: ['borders'],
  fill: ['fill', 'fillPattern', 'fillPatternColor'],
  protection: ['locked', 'formulaHidden'],
};

const ALL_BORDER_SIDES: readonly (keyof NonNullable<CellFormat['borders']>)[] = [
  'top',
  'right',
  'bottom',
  'left',
  'diagonalDown',
  'diagonalUp',
];

const inferCellStyleGroups = (format: Partial<CellFormat>): readonly CellStyleFormatGroup[] => {
  const groups: CellStyleFormatGroup[] = [];
  for (const [group, fields] of Object.entries(CELL_STYLE_GROUP_FIELDS) as [
    CellStyleFormatGroup,
    readonly (keyof CellFormat)[],
  ][]) {
    if (fields.some((field) => Object.hasOwn(format, field))) groups.push(group);
  }
  return groups;
};

export function cellStyleGroups(
  style: Pick<CellStyleDef, 'format' | 'includedGroups'> | Pick<CustomCellStyle, 'format'>,
): readonly CellStyleFormatGroup[] {
  return 'includedGroups' in style && style.includedGroups
    ? style.includedGroups
    : inferCellStyleGroups(style.format);
}

export const formatForCellStyleGroups = (
  format: Partial<CellFormat>,
  groups: readonly CellStyleFormatGroup[],
): Partial<CellFormat> => {
  const included = new Set(groups);
  const patch: Partial<CellFormat> = {};
  for (const group of groups) {
    for (const field of CELL_STYLE_GROUP_FIELDS[group]) {
      if (field === 'borders') {
        patch.borders = Object.fromEntries(
          ALL_BORDER_SIDES.map((side) => [side, undefined]),
        ) as CellFormat['borders'];
      } else {
        patch[field] = undefined;
      }
    }
  }
  for (const [key, value] of Object.entries(format) as [keyof CellFormat, unknown][]) {
    const group = (
      Object.entries(CELL_STYLE_GROUP_FIELDS) as [
        CellStyleFormatGroup,
        readonly (keyof CellFormat)[],
      ][]
    ).find(([, fields]) => fields.includes(key));
    if (!group || !included.has(group[0])) continue;
    if (key === 'borders') {
      patch.borders = {
        ...(patch.borders ?? {}),
        ...(value as CellFormat['borders']),
      };
    } else {
      (patch as Record<string, unknown>)[key] = value;
    }
  }
  return patch;
};

const formatForGroups = formatForCellStyleGroups;

export function getCellStyle(id: CellStyleId): CellStyleDef | undefined {
  return STYLE_BY_ID.get(id);
}

const resolvedCellStyle = (
  state: { format: { customCellStyles?: readonly CustomCellStyle[] } },
  id: string,
): {
  id: string;
  label: string;
  format: Partial<CellFormat>;
  groups: readonly CellStyleFormatGroup[];
  builtinId?: number;
} | null => {
  const builtin =
    STYLE_BY_ID.get(id as CellStyleId) ?? CELL_STYLES.find((style) => style.label === id);
  if (builtin) {
    return {
      id: builtin.id,
      label: builtin.label,
      format: builtin.format,
      groups: cellStyleGroups(builtin),
      ...(builtin.builtinId === undefined ? {} : { builtinId: builtin.builtinId }),
    };
  }
  const custom =
    state.format.customCellStyles?.find((style) => style.id === id) ??
    state.format.customCellStyles?.find((style) => style.label === id);
  if (!custom) return null;
  return {
    id: custom.id,
    label: custom.label,
    format: custom.format,
    groups: cellStyleGroups(custom),
  };
};

/** Resolve the style represented by the active cell, following a merge to its anchor. */
export function activeCellStyleId(state: {
  selection: State['selection'];
  merges: State['merges'];
  format: State['format'];
}): string {
  const active = mergeAnchorOf(state as State, state.selection.active);
  const raw = state.format.formats.get(addrKey(active))?.cellStyle;
  if (!raw) return 'normal';
  const resolved = resolvedCellStyle(state, raw);
  return resolved?.id ?? raw;
}

export interface CellStyleCommandContext {
  origin?: InteractionOrigin;
  commandId?: string;
  getWorkbook?: () => WorkbookHandle | null;
  getFallbackProfile?: () => CellStyleFallbackProfile;
}

const sameStyleValue = (left: unknown, right: unknown): boolean => {
  if (Object.is(left, right)) return true;
  if (left === null || right === null || typeof left !== 'object' || typeof right !== 'object') {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    return (
      Array.isArray(left) &&
      Array.isArray(right) &&
      left.length === right.length &&
      left.every((item, index) => sameStyleValue(item, right[index]))
    );
  }
  const leftRecord = left as Record<string, unknown>;
  const rightRecord = right as Record<string, unknown>;
  const keys = Object.keys(leftRecord);
  return (
    keys.length === Object.keys(rightRecord).length &&
    keys.every(
      (key) => Object.hasOwn(rightRecord, key) && sameStyleValue(leftRecord[key], rightRecord[key]),
    )
  );
};

const normalizedBorderSide = (
  side: CellFormat['borders'] extends infer B ? (B extends object ? B[keyof B] : never) : never,
): unknown => {
  // The store accepts both omitted and explicit false for an absent border.
  // Treat them as the same semantic default while retaining real styles and
  // their color selectors below.
  if (side === undefined || side === false) return undefined;
  if (side === true) return { style: 'thin' };
  if (typeof side !== 'object' || side === null) return side;
  const record = side as { style?: unknown; color?: unknown };
  return { style: record.style, ...(record.color === undefined ? {} : { color: record.color }) };
};

const normalizedStyleField = (field: keyof CellFormat, value: unknown): unknown => {
  switch (field) {
    case 'bold':
    case 'italic':
    case 'strike':
    case 'wrap':
    case 'justifyLastLine':
    case 'shrinkToFit':
    case 'formulaHidden':
      return value === true;
    case 'underline':
      return value === true ? 'single' : value === false ? undefined : value;
    case 'align':
      return value ?? 'general';
    case 'vAlign':
      return value ?? 'bottom';
    case 'indent':
    case 'rotation':
      return value ?? 0;
    case 'textDirection':
      return value ?? 'context';
    case 'locked':
      return value === undefined ? true : value === true;
    case 'numFmt':
      return (value as CellFormat['numFmt'] | undefined)?.kind === 'general' ? undefined : value;
    default:
      return value;
  }
};

export const stylePayloadMatches = (
  native: Partial<CellFormat>,
  stored: Partial<CellFormat>,
  groups: readonly CellStyleFormatGroup[],
): boolean => {
  const expected = formatForGroups(stored, groups);
  const actual = formatForGroups(native, groups);
  for (const group of groups) {
    for (const field of CELL_STYLE_GROUP_FIELDS[group]) {
      if (field === 'borders') {
        for (const side of ALL_BORDER_SIDES) {
          if (
            !sameStyleValue(
              normalizedBorderSide(actual.borders?.[side]),
              normalizedBorderSide(expected.borders?.[side]),
            )
          ) {
            return false;
          }
        }
      } else if (
        !sameStyleValue(
          normalizedStyleField(field, actual[field]),
          normalizedStyleField(field, expected[field]),
        )
      ) {
        return false;
      }
    }
  }
  return true;
};

export const cellStyleGroupMatches = (
  left: Partial<CellFormat>,
  right: Partial<CellFormat>,
  group: CellStyleFormatGroup,
): boolean => stylePayloadMatches(left, right, [group]);

/** A literal store color must not reuse a native theme/indexed/auto selector
 * merely because the engine exposes the same fallback RGB. */
export const stylePayloadMatchesNative = (
  workbook: WorkbookHandle,
  nativeXf: CellXf,
  native: Partial<CellFormat>,
  stored: Partial<CellFormat>,
  groups: readonly CellStyleFormatGroup[],
): boolean => {
  if (!stylePayloadMatches(native, stored, groups)) return false;
  const literalOrFallback = (kind: number): boolean => kind === 0 || kind === 1;
  if (groups.includes('font') && Object.hasOwn(stored, 'color')) {
    const font = workbook.getFontRecord(nativeXf.fontIndex);
    if (font?.color && !literalOrFallback(font.color.kind)) return false;
  }
  if (groups.includes('fill')) {
    const fill = workbook.getFillRecord(nativeXf.fillIndex);
    if (fill?.fg && Object.hasOwn(stored, 'fill') && !literalOrFallback(fill.fg.kind)) return false;
    if (fill?.fg && Object.hasOwn(stored, 'fillPatternColor') && !literalOrFallback(fill.fg.kind)) {
      return false;
    }
  }
  if (groups.includes('border') && stored.borders) {
    const border = workbook.getBorderRecord(nativeXf.borderIndex);
    if (border) {
      const nativeSides = {
        top: border.top,
        right: border.right,
        bottom: border.bottom,
        left: border.left,
        diagonalDown: border.diagonal,
        diagonalUp: border.diagonal,
      };
      for (const side of ALL_BORDER_SIDES) {
        const value = stored.borders[side];
        const nativeSide = nativeSides[side];
        if (value && typeof value === 'object' && Object.hasOwn(value, 'color') && nativeSide) {
          if (nativeSide.color && !literalOrFallback(nativeSide.color.kind)) return false;
        }
      }
    }
  }
  return true;
};

const resolveWorkbookStyleFormat = (
  workbook: WorkbookHandle | null,
  style: ReturnType<typeof resolvedCellStyle>,
  profile: CellStyleFallbackProfile = 'default',
): Partial<CellFormat> => {
  if (!style) return {};
  const fallback =
    style.builtinId === undefined
      ? { ...style.format }
      : cellStyleFallbackFormat(style.id as CellStyleId, profile);
  if (!workbook?.capabilities.cellStyles) return fallback;
  const named = workbook.getNamedCellStyles();
  const native =
    style.builtinId === undefined
      ? named.find((entry) => entry.name.trim().toLowerCase() === style.label.trim().toLowerCase())
      : named.find((entry) => entry.builtinId === style.builtinId);
  if (!native) return fallback;
  const xf = workbook.getCellStyleXf(native.xfId);
  if (!xf) return fallback;
  const nativeFormat = cellFormatFromXf(workbook, xf, workbook.workbookDefaultFont);
  if (
    style.builtinId === undefined &&
    !stylePayloadMatchesNative(workbook, xf, nativeFormat, style.format, style.groups)
  ) {
    return fallback;
  }
  return nativeFormat;
};

/** Resolve a built-in style for previews and application, preferring the
 * workbook's raw named-style XF over the selected portable fallback. */
export function resolveCellStyleFormat(
  id: CellStyleId,
  workbook: WorkbookHandle | null,
  profile: CellStyleFallbackProfile = 'default',
): Partial<CellFormat> {
  const style = getCellStyle(id);
  if (!style) return {};
  return resolveWorkbookStyleFormat(
    workbook,
    {
      id: style.id,
      label: style.label,
      format: style.format,
      groups: cellStyleGroups(style),
      ...(style.builtinId === undefined ? {} : { builtinId: style.builtinId }),
    },
    profile,
  );
}

const applyCellStyleToRange = (
  store: SpreadsheetStore,
  history: History | null,
  range: Range,
  patch: Partial<CellFormat>,
  repeat: () => void,
): boolean => {
  const state = store.getState();
  const scopedState: State = {
    ...state,
    selection: {
      ...state.selection,
      range: { ...range },
      extraRanges: [],
    },
  };
  const plan = planSelectionFormat(scopedState);
  if (!plan || plan.ranges.length === 0) return false;
  return recordDialogFormatChange({
    history,
    store,
    workbook: null,
    sheet: range.sheet,
    targets: plan.cells,
    pendingBefore: state.ui.pendingFormat,
    mutate: () =>
      applySelectionFormatAction(
        scopedState,
        store,
        { patch },
        {
          allowPending: false,
          origin: 'instanceApi',
          commandId: 'cellStyles',
        },
      ),
    repeat,
  });
};

/** Apply an existing built-in or registered custom style to the full selection union. */
export function applyCellStyleToSelection(
  store: SpreadsheetStore,
  history: History | null,
  id: string,
  context: CellStyleCommandContext = {},
): boolean {
  const state = store.getState();
  const style = resolvedCellStyle(state, id);
  if (!style) return false;
  const plan = planSelectionFormat(state);
  if (!plan || plan.ranges.length === 0) return false;
  const workbook = context.getWorkbook?.() ?? null;
  const profile = context.getFallbackProfile?.() ?? 'default';
  const payload = resolveWorkbookStyleFormat(workbook, style, profile);
  const patch = formatForGroups(payload, style.groups);
  patch.cellStyle =
    style.id === 'normal' ? undefined : style.builtinId !== undefined ? style.id : style.label;
  const commandId = context.commandId ?? 'cellStyles';
  const origin = context.origin ?? 'instanceApi';
  return recordDialogFormatChange({
    history,
    store,
    workbook,
    sheet: plan.ranges[0]?.sheet ?? state.selection.range.sheet,
    targets: plan.cells,
    pendingBefore: state.ui.pendingFormat,
    mutate: () =>
      applySelectionFormatAction(
        store.getState(),
        store,
        { patch },
        {
          allowPending: false,
          origin,
          commandId,
        },
      ),
    repeat: () => applyCellStyleToSelection(store, history, style.id, context),
  });
}

export function customCellStyleId(name: string): string {
  return `${CUSTOM_STYLE_PREFIX}${name.trim()}`;
}

export function listCustomCellStyles(state: {
  format: { customCellStyles?: readonly CustomCellStyle[] };
}): readonly CustomCellStyle[] {
  return state.format.customCellStyles ?? [];
}

export function customCellStyleById(
  state: { format: { customCellStyles?: readonly CustomCellStyle[] } },
  id: string,
): CustomCellStyle | null {
  return (state.format.customCellStyles ?? []).find((style) => style.id === id) ?? null;
}

const DEFAULT_CELL_STYLE_INCLUDE: Required<CellStyleIncludeOptions> = {
  number: true,
  alignment: true,
  font: true,
  border: true,
  fill: true,
  protection: true,
};

export function filterCellStyleFormat(
  format: Partial<CellFormat>,
  include: CellStyleIncludeOptions = DEFAULT_CELL_STYLE_INCLUDE,
): Partial<CellFormat> {
  const opts = { ...DEFAULT_CELL_STYLE_INCLUDE, ...include };
  const filtered: Partial<CellFormat> = {};
  const copy = <K extends keyof CellFormat>(key: K): void => {
    if (Object.hasOwn(format, key)) {
      filtered[key] = format[key];
    }
  };
  if (opts.number) {
    copy('numFmt');
  }
  if (opts.alignment) {
    copy('align');
    copy('vAlign');
    copy('wrap');
    copy('justifyLastLine');
    copy('shrinkToFit');
    copy('indent');
    copy('rotation');
    copy('textDirection');
  }
  if (opts.font) {
    copy('bold');
    copy('italic');
    copy('underline');
    copy('strike');
    copy('fontVertAlign');
    copy('color');
    copy('fontFamily');
    copy('fontSize');
  }
  if (opts.border) {
    copy('borders');
  }
  if (opts.fill) {
    copy('fill');
    copy('fillPattern');
    copy('fillPatternColor');
  }
  if (opts.protection) {
    copy('locked');
    copy('formulaHidden');
  }
  return filtered;
}

/** Apply a named style to `range`. Wraps the format mutation in a single
 *  history entry so Cmd+Z reverts the whole gallery click. The `normal`
 *  style is a clear — it strips every format field instead of merging. */
export function applyCellStyle(
  store: SpreadsheetStore,
  history: History | null,
  range: Range,
  id: CellStyleId,
): void {
  const def = STYLE_BY_ID.get(id);
  if (!def) return;
  const patch = formatForGroups(def.format, cellStyleGroups(def));
  patch.cellStyle = id === 'normal' ? undefined : id;
  applyCellStyleToRange(store, history, range, patch, () => {
    applyCellStyle(store, history, store.getState().selection.range, id);
  });
}

export function applyCellStyleByName(
  store: SpreadsheetStore,
  history: History | null,
  range: Range,
  id: string,
): boolean {
  if (STYLE_BY_ID.has(id as CellStyleId)) {
    const builtin = STYLE_BY_ID.get(id as CellStyleId);
    if (!builtin) return false;
    const patch = formatForGroups(builtin.format, cellStyleGroups(builtin));
    patch.cellStyle = builtin.id === 'normal' ? undefined : builtin.id;
    return applyCellStyleToRange(store, history, range, patch, () => {
      applyCellStyle(store, history, store.getState().selection.range, builtin.id);
    });
  }
  const custom = customCellStyleById(store.getState(), id);
  if (!custom) return false;
  const patch = formatForGroups(custom.format, cellStyleGroups(custom));
  patch.cellStyle = custom.label;
  return applyCellStyleToRange(store, history, range, patch, () => {
    applyCellStyleByName(store, history, store.getState().selection.range, id);
  });
}

/** Create an ad-hoc named style from the active cell's current formatting and
 *  apply it to `range`. This mirrors Excel's "New Cell Style..." default of
 *  starting from the selected cell while keeping the implementation session
 *  scoped until a full OOXML style registry is available. */
export function createCellStyleFromActiveFormat(
  store: SpreadsheetStore,
  history: History | null,
  range: Range,
  name: string,
  options: CreateCellStyleOptions = {},
): boolean {
  const styleName = name.trim();
  if (!styleName) return false;
  let applied = false;
  recordFormatChangeWithRepeat(
    history,
    store,
    () => {
      const state = store.getState();
      const { cellStyle: _cellStyle, ...activeFormat } =
        state.format.formats.get(addrKey(state.selection.active)) ?? {};
      const styleFormat = filterCellStyleFormat(activeFormat, options.include);
      const patch: Partial<CellFormat> = { ...styleFormat, cellStyle: styleName };
      mutators.upsertCustomCellStyle(store, {
        id: customCellStyleId(styleName),
        label: styleName,
        format: styleFormat,
      });
      applied = applyFormatPatch(state, store, range, patch);
    },
    // Repeating "New Cell Style" applies the style it just created rather
    // than defining another one from whatever is now selected.
    () => {
      applyCellStyleByName(
        store,
        history,
        store.getState().selection.range,
        customCellStyleId(styleName),
      );
    },
  );
  return applied;
}

export function mergeCellStylesFromWorkbook(
  store: SpreadsheetStore,
  history: History | null,
  workbook: WorkbookHandle,
): MergeCellStylesResult {
  const namedStyles = workbook.getNamedCellStyles();
  const merged: CustomCellStyle[] = [];
  let skipped = 0;
  for (const style of namedStyles) {
    const label = style.name.trim();
    if (!label) {
      skipped += 1;
      continue;
    }
    if (
      BUILT_IN_STYLE_NAMES.has(label.toLowerCase()) ||
      CELL_STYLES.some((builtin) => builtin.builtinId === style.builtinId)
    ) {
      skipped += 1;
      continue;
    }
    const xf = workbook.getCellStyleXf(style.xfId);
    if (!xf) {
      skipped += 1;
      continue;
    }
    merged.push({
      id: customCellStyleId(label),
      label,
      format: cellFormatFromXf(workbook, xf),
    });
  }
  if (merged.length === 0) return { imported: 0, skipped };
  recordFormatChange(history, store, () => {
    for (const style of merged) {
      mutators.upsertCustomCellStyle(store, style);
    }
  });
  return { imported: merged.length, skipped };
}
