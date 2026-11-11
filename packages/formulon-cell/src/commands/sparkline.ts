import type { Color } from '@libraz/formulon';
import { BlanksAs, ColorKind, SparklineAxisType, SparklineType } from '@libraz/formulon';
import { addrKey, parseAddrKey, quoteSheetName } from '../engine/address.js';
import { argbFromCssColor } from '../engine/theme-color.js';
import type { Addr, Range } from '../engine/types.js';
import { DEFAULT_NEGATIVE_COLOR, DEFAULT_SPARKLINE_COLOR } from '../render/painters/controls.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import { cloneSparklineSettings, withSparklineMembers } from '../store/sparkline-slice.js';
import {
  mutators,
  type Sparkline,
  type SparklineGroupSettings,
  type SparklineKind,
  type SparklineSlice,
  type SpreadsheetStore,
} from '../store/store.js';
import { authorizedWrite } from './authorized-write.js';
import { adjustInterval } from './axis-shift.js';
import { adjustFormulaForRowColEdit } from './formula-refs.js';
import { History } from './history.js';
import { type ObjectActionContext, objectWrite } from './object-actions.js';
import { isCellWritable, isSheetProtected } from './protection.js';
import { recordSparklineChange } from './slice-history.js';

export interface SparklineEntry {
  addr: Addr;
  spec: Sparkline;
}

/** A sparkline to create: its host cell and data range. */
export interface SparklineInsert {
  addr: Addr;
  source: string;
}

type SparklineState = { sparkline: SparklineSlice };

const TYPE_OF_KIND: Record<SparklineKind, SparklineType> = {
  line: SparklineType.Line,
  column: SparklineType.Column,
  'win-loss': SparklineType.Stacked,
};

const KIND_OF_TYPE: Record<SparklineType, SparklineKind> = {
  [SparklineType.Line]: 'line',
  [SparklineType.Column]: 'column',
  [SparklineType.Stacked]: 'win-loss',
};

export const sparklineTypeOf = (kind: SparklineKind): SparklineType => TYPE_OF_KIND[kind];

export const sparklineKindOf = (type: SparklineType): SparklineKind => KIND_OF_TYPE[type] ?? 'line';

/** An opaque RGB engine colour of `css` (`#rgb` / `#rrggbb`). */
export const rgbColorSpec = (css: string): Color => ({
  kind: ColorKind.Rgb,
  argb: argbFromCssColor(css) ?? 0xff000000,
  theme: 0,
  tint: 0,
  indexed: 0,
});

/** Settings of a new group of `type`: the series and negative colours of the overlay,
 *  negatives shown for column and win/loss. */
export function defaultSparklineGroupSettings(type: SparklineType): SparklineGroupSettings {
  const mark = rgbColorSpec(DEFAULT_NEGATIVE_COLOR);
  return {
    type,
    colorSeries: rgbColorSpec(DEFAULT_SPARKLINE_COLOR),
    colorNegative: mark,
    colorAxis: rgbColorSpec('#000000'),
    colorMarkers: mark,
    colorFirst: mark,
    colorLast: mark,
    colorHigh: mark,
    colorLow: mark,
    markers: false,
    high: false,
    low: false,
    first: false,
    last: false,
    negative: type !== SparklineType.Line,
    displayXAxis: false,
    displayHidden: false,
    minAxisType: SparklineAxisType.Individual,
    maxAxisType: SparklineAxisType.Individual,
    manualMin: null,
    manualMax: null,
    dateAxis: false,
    dateAxisRange: null,
    lineWidthPt: 0.75,
    emptyCellsAs: BlanksAs.Gap,
  };
}

export function listSparklines(state: {
  sparkline: { sparklines: ReadonlyMap<string, Sparkline> };
}): readonly SparklineEntry[] {
  const out: SparklineEntry[] = [];
  for (const [key, spec] of state.sparkline.sparklines) {
    const addr = parseAddrKey(key);
    if (addr) out.push({ addr, spec: { ...spec } });
  }
  return out.sort(
    (a, b) => a.addr.sheet - b.addr.sheet || a.addr.row - b.addr.row || a.addr.col - b.addr.col,
  );
}

export function sparklineAt(
  state: { sparkline: { sparklines: ReadonlyMap<string, Sparkline> } },
  addr: Addr,
): Sparkline | null {
  const spec = state.sparkline.sparklines.get(addrKey(addr));
  return spec ? { ...spec } : null;
}

/** The group of the sparkline at `addr`, with a copy of its settings; null when there is none. */
export function groupOfCell(
  state: SparklineState,
  addr: Addr,
): { id: string; settings: SparklineGroupSettings } | null {
  const member = state.sparkline.sparklines.get(addrKey(addr));
  const settings = member && state.sparkline.groups.get(member.group);
  return member && settings
    ? { id: member.group, settings: cloneSparklineSettings(settings) }
    : null;
}

/** The members of group `groupId` in row-major order. */
export function membersOfGroup(state: SparklineState, groupId: string): readonly SparklineEntry[] {
  return listSparklines(state).filter(({ spec }) => spec.group === groupId);
}

let groupSeq = 0;

/** A group id no other sparkline of this session shares. */
export const newSparklineGroupId = (): string => {
  groupSeq += 1;
  return `sparkline-group-${groupSeq}`;
};

/** Patch the settings of group `groupId`; false when the group does not exist. Store-level:
 *  callers wrap it in `recordSparklineChange`. */
export function updateSparklineGroup(
  store: SpreadsheetStore,
  groupId: string,
  patch: Partial<SparklineGroupSettings>,
): boolean {
  const current = store.getState().sparkline.groups.get(groupId);
  if (!current) return false;
  mutators.setSparklineGroup(store, groupId, { ...current, ...structuredClone(patch) });
  return true;
}

/** Merge the groups of the sparklines at `addrs` (those on `active`'s sheet) into one new
 *  group that takes the settings of `active`'s group, or of the first cell's. Returns the new
 *  group id, or null when no cell hosts a sparkline. */
export function groupSparklines(
  store: SpreadsheetStore,
  addrs: readonly Addr[],
  active: Addr | undefined = addrs[0],
): string | null {
  const state = store.getState();
  const sheet = (active ?? addrs[0])?.sheet;
  if (sheet === undefined) return null;
  const touched = new Set<string>();
  for (const addr of addrs) {
    const member = addr.sheet === sheet ? state.sparkline.sparklines.get(addrKey(addr)) : undefined;
    if (member) touched.add(member.group);
  }
  if (touched.size === 0) return null;
  const activeGroup = active ? state.sparkline.sparklines.get(addrKey(active))?.group : undefined;
  const settings = state.sparkline.groups.get(
    activeGroup && touched.has(activeGroup) ? activeGroup : ([...touched][0] as string),
  );
  if (!settings) return null;
  const id = newSparklineGroupId();
  store.setState((s) => {
    const sparklines = new Map(
      [...s.sparkline.sparklines].map(([key, member]) => [
        key,
        touched.has(member.group) ? { ...member, group: id } : member,
      ]),
    );
    const groups = new Map(s.sparkline.groups);
    groups.set(id, cloneSparklineSettings(settings));
    return { ...s, sparkline: withSparklineMembers({ sparklines, groups }, sparklines) };
  });
  return id;
}

/** Give each member of the groups of the sparklines at `addrs` its own group with a copy of
 *  the group's settings. Returns how many sparklines were regrouped. */
export function ungroupSparklines(store: SpreadsheetStore, addrs: readonly Addr[]): number {
  const state = store.getState();
  const touched = new Set<string>();
  for (const addr of addrs) {
    const member = state.sparkline.sparklines.get(addrKey(addr));
    if (member) touched.add(member.group);
  }
  let count = 0;
  store.setState((s) => {
    const sparklines = new Map(s.sparkline.sparklines);
    const groups = new Map(s.sparkline.groups);
    for (const [key, member] of s.sparkline.sparklines) {
      const settings = touched.has(member.group) ? s.sparkline.groups.get(member.group) : undefined;
      if (!settings) continue;
      const id = newSparklineGroupId();
      sparklines.set(key, { ...member, group: id });
      groups.set(id, cloneSparklineSettings(settings));
      count += 1;
    }
    return { ...s, sparkline: withSparklineMembers({ sparklines, groups }, sparklines) };
  });
  return count;
}

/** Attach several sparklines as one new group with `settings` from the entry point `ctx`
 *  names: one engine group and one undo step, an object change of the sheet and the location
 *  cells. Any sheet protection refuses it, Edit objects included. */
export function setSparklinesAction(
  ctx: ObjectActionContext,
  entries: readonly SparklineInsert[],
  settings: SparklineGroupSettings,
): boolean {
  const cells = entries.map(({ addr }) => addr);
  const first = cells[0];
  if (!first) return false;
  return objectWrite(
    ctx,
    first.sheet,
    false,
    () => writeSparklines(ctx.store, entries, settings, ctx.history),
    [{ kind: 'cells', cells }],
  );
}

/** Attach several sparklines as one group, which is one engine group and one
 *  undo step. Nothing is written on a protected sheet. */
export function setSparklines(
  store: SpreadsheetStore,
  entries: readonly SparklineInsert[],
  settings: SparklineGroupSettings,
  history: History | null = null,
): boolean {
  return setSparklinesAction(
    { store, history, origin: 'instanceApi', commandId: 'mac.insert.sparkline' },
    entries,
    settings,
  );
}

const writeSparklines = (
  store: SpreadsheetStore,
  entries: readonly SparklineInsert[],
  settings: SparklineGroupSettings,
  history: History | null,
): boolean => {
  // Inserting or regrouping sparklines is refused under any protection, silently.
  if (entries.some(({ addr }) => isSheetProtected(store.getState(), addr.sheet))) return false;
  const group = newSparklineGroupId();
  recordSparklineChange(history, store, () => {
    // Every cell leaves its previous group first, so a replaced member never lingers in it.
    for (const { addr, source } of entries) mutators.setSparkline(store, addr, { source, group });
    mutators.setSparklineGroup(store, group, settings);
  });
  return true;
};

export function setSparkline(
  store: SpreadsheetStore,
  addr: Addr,
  source: string,
  settings: SparklineGroupSettings,
  history: History | null = null,
): boolean {
  return setSparklines(store, [{ addr, source }], settings, history);
}

export function clearSparkline(
  store: SpreadsheetStore,
  addr: Addr,
  history: History | null = null,
): boolean {
  const range = { sheet: addr.sheet, r0: addr.row, c0: addr.col, r1: addr.row, c1: addr.col };
  return (
    clearSparklinesAction(
      { store, history, origin: 'instanceApi', commandId: 'sparklineClear' },
      range,
    ) > 0
  );
}

export function clearSparklinesInRange(
  store: SpreadsheetStore,
  range: Range,
  history: History | null = null,
): number {
  return clearSparklinesAction(
    { store, history, origin: 'instanceApi', commandId: 'sparklineClear' },
    range,
  );
}

/**
 * Remove the sparklines of the cells in `range` from the entry point `ctx` names, as Clear All
 * does: a clear of those cells, so each cell's lock decides it (Edit objects does not) and a
 * locked cell keeps its sparkline. Only the cleared cells lose theirs; the rest of a group stays
 * one group, and a group left with no member is dropped. Returns how many were removed.
 */
export function clearSparklinesAction(ctx: ObjectActionContext, range: Range): number {
  return clearSparklineCellsAction(
    ctx,
    listSparklines(ctx.store.getState())
      .filter(({ addr }) => rangeContainsAddr(range, addr))
      .map(({ addr }) => addr),
  );
}

/** Remove the sparklines of `cells` as {@link clearSparklinesAction} does: one undo step, each
 *  cell's lock deciding it. Returns how many were removed. */
export function clearSparklineCellsAction(
  ctx: ObjectActionContext,
  cells: readonly Addr[],
): number {
  const { store, history } = ctx;
  const state = store.getState();
  const writable = cells.filter(
    (addr) => state.sparkline.sparklines.has(addrKey(addr)) && isCellWritable(state, addr),
  );
  if (writable.length === 0) return 0;
  const result = authorizedWrite({
    store,
    history: history ?? new History(),
    intents: (origin) => [
      {
        operation: 'clear',
        origin,
        commandId: ctx.commandId,
        effects: [{ kind: 'cells', cells: writable }],
      },
    ],
    origin: ctx.origin,
    write: () =>
      recordSparklineChange(history, store, () => {
        for (const addr of writable) mutators.clearSparkline(store, addr);
      }),
  });
  return result.ok ? writable.length : 0;
}

/** Split `Sheet2!B2:B9` / `'My Sheet'!B2:B9` into its qualifier (with the `!`)
 *  and the A1 body; the qualifier is empty for a bare range. */
export const splitSourceQualifier = (source: string): { qualifier: string; body: string } => {
  const bang = source.lastIndexOf('!');
  return bang < 0
    ? { qualifier: '', body: source }
    : { qualifier: source.slice(0, bang + 1), body: source.slice(bang + 1) };
};

export const unquoteSheetName = (qualifier: string): string => {
  const name = qualifier.slice(0, -1);
  return name.startsWith("'") && name.endsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
};

/** Re-point sparklines after a row/column insert or delete on `sheet`: host
 *  cells move with their cells (a deleted host drops its sparkline) and a
 *  source range on that sheet follows its data. `sheetName` is the edited
 *  sheet's name, used to recognise qualified sources. */
export function shiftSparklinesForAxisEdit(
  sparklines: ReadonlyMap<string, Sparkline>,
  sheet: number,
  sheetName: string,
  axis: 'row' | 'col',
  split: number,
  delta: number,
): Map<string, Sparkline> {
  const out = new Map<string, Sparkline>();
  const target = sheetName.toLowerCase();
  for (const [key, spec] of sparklines) {
    const host = parseAddrKey(key);
    if (!host) {
      out.set(key, spec);
      continue;
    }
    const { qualifier, body } = splitSourceQualifier(spec.source);
    const sourceOnSheet = qualifier
      ? unquoteSheetName(qualifier).toLowerCase() === target
      : host.sheet === sheet;
    const source = sourceOnSheet
      ? `${qualifier}${adjustFormulaForRowColEdit(`=${body}`, axis, split, delta).slice(1)}`
      : spec.source;
    const next = source === spec.source ? spec : { ...spec, source };
    if (host.sheet !== sheet) {
      out.set(key, next);
      continue;
    }
    const at = axis === 'row' ? host.row : host.col;
    const moved = adjustInterval(at, at, split, delta);
    if (!moved) continue;
    out.set(
      addrKey({
        ...host,
        row: axis === 'row' ? moved[0] : host.row,
        col: axis === 'col' ? moved[0] : host.col,
      }),
      next,
    );
  }
  return out;
}

/** Rewrite sparkline sources qualified with a renamed sheet's old name. */
export function renameSheetInSparklineSources(
  sparklines: ReadonlyMap<string, Sparkline>,
  oldName: string,
  newName: string,
): Map<string, Sparkline> | null {
  const old = oldName.toLowerCase();
  let changed = false;
  const out = new Map<string, Sparkline>();
  for (const [key, spec] of sparklines) {
    const { qualifier, body } = splitSourceQualifier(spec.source);
    if (qualifier && unquoteSheetName(qualifier).toLowerCase() === old) {
      changed = true;
      out.set(key, { ...spec, source: `${quoteSheetName(newName)}!${body}` });
    } else {
      out.set(key, spec);
    }
  }
  return changed ? out : null;
}
