import { SparklineAxisType, SparklineType } from '@libraz/formulon';
import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr } from '../engine/types.js';
import { rangeContainsAddr } from '../store/selection-geometry.js';
import {
  mutators,
  type SparklineGroupSettings,
  type SparklineSlice,
  type State,
} from '../store/store.js';
import { type ObjectActionContext, objectWrite } from './object-actions.js';
import { isSheetProtected } from './protection.js';
import { recordSparklineChange } from './slice-history.js';
import {
  clearSparklineCellsAction,
  groupOfCell,
  groupSparklines,
  membersOfGroup,
  rgbColorSpec,
  type SparklineInsert,
  ungroupSparklines,
  updateSparklineGroup,
} from './sparkline.js';
import { switchSparklineSources } from './sparkline-layout.js';

/** The on/off settings of a group that the Show group toggles. */
export type SparklineFlag = 'high' | 'low' | 'negative' | 'first' | 'last' | 'markers';

/** The colours of a group, by the point or part each paints. */
export type SparklineColorPart =
  | 'series'
  | 'negative'
  | 'markers'
  | 'high'
  | 'low'
  | 'first'
  | 'last';

export const SPARKLINE_COLOR_KEYS = {
  series: 'colorSeries',
  negative: 'colorNegative',
  markers: 'colorMarkers',
  high: 'colorHigh',
  low: 'colorLow',
  first: 'colorFirst',
  last: 'colorLast',
} as const satisfies Readonly<Record<SparklineColorPart, keyof SparklineGroupSettings>>;

/** Ribbon command id → the sparkline type it sets. */
export const SPARKLINE_TYPE_COMMANDS: Readonly<Record<string, SparklineType>> = {
  sparklineTypeLine: SparklineType.Line,
  sparklineTypeColumn: SparklineType.Column,
  sparklineTypeWinLoss: SparklineType.Stacked,
};

/** Ribbon command id → the flag it toggles. */
export const SPARKLINE_SHOW_COMMANDS: Readonly<Record<string, SparklineFlag>> = {
  sparklineShowHigh: 'high',
  sparklineShowLow: 'low',
  sparklineShowNegative: 'negative',
  sparklineShowFirst: 'first',
  sparklineShowLast: 'last',
  sparklineShowMarkers: 'markers',
};

type SparklineState = Pick<State, 'selection'> & { sparkline: SparklineSlice };

/** The sparkline cells inside the selection ranges, row-major. The active cell stands in when
 *  it hosts a sparkline the ranges somehow miss. */
export function selectedSparklineCells(state: SparklineState): Addr[] {
  const ranges = [state.selection.range, ...(state.selection.extraRanges ?? [])];
  const out: Addr[] = [];
  for (const key of state.sparkline.sparklines.keys()) {
    const addr = parseAddrKey(key);
    if (addr && ranges.some((range) => rangeContainsAddr(range, addr))) out.push(addr);
  }
  const active = state.selection.active;
  if (out.length === 0 && state.sparkline.sparklines.has(addrKey(active))) out.push(active);
  return out.sort((a, b) => a.row - b.row || a.col - b.col);
}

/** Ids of the groups the sparklines at `cells` belong to. */
export const sparklineGroupIds = (state: SparklineState, cells: readonly Addr[]): string[] => {
  const ids = new Set<string>();
  for (const addr of cells) {
    const member = state.sparkline.sparklines.get(addrKey(addr));
    if (member) ids.add(member.group);
  }
  return [...ids];
};

/** The group of the active cell, with a copy of its settings. */
export const activeSparklineGroup = (state: SparklineState) =>
  groupOfCell(state, state.selection.active);

const designWrite = (
  ctx: ObjectActionContext,
  cells: readonly Addr[],
  write: () => void,
  writesCells = false,
): boolean => {
  const first = cells[0];
  if (!first) return false;
  const state = ctx.store.getState();
  // A protected sheet refuses every sparkline change, Edit objects included.
  if (cells.some((addr) => isSheetProtected(state, addr.sheet))) return false;
  return objectWrite(
    ctx,
    first.sheet,
    false,
    () => {
      write();
      return true;
    },
    writesCells ? [{ kind: 'cells', cells }] : [],
  );
};

/** Patch the settings of the groups of the selected sparklines. */
export function patchSelectedSparklineGroupsAction(
  ctx: ObjectActionContext,
  patch: Partial<SparklineGroupSettings>,
): boolean {
  const state = ctx.store.getState();
  const cells = selectedSparklineCells(state);
  const groups = sparklineGroupIds(state, cells);
  return designWrite(ctx, cells, () =>
    recordSparklineChange(ctx.history, ctx.store, () => {
      for (const id of groups) updateSparklineGroup(ctx.store, id, patch);
    }),
  );
}

/** Set the colour of `part` on the selected groups to the CSS colour `css`. */
export const setSparklineColorAction = (
  ctx: ObjectActionContext,
  part: SparklineColorPart,
  css: string,
): boolean =>
  patchSelectedSparklineGroupsAction(ctx, { [SPARKLINE_COLOR_KEYS[part]]: rgbColorSpec(css) });

/** Flip `flag` on the selected groups to the opposite of the active group's value. */
export function toggleSparklineFlagAction(ctx: ObjectActionContext, flag: SparklineFlag): boolean {
  const active = activeSparklineGroup(ctx.store.getState());
  return active
    ? patchSelectedSparklineGroupsAction(ctx, { [flag]: !active.settings[flag] })
    : false;
}

/** Set one end of the vertical axis on the selected groups. */
export function setSparklineAxisBoundAction(
  ctx: ObjectActionContext,
  end: 'min' | 'max',
  mode: 'individual' | 'group' | 'custom',
  value?: number,
): boolean {
  const type =
    mode === 'individual'
      ? SparklineAxisType.Individual
      : mode === 'group'
        ? SparklineAxisType.Group
        : SparklineAxisType.Custom;
  return patchSelectedSparklineGroupsAction(
    ctx,
    end === 'min'
      ? { minAxisType: type, ...(value === undefined ? {} : { manualMin: value }) }
      : { maxAxisType: type, ...(value === undefined ? {} : { manualMax: value }) },
  );
}

/** Merge the groups of the selected sparklines into one with the active group's settings. */
export function groupSelectedSparklinesAction(ctx: ObjectActionContext): boolean {
  const state = ctx.store.getState();
  const cells = selectedSparklineCells(state);
  return designWrite(
    ctx,
    cells,
    () =>
      recordSparklineChange(ctx.history, ctx.store, () => {
        groupSparklines(ctx.store, cells, state.selection.active);
      }),
    true,
  );
}

/** Split the groups of the selected sparklines into single-sparkline groups. */
export function ungroupSelectedSparklinesAction(ctx: ObjectActionContext): boolean {
  const cells = selectedSparklineCells(ctx.store.getState());
  return designWrite(
    ctx,
    cells,
    () =>
      recordSparklineChange(ctx.history, ctx.store, () => {
        ungroupSparklines(ctx.store, cells);
      }),
    true,
  );
}

/** Clear the selected sparklines; their groups keep the other members. */
export function clearSelectedSparklinesAction(ctx: ObjectActionContext): number {
  const state = ctx.store.getState();
  const cells = selectedSparklineCells(state);
  if (cells.some((addr) => isSheetProtected(state, addr.sheet))) return 0;
  return clearSparklineCellsAction(ctx, cells);
}

/** Clear every sparkline of the groups of the selected sparklines. */
export function clearSelectedSparklineGroupsAction(ctx: ObjectActionContext): number {
  const state = ctx.store.getState();
  const cells = selectedSparklineCells(state);
  if (cells.some((addr) => isSheetProtected(state, addr.sheet))) return 0;
  const groups = new Set(sparklineGroupIds(state, cells));
  const members = [...state.sparkline.sparklines]
    .filter(([, member]) => groups.has(member.group))
    .flatMap(([key]) => parseAddrKey(key) ?? []);
  return clearSparklineCellsAction(ctx, members);
}

/** Replace the members of group `groupId` with `entries`, keeping its settings. */
export function replaceSparklineGroupAction(
  ctx: ObjectActionContext,
  groupId: string,
  entries: readonly SparklineInsert[],
): boolean {
  const state = ctx.store.getState();
  const old = membersOfGroup(state, groupId).map(({ addr }) => addr);
  const kept = new Set(entries.map(({ addr }) => addrKey(addr)));
  return designWrite(
    ctx,
    [...entries.map(({ addr }) => addr), ...old],
    () =>
      recordSparklineChange(ctx.history, ctx.store, () => {
        // Members first, so the group never loses its last one midway.
        for (const { addr, source } of entries) {
          mutators.setSparkline(ctx.store, addr, { source, group: groupId });
        }
        for (const addr of old)
          if (!kept.has(addrKey(addr))) mutators.clearSparkline(ctx.store, addr);
      }),
    true,
  );
}

/** Point the sparkline at `addr` at another data range, leaving it in its group. */
export function setSparklineSourceAction(
  ctx: ObjectActionContext,
  addr: Addr,
  source: string,
): boolean {
  const member = ctx.store.getState().sparkline.sparklines.get(addrKey(addr));
  if (!member) return false;
  return designWrite(
    ctx,
    [addr],
    () =>
      recordSparklineChange(ctx.history, ctx.store, () => {
        mutators.setSparkline(ctx.store, addr, { ...member, source });
      }),
    true,
  );
}

/** The new sources of the active group's members after Switch Row/Column; null when the group
 *  has no other split. */
export function switchedSparklineSources(
  state: SparklineState,
): { addr: Addr; source: string }[] | null {
  const group = activeSparklineGroup(state);
  if (!group) return null;
  const members = membersOfGroup(state, group.id);
  const sources = switchSparklineSources(members.map(({ spec }) => spec.source));
  return sources ? members.map(({ addr }, i) => ({ addr, source: sources[i] as string })) : null;
}

/** Re-split the active group's data along the other axis. False when it has no other split. */
export function switchSparklineRowColumnAction(ctx: ObjectActionContext): boolean {
  const switched = switchedSparklineSources(ctx.store.getState());
  if (!switched) return false;
  const group = activeSparklineGroup(ctx.store.getState());
  if (!group) return false;
  return replaceSparklineGroupAction(ctx, group.id, switched);
}
