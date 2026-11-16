import { DrawingObjectKind } from '@libraz/formulon';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from './types.js';

/** One drawing of a sheet's stack; `objectId` is set once the engine holds it. */
export interface DrawingStackEntry {
  readonly id: string;
  readonly chart: boolean;
  readonly objectId: number | null;
}

export type DrawingState = Pick<State, 'charts' | 'illustrations' | 'drawings'>;

type DrawingEngine = Pick<
  WorkbookHandle,
  'capabilities' | 'listDrawingObjects' | 'setDrawingZOrder'
>;

/** `order` with the ids it lacks appended. */
const withIds = (order: readonly string[], ids: Iterable<string>): readonly string[] => {
  const known = new Set(order);
  const added: string[] = [];
  for (const id of ids) {
    if (known.has(id)) continue;
    known.add(id);
    added.push(id);
  }
  return added.length === 0 ? order : [...order, ...added];
};

/** `order` with every item it lacks listed right under the next listed item of
 *  its slice, as the slice holds it, or on top when none follows. */
const withItems = (
  order: readonly string[],
  items: readonly { id: string }[],
): readonly string[] => {
  const known = new Set(order);
  if (items.every((item) => known.has(item.id))) return order;
  const next = [...order];
  let pending: string[] = [];
  for (const { id } of items) {
    if (!known.has(id)) {
      known.add(id);
      pending.push(id);
    } else if (pending.length > 0) {
      next.splice(next.indexOf(id), 0, ...pending);
      pending = [];
    }
  }
  return pending.length > 0 ? [...next, ...pending] : next;
};

/** `order` without the ids no chart or illustration of `state` has. */
export const liveOrder = (state: DrawingState, order: readonly string[]): readonly string[] => {
  const live = new Set([
    ...state.charts.charts.map((chart) => chart.id),
    ...state.illustrations.illustrations.map((item) => item.id),
  ]);
  return order.every((id) => live.has(id)) ? order : order.filter((id) => live.has(id));
};

/** `base`, or `base` with a suffix when a drawing of `state` or an entry of its
 *  order already has it. A minted id is kept for as long as the object lives. */
export const mintDrawingId = (
  state: DrawingState,
  base: string,
  taken: ReadonlySet<string> = new Set(),
): string => {
  const used = (id: string): boolean =>
    taken.has(id) ||
    state.drawings.order.includes(id) ||
    state.charts.charts.some((chart) => chart.id === id) ||
    state.illustrations.illustrations.some((item) => item.id === id);
  let id = base;
  for (let n = 2; used(id); n += 1) id = `${base}-${n}`;
  return id;
};

/** `items` stably sorted by `rank`; the same array when already sorted. */
const sortedBy = <T extends { id: string }>(
  items: readonly T[],
  rank: ReadonlyMap<string, number>,
): readonly T[] => {
  const key = (item: T): number => rank.get(item.id) ?? Number.POSITIVE_INFINITY;
  const sorted = items
    .map((item, index) => ({ item, index }))
    .sort((a, b) => key(a.item) - key(b.item) || a.index - b.index)
    .map(({ item }) => item);
  return sorted.every((item, index) => item === items[index]) ? items : sorted;
};

/** `state` with `order` as its drawing order: every chart and illustration is
 *  listed, and both slices are kept sorted by it. An id whose object is gone
 *  stays until a removal or a hydration prunes it, so a hydration that drops
 *  and re-reads an object keeps its slot. */
export const withDrawingOrder = <S extends DrawingState>(state: S, order: readonly string[]): S => {
  const full = withItems(withItems(order, state.charts.charts), state.illustrations.illustrations);
  const rank = new Map(full.map((id, index) => [id, index] as const));
  const charts = sortedBy(state.charts.charts, rank);
  const illustrations = sortedBy(state.illustrations.illustrations, rank);
  if (
    full === state.drawings.order &&
    charts === state.charts.charts &&
    illustrations === state.illustrations.illustrations
  ) {
    return state;
  }
  return {
    ...state,
    drawings: { order: full },
    charts: charts === state.charts.charts ? state.charts : { charts },
    illustrations:
      illustrations === state.illustrations.illustrations ? state.illustrations : { illustrations },
  };
};

/** `order` with the ids of `stack` refilled, in the order of `stack`, into the
 *  slots they hold. */
export const restacked = (
  order: readonly string[],
  stack: readonly string[],
): readonly string[] => {
  const full = withIds(order, stack);
  const members = new Set(stack);
  let next = 0;
  return full.map((id) => (members.has(id) ? (stack[next++] as string) : id));
};

/**
 * `order` with the saved ids in `saved`'s order. Every other id keeps its place
 * over the saved id it lay over, or under every saved id when none lay under it.
 */
export const followingSaved = (
  order: readonly string[],
  saved: readonly string[],
): readonly string[] => {
  const savedIds = new Set(saved);
  const over = new Map<string | null, string[]>([[null, []]]);
  let below: string | null = null;
  for (const id of order) {
    if (savedIds.has(id)) {
      below = id;
      continue;
    }
    const group = over.get(below);
    if (group) group.push(id);
    else over.set(below, [id]);
  }
  const next = [...(over.get(null) ?? [])];
  for (const id of new Set(saved)) next.push(id, ...(over.get(id) ?? []));
  return next.length === order.length && next.every((id, i) => id === order[i]) ? order : next;
};

/** A sheet's charts, pictures and shapes in the store's drawing order, back to front. */
export function sheetDrawingStack(state: DrawingState, sheet: number): DrawingStackEntry[] {
  const entries = new Map<string, DrawingStackEntry>();
  for (const chart of state.charts.charts) {
    if (chart.sheet !== sheet) continue;
    entries.set(chart.id, { id: chart.id, chart: true, objectId: chart.nativeObjectId });
  }
  for (const item of state.illustrations.illustrations) {
    if (item.sheet !== sheet) continue;
    entries.set(item.id, { id: item.id, chart: false, objectId: item.nativeObjectId ?? null });
  }
  return withItems(
    withItems(state.drawings.order, state.charts.charts),
    state.illustrations.illustrations,
  ).flatMap((id) => entries.get(id) ?? []);
}

/** The store id of each object of `sheet` the engine lists, in the engine's order. */
export function savedDrawingIds(state: DrawingState, sheet: number, wb: DrawingEngine): string[] {
  const charts = new Map<number, string>();
  const others = new Map<number, string>();
  for (const entry of sheetDrawingStack(state, sheet)) {
    if (entry.objectId !== null) (entry.chart ? charts : others).set(entry.objectId, entry.id);
  }
  return wb
    .listDrawingObjects(sheet)
    .flatMap(
      (object) =>
        (object.kind === DrawingObjectKind.Chart ? charts : others).get(object.objectId) ?? [],
    );
}

/** Restack the engine's objects of `sheet` in the store's order. Objects the
 *  store does not hold keep their engine slots. False when the engine refuses. */
export function writeDrawingOrder(wb: DrawingEngine, state: DrawingState, sheet: number): boolean {
  if (!wb.capabilities.drawings) return true;
  const objects = wb.listDrawingObjects(sheet).map((object) => object.objectId);
  const listed = new Set(objects);
  const wanted = [
    ...new Set(
      sheetDrawingStack(state, sheet).flatMap((entry) =>
        entry.objectId !== null && listed.has(entry.objectId) ? [entry.objectId] : [],
      ),
    ),
  ];
  const held = new Set(wanted);
  let next = 0;
  const target = objects.map((objectId) =>
    held.has(objectId) ? (wanted[next++] as number) : objectId,
  );
  const current = [...objects];
  for (let index = 0; index < target.length; index += 1) {
    const objectId = target[index] as number;
    if (current[index] === objectId) continue;
    if (!wb.setDrawingZOrder(sheet, objectId, index)) return false;
    current.splice(current.indexOf(objectId), 1);
    current.splice(index, 0, objectId);
  }
  return true;
}
