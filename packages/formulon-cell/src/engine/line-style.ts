import type { ChartLineEnd, ChartLineEndInput } from '@libraz/formulon';
import { ChartLineDash, ChartLineEndSize, ChartLineEndType } from '@libraz/formulon';
import type {
  SessionLineEnd,
  SessionLineEndSize,
  SessionLineEndType,
  SessionShapeLineDash,
} from '../store/types.js';

/** The engine dash each dash style is stored as: the Dashes entries are the preset dashes. */
const ENGINE_DASH: Readonly<Record<SessionShapeLineDash, ChartLineDash>> = {
  solid: ChartLineDash.Solid,
  roundDot: ChartLineDash.SysDot,
  squareDot: ChartLineDash.SysDash,
  dash: ChartLineDash.Dash,
  dashDot: ChartLineDash.DashDot,
  longDash: ChartLineDash.LgDash,
  longDashDot: ChartLineDash.LgDashDot,
  longDashDotDot: ChartLineDash.LgDashDotDot,
  dot: ChartLineDash.Dot,
  sysDashDot: ChartLineDash.SysDashDot,
  sysDashDotDot: ChartLineDash.SysDashDotDot,
};

const DASH_OF_ENGINE: ReadonlyMap<ChartLineDash, SessionShapeLineDash> = new Map(
  (Object.keys(ENGINE_DASH) as SessionShapeLineDash[]).map((dash) => [ENGINE_DASH[dash], dash]),
);

/** The engine dash of a dash style. */
export const engineDashOf = (dash: SessionShapeLineDash): ChartLineDash => ENGINE_DASH[dash];

/** The dash style of an engine dash; none for one the overlay has no style for (a custom dash). */
export const dashStyleOf = (dash: ChartLineDash): SessionShapeLineDash | undefined =>
  DASH_OF_ENGINE.get(dash);

const END_TYPE: Readonly<Record<SessionLineEndType, ChartLineEndType>> = {
  triangle: ChartLineEndType.Triangle,
  stealth: ChartLineEndType.Stealth,
  diamond: ChartLineEndType.Diamond,
  oval: ChartLineEndType.Oval,
  arrow: ChartLineEndType.Arrow,
};

const END_SIZE: Readonly<Record<SessionLineEndSize, ChartLineEndSize>> = {
  small: ChartLineEndSize.Small,
  medium: ChartLineEndSize.Medium,
  large: ChartLineEndSize.Large,
};

const keyOf = <K extends string, V>(table: Readonly<Record<K, V>>, value: V): K | undefined =>
  (Object.keys(table) as K[]).find((key) => table[key] === value);

/** A line end as the overlay holds it; none when the end is absent or states no decoration. */
export const lineEndOf = (end: ChartLineEnd | null): SessionLineEnd | undefined => {
  const type = end ? keyOf(END_TYPE, end.type) : undefined;
  if (!end || !type) return undefined;
  const width = end.width === null ? undefined : keyOf(END_SIZE, end.width);
  const length = end.length === null ? undefined : keyOf(END_SIZE, end.length);
  return {
    type,
    ...(width && width !== 'medium' ? { width } : {}),
    ...(length && length !== 'medium' ? { length } : {}),
  };
};

/** A line end as the engine takes it. */
export const engineLineEnd = (end: SessionLineEnd): ChartLineEndInput => ({
  type: END_TYPE[end.type],
  ...(end.width ? { width: END_SIZE[end.width] } : {}),
  ...(end.length ? { length: END_SIZE[end.length] } : {}),
});

/** The ends a line shows: an arrow stands for a triangle at the end when it states none. */
export const lineEndsOf = (item: {
  shape?: string;
  headEnd?: SessionLineEnd;
  tailEnd?: SessionLineEnd;
}): { head?: SessionLineEnd; tail?: SessionLineEnd } => ({
  ...(item.headEnd ? { head: item.headEnd } : {}),
  ...(item.tailEnd
    ? { tail: item.tailEnd }
    : item.shape === 'arrow' && !item.headEnd
      ? { tail: { type: 'triangle' as const } }
      : {}),
});
