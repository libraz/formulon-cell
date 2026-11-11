import type { Sparkline, SparklineGroupSettings, SparklineSlice } from './types.js';

/** A detached copy of a group's settings. */
export const cloneSparklineSettings = (settings: SparklineGroupSettings): SparklineGroupSettings =>
  structuredClone(settings);

/** A detached copy of the slice. */
export const cloneSparklineSlice = (slice: SparklineSlice): SparklineSlice => ({
  sparklines: new Map(Array.from(slice.sparklines, ([key, member]) => [key, { ...member }])),
  groups: new Map(
    Array.from(slice.groups, ([id, settings]) => [id, cloneSparklineSettings(settings)]),
  ),
});

/** The slice with `sparklines` as its members and only the groups they still use. */
export const withSparklineMembers = (
  slice: SparklineSlice,
  sparklines: Map<string, Sparkline>,
): SparklineSlice => {
  const used = new Set<string>();
  for (const member of sparklines.values()) used.add(member.group);
  let groups = slice.groups;
  if (groups.size !== used.size || [...groups.keys()].some((id) => !used.has(id))) {
    groups = new Map([...groups].filter(([id]) => used.has(id)));
  }
  return { sparklines, groups };
};

const sortedKeys = (_key: string, value: unknown): unknown =>
  value !== null && typeof value === 'object' && !Array.isArray(value)
    ? Object.fromEntries(Object.entries(value).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0)))
    : value;

/** Whether two groups' settings are equal, whatever the order of their keys. */
export const sameSparklineSettings = (
  a: SparklineGroupSettings,
  b: SparklineGroupSettings,
): boolean => a === b || JSON.stringify(a, sortedKeys) === JSON.stringify(b, sortedKeys);
