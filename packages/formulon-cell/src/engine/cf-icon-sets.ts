import { type CfIconInput, IconSetName } from '@libraz/formulon';
import { iconSetSlotCount } from '../render/conditional.js';
import type { ConditionalIconRef, ConditionalIconSet } from '../store/store.js';

/** The set's `IconSetName`; null for a family with no file form. */
const SET_NAME: Readonly<Record<ConditionalIconSet, IconSetName | null>> = {
  arrows3: IconSetName.ThreeArrows,
  arrowsGray3: IconSetName.ThreeArrowsGray,
  arrows4: IconSetName.FourArrows,
  arrowsGray4: IconSetName.FourArrowsGray,
  arrows5: IconSetName.FiveArrows,
  arrowsGray5: IconSetName.FiveArrowsGray,
  triangles3: IconSetName.ThreeTriangles,
  traffic3: IconSetName.ThreeTrafficLights1,
  trafficRim3: IconSetName.ThreeTrafficLights2,
  signs3: IconSetName.ThreeSigns,
  traffic4: IconSetName.FourTrafficLights,
  redToBlack4: IconSetName.FourRedToBlack,
  symbols3: IconSetName.ThreeSymbols2,
  symbolsCircled3: IconSetName.ThreeSymbols,
  flags3: IconSetName.ThreeFlags,
  stars3: IconSetName.ThreeStars,
  ratings4: IconSetName.FourRating,
  quarters5: IconSetName.FiveQuarters,
  ratings5: IconSetName.FiveRating,
  bars5: null,
  boxes5: IconSetName.FiveBoxes,
};

/** The x14-only sets have no base-set ordinal; they are written as custom
 *  icons over a modelled set with the same icon count. */
const X14_BASE: Readonly<Partial<Record<ConditionalIconSet, ConditionalIconSet>>> = {
  triangles3: 'arrows3',
  stars3: 'arrows3',
  boxes5: 'arrows5',
};

/** A rule's own set is one of the built-in sets up to this ordinal. */
const isBaseSetName = (name: number): boolean =>
  Number.isInteger(name) && name >= 0 && name <= IconSetName.FiveQuarters;

const FAMILY_BY_NAME: ReadonlyMap<number, ConditionalIconSet> = new Map(
  (Object.entries(SET_NAME) as [ConditionalIconSet, IconSetName | null][])
    .filter((entry): entry is [ConditionalIconSet, IconSetName] => entry[1] !== null)
    .map(([family, name]) => [name, family]),
);

const familyIcons = (set: ConditionalIconSet): ConditionalIconRef[] =>
  Array.from({ length: iconSetSlotCount(set) }, (_, slot) => ({ set, slot }));

const sameIcons = (
  left: readonly (ConditionalIconRef | null)[],
  right: readonly (ConditionalIconRef | null)[],
): boolean =>
  left.length === right.length &&
  left.every((icon, i) => {
    const other = right[i];
    return icon === null || other === null || other === undefined
      ? icon === other
      : icon.set === other.set && icon.slot === other.slot;
  });

/** The icon a custom level draws; null for `NoIcons` or an unknown set. */
export const customIconFromEngine = (icon: CfIconInput | undefined): ConditionalIconRef | null => {
  const set = icon ? FAMILY_BY_NAME.get(icon.iconSet) : undefined;
  return set && icon && icon.iconIndex >= 0 && icon.iconIndex < iconSetSlotCount(set)
    ? { set, slot: icon.iconIndex }
    : null;
};

/** Engine `iconSet` name + optional per-level icons for a family; null when
 *  the family or a custom level has no file form. */
export function iconSetToEngine(
  icons: ConditionalIconSet,
  customIcons?: readonly (ConditionalIconRef | null)[],
): { name: IconSetName; icons?: CfIconInput[] } | null {
  const base = X14_BASE[icons] ?? icons;
  const name = SET_NAME[base];
  if (name === null || !isBaseSetName(name) || SET_NAME[icons] === null) return null;
  const levels = customIcons ?? (base === icons ? undefined : familyIcons(icons));
  if (!levels || sameIcons(levels, familyIcons(base))) return { name };
  if (levels.length !== iconSetSlotCount(base)) return null;
  const out: CfIconInput[] = [];
  for (const icon of levels) {
    if (icon === null) {
      out.push({ iconSet: IconSetName.NoIcons, iconIndex: 0 });
      continue;
    }
    const setName = SET_NAME[icon.set];
    if (setName === null || icon.slot < 0 || icon.slot >= iconSetSlotCount(icon.set)) return null;
    out.push({ iconSet: setName, iconIndex: icon.slot });
  }
  return { name, icons: out };
}

/** Family + per-level icons of an engine icon-set rule; null for an unknown
 *  base set. Custom icons that spell out an x14-only family collapse to it. */
export function iconSetFromEngine(
  name: number,
  icons?: readonly CfIconInput[],
): { icons: ConditionalIconSet; customIcons?: (ConditionalIconRef | null)[] } | null {
  const base = isBaseSetName(name) ? FAMILY_BY_NAME.get(name) : undefined;
  if (!base) return null;
  if (!icons || icons.length === 0) return { icons: base };
  const levels = icons.map(customIconFromEngine);
  if (sameIcons(levels, familyIcons(base))) return { icons: base };
  for (const [family, x14Base] of Object.entries(X14_BASE) as [
    ConditionalIconSet,
    ConditionalIconSet,
  ][]) {
    if (x14Base === base && sameIcons(levels, familyIcons(family))) return { icons: family };
  }
  return { icons: base, customIcons: levels };
}

/** Family of a base-set ordinal reported by CF evaluation. */
export const iconSetFamilyForOrdinal = (name: number): ConditionalIconSet | null =>
  iconSetFromEngine(name)?.icons ?? null;
