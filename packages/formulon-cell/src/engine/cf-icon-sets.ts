import type { CfIconInput } from '@libraz/formulon';
import { iconSetSlotCount } from '../render/conditional.js';
import type { ConditionalIconRef, ConditionalIconSet } from '../store/store.js';

/** `formulon::cf::IconSetName` ordinals, as OOXML `ST_IconSetType` spellings. */
const ENGINE_ICON_SET_NAMES = [
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '4Arrows',
  '4ArrowsGray',
  '4RedToBlack',
  '4Rating',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Rating',
  '5Quarters',
] as const;

/** The set's file spelling; null for a family with no file form. */
const SET_NAME: Readonly<Record<ConditionalIconSet, string | null>> = {
  arrows3: '3Arrows',
  arrowsGray3: '3ArrowsGray',
  arrows4: '4Arrows',
  arrowsGray4: '4ArrowsGray',
  arrows5: '5Arrows',
  arrowsGray5: '5ArrowsGray',
  triangles3: '3Triangles',
  traffic3: '3TrafficLights1',
  trafficRim3: '3TrafficLights2',
  signs3: '3Signs',
  traffic4: '4TrafficLights',
  redToBlack4: '4RedToBlack',
  symbols3: '3Symbols2',
  symbolsCircled3: '3Symbols',
  flags3: '3Flags',
  stars3: '3Stars',
  ratings4: '4Rating',
  quarters5: '5Quarters',
  ratings5: '5Rating',
  bars5: null,
  boxes5: '5Boxes',
};

/** The x14-only sets have no base-set ordinal; they are written as custom
 *  icons over a modelled set with the same icon count. */
const X14_BASE: Readonly<Partial<Record<ConditionalIconSet, ConditionalIconSet>>> = {
  triangles3: 'arrows3',
  stars3: 'arrows3',
  boxes5: 'arrows5',
};

const NO_ICONS = 'NoIcons';

const FAMILY_BY_NAME: ReadonlyMap<string, ConditionalIconSet> = new Map(
  (Object.entries(SET_NAME) as [ConditionalIconSet, string | null][])
    .filter((entry): entry is [ConditionalIconSet, string] => entry[1] !== null)
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
  return set && icon && icon.iconId >= 0 && icon.iconId < iconSetSlotCount(set)
    ? { set, slot: icon.iconId }
    : null;
};

/** Engine `iconSet` name + optional per-level icons for a family; null when
 *  the family or a custom level has no file form. */
export function iconSetToEngine(
  icons: ConditionalIconSet,
  customIcons?: readonly (ConditionalIconRef | null)[],
): { name: number; icons?: CfIconInput[] } | null {
  const base = X14_BASE[icons] ?? icons;
  const baseName = SET_NAME[base];
  const name = baseName === null ? -1 : ENGINE_ICON_SET_NAMES.indexOf(baseName as never);
  if (name < 0 || SET_NAME[icons] === null) return null;
  const levels = customIcons ?? (base === icons ? undefined : familyIcons(icons));
  if (!levels || sameIcons(levels, familyIcons(base))) return { name };
  if (levels.length !== iconSetSlotCount(base)) return null;
  const out: CfIconInput[] = [];
  for (const icon of levels) {
    if (icon === null) {
      out.push({ iconSet: NO_ICONS, iconId: 0 });
      continue;
    }
    const setName = SET_NAME[icon.set];
    if (setName === null || icon.slot < 0 || icon.slot >= iconSetSlotCount(icon.set)) return null;
    out.push({ iconSet: setName, iconId: icon.slot });
  }
  return { name, icons: out };
}

/** Family + per-level icons of an engine icon-set rule; null for an unknown
 *  base set. Custom icons that spell out an x14-only family collapse to it. */
export function iconSetFromEngine(
  name: number,
  icons?: readonly CfIconInput[],
): { icons: ConditionalIconSet; customIcons?: (ConditionalIconRef | null)[] } | null {
  const baseName = Number.isInteger(name) ? ENGINE_ICON_SET_NAMES[name] : undefined;
  const base = baseName === undefined ? undefined : FAMILY_BY_NAME.get(baseName);
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
