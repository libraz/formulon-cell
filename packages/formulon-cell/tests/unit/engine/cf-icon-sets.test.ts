import { describe, expect, it } from 'vitest';

import { iconSetFromEngine, iconSetToEngine } from '../../../src/engine/cf-icon-sets.js';
import type { ConditionalIconSet } from '../../../src/store/store.js';

// `formulon::cf::IconSetName` ordinals, in engine order.
const ENGINE_NAMES = [
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

const FAMILY_BY_NAME: Record<(typeof ENGINE_NAMES)[number], ConditionalIconSet> = {
  '3Arrows': 'arrows3',
  '3ArrowsGray': 'arrowsGray3',
  '3Flags': 'flags3',
  '3TrafficLights1': 'traffic3',
  '3TrafficLights2': 'trafficRim3',
  '3Signs': 'signs3',
  '3Symbols': 'symbolsCircled3',
  '3Symbols2': 'symbols3',
  '4Arrows': 'arrows4',
  '4ArrowsGray': 'arrowsGray4',
  '4RedToBlack': 'redToBlack4',
  '4Rating': 'ratings4',
  '4TrafficLights': 'traffic4',
  '5Arrows': 'arrows5',
  '5ArrowsGray': 'arrowsGray5',
  '5Rating': 'ratings5',
  '5Quarters': 'quarters5',
};

describe('cf icon-set mapping', () => {
  it('maps every engine base set ordinal to its family and back', () => {
    ENGINE_NAMES.forEach((name, ordinal) => {
      const family = FAMILY_BY_NAME[name];
      expect(iconSetFromEngine(ordinal)).toEqual({ icons: family });
      expect(iconSetToEngine(family)).toEqual({ name: ordinal });
    });
  });

  it('writes the x14-only families as custom icons over a modelled base set', () => {
    expect(iconSetToEngine('triangles3')).toEqual({
      name: 0,
      icons: [0, 1, 2].map((iconId) => ({ iconSet: '3Triangles', iconId })),
    });
    expect(iconSetToEngine('stars3')).toEqual({
      name: 0,
      icons: [0, 1, 2].map((iconId) => ({ iconSet: '3Stars', iconId })),
    });
    expect(iconSetToEngine('boxes5')).toEqual({
      name: 13,
      icons: [0, 1, 2, 3, 4].map((iconId) => ({ iconSet: '5Boxes', iconId })),
    });
    expect(
      iconSetFromEngine(
        0,
        [0, 1, 2].map((iconId) => ({ iconSet: '3Triangles', iconId })),
      ),
    ).toEqual({ icons: 'triangles3' });
  });

  it('has no file form for a family without a desktop counterpart', () => {
    expect(iconSetToEngine('bars5')).toBeNull();
  });

  it('round-trips per-level custom icons, with NoIcons as an empty level', () => {
    const icons = [
      { iconSet: '3Stars', iconId: 2 },
      { iconSet: 'NoIcons', iconId: 0 },
      { iconSet: '4RedToBlack', iconId: 3 },
    ];
    const hydrated = iconSetFromEngine(3, icons);
    expect(hydrated).toEqual({
      icons: 'traffic3',
      customIcons: [{ set: 'stars3', slot: 2 }, null, { set: 'redToBlack4', slot: 3 }],
    });
    expect(iconSetToEngine('traffic3', hydrated?.customIcons)).toEqual({ name: 3, icons });
  });

  it('drops custom icons that only restate the family default', () => {
    expect(
      iconSetToEngine('arrows3', [
        { set: 'arrows3', slot: 0 },
        { set: 'arrows3', slot: 1 },
        { set: 'arrows3', slot: 2 },
      ]),
    ).toEqual({ name: 0 });
    expect(
      iconSetFromEngine(
        0,
        [0, 1, 2].map((iconId) => ({ iconSet: '3Arrows', iconId })),
      ),
    ).toEqual({ icons: 'arrows3' });
  });

  it('rejects unknown ordinals and custom icons of the wrong length', () => {
    expect(iconSetFromEngine(99)).toBeNull();
    expect(iconSetToEngine('arrows3', [{ set: 'arrows3', slot: 0 }])).toBeNull();
  });
});
