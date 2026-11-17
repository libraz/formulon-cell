import type { MacRibbonMenuItem } from './menu-manifest.js';

/** A gallery section: how many shapes it lists, and the ones the host draws by position. Every
 * caption and shape name is the generated `gallery:` label of its menu id. */
interface ShapeSection {
  readonly key: string;
  readonly size: number;
  /** Position of a drawable shape mapped to its icon, which also names its command. */
  readonly drawn?: Readonly<Record<number, string>>;
}

const SECTIONS: readonly ShapeSection[] = [
  {
    key: 'basic',
    size: 35,
    drawn: {
      2: 'shapeRectangle',
      3: 'shapeRoundedRectangle',
      4: 'shapeOval',
      12: 'shapeTriangle',
      16: 'shapeDiamond',
    },
  },
  { key: 'lines', size: 19, drawn: { 0: 'shapeLine', 1: 'shapeArrow' } },
  { key: 'arrows', size: 27 },
  { key: 'callouts', size: 17 },
  { key: 'symbols', size: 15 },
  { key: 'flowchart', size: 28 },
  { key: 'stars', size: 20 },
];

/** The Shapes gallery: every Excel section, with the shapes the host draws enabled. */
export const MAC_SHAPE_GALLERY: readonly MacRibbonMenuItem[] = SECTIONS.flatMap((group) => {
  const headingId = `mac.insert.heading.shape.${group.key}`;
  return [
    { id: headingId, heading: true, own: true },
    ...Array.from({ length: group.size }, (_, index): MacRibbonMenuItem => {
      const icon = group.drawn?.[index];
      const id = icon ? `mac.insert.${icon}` : `mac.insert.shape.${group.key}.${index}`;
      return {
        id,
        own: true,
        ...(icon ? { icon } : {}),
        disabled: !icon,
        disabledReason: icon ? undefined : 'notImplemented',
      };
    }),
  ];
});
