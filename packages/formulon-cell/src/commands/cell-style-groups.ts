import type { CellFormat } from '../store/store.js';
import type { CellStyleFormatGroup } from './cell-style-catalog.js';

/** The CellFormat fields each style group owns. A leaf module: the cell-style
 *  commands and the engine format sync both read it. */
export const CELL_STYLE_GROUP_FIELDS: Record<CellStyleFormatGroup, readonly (keyof CellFormat)[]> =
  {
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
      'fontOutline',
      'fontShadow',
      'fontCondense',
      'fontExtend',
    ],
    border: ['borders'],
    fill: ['fill', 'fillPattern', 'fillPatternColor', 'fillGradient'],
    protection: ['locked', 'formulaHidden'],
  };

export const ALL_BORDER_SIDES: readonly (keyof NonNullable<CellFormat['borders']>)[] = [
  'top',
  'right',
  'bottom',
  'left',
  'diagonalDown',
  'diagonalUp',
];

export const inferCellStyleGroups = (
  format: Partial<CellFormat>,
): readonly CellStyleFormatGroup[] => {
  const groups: CellStyleFormatGroup[] = [];
  for (const [group, fields] of Object.entries(CELL_STYLE_GROUP_FIELDS) as [
    CellStyleFormatGroup,
    readonly (keyof CellFormat)[],
  ][]) {
    if (fields.some((field) => Object.hasOwn(format, field))) groups.push(group);
  }
  return groups;
};

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
    case 'fontOutline':
    case 'fontShadow':
    case 'fontCondense':
    case 'fontExtend':
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
  const expected = formatForCellStyleGroups(stored, groups);
  const actual = formatForCellStyleGroups(native, groups);
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
