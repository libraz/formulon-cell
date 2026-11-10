import {
  type Color,
  ColorKind,
  type BorderInput as EngineBorderInput,
  type BorderSideInput as EngineBorderSideInput,
  type DxfInput as EngineDxfInput,
  type FillInput as EngineFillInput,
  type FontInput as EngineFontInput,
} from '@libraz/formulon';
import type { BorderRecord, BorderSide, DxfRecord, FillRecord, FontRecord } from './types.js';

/** No `<color>` element: an automatic font colour, or no fill / border colour. */
export const noColor = (): Color => ({
  kind: ColorKind.None,
  argb: 0,
  theme: 0,
  tint: 0,
  indexed: 0,
});

/** A literal colour the cell layer authored from UI state. ARGB 0 is "no
 * colour", which the store cannot distinguish from an absent element. */
export const literalColor = (argb: number): Color =>
  argb === 0 ? noColor() : { kind: ColorKind.Rgb, argb, theme: 0, tint: 0, indexed: 0 };

/** ARGB a colour displays as: the literal value, or the value a resolved
 * record carries for a theme / indexed / automatic selector; 0 for none. */
export const colorArgb = (color: Color | null | undefined): number =>
  !color || color.kind === ColorKind.None ? 0 : color.argb;

/** Fill in the fields the engine expects on every font record. Theme,
 * indexed and auto selectors exist solely on records read back from a file;
 * a record the cell layer authored carries no `<color>` selector of its own. */
export function completeFontRecord(record: FontRecord): EngineFontInput {
  return {
    ...record,
    vertAlign: record.vertAlign ?? 0,
    family: record.family ?? null,
    charset: record.charset ?? null,
    outline: record.outline ?? null,
    shadow: record.shadow ?? null,
    condense: record.condense ?? null,
    extend: record.extend ?? null,
    // A `<scheme>` link makes the font track the workbook theme. Only a record
    // read back out of a file carries one; a font the cell layer authored from
    // a family name is a literal typeface, so it stays unlinked.
    scheme: record.scheme ?? 0,
    color: record.color ?? noColor(),
  };
}

export function completeFillRecord(record: FillRecord): EngineFillInput {
  const { gradient, ...rest } = record;
  return {
    ...rest,
    foreground: record.foreground ?? noColor(),
    background: record.background ?? noColor(),
    gradient: gradient
      ? {
          ...gradient,
          stops: gradient.stops.map((stop) => ({ ...stop, color: stop.color ?? noColor() })),
        }
      : null,
  };
}

const completeBorderSide = (side: BorderSide): EngineBorderSideInput => ({
  ...side,
  color: side.color ?? noColor(),
});

export function completeBorderRecord(record: BorderRecord): EngineBorderInput {
  return {
    ...record,
    left: completeBorderSide(record.left),
    right: completeBorderSide(record.right),
    top: completeBorderSide(record.top),
    bottom: completeBorderSide(record.bottom),
    diagonal: completeBorderSide(record.diagonal),
  };
}

/** A differential format states only the properties it overrides, so absent
 * sub-records stay absent rather than becoming an explicit `undefined`. */
export function completeDxfRecord(record: DxfRecord): EngineDxfInput {
  const { font, fill, border, ...rest } = record;
  return {
    ...rest,
    ...(font ? { font: completeFontRecord(font) } : {}),
    ...(fill ? { fill: completeFillRecord(fill) } : {}),
    ...(border ? { border: completeBorderRecord(border) } : {}),
  };
}
