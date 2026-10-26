import type {
  ColorSpec,
  BorderRecord as EngineBorderRecord,
  BorderSide as EngineBorderSide,
  DxfRecord as EngineDxfRecord,
  FillRecord as EngineFillRecord,
  FontRecord as EngineFontRecord,
} from '@libraz/formulon';
import type { BorderRecord, BorderSide, DxfRecord, FillRecord, FontRecord } from './types.js';

/** A `<color>` selector for a record the cell layer authored from UI state.
 * `kind: 0` tells the writer to emit the record's sibling `*Argb` field as
 * literal rgb, which is the only kind of colour the store can hold — theme,
 * indexed and auto selectors exist solely on records read back from a file. */
const authoredColorSpec = (): ColorSpec => ({ kind: 0, rgb: 0, theme: 0, tint: 0, indexed: 0 });

/** Fill in the fields the engine expects on every font record. A caller that
 * states `bold` without stating `hasBold` means it — an omitted flag defaults
 * to "the element is present", which is what a full cell font always wants.
 * Differential fonts pass the flags explicitly so an unset property stays
 * absent and leaves the underlying cell's attribute alone. */
export function completeFontRecord(record: FontRecord): EngineFontRecord {
  return {
    ...record,
    vertAlign: record.vertAlign ?? 0,
    hasBold: record.hasBold ?? true,
    hasItalic: record.hasItalic ?? true,
    hasStrike: record.hasStrike ?? true,
    hasSize: record.hasSize ?? true,
    hasFamily: record.hasFamily ?? false,
    family: record.family ?? 0,
    hasCharset: record.hasCharset ?? false,
    charset: record.charset ?? 0,
    outline: record.outline ?? false,
    shadow: record.shadow ?? false,
    condense: record.condense ?? false,
    extend: record.extend ?? false,
    hasOutline: record.hasOutline ?? false,
    hasShadow: record.hasShadow ?? false,
    hasCondense: record.hasCondense ?? false,
    hasExtend: record.hasExtend ?? false,
    // A `<scheme>` link makes the font track the workbook theme. Only a record
    // read back out of a file carries one; a font the cell layer authored from
    // a family name is a literal typeface, so it stays unlinked.
    scheme: record.scheme ?? 0,
    color: record.color ?? authoredColorSpec(),
  };
}

export function completeFillRecord(record: FillRecord): EngineFillRecord {
  const { gradient, ...rest } = record;
  const out: EngineFillRecord = {
    ...rest,
    fg: record.fg ?? authoredColorSpec(),
    bg: record.bg ?? authoredColorSpec(),
  };
  if (gradient) {
    out.gradient = {
      ...gradient,
      stops: gradient.stops.map((stop) => ({ ...stop, color: stop.color ?? authoredColorSpec() })),
    };
  }
  return out;
}

const completeBorderSide = (side: BorderSide): EngineBorderSide => ({
  ...side,
  color: side.color ?? authoredColorSpec(),
});

export function completeBorderRecord(record: BorderRecord): EngineBorderRecord {
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
export function completeDxfRecord(record: DxfRecord): EngineDxfRecord {
  const { font, fill, border, ...rest } = record;
  const out: EngineDxfRecord = { ...rest };
  if (font) out.font = completeFontRecord(font);
  if (fill) out.fill = completeFillRecord(fill);
  if (border) out.border = completeBorderRecord(border);
  return out;
}
