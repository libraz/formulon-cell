import type { CellStyleFormatGroup } from '../commands/cell-style-catalog.js';
import { CELL_STYLE_GROUP_FIELDS, cellStyleGroupMatches } from '../commands/cell-style-groups.js';
import { withFillReplacement } from '../store/format-merge.js';
import type { CellFormat } from '../store/store.js';
import { cellFormatFromXf, cellFormatFromXfWithStatus } from './cell-format-projection.js';
import type { NamedStyleXfResolution } from './cell-style-writeback.js';
import { fontCssPxToPoints } from './font-units.js';
import {
  BUILTIN_NUM_FMT_GENERAL,
  borderRecordFromFormat,
  buildXfRecord,
  cssColorToArgb,
  fillRecordFromFormat,
  fontRecordFromFormat,
  numFmtToFormatCode,
  readingOrderToTextDirection,
  textRotationToCellRotation,
} from './format-writeback.js';
import type { CellXf, ColorSpec, FontRecord, SheetFontPatch } from './types.js';
import { formatWithSheetFontPatch } from './whole-sheet-font-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** A raw alignment group can only be reused when every authored extended
 *  field has a representable source value. Without this guard the normalised
 *  style matcher would treat absent/255/explicit-default values as equal and
 *  copy stale raw metadata over a deliberate reset. */
function rawAlignmentHasAuthoredDifference(fmt: CellFormat, raw: CellXf): boolean {
  if (raw.hasAlignment === false) {
    return (
      fmt.align !== undefined ||
      fmt.vAlign !== undefined ||
      fmt.wrap !== undefined ||
      fmt.justifyLastLine !== undefined ||
      fmt.rotation !== undefined ||
      fmt.indent !== undefined ||
      fmt.shrinkToFit !== undefined ||
      fmt.textDirection !== undefined
    );
  }
  if (fmt.align !== undefined && raw.hasHorizontalAlign !== true) return true;
  if (fmt.vAlign !== undefined && raw.hasVerticalAlign !== true) return true;
  if (fmt.wrap !== undefined && raw.hasWrapText !== true) return true;
  if (fmt.justifyLastLine !== undefined && raw.hasJustifyLastLine !== true) return true;
  if (fmt.rotation !== undefined) {
    if (raw.textRotation === undefined) return true;
    if (raw.textRotation === 255) return true;
    if (textRotationToCellRotation(raw.textRotation) !== fmt.rotation) return true;
  }
  if (fmt.indent !== undefined) {
    if (raw.indent === undefined || Math.round(fmt.indent) !== raw.indent) return true;
  }
  if (fmt.shrinkToFit !== undefined && raw.shrinkToFit !== fmt.shrinkToFit) return true;
  if (fmt.textDirection !== undefined) {
    if (raw.readingOrder === undefined) return true;
    if (readingOrderToTextDirection(raw.readingOrder) !== fmt.textDirection) return true;
  }
  return false;
}

/** Assemble the XF record for a CellFormat, ensuring every component record
 *  exists first (the engine dedups on add). Returns null on engine failure. */
function buildXfForFormat(
  wb: WorkbookHandle,
  fmt: CellFormat,
  sheetFontPatch?: SheetFontPatch,
  workbookDefaultFont?: Pick<FontRecord, 'name' | 'size' | 'scheme'> | null,
): CellXf | null {
  const effectiveFormat = formatWithSheetFontPatch(fmt, sheetFontPatch);
  const defaultFont =
    workbookDefaultFont === undefined ? wb.workbookDefaultFont : workbookDefaultFont;
  const fontIndex = wb.addFontRecord(fontRecordFromFormat(effectiveFormat, defaultFont));
  if (fontIndex < 0) return null;
  const fillIndex = wb.addFillRecord(fillRecordFromFormat(fmt));
  if (fillIndex < 0) return null;
  const borderIndex = wb.addBorderRecord(borderRecordFromFormat(fmt));
  if (borderIndex < 0) return null;
  const code = numFmtToFormatCode(fmt.numFmt, wb);
  let numFmtId = BUILTIN_NUM_FMT_GENERAL;
  if (code !== null) {
    const id = wb.addNumFmtCode(code);
    if (id < 0) return null;
    numFmtId = id;
  }
  return buildXfRecord(fontIndex, fillIndex, borderIndex, numFmtId, fmt);
}

const RAW_ALIGNMENT_OPTIONAL_FIELDS = [
  'justifyLastLine',
  'hasAlignment',
  'hasHorizontalAlign',
  'hasVerticalAlign',
  'hasWrapText',
  'hasJustifyLastLine',
  'textRotation',
  'indent',
  'relativeIndent',
  'shrinkToFit',
  'readingOrder',
] as const;

const formatHasGroupFields = (fmt: CellFormat, group: CellStyleFormatGroup): boolean =>
  CELL_STYLE_GROUP_FIELDS[group].some((field) => Object.hasOwn(fmt, field));

const FONT_EFFECT_FIELDS = [
  ['fontOutline', 'outline', 'hasOutline'],
  ['fontShadow', 'shadow', 'hasShadow'],
  ['fontCondense', 'condense', 'hasCondense'],
  ['fontExtend', 'extend', 'hasExtend'],
] as const;

const sameFontNonAxisField = (
  format: CellFormat,
  current: Partial<CellFormat>,
  field: keyof CellFormat,
): boolean => {
  // A sparse direct format changes only authored fields. An omitted bold,
  // colour, or vertical-alignment axis inherits the raw source value and
  // therefore must not force a generic font rebuild that drops native
  // metadata. An explicitly authored `false`/`undefined` remains a real
  // reset and is compared below.
  if (!Object.hasOwn(format, field)) return true;
  const left = format[field];
  const right = current[field];
  if (
    field === 'bold' ||
    field === 'italic' ||
    field === 'strike' ||
    FONT_EFFECT_FIELDS.some(([key]) => key === field)
  ) {
    return (left === true) === (right === true);
  }
  return Object.is(left, right);
};

const fontNonAxisMatches = (format: CellFormat, current: Partial<CellFormat>): boolean =>
  CELL_STYLE_GROUP_FIELDS.font.every(
    (field) =>
      field === 'fontFamily' ||
      field === 'fontSize' ||
      sameFontNonAxisField(format, current, field),
  );

/** Preserve the raw inherited font record when a direct write changes only
 *  family and/or size. The raw record carries theme, charset, colour and
 *  apply metadata that CellFormat intentionally does not model. */
const preserveRawFontForAxisEdit = (
  wb: WorkbookHandle,
  record: CellXf,
  source: CellXf,
  currentProjected: Partial<CellFormat>,
  effectiveFormat: CellFormat,
  workbookDefaultFont: Pick<FontRecord, 'name' | 'size' | 'scheme'> | null,
): boolean => {
  if (!fontNonAxisMatches(effectiveFormat, currentProjected)) return true;
  const familyChanged =
    Object.hasOwn(effectiveFormat, 'fontFamily') &&
    effectiveFormat.fontFamily !== currentProjected.fontFamily;
  const sizeChanged =
    Object.hasOwn(effectiveFormat, 'fontSize') &&
    effectiveFormat.fontSize !== currentProjected.fontSize;
  if (!familyChanged && !sizeChanged) return true;
  const baseFont = wb.getFontRecord(source.fontIndex);
  if (!baseFont) return false;
  const font = { ...baseFont };
  if (familyChanged) {
    const name = effectiveFormat.fontFamily;
    if (name !== undefined) font.name = name;
    font.scheme = name === workbookDefaultFont?.name ? workbookDefaultFont?.scheme : undefined;
  }
  if (sizeChanged && effectiveFormat.fontSize !== undefined) {
    font.size = fontCssPxToPoints(effectiveFormat.fontSize);
  }
  if (!font.name || !Number.isFinite(font.size) || font.size <= 0) return false;
  const fontIndex = wb.addFontRecord(font);
  if (fontIndex < 0) return false;
  record.fontIndex = fontIndex;
  record.applyFont = source.applyFont;
  return true;
};

const copyRawGroup = (target: CellXf, source: CellXf, group: CellStyleFormatGroup): void => {
  if (group === 'font') {
    target.fontIndex = source.fontIndex;
    target.applyFont = source.applyFont;
  } else if (group === 'fill') {
    target.fillIndex = source.fillIndex;
    target.applyFill = source.applyFill;
  } else if (group === 'border') {
    target.borderIndex = source.borderIndex;
    target.applyBorder = source.applyBorder;
  } else if (group === 'number') {
    target.numFmtId = source.numFmtId;
    target.applyNumberFormat = source.applyNumberFormat;
  } else if (group === 'alignment') {
    target.horizontalAlign = source.horizontalAlign;
    target.verticalAlign = source.verticalAlign;
    target.wrapText = source.wrapText;
    target.applyAlignment = source.applyAlignment;
    const targetRecord = target as unknown as Record<string, unknown>;
    const sourceRecord = source as unknown as Record<string, unknown>;
    for (const field of RAW_ALIGNMENT_OPTIONAL_FIELDS) {
      if (Object.hasOwn(source, field)) {
        targetRecord[field] = sourceRecord[field];
      } else {
        delete targetRecord[field];
      }
    }
  } else if (group === 'protection') {
    target.applyProtection = source.applyProtection;
    target.hasProtection = source.hasProtection;
    target.locked = source.locked;
    target.hidden = source.hidden;
  }
};

const literalColorSpec = (): ColorSpec => ({ kind: 0, rgb: 0, theme: 0, tint: 0, indexed: 0 });

const copyRawAlignmentFieldsExcept = (
  target: CellXf,
  source: CellXf,
  patch: Partial<CellFormat>,
): void => {
  const targetRecord = target as unknown as Record<string, unknown>;
  const sourceRecord = source as unknown as Record<string, unknown>;
  const owners: Readonly<Record<string, keyof CellFormat | undefined>> = {
    horizontalAlign: 'align',
    hasHorizontalAlign: 'align',
    verticalAlign: 'vAlign',
    hasVerticalAlign: 'vAlign',
    wrapText: 'wrap',
    hasWrapText: 'wrap',
    justifyLastLine: 'justifyLastLine',
    hasJustifyLastLine: 'justifyLastLine',
    textRotation: 'rotation',
    indent: 'indent',
    shrinkToFit: 'shrinkToFit',
    readingOrder: 'textDirection',
  };
  for (const [field, owner] of Object.entries(owners)) {
    // `hasAlignment` is derived from all alignment fields. Once any one of
    // them is authored, the target builder owns that flag; copying a source
    // false would otherwise suppress the newly authored field.
    if (Object.hasOwn(patch, owner as keyof CellFormat)) continue;
    if (Object.hasOwn(sourceRecord, field)) targetRecord[field] = sourceRecord[field];
    else delete targetRecord[field];
  }
  const alignmentPatched = [
    'align',
    'vAlign',
    'wrap',
    'justifyLastLine',
    'shrinkToFit',
    'indent',
    'rotation',
    'textDirection',
  ].some((field) => Object.hasOwn(patch, field));
  if (!alignmentPatched) {
    if (Object.hasOwn(sourceRecord, 'hasAlignment')) {
      targetRecord.hasAlignment = sourceRecord.hasAlignment;
    } else {
      delete targetRecord.hasAlignment;
    }
  } else if (targetRecord.hasAlignment === false && sourceRecord.hasAlignment === false) {
    // Keep an explicit empty alignment child when the patch itself clears the
    // last authored alignment field. Otherwise the target builder's `true`
    // presence bit remains authoritative for the newly patched field.
    targetRecord.hasAlignment = false;
  }
  for (const field of RAW_ALIGNMENT_OPTIONAL_FIELDS) {
    if (Object.hasOwn(owners, field)) continue;
    if (field === 'hasAlignment') continue;
    if (Object.hasOwn(sourceRecord, field)) targetRecord[field] = sourceRecord[field];
    else delete targetRecord[field];
  }
};

/**
 * Build a new XF by changing only the fields in `patch` while retaining the
 * raw source record for every untouched style group.  Axis styles are native
 * cell-style XFs rather than CellFormat entries, so routing them through the
 * ordinary sparse writer would otherwise lose apply flags, charset/theme
 * selectors, quotePrefix, or extended alignment metadata.
 */
export function patchXfIndexFromBase(
  wb: WorkbookHandle,
  baseXfIndex: number,
  patch: Partial<CellFormat>,
): number {
  const source = wb.getCellXf(baseXfIndex);
  if (!source) return -1;
  const base = cellFormatFromXf(wb, source, wb.workbookDefaultFont);
  const effective: CellFormat = { ...base, ...patch };
  if (patch.borders) effective.borders = { ...(base.borders ?? {}), ...patch.borders };
  for (const key of Object.keys(patch) as Array<keyof CellFormat>) {
    if (patch[key] === undefined) delete (effective as Partial<CellFormat>)[key];
  }
  withFillReplacement(effective, patch);

  const target = buildXfForFormat(wb, effective);
  if (!target) return -1;
  const has = (key: keyof CellFormat): boolean => Object.hasOwn(patch, key);
  const groups: readonly CellStyleFormatGroup[] = [
    'number',
    'alignment',
    'font',
    'border',
    'fill',
    'protection',
  ];
  const touched = new Set<CellStyleFormatGroup>();
  for (const group of groups) {
    if (CELL_STYLE_GROUP_FIELDS[group].some((key) => has(key))) touched.add(group);
  }
  for (const group of groups) {
    if (!touched.has(group)) copyRawGroup(target, source, group);
  }

  if (touched.has('font')) {
    const sourceFont = wb.getFontRecord(source.fontIndex);
    if (!sourceFont) return -1;
    const font = { ...sourceFont };
    if (has('fontFamily')) {
      const defaultFont = wb.workbookDefaultFont;
      font.name = patch.fontFamily ?? defaultFont?.name ?? font.name;
      font.scheme =
        font.name === wb.workbookDefaultFont?.name ? wb.workbookDefaultFont?.scheme : undefined;
    }
    if (has('fontSize')) {
      const size = patch.fontSize;
      if (size !== undefined && Number.isFinite(size) && size > 0) {
        font.size = fontCssPxToPoints(size);
      } else if (wb.workbookDefaultFont) {
        font.size = wb.workbookDefaultFont.size;
      }
    }
    if (has('bold')) {
      font.bold = patch.bold === true;
      font.hasBold = true;
    }
    if (has('italic')) {
      font.italic = patch.italic === true;
      font.hasItalic = true;
    }
    if (has('strike')) {
      font.strike = patch.strike === true;
      font.hasStrike = true;
    }
    if (has('underline')) {
      font.underline =
        patch.underline === true || patch.underline === 'single'
          ? 1
          : patch.underline === 'double'
            ? 2
            : patch.underline === 'singleAccounting'
              ? 3
              : patch.underline === 'doubleAccounting'
                ? 4
                : 0;
    }
    if (has('fontVertAlign')) {
      font.vertAlign =
        patch.fontVertAlign === 'superscript' ? 1 : patch.fontVertAlign === 'subscript' ? 2 : 0;
    }
    if (has('color')) {
      font.colorArgb =
        patch.color === undefined ? 0xff000000 : (cssColorToArgb(patch.color) ?? 0xff000000);
      // A source theme/indexed selector is authoritative over colorArgb in
      // the native writer. Replacing it with kind=0 makes an explicit UI
      // colour win while leaving every other raw font field intact.
      font.color = literalColorSpec();
    }
    for (const [key, flag, presence] of FONT_EFFECT_FIELDS) {
      if (!has(key)) continue;
      font[flag] = patch[key] === true;
      font[presence] = true;
    }
    const fontIndex = wb.addFontRecord(font);
    if (fontIndex < 0) return -1;
    target.fontIndex = fontIndex;
    target.applyFont = true;
  }
  if (touched.has('fill')) {
    const sourceFill = wb.getFillRecord(source.fillIndex);
    const targetFill = wb.getFillRecord(target.fillIndex);
    if (!sourceFill || !targetFill) return -1;
    // A gradient replaces the whole fill, and a solid or pattern fill
    // replaces a gradient, so either change takes the rebuilt record.
    const replacesFill =
      patch.fillGradient !== undefined ||
      (sourceFill.gradient !== undefined &&
        (has('fill') || has('fillPattern') || has('fillGradient')));
    const fill = replacesFill ? { ...targetFill } : { ...sourceFill };
    const fillRecord = fill as unknown as Record<string, unknown>;
    const setLiteral = (slot: 'fg' | 'bg', argb: number): void => {
      fillRecord[slot] = literalColorSpec();
      fillRecord[`${slot}Argb`] = argb;
    };
    if (!replacesFill) {
      if (has('fillPattern') || has('fill')) fill.pattern = targetFill.pattern;
      if (has('fillPatternColor')) setLiteral('fg', targetFill.fgArgb);
      // `fill` is the visible solid foreground for a solid/empty pattern and
      // the background for a patterned fill. Change only that owned colour
      // pair; the other selector stays exactly as it was in the source record.
      if (has('fill'))
        setLiteral(
          targetFill.pattern > 1 ? 'bg' : 'fg',
          targetFill.pattern > 1 ? targetFill.bgArgb : targetFill.fgArgb,
        );
    }
    const fillIndex = wb.addFillRecord(fill);
    if (fillIndex < 0) return -1;
    target.fillIndex = fillIndex;
    target.applyFill = true;
  }
  if (touched.has('border')) {
    const sourceBorder = wb.getBorderRecord(source.borderIndex);
    const targetBorder = wb.getBorderRecord(target.borderIndex);
    if (!sourceBorder || !targetBorder) return -1;
    const border = { ...sourceBorder };
    const borderPatch = patch.borders ?? {};
    for (const side of ['left', 'right', 'top', 'bottom'] as const) {
      if (Object.hasOwn(borderPatch, side)) border[side] = targetBorder[side];
    }
    if (Object.hasOwn(borderPatch, 'diagonalDown')) {
      border.diagonalDown = targetBorder.diagonalDown;
      if (borderPatch.diagonalDown) border.diagonal = targetBorder.diagonal;
    }
    if (Object.hasOwn(borderPatch, 'diagonalUp')) {
      border.diagonalUp = targetBorder.diagonalUp;
      if (borderPatch.diagonalUp) border.diagonal = targetBorder.diagonal;
    }
    const borderIndex = wb.addBorderRecord(border);
    if (borderIndex < 0) return -1;
    target.borderIndex = borderIndex;
    target.applyBorder = true;
  }
  if (touched.has('alignment')) copyRawAlignmentFieldsExcept(target, source, patch);
  // These fields are not style groups, but they are meaningful OOXML
  // metadata and must survive an axis edit verbatim.
  const targetRecord = target as unknown as Record<string, unknown>;
  const sourceRecord = source as unknown as Record<string, unknown>;
  for (const key of ['quotePrefix', 'xfId'] as const) {
    if (Object.hasOwn(sourceRecord, key)) targetRecord[key] = sourceRecord[key];
    else delete targetRecord[key];
  }
  return wb.addXfRecord(target);
}

/** Resolve a CellFormat to an engine xfIndex. `styleXfIds` maps a named style
 *  to its `<cellStyleXfs>` row, so a styled cell records which style it came
 *  from instead of collapsing into anonymous direct formatting. Returns -1 on
 *  engine failure. */
export function resolveXfForFormat(
  wb: WorkbookHandle,
  fmt: CellFormat,
  styleXfIds: ReadonlyMap<string, number>,
  namedStyleXfs: ReadonlyMap<string, NamedStyleXfResolution>,
  strict: boolean,
  currentXfIndex: number | null,
  sheetFontPatch?: SheetFontPatch,
): { readonly xfIndex: number; readonly cacheable: boolean } {
  const effectiveFormat = formatWithSheetFontPatch(fmt, sheetFontPatch);
  const workbookDefaultFont = wb.workbookDefaultFont;
  const record = buildXfForFormat(wb, fmt, sheetFontPatch, workbookDefaultFont);
  if (!record) return { xfIndex: -1, cacheable: false };
  const namedKey = fmt.cellStyle ?? 'normal';
  const namedXf = namedStyleXfs.get(namedKey);
  const mappedXfId =
    fmt.cellStyle === undefined ? styleXfIds.get('normal') : styleXfIds.get(fmt.cellStyle);
  if (strict && mappedXfId !== undefined && !namedXf) return { xfIndex: -1, cacheable: false };

  const currentRaw = currentXfIndex === null ? null : (wb.getCellXf(currentXfIndex) ?? null);
  const currentProjection = currentRaw
    ? cellFormatFromXfWithStatus(wb, currentRaw, workbookDefaultFont, null)
    : null;
  const currentProjected = currentProjection?.format ?? null;
  let fontPreservationComplete = true;
  if (currentRaw && currentProjected) {
    fontPreservationComplete = preserveRawFontForAxisEdit(
      wb,
      record,
      currentRaw,
      currentProjected,
      effectiveFormat,
      workbookDefaultFont,
    );
  }
  // quotePrefix has no CellFormat counterpart. Carry it through an edit as
  // opaque XF metadata while the current cell still points at its source row.
  if (currentRaw && Object.hasOwn(currentRaw, 'quotePrefix')) {
    record.quotePrefix = currentRaw.quotePrefix;
  }
  const groups: readonly CellStyleFormatGroup[] = [
    'number',
    'alignment',
    'font',
    'border',
    'fill',
    'protection',
  ];
  for (const group of groups) {
    const namedMatch =
      namedXf?.groups.includes(group) === true &&
      !(group === 'alignment' && rawAlignmentHasAuthoredDifference(effectiveFormat, namedXf.raw)) &&
      cellStyleGroupMatches(effectiveFormat, namedXf.projected, group);
    const currentMatch =
      currentRaw !== null &&
      currentProjected !== null &&
      !(group === 'alignment' && rawAlignmentHasAuthoredDifference(effectiveFormat, currentRaw)) &&
      !(
        group === 'protection' &&
        formatHasGroupFields(fmt, 'protection') &&
        currentRaw.hasProtection !== true
      ) &&
      cellStyleGroupMatches(effectiveFormat, currentProjected, group);
    // An explicitly named style owns the groups it declares, so prefer its
    // raw selectors. An unstyled cell still inherits Normal semantically, but
    // its current XF may carry meaningful source metadata (apply flags,
    // explicit defaults), which must survive an unrelated edit.
    if (fmt.cellStyle !== undefined && namedMatch) copyRawGroup(record, namedXf.raw, group);
    else if (!formatHasGroupFields(effectiveFormat, group) && currentRaw !== null)
      copyRawGroup(record, currentRaw, group);
    else if (currentMatch) copyRawGroup(record, currentRaw, group);
    else if (namedMatch && !formatHasGroupFields(effectiveFormat, group))
      copyRawGroup(record, namedXf.raw, group);
  }
  const xfId = mappedXfId;
  const xfIndex = wb.addXfRecord(xfId === undefined ? record : { ...record, xfId });
  return {
    xfIndex,
    cacheable:
      xfIndex >= 0 &&
      currentRaw !== null &&
      currentXfIndex !== null &&
      Number.isInteger(currentXfIndex) &&
      currentXfIndex >= 0 &&
      currentProjection?.complete === true &&
      fontPreservationComplete &&
      workbookDefaultFont !== null &&
      workbookDefaultFont !== undefined,
  };
}

/** Resolve a named style's own formatting to a `<cellStyleXfs>` row. */
export function resolveStyleXf(wb: WorkbookHandle, format: CellFormat): number {
  const record = buildXfForFormat(wb, format);
  return record ? wb.addCellStyleXfRecord(record) : -1;
}
