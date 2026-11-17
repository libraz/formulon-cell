import { type CellXfInput, FontScheme } from '@libraz/formulon';
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
  FONT_SCHEME_BY_NAME,
  fillRecordFromFormat,
  fontRecordFromFormat,
  numFmtToFormatCode,
  readingOrderToTextDirection,
  textRotationToCellRotation,
  type WorkbookDefaultFont,
} from './format-writeback.js';
import { colorArgb, literalColor } from './style-records.js';
import type { CellXf, SheetFontPatch } from './types.js';
import { formatWithSheetFontPatch } from './whole-sheet-font-sync.js';
import type { WorkbookHandle } from './workbook-handle.js';

/** A raw alignment group can only be reused when every authored extended
 *  field has a representable source value. Without this guard the normalised
 *  style matcher would treat absent/255/explicit-default values as equal and
 *  copy stale raw metadata over a deliberate reset. */
function rawAlignmentHasAuthoredDifference(fmt: CellFormat, raw: CellXf): boolean {
  if (fmt.align !== undefined && raw.horizontal === null) return true;
  if (fmt.vAlign !== undefined && raw.vertical === null) return true;
  if (fmt.wrap !== undefined && raw.wrapText === null) return true;
  if (fmt.justifyLastLine !== undefined && raw.justifyLastLine === null) return true;
  if (fmt.rotation !== undefined) {
    if (raw.textRotation === null) return true;
    if (textRotationToCellRotation(raw.textRotation) !== fmt.rotation) return true;
  }
  if (fmt.indent !== undefined) {
    if (raw.indent === null || Math.round(fmt.indent) !== raw.indent) return true;
  }
  if (fmt.shrinkToFit !== undefined && raw.shrinkToFit !== fmt.shrinkToFit) return true;
  if (fmt.textDirection !== undefined) {
    if (raw.readingOrder === null) return true;
    if (readingOrderToTextDirection(raw.readingOrder) !== fmt.textDirection) return true;
  }
  return false;
}

/** An XF record under construction. */
type XfDraft = { -readonly [K in keyof CellXfInput]: CellXfInput[K] };

/** Assemble the XF record for a CellFormat, ensuring every component record
 *  exists first (the engine dedups on add). Returns null on engine failure. */
function buildXfForFormat(
  wb: WorkbookHandle,
  fmt: CellFormat,
  sheetFontPatch?: SheetFontPatch,
  workbookDefaultFont?: WorkbookDefaultFont | null,
  sourceFontIndex?: number,
): XfDraft | null {
  const effectiveFormat = formatWithSheetFontPatch(fmt, sheetFontPatch);
  const defaultFont =
    workbookDefaultFont === undefined ? wb.workbookDefaultFont : workbookDefaultFont;
  const fontRecord = fontRecordFromFormat(effectiveFormat, defaultFont);
  // Only a face change unlinks a font from the theme; the same face keeps its link.
  const sourceFont = sourceFontIndex === undefined ? undefined : wb.getFontRecord(sourceFontIndex);
  if (
    sourceFont &&
    sourceFont.name === fontRecord.name &&
    (sourceFont.scheme === FontScheme.Minor || sourceFont.scheme === FontScheme.Major)
  ) {
    fontRecord.scheme = sourceFont.scheme;
  }
  const fontIndex = wb.addFontRecord(fontRecord);
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
  return { ...buildXfRecord(fontIndex, fillIndex, borderIndex, numFmtId, fmt) };
}

const RAW_ALIGNMENT_FIELDS = [
  'horizontal',
  'vertical',
  'wrapText',
  'justifyLastLine',
  'textRotation',
  'indent',
  'relativeIndent',
  'shrinkToFit',
  'readingOrder',
] as const;

const formatHasGroupFields = (fmt: CellFormat, group: CellStyleFormatGroup): boolean =>
  CELL_STYLE_GROUP_FIELDS[group].some((field) => Object.hasOwn(fmt, field));

const FONT_EFFECT_FIELDS = [
  ['fontOutline', 'outline'],
  ['fontShadow', 'shadow'],
  ['fontCondense', 'condense'],
  ['fontExtend', 'extend'],
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
  record: XfDraft,
  source: CellXf,
  currentProjected: Partial<CellFormat>,
  effectiveFormat: CellFormat,
  workbookDefaultFont: WorkbookDefaultFont | null,
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
    font.scheme = effectiveFormat.fontScheme
      ? FONT_SCHEME_BY_NAME[effectiveFormat.fontScheme]
      : name === workbookDefaultFont?.name
        ? workbookDefaultFont?.scheme
        : undefined;
  }
  if (sizeChanged && effectiveFormat.fontSize !== undefined) {
    font.sizePt = fontCssPxToPoints(effectiveFormat.fontSize);
  }
  if (!font.name || font.sizePt === null || !Number.isFinite(font.sizePt) || font.sizePt <= 0)
    return false;
  const fontIndex = wb.addFontRecord(font);
  if (fontIndex < 0) return false;
  record.fontIndex = fontIndex;
  record.applyFont = source.applyFont;
  return true;
};

const copyRawGroup = (target: XfDraft, source: CellXf, group: CellStyleFormatGroup): void => {
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
    target.applyAlignment = source.applyAlignment;
    const targetRecord = target as Record<string, unknown>;
    for (const field of RAW_ALIGNMENT_FIELDS) targetRecord[field] = source[field];
  } else if (group === 'protection') {
    target.applyProtection = source.applyProtection;
    target.locked = source.locked;
    target.hidden = source.hidden;
  }
};

const copyRawAlignmentFieldsExcept = (
  target: XfDraft,
  source: CellXf,
  patch: Partial<CellFormat>,
): void => {
  const targetRecord = target as Record<string, unknown>;
  const owners: Readonly<Record<(typeof RAW_ALIGNMENT_FIELDS)[number], keyof CellFormat | null>> = {
    horizontal: 'align',
    vertical: 'vAlign',
    wrapText: 'wrap',
    justifyLastLine: 'justifyLastLine',
    textRotation: 'rotation',
    indent: 'indent',
    relativeIndent: null,
    shrinkToFit: 'shrinkToFit',
    readingOrder: 'textDirection',
  };
  for (const field of RAW_ALIGNMENT_FIELDS) {
    const owner = owners[field];
    if (owner !== null && Object.hasOwn(patch, owner)) continue;
    targetRecord[field] = source[field];
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
      font.scheme = patch.fontScheme
        ? FONT_SCHEME_BY_NAME[patch.fontScheme]
        : font.name === wb.workbookDefaultFont?.name
          ? wb.workbookDefaultFont?.scheme
          : undefined;
    }
    if (has('fontSize')) {
      const size = patch.fontSize;
      if (size !== undefined && Number.isFinite(size) && size > 0) {
        font.sizePt = fontCssPxToPoints(size);
      } else if (wb.workbookDefaultFont) {
        font.sizePt = wb.workbookDefaultFont.size;
      }
    }
    if (has('bold')) font.bold = patch.bold === true;
    if (has('italic')) font.italic = patch.italic === true;
    if (has('strike')) font.strike = patch.strike === true;
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
      // A literal selector replaces a source theme/indexed one, so an explicit
      // UI colour wins while every other raw font field stays intact.
      font.color = literalColor(
        patch.color === undefined ? 0xff000000 : (cssColorToArgb(patch.color) ?? 0xff000000),
      );
    }
    for (const [key, flag] of FONT_EFFECT_FIELDS) {
      if (!has(key)) continue;
      font[flag] = patch[key] === true;
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
      (sourceFill.gradient != null && (has('fill') || has('fillPattern') || has('fillGradient')));
    const fill = replacesFill ? { ...targetFill } : { ...sourceFill };
    const setLiteral = (slot: 'foreground' | 'background', argb: number): void => {
      fill[slot] = literalColor(argb);
    };
    if (!replacesFill) {
      if (has('fillPattern') || has('fill')) fill.pattern = targetFill.pattern;
      if (has('fillPatternColor')) setLiteral('foreground', colorArgb(targetFill.foreground));
      // `fill` is the visible solid foreground for a solid/empty pattern and
      // the background for a patterned fill. Change only that owned colour
      // pair; the other selector stays exactly as it was in the source record.
      if (has('fill'))
        setLiteral(
          targetFill.pattern > 1 ? 'background' : 'foreground',
          colorArgb(targetFill.pattern > 1 ? targetFill.background : targetFill.foreground),
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
  target.quotePrefix = source.quotePrefix;
  if (source.styleXfIndex === undefined) delete target.styleXfIndex;
  else target.styleXfIndex = source.styleXfIndex;
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
  const currentRaw = currentXfIndex === null ? null : (wb.getCellXf(currentXfIndex) ?? null);
  const record = buildXfForFormat(
    wb,
    fmt,
    sheetFontPatch,
    workbookDefaultFont,
    currentRaw?.fontIndex,
  );
  if (!record) return { xfIndex: -1, cacheable: false };
  const namedKey = fmt.cellStyle ?? 'normal';
  const namedXf = namedStyleXfs.get(namedKey);
  const mappedXfId =
    fmt.cellStyle === undefined ? styleXfIds.get('normal') : styleXfIds.get(fmt.cellStyle);
  if (strict && mappedXfId !== undefined && !namedXf) return { xfIndex: -1, cacheable: false };

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
  // A cell returning to Normal from another named style takes Normal's raw group
  // wherever the format leaves the group unset, never the old style's.
  const leavesNamedStyle =
    fmt.cellStyle === undefined &&
    currentRaw !== null &&
    mappedXfId !== undefined &&
    (currentRaw.styleXfIndex ?? 0) !== mappedXfId;
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
        currentRaw.locked === null &&
        currentRaw.hidden === null
      ) &&
      cellStyleGroupMatches(effectiveFormat, currentProjected, group);
    // An explicitly named style owns the groups it declares, so prefer its
    // raw selectors. An unstyled cell still inherits Normal semantically, but
    // its current XF may carry meaningful source metadata (apply flags,
    // explicit defaults), which must survive an unrelated edit.
    if (fmt.cellStyle !== undefined && namedMatch) copyRawGroup(record, namedXf.raw, group);
    else if (
      leavesNamedStyle &&
      namedXf?.groups.includes(group) === true &&
      !formatHasGroupFields(effectiveFormat, group)
    )
      copyRawGroup(record, namedXf.raw, group);
    else if (!formatHasGroupFields(effectiveFormat, group) && currentRaw !== null)
      copyRawGroup(record, currentRaw, group);
    else if (currentMatch) copyRawGroup(record, currentRaw, group);
    else if (namedMatch && !formatHasGroupFields(effectiveFormat, group))
      copyRawGroup(record, namedXf.raw, group);
  }
  const xfId = mappedXfId;
  const xfIndex = wb.addXfRecord(xfId === undefined ? record : { ...record, styleXfIndex: xfId });
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
