import { addrKey } from '../engine/address.js';
import { fontPointsToCssPx } from '../engine/font-units.js';
import { cssColorFromArgb } from '../engine/theme-color.js';
import type { CellXf } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { CellFormat, State } from '../store/store.js';
import { hyperlinkCoverAt } from './hyperlinks.js';
import { isCellWritable, isSheetProtected, protectedSheetPermissions } from './protection.js';

const HYPERLINK_BUILTIN_ID = 8;
const CUSTOM_BUILTIN_ID = 0xffffffff;
export const AUTO_HYPERLINK_COMMAND_ID = 'editor.autoHyperlink';

/** Return the user's exact text when it is a complete absolute web URL. */
export function detectTypedWebUrl(raw: string): string | null {
  if (raw.startsWith("'") || /\s/.test(raw) || !/^https?:\/\//i.test(raw)) return null;

  try {
    const url = new URL(raw);
    if ((url.protocol === 'http:' || url.protocol === 'https:') && url.hostname !== '') {
      return raw;
    }
  } catch {
    return null;
  }

  return null;
}

type NamedCellStyle = NonNullable<ReturnType<WorkbookHandle['getCellStyle']>>;

type HyperlinkStyleLookup = {
  readonly found: boolean;
  readonly resolved: { name: string; xfId: number; builtinId: number; xf: CellXf } | null;
};

function findHyperlinkStyle(wb: WorkbookHandle): HyperlinkStyleLookup {
  if (wb.capabilities.cellStyles !== true) return { found: false, resolved: null };

  const styleCount = wb.cellStyleCount();
  let firstCustomNameMatch: NamedCellStyle | null = null;
  let firstBuiltinHyperlink: NamedCellStyle | null = null;

  for (let index = 0; index < styleCount; index += 1) {
    const style = wb.getCellStyle(index);
    if (!style) continue;

    if (style.builtinId === HYPERLINK_BUILTIN_ID) {
      firstBuiltinHyperlink = style;
      break;
    }

    const normalizedName = style.name.trim().toLowerCase();
    const isNativeCustomStyle = style.builtinId === CUSTOM_BUILTIN_ID || style.builtinId === -1;
    if (firstCustomNameMatch === null && normalizedName === 'hyperlink' && isNativeCustomStyle) {
      firstCustomNameMatch = style;
    }
  }

  const chosen = firstBuiltinHyperlink ?? firstCustomNameMatch;
  if (!chosen) return { found: false, resolved: null };
  if (!Number.isInteger(chosen.xfId) || chosen.xfId < 0) return { found: true, resolved: null };

  const xf = wb.getCellStyleXf(chosen.xfId);
  if (!xf) return { found: true, resolved: null };

  return {
    found: true,
    resolved: {
      name: chosen.name,
      xfId: chosen.xfId,
      builtinId: chosen.builtinId,
      xf,
    },
  };
}

/** Resolve the workbook's preferred native Hyperlink named-style XF. */
export function resolveHyperlinkStyleXf(
  wb: WorkbookHandle,
): { name: string; xfId: number; builtinId: number; xf: CellXf } | null {
  return findHyperlinkStyle(wb).resolved;
}

const underlineFromOrdinal = (ordinal: number | undefined): CellFormat['underline'] => {
  switch (ordinal) {
    case 1:
      return 'single';
    case 2:
      return 'double';
    case 3:
      return 'singleAccounting';
    case 4:
      return 'doubleAccounting';
    default:
      return false;
  }
};

function fontPatchFromNamedStyle(wb: WorkbookHandle, xf: CellXf): Partial<CellFormat> | null {
  const font = wb.getResolvedFontRecord(xf.fontIndex);
  if (
    !font?.name ||
    !Number.isFinite(font.size) ||
    font.size <= 0 ||
    !Number.isFinite(font.colorArgb)
  ) {
    return null;
  }
  const color = cssColorFromArgb(font.colorArgb);
  return {
    fontFamily: font.name,
    fontSize: fontPointsToCssPx(font.size),
    bold: font.bold ?? false,
    italic: font.italic ?? false,
    strike: font.strike ?? false,
    underline: underlineFromOrdinal(font.underline),
    fontVertAlign:
      font.vertAlign === 1 ? 'superscript' : font.vertAlign === 2 ? 'subscript' : undefined,
    fontOutline: font.outline ?? false,
    fontShadow: font.shadow ?? false,
    fontCondense: font.condense ?? false,
    fontExtend: font.extend ?? false,
    color,
  };
}

function defaultHyperlinkFontPatch(wb: WorkbookHandle): Partial<CellFormat> | null {
  const resolvedDefault = wb.getResolvedFontRecord(0);
  const rawDefault = resolvedDefault ?? wb.getFontRecord(0);
  const displayFont = wb.workbookDisplayFont;
  const name = displayFont?.name ?? rawDefault?.name;
  const size = displayFont?.size ?? rawDefault?.size;
  const hyperlinkColor = wb.getTheme()?.colors[10];
  if (
    !name ||
    typeof size !== 'number' ||
    !Number.isFinite(size) ||
    size <= 0 ||
    typeof hyperlinkColor !== 'number' ||
    !Number.isFinite(hyperlinkColor)
  ) {
    return null;
  }
  return {
    fontFamily: name,
    fontSize: fontPointsToCssPx(size),
    bold: resolvedDefault?.bold ?? rawDefault?.bold ?? false,
    italic: resolvedDefault?.italic ?? rawDefault?.italic ?? false,
    strike: resolvedDefault?.strike ?? rawDefault?.strike ?? false,
    underline: 'single',
    fontVertAlign:
      resolvedDefault?.vertAlign === 1 || rawDefault?.vertAlign === 1
        ? 'superscript'
        : resolvedDefault?.vertAlign === 2 || rawDefault?.vertAlign === 2
          ? 'subscript'
          : undefined,
    fontOutline: resolvedDefault?.outline ?? rawDefault?.outline ?? false,
    fontShadow: resolvedDefault?.shadow ?? rawDefault?.shadow ?? false,
    fontCondense: resolvedDefault?.condense ?? rawDefault?.condense ?? false,
    fontExtend: resolvedDefault?.extend ?? rawDefault?.extend ?? false,
    color: cssColorFromArgb(hyperlinkColor),
  };
}

/** Build an exact one-cell font/link patch only while the typed value still matches. */
export function getTypedWebUrlFormatPatch(
  wb: WorkbookHandle,
  state: State,
  addr: { sheet: number; row: number; col: number },
  raw: string,
): Partial<CellFormat> | null {
  try {
    if (detectTypedWebUrl(raw) !== raw) return null;
    if (!isCellWritable(state, addr)) return null;
    if (isSheetProtected(state, addr.sheet)) {
      const permissions = protectedSheetPermissions(state, addr.sheet);
      if (permissions?.formatCells !== true || permissions.insertHyperlinks !== true) return null;
    }

    const currentValue = wb.getValue(addr);
    if (currentValue.kind !== 'text' || currentValue.value !== raw || wb.cellFormula(addr) !== null)
      return null;

    const coveringLink = hyperlinkCoverAt(state, addr);
    if (coveringLink) {
      if (addrKey(coveringLink.addr) !== addrKey(addr)) return null;
      const span = coveringLink.fmt.hyperlinkSpan;
      if (span && (span.lastRow > addr.row || span.lastCol > addr.col)) return null;
    }

    const lookup = findHyperlinkStyle(wb);
    const fontPatch = lookup.found
      ? lookup.resolved
        ? fontPatchFromNamedStyle(wb, lookup.resolved.xf)
        : null
      : defaultHyperlinkFontPatch(wb);
    if (!fontPatch) return null;

    return {
      ...fontPatch,
      hyperlink: raw,
      hyperlinkLocation: undefined,
      hyperlinkDisplay: undefined,
      hyperlinkTooltip: undefined,
      hyperlinkSpan: undefined,
    };
  } catch {
    return null;
  }
}
