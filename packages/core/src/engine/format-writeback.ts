import type { CellBorderSide, CellBorders, CellFormat, NumFmt } from '../store/store.js';
import type { BorderRecord, BorderSide, CellXf, FillRecord, FontRecord } from './types.js';

/** Default font name + size used when a CellFormat omits these. Mirrors the
 *  workbook's default font; matches Excel's "Calibri 11". */
const DEFAULT_FONT_NAME = 'Calibri';
const DEFAULT_FONT_SIZE = 11;

/** ARGB packed value for opaque black. */
const BLACK_ARGB = 0xff000000;
/** ARGB packed value for "auto" (transparent — caller-defined default). */
const NO_COLOR_ARGB = 0;

/** OOXML border-style ordinal — 0=none, 1=thin, 2=medium, 3=dashed, 4=dotted,
 *  5=thick, 6=double. The full upstream ordinal list extends to 13 (slantDashDot)
 *  but we only emit the styles the UI exposes. */
const BORDER_STYLE_NONE = 0;
const BORDER_STYLE_THIN = 1;
const BORDER_STYLE_MEDIUM = 2;
const BORDER_STYLE_DASHED = 3;
const BORDER_STYLE_DOTTED = 4;
const BORDER_STYLE_THICK = 5;
const BORDER_STYLE_DOUBLE = 6;

/** OOXML alignment ordinals. Mirrors the JSDoc on `CellXf.horizontalAlign` /
 *  `CellXf.verticalAlign`. */
const HALIGN_GENERAL = 0;
const HALIGN_LEFT = 1;
const HALIGN_CENTER = 2;
const HALIGN_RIGHT = 3;
const VALIGN_TOP = 0;
const VALIGN_CENTER = 1;
const VALIGN_BOTTOM = 2;

/** Built-in numFmt id for "General" — every engine reserves 0 for this. */
const BUILTIN_NUM_FMT_GENERAL = 0;

/** Build a FontRecord from the cell's UI format fields. Missing fields fall
 *  back to the workbook default font. */
export function fontRecordFromFormat(fmt: CellFormat): FontRecord {
  return {
    name: fmt.fontFamily ?? DEFAULT_FONT_NAME,
    size: fmt.fontSize ?? DEFAULT_FONT_SIZE,
    bold: fmt.bold === true,
    italic: fmt.italic === true,
    strike: fmt.strike === true,
    underline: fmt.underline === true ? 1 : 0,
    colorArgb: fmt.color ? (cssColorToArgb(fmt.color) ?? BLACK_ARGB) : BLACK_ARGB,
  };
}

/** Build a FillRecord from the cell's `fill` field. No fill → solid pattern
 *  with auto color (engine treats this as "no fill" in OOXML). */
export function fillRecordFromFormat(fmt: CellFormat): FillRecord {
  if (!fmt.fill) {
    return { pattern: 0, fgArgb: NO_COLOR_ARGB, bgArgb: NO_COLOR_ARGB };
  }
  const fg = cssColorToArgb(fmt.fill) ?? NO_COLOR_ARGB;
  return { pattern: 1, fgArgb: fg, bgArgb: NO_COLOR_ARGB };
}

/** Build a BorderRecord from the cell's `borders` field. Missing sides → none. */
export function borderRecordFromFormat(fmt: CellFormat): BorderRecord {
  const b: CellBorders = fmt.borders ?? {};
  return {
    left: borderSideToRecord(b.left),
    right: borderSideToRecord(b.right),
    top: borderSideToRecord(b.top),
    bottom: borderSideToRecord(b.bottom),
    diagonal: borderSideToRecord(b.diagonalUp ?? b.diagonalDown),
    diagonalUp: !!b.diagonalUp,
    diagonalDown: !!b.diagonalDown,
  };
}

function borderSideToRecord(side: CellBorderSide | undefined): BorderSide {
  if (!side) return { style: BORDER_STYLE_NONE, colorArgb: BLACK_ARGB };
  if (side === true) return { style: BORDER_STYLE_THIN, colorArgb: BLACK_ARGB };
  let style = BORDER_STYLE_THIN;
  switch (side.style) {
    case 'medium':
      style = BORDER_STYLE_MEDIUM;
      break;
    case 'thick':
      style = BORDER_STYLE_THICK;
      break;
    case 'dashed':
      style = BORDER_STYLE_DASHED;
      break;
    case 'dotted':
      style = BORDER_STYLE_DOTTED;
      break;
    case 'double':
      style = BORDER_STYLE_DOUBLE;
      break;
    default:
      style = BORDER_STYLE_THIN;
  }
  const colorArgb = side.color ? (cssColorToArgb(side.color) ?? BLACK_ARGB) : BLACK_ARGB;
  return { style, colorArgb };
}

/** Translate a UI NumFmt to the format-code string Excel persists. Returns
 *  null when the format is "general" — caller should use built-in numFmtId 0. */
export function numFmtToFormatCode(fmt: NumFmt | undefined): string | null {
  if (!fmt || fmt.kind === 'general') return null;
  switch (fmt.kind) {
    case 'fixed': {
      const dec = '0'.repeat(fmt.decimals);
      const body = dec ? `0.${dec}` : '0';
      return fmt.thousands ? `#,##${body}` : body;
    }
    case 'currency': {
      const dec = '0'.repeat(fmt.decimals);
      const body = dec ? `0.${dec}` : '0';
      const sym = fmt.symbol ?? '$';
      return `"${sym}"#,##${body}`;
    }
    case 'percent': {
      const dec = '0'.repeat(fmt.decimals);
      return dec ? `0.${dec}%` : '0%';
    }
    case 'scientific': {
      const dec = '0'.repeat(fmt.decimals);
      return dec ? `0.${dec}E+00` : '0E+00';
    }
    case 'accounting': {
      const dec = '0'.repeat(fmt.decimals);
      const body = dec ? `0.${dec}` : '0';
      const sym = fmt.symbol ?? '$';
      return `_-"${sym}"* #,##${body}_-;-"${sym}"* #,##${body}_-;_-"${sym}"* "-"??_-;_-@_-`;
    }
    case 'date':
    case 'time':
    case 'datetime':
    case 'custom':
      return fmt.pattern;
    case 'text':
      return '@';
    default:
      return null;
  }
}

/** Map a UI horizontal-align string to its OOXML ordinal. */
export function halignOrdinal(align: CellFormat['align']): number {
  switch (align) {
    case 'left':
      return HALIGN_LEFT;
    case 'center':
      return HALIGN_CENTER;
    case 'right':
      return HALIGN_RIGHT;
    default:
      return HALIGN_GENERAL;
  }
}

/** Map a UI vertical-align string to its OOXML ordinal. Excel's default is
 *  bottom; we treat "middle" as ordinal 1 (center). */
export function valignOrdinal(vAlign: CellFormat['vAlign']): number {
  switch (vAlign) {
    case 'top':
      return VALIGN_TOP;
    case 'middle':
      return VALIGN_CENTER;
    default:
      return VALIGN_BOTTOM;
  }
}

/** Build a complete CellXf record from a CellFormat by resolving every
 *  sub-record. The caller passes pre-resolved indices for each sub-record. */
export function buildXfRecord(
  fontIndex: number,
  fillIndex: number,
  borderIndex: number,
  numFmtId: number,
  fmt: CellFormat,
): CellXf {
  return {
    fontIndex,
    fillIndex,
    borderIndex,
    numFmtId,
    horizontalAlign: halignOrdinal(fmt.align),
    verticalAlign: valignOrdinal(fmt.vAlign),
    wrapText: fmt.wrap === true,
  };
}

/* ---------- Inverse translators (used during hydrate) ---------- */

/** Translate an engine FontRecord back into the CellFormat font fields. */
export function fontRecordToFormat(rec: FontRecord): Partial<CellFormat> {
  const out: Partial<CellFormat> = {};
  if (rec.name !== DEFAULT_FONT_NAME) out.fontFamily = rec.name;
  if (rec.size !== DEFAULT_FONT_SIZE) out.fontSize = rec.size;
  if (rec.bold) out.bold = true;
  if (rec.italic) out.italic = true;
  if (rec.strike) out.strike = true;
  if (rec.underline > 0) out.underline = true;
  if (rec.colorArgb !== BLACK_ARGB && rec.colorArgb !== NO_COLOR_ARGB) {
    out.color = argbToCssColor(rec.colorArgb);
  }
  return out;
}

/** Translate an engine FillRecord back into the CellFormat fill field. */
export function fillRecordToFormat(rec: FillRecord): Partial<CellFormat> {
  if (rec.pattern === 0 || rec.fgArgb === NO_COLOR_ARGB) return {};
  return { fill: argbToCssColor(rec.fgArgb) };
}

/** Translate an engine BorderRecord back into the CellFormat borders field. */
export function borderRecordToFormat(rec: BorderRecord): Partial<CellFormat> {
  const sides: CellBorders = {};
  const map = (s: BorderSide): CellBorderSide | undefined => {
    if (s.style === BORDER_STYLE_NONE) return undefined;
    const style: NonNullable<Exclude<CellBorderSide, boolean | undefined>>['style'] = (() => {
      switch (s.style) {
        case BORDER_STYLE_MEDIUM:
          return 'medium';
        case BORDER_STYLE_THICK:
          return 'thick';
        case BORDER_STYLE_DASHED:
          return 'dashed';
        case BORDER_STYLE_DOTTED:
          return 'dotted';
        case BORDER_STYLE_DOUBLE:
          return 'double';
        default:
          return 'thin';
      }
    })();
    const out: { style: typeof style; color?: string } = { style };
    if (s.colorArgb !== BLACK_ARGB && s.colorArgb !== NO_COLOR_ARGB) {
      out.color = argbToCssColor(s.colorArgb);
    }
    return out;
  };
  const left = map(rec.left);
  const right = map(rec.right);
  const top = map(rec.top);
  const bottom = map(rec.bottom);
  if (left) sides.left = left;
  if (right) sides.right = right;
  if (top) sides.top = top;
  if (bottom) sides.bottom = bottom;
  if (rec.diagonalUp) sides.diagonalUp = map(rec.diagonal) ?? true;
  if (rec.diagonalDown) sides.diagonalDown = map(rec.diagonal) ?? true;
  if (Object.keys(sides).length === 0) return {};
  return { borders: sides };
}

/** Translate the OOXML format-code string back into a UI NumFmt. Conservative:
 *  unknown patterns surface as `custom`. */
export function formatCodeToNumFmt(code: string): NumFmt | null {
  if (!code || code === 'General') return null;
  if (code === '@') return { kind: 'text' };
  // Percent
  const pctMatch = code.match(/^0(?:\.(0+))?%$/);
  if (pctMatch) {
    return { kind: 'percent', decimals: pctMatch[1] ? pctMatch[1].length : 0 };
  }
  // Scientific
  const sciMatch = code.match(/^0(?:\.(0+))?E\+0+$/);
  if (sciMatch) {
    return { kind: 'scientific', decimals: sciMatch[1] ? sciMatch[1].length : 0 };
  }
  // Currency: "$#,##0.00" or "¥#,##0"
  const curMatch = code.match(/^"([^"]+)"#,##0(?:\.(0+))?$/);
  if (curMatch) {
    const symbol = curMatch[1] ?? '$';
    return { kind: 'currency', decimals: curMatch[2] ? curMatch[2].length : 0, symbol };
  }
  // Fixed with thousands
  const tFixMatch = code.match(/^#,##0(?:\.(0+))?$/);
  if (tFixMatch) {
    return { kind: 'fixed', decimals: tFixMatch[1] ? tFixMatch[1].length : 0, thousands: true };
  }
  // Plain fixed
  const fixMatch = code.match(/^0(?:\.(0+))?$/);
  if (fixMatch) {
    return { kind: 'fixed', decimals: fixMatch[1] ? fixMatch[1].length : 0 };
  }
  // Date / time tokens
  if (/[ymdhs]/i.test(code)) {
    if (/[ymd]/.test(code) && /[hs]/.test(code)) return { kind: 'datetime', pattern: code };
    if (/[hs]/.test(code)) return { kind: 'time', pattern: code };
    return { kind: 'date', pattern: code };
  }
  return { kind: 'custom', pattern: code };
}

/* ---------- Color helpers ---------- */

/** Parse a CSS color into an opaque ARGB packed integer. Supports #rgb,
 *  #rrggbb, #rrggbbaa, rgb()/rgba(), and a small set of named colors. Returns
 *  null when the color cannot be parsed. */
export function cssColorToArgb(color: string): number | null {
  const c = color.trim().toLowerCase();
  if (!c) return null;
  // Named-color shortcuts.
  const named = NAMED_COLORS[c];
  if (named !== undefined) return named;
  // #rgb / #rrggbb / #rrggbbaa
  if (c.startsWith('#')) {
    const hex = c.slice(1);
    if (hex.length === 3) {
      const h0 = hex[0] ?? '0';
      const h1 = hex[1] ?? '0';
      const h2 = hex[2] ?? '0';
      const r = Number.parseInt(h0 + h0, 16);
      const g = Number.parseInt(h1 + h1, 16);
      const b = Number.parseInt(h2 + h2, 16);
      return packArgb(0xff, r, g, b);
    }
    if (hex.length === 6) {
      const r = Number.parseInt(hex.slice(0, 2), 16);
      const g = Number.parseInt(hex.slice(2, 4), 16);
      const b = Number.parseInt(hex.slice(4, 6), 16);
      return packArgb(0xff, r, g, b);
    }
    if (hex.length === 8) {
      const r = Number.parseInt(hex.slice(0, 2), 16);
      const g = Number.parseInt(hex.slice(2, 4), 16);
      const b = Number.parseInt(hex.slice(4, 6), 16);
      const a = Number.parseInt(hex.slice(6, 8), 16);
      return packArgb(a, r, g, b);
    }
    return null;
  }
  // rgb()/rgba()
  const rgbMatch = c.match(/^rgba?\(([^)]+)\)$/);
  if (rgbMatch) {
    const inner = rgbMatch[1] ?? '';
    const parts = inner.split(',').map((s) => s.trim());
    if (parts.length < 3) return null;
    const r = clamp255(Number.parseFloat(parts[0] ?? '0'));
    const g = clamp255(Number.parseFloat(parts[1] ?? '0'));
    const b = clamp255(Number.parseFloat(parts[2] ?? '0'));
    const a = parts[3] !== undefined ? clamp01(Number.parseFloat(parts[3])) : 1;
    return packArgb(Math.round(a * 255), r, g, b);
  }
  return null;
}

/** Render an ARGB packed integer as a CSS color string. */
export function argbToCssColor(argb: number): string {
  const a = (argb >>> 24) & 0xff;
  const r = (argb >>> 16) & 0xff;
  const g = (argb >>> 8) & 0xff;
  const b = argb & 0xff;
  if (a === 0xff) {
    return `#${hex2(r)}${hex2(g)}${hex2(b)}`;
  }
  return `rgba(${r}, ${g}, ${b}, ${(a / 255).toFixed(2)})`;
}

function packArgb(a: number, r: number, g: number, b: number): number {
  return (((a & 0xff) << 24) | ((r & 0xff) << 16) | ((g & 0xff) << 8) | (b & 0xff)) >>> 0;
}

function clamp255(n: number): number {
  if (Number.isNaN(n)) return 0;
  return Math.max(0, Math.min(255, Math.round(n)));
}

function clamp01(n: number): number {
  if (Number.isNaN(n)) return 1;
  return Math.max(0, Math.min(1, n));
}

function hex2(n: number): string {
  return n.toString(16).padStart(2, '0');
}

/** Built-in numFmt id constant — exported so writeback can pass it as the
 *  fallback for `general` cells without a separate addNumFmt call. */
export { BUILTIN_NUM_FMT_GENERAL };

/** A compact subset of CSS named colors. We only ship the common Excel
 *  picker colors; unknown names fall through to null. */
const NAMED_COLORS: Record<string, number> = {
  black: 0xff000000,
  white: 0xffffffff,
  red: 0xffff0000,
  green: 0xff008000,
  blue: 0xff0000ff,
  yellow: 0xffffff00,
  cyan: 0xff00ffff,
  magenta: 0xffff00ff,
  gray: 0xff808080,
  grey: 0xff808080,
  orange: 0xffffa500,
  purple: 0xff800080,
  pink: 0xffffc0cb,
  brown: 0xffa52a2a,
  transparent: 0x00000000,
};
