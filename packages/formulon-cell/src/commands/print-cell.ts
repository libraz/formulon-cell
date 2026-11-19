import { numFmtToFormatCode } from '../engine/format-writeback.js';
import type { CellValue } from '../engine/types.js';
import { formatCell } from '../engine/value.js';
import { safeFontFamily } from '../format/font-family.js';
import { numberFormatPaint } from '../format/number-format.js';
import { gradientFillCss } from '../render/gradient-fill-css.js';
import { cellFontFamily } from '../render/painters/text.js';
import type { CellFormat, PrintCellErrorsMode } from '../store/store.js';
import { cssString } from '../text-escape.js';
import { formatA1FormulaAsR1C1 } from './refs.js';

const PRINT_FONT_FALLBACK = 'system-ui, -apple-system, "Segoe UI", sans-serif';

export const cellDisplay = (
  value: CellValue,
  formula: string | null,
  showFormulas: boolean,
  addr?: { row: number; col: number },
  r1c1 = false,
  cellErrorsAs: PrintCellErrorsMode = 'displayed',
  nativeR1C1Formula?: string | null,
  nativeDisplay?: string | null,
  format?: CellFormat,
): string => {
  if (showFormulas && formula) {
    if (!r1c1) return formula;
    return nativeR1C1Formula ?? (addr ? formatA1FormulaAsR1C1(formula, addr) : formula);
  }
  if (value.kind === 'blank') return format?.hyperlinkDisplay ?? '';
  if (value.kind === 'error') {
    if (cellErrorsAs === 'blank') return '';
    if (cellErrorsAs === 'dash') return '--';
    if (cellErrorsAs === 'na') return '#N/A';
  }
  if (nativeDisplay !== null && nativeDisplay !== undefined) return nativeDisplay;
  return formatCell(value);
};

const fillPatternCss = (pattern: CellFormat['fillPattern'], color = '#000000'): string => {
  switch (pattern) {
    case 'gray0625':
      return `radial-gradient(${color} 0.4px, transparent 0.4px)`;
    case 'gray125':
      return `radial-gradient(${color} 0.6px, transparent 0.6px)`;
    case 'gray25':
      return `radial-gradient(${color} 1px, transparent 1px)`;
    case 'gray50':
      return `repeating-linear-gradient(45deg, ${color} 0 2px, transparent 2px 4px)`;
    case 'gray75':
      return `repeating-linear-gradient(45deg, ${color} 0 3px, transparent 3px 4px)`;
    case 'horizontal':
    case 'darkHorizontal':
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'lightHorizontal':
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px 7px)`;
    case 'vertical':
    case 'darkVertical':
      return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 4px)`;
    case 'lightVertical':
      return `repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px 7px)`;
    case 'diagonalDown':
    case 'darkDown':
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 5px)`;
    case 'diagonalUp':
    case 'darkUp':
      return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 5px)`;
    case 'lightDown':
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px 9px)`;
    case 'lightUp':
      return `repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px 9px)`;
    case 'darkGrid':
    case 'lightGrid': {
      const step = pattern === 'darkGrid' ? 4 : 7;
      return `repeating-linear-gradient(0deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(90deg, ${color} 0 1px, transparent 1px ${step}px)`;
    }
    case 'darkTrellis':
    case 'lightTrellis': {
      const step = pattern === 'darkTrellis' ? 6 : 9;
      return `repeating-linear-gradient(45deg, ${color} 0 1px, transparent 1px ${step}px), repeating-linear-gradient(135deg, ${color} 0 1px, transparent 1px ${step}px)`;
    }
    default:
      return '';
  }
};

/** Translate a `CellFormat` into an inline-style string. Only the subset that
 *  matters for print is included — borders are emitted on the cell itself
 *  so colspan/rowspan merges retain their outline. */
export function inlineCellStyle(
  fmt: CellFormat | undefined,
  value: CellValue,
  showGridlines: boolean,
  blackAndWhite: boolean,
  scale = 1,
  fontEffects = false,
): string {
  const parts: string[] = [];
  if (showGridlines) {
    // Default hairline grid for unformatted cells. Per-side overrides below.
    parts.push(`border:${scale}px solid #c8c8c8`);
  }
  if (!fmt) return parts.join(';');
  if (fmt.bold) parts.push('font-weight:600');
  if (fmt.italic) parts.push('font-style:italic');
  if (fontEffects && fmt.fontShadow)
    parts.push(`text-shadow:${scale}px ${scale}px 0 rgba(0,0,0,0.35)`);
  if (fontEffects && fmt.fontOutline) {
    parts.push(`-webkit-text-stroke:${0.75 * scale}px currentColor`);
    parts.push('-webkit-text-fill-color:transparent');
  }
  if (fmt.underline || fmt.strike) {
    const decos: string[] = [];
    if (fmt.underline) decos.push('underline');
    if (fmt.strike) decos.push('line-through');
    parts.push(`text-decoration:${decos.join(' ')}`);
  }
  if (fmt.align) {
    const align =
      fmt.align === 'centerContinuous'
        ? 'center'
        : fmt.align === 'justify' || fmt.align === 'distributed'
          ? 'justify'
          : fmt.align === 'fill'
            ? 'left'
            : fmt.align;
    parts.push(`text-align:${align}`);
  }
  if (fmt.vAlign) {
    const v = fmt.vAlign === 'justify' || fmt.vAlign === 'distributed' ? 'middle' : fmt.vAlign;
    parts.push(`vertical-align:${v}`);
  }
  if (fmt.wrap) parts.push('white-space:pre-wrap');
  else parts.push('white-space:nowrap');
  if (typeof fmt.indent === 'number' && fmt.indent > 0) {
    parts.push(`padding-left:${(4 + fmt.indent * 8) * scale}px`);
  }
  if (typeof fmt.rotation === 'number' && fmt.rotation !== 0) {
    parts.push(`transform:rotate(${-fmt.rotation}deg);transform-origin:center`);
  }
  if (fmt.textDirection === 'rtl') parts.push('direction:rtl');
  else if (fmt.textDirection === 'ltr') parts.push('direction:ltr');
  const formatCode = value.kind === 'number' ? numFmtToFormatCode(fmt.numFmt) : null;
  const formatColor =
    formatCode && value.kind === 'number'
      ? numberFormatPaint(formatCode, value.value).color
      : undefined;
  const color = formatColor ?? fmt.color;
  if (color && !blackAndWhite) parts.push(`color:${color}`);
  if (fmt.fillGradient && !blackAndWhite) {
    parts.push(`background:${gradientFillCss(fmt.fillGradient)}`);
  } else if (fmt.fill && !blackAndWhite) parts.push(`background:${fmt.fill}`);
  if (fmt.fillPattern && !fmt.fillGradient && !blackAndWhite) {
    const image = fillPatternCss(fmt.fillPattern, fmt.fillPatternColor);
    if (image) parts.push(`background-image:${image}`);
    if (fmt.fillPattern === 'gray125' || fmt.fillPattern === 'gray25') {
      parts.push(`background-size:${4 * scale}px ${4 * scale}px`);
    }
  }
  if (fmt.fontFamily) {
    parts.push(
      `font-family:${safeFontFamily(cellFontFamily(fmt.fontFamily, PRINT_FONT_FALLBACK), cssString)}`,
    );
  }
  if (typeof fmt.fontSize === 'number') parts.push(`font-size:${fmt.fontSize * scale}px`);
  if (fmt.borders) {
    const sideToCss = (side: 'top' | 'right' | 'bottom' | 'left'): void => {
      const b = fmt.borders?.[side];
      if (!b) return;
      if (b === true) {
        parts.push(`border-${side}:${scale}px solid #000`);
        return;
      }
      const width =
        b.style === 'thick'
          ? 2
          : b.style === 'medium'
            ? 1.5
            : b.style === 'double'
              ? 3
              : b.style === 'dashed' || b.style === 'mediumDashed' || b.style === 'dashDot'
                ? 1
                : b.style === 'dotted'
                  ? 1
                  : 1;
      const style =
        b.style === 'double'
          ? 'double'
          : b.style === 'dashed' || b.style === 'mediumDashed' || b.style === 'dashDot'
            ? 'dashed'
            : 'solid';
      parts.push(
        `border-${side}:${width * scale}px ${style} ${blackAndWhite ? '#000' : (b.color ?? '#000')}`,
      );
    };
    sideToCss('top');
    sideToCss('right');
    sideToCss('bottom');
    sideToCss('left');
  }
  return parts.join(';');
}
