import { effectiveCellFormat } from '../engine/cell-format-projection.js';
import type { Range } from '../engine/types.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { State } from '../store/store.js';

/** The rules of the desktop accessibility checker, in its pane order. */
export type AccessibilityRuleId =
  | 'contrast'
  | 'redFormat'
  | 'altText'
  | 'tableHeader'
  | 'mergedCells'
  | 'defaultSheetName'
  | 'restrictedAccess';

export type AccessibilitySectionId = 'colorContrast' | 'media' | 'tables' | 'structure' | 'access';

export const ACCESSIBILITY_SECTIONS: readonly {
  readonly id: AccessibilitySectionId;
  readonly rules: readonly AccessibilityRuleId[];
}[] = [
  { id: 'colorContrast', rules: ['contrast', 'redFormat'] },
  { id: 'media', rules: ['altText'] },
  { id: 'tables', rules: ['tableHeader', 'mergedCells'] },
  { id: 'structure', rules: ['defaultSheetName'] },
  { id: 'access', rules: ['restrictedAccess'] },
];

/** One issue: the cells to select, plus what the rule's fixes need. */
export interface AccessibilityFinding {
  readonly rule: AccessibilityRuleId;
  readonly range: Range;
  /** Effective text and fill colours of a low-contrast cell. */
  readonly foreground?: string;
  readonly background?: string;
  /** The table a header finding is about. */
  readonly tableId?: string;
}

export type AccessibilityFindings = Readonly<
  Record<AccessibilityRuleId, readonly AccessibilityFinding[]>
>;

/** Minimum text contrast before a cell is flagged: WCAG AA for body text. */
export const MIN_TEXT_CONTRAST = 4.5;
/** The stronger WCAG AAA ratio, the second suggested colour. */
const ENHANCED_TEXT_CONTRAST = 7;

const DEFAULT_TEXT = '#000000';
const DEFAULT_FILL = '#ffffff';

const parseHex = (color: string): [number, number, number] | null => {
  const hex = color.trim().replace(/^#/, '');
  const full =
    hex.length === 3
      ? hex
          .split('')
          .map((c) => c + c)
          .join('')
      : hex;
  if (!/^[0-9a-f]{6}$/i.test(full)) return null;
  return [0, 2, 4].map((i) => Number.parseInt(full.slice(i, i + 2), 16)) as [
    number,
    number,
    number,
  ];
};

const toHex = (rgb: readonly number[]): string =>
  `#${rgb
    .map((v) =>
      Math.round(Math.min(255, Math.max(0, v)))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;

const luminance = (rgb: readonly number[]): number => {
  const [r, g, b] = rgb.map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4;
  }) as [number, number, number];
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
};

/** WCAG contrast ratio of two colours, 1 to 21; 21 when either is unreadable. */
export function contrastRatio(foreground: string, background: string): number {
  const fg = parseHex(foreground);
  const bg = parseHex(background);
  if (!fg || !bg) return 21;
  const [hi, lo] = [luminance(fg), luminance(bg)].sort((a, b) => b - a) as [number, number];
  return (hi + 0.05) / (lo + 0.05);
}

/** Two text colours of the same hue that reach the AA and AAA ratios on
 *  `background`: darker on a light fill, lighter on a dark one. */
export function suggestedTextColors(foreground: string, background: string): string[] {
  const fg = parseHex(foreground);
  const bg = parseHex(background);
  if (!fg || !bg) return [];
  const lighten = luminance(bg) < 0.18;
  const reach = (target: number): string => {
    for (let step = 0; step <= 100; step += 1) {
      const t = step / 100;
      const candidate = fg.map((v) => (lighten ? v + (255 - v) * t : v * (1 - t)));
      const hex = toHex(candidate);
      if (contrastRatio(hex, background) >= target) return hex;
    }
    return lighten ? '#ffffff' : '#000000';
  };
  return [reach(MIN_TEXT_CONTRAST), reach(ENHANCED_TEXT_CONTRAST)];
}

const cellRange = (sheet: number, row: number, col: number): Range => ({
  sheet,
  r0: row,
  c0: col,
  r1: row,
  c1: col,
});

/**
 * Run the accessibility rules over the workbook. Only text contrast and
 * header-less tables are flagged: those are the rules measured failing on the
 * desktop. The others report nothing until their failing cases are measured
 * (red text, a lone "Sheet1" and a merge outside a table all passed there).
 */
export function checkAccessibility(state: State, wb: WorkbookHandle): AccessibilityFindings {
  const contrast: AccessibilityFinding[] = [];
  const cache = new Map();
  for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
    for (const cell of wb.cells(sheet)) {
      if (cell.value.kind === 'blank') continue;
      if (cell.value.kind === 'text' && cell.value.value.trim() === '') continue;
      const format = effectiveCellFormat(wb, state, cell.addr, cache);
      const foreground = format?.color ?? DEFAULT_TEXT;
      const background = format?.fill ?? DEFAULT_FILL;
      if (contrastRatio(foreground, background) >= MIN_TEXT_CONTRAST) continue;
      contrast.push({
        rule: 'contrast',
        range: cellRange(sheet, cell.addr.row, cell.addr.col),
        foreground,
        background,
      });
    }
  }
  const tableHeader: AccessibilityFinding[] = state.tables.tables
    .filter((table) => !table.showHeader)
    .map((table) => ({ rule: 'tableHeader', range: { ...table.range }, tableId: table.id }));
  return {
    contrast,
    redFormat: [],
    altText: [],
    tableHeader,
    mergedCells: [],
    defaultSheetName: [],
    restrictedAccess: [],
  };
}

/** Whether any rule has a finding. */
export const hasAccessibilityIssues = (findings: AccessibilityFindings): boolean =>
  Object.values(findings).some((list) => list.length > 0);
