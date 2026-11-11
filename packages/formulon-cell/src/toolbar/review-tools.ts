import { formatA1Cell, parseAddrKey } from '../engine/address.js';
import type { Locale } from '../i18n/strings.js';
import { dictionaries, type Strings } from '../i18n/strings.js';
import type { State } from '../store/types.js';

export interface RibbonReportItem {
  severity: 'info' | 'warning';
  label: string;
  detail: string;
}

export type ReviewCellValue =
  | { kind: 'blank' }
  | { kind: 'text'; value: string }
  | { kind: 'number' }
  | { kind: 'bool' }
  | { kind: 'error'; text: string };

export interface ReviewCell {
  label: string;
  value: ReviewCellValue;
  formula?: string | null;
  source?: 'cell' | 'comment';
}

export type ScriptCommand = 'uppercase' | 'lowercase' | 'trim' | 'clear';
export type RibbonReportLang = Locale;
type ReviewReportText = Strings['reviewReports'];

interface ReviewRange {
  sheet: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export function parseScriptCommand(value: string): ScriptCommand | null {
  const normalized = value.trim().toLowerCase();
  return normalized === 'uppercase' ||
    normalized === 'lowercase' ||
    normalized === 'trim' ||
    normalized === 'clear'
    ? normalized
    : null;
}

export function applyTextScript(value: string, command: Exclude<ScriptCommand, 'clear'>): string {
  switch (command) {
    case 'uppercase':
      return value.toUpperCase();
    case 'lowercase':
      return value.toLowerCase();
    case 'trim':
      return value.trim().replace(/\s+/g, ' ');
  }
}

/** The cells and notes a review reads, by rows; a note follows its cell. */
export function reviewCellsFromState(
  state: State,
  sheet = state.data.sheetIndex,
  range?: ReviewRange,
): ReviewCell[] {
  const entries: { row: number; col: number; cell: ReviewCell }[] = [];
  const inScope = (addr: { sheet: number; row: number; col: number }): boolean =>
    addr.sheet === sheet &&
    (!range ||
      (addr.sheet === range.sheet &&
        addr.row >= range.r0 &&
        addr.row <= range.r1 &&
        addr.col >= range.c0 &&
        addr.col <= range.c1));
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr || !inScope(addr)) continue;
    const label = formatA1Cell(addr.row, addr.col);
    const value: ReviewCellValue =
      cell.value.kind === 'text'
        ? { kind: 'text', value: cell.value.value }
        : cell.value.kind === 'error'
          ? { kind: 'error', text: cell.value.text }
          : cell.value.kind === 'number'
            ? { kind: 'number' }
            : cell.value.kind === 'bool'
              ? { kind: 'bool' }
              : { kind: 'blank' };
    entries.push({ ...addr, cell: { label, value, formula: cell.formula, source: 'cell' } });
  }
  for (const [key, fmt] of state.format.formats) {
    if (typeof fmt.comment !== 'string' || fmt.comment.trim().length === 0) continue;
    const addr = parseAddrKey(key);
    if (!addr || !inScope(addr)) continue;
    const label = `${formatA1Cell(addr.row, addr.col)} comment`;
    entries.push({
      ...addr,
      cell: { label, value: { kind: 'text', value: fmt.comment }, source: 'comment' },
    });
  }
  return entries
    .sort(
      (a, b) =>
        a.row - b.row ||
        a.col - b.col ||
        (a.cell.source === 'comment' ? 1 : 0) - (b.cell.source === 'comment' ? 1 : 0),
    )
    .map((entry) => entry.cell);
}

const reportText = (lang: RibbonReportLang): ReviewReportText => dictionaries[lang].reviewReports;

const reviewText = reportText;

const interpolate = (template: string, vars: Record<string, string | number>): string =>
  template.replace(/\{(\w+)\}/g, (_, key) => String(vars[key] ?? ''));

export function formatRibbonReport(
  title: string,
  items: readonly RibbonReportItem[],
  lang: RibbonReportLang = 'en',
): string {
  const text = reportText(lang);
  if (items.length === 0) return `${title}\n${text.noIssues}`;
  const lines = items
    .slice(0, 20)
    .map(
      (item) =>
        `${item.severity === 'warning' ? text.warning : text.info} - ${item.label}: ${item.detail}`,
    );
  if (items.length > lines.length)
    lines.push(
      `${text.info} - ${text.moreLabel}: ${interpolate(text.more, {
        count: items.length - lines.length,
      })}`,
    );
  return `${title}\n${lines.join('\n')}`;
}

export function buildTranslationReviewItems(
  cells: readonly ReviewCell[],
  lang: RibbonReportLang = 'en',
): RibbonReportItem[] {
  const text = reportText(lang);
  const items = cells
    .filter((entry) => entry.value.kind === 'text' && entry.value.value.trim().length > 0)
    .map<RibbonReportItem>((entry) => {
      const value =
        entry.value.kind === 'text' ? entry.value.value.trim().replace(/\s+/g, ' ') : '';
      const snippet = value.length > 80 ? `${value.slice(0, 77)}...` : value;
      return {
        severity: 'info',
        label: entry.label,
        detail: `${text.translateReady}: "${snippet}"`,
      };
    });
  if (items.length > 0) return items;
  return [{ severity: 'info', label: text.translateEmptyLabel, detail: text.translateEmptyDetail }];
}

export function analyzeAccessibilityCells(
  cells: readonly ReviewCell[],
  lang: RibbonReportLang = 'en',
): RibbonReportItem[] {
  const text = reviewText(lang);
  const items: RibbonReportItem[] = [];
  const nonBlank = cells.filter((entry) => entry.value.kind !== 'blank' || !!entry.formula);
  const formulas = nonBlank.filter((entry) => !!entry.formula).length;
  const textCells = nonBlank.filter((entry) => entry.value.kind === 'text').length;

  if (nonBlank.length === 0) {
    items.push({
      severity: 'info',
      label: text.emptySheetLabel,
      detail: text.emptySheetDetail,
    });
  }

  for (const entry of nonBlank) {
    if (entry.value.kind === 'error') {
      items.push({
        severity: 'warning',
        label: entry.label,
        detail: interpolate(text.errorDetail, { text: entry.value.text }),
      });
    }
    if (entry.formula?.includes('#REF!')) {
      items.push({
        severity: 'warning',
        label: entry.label,
        detail: text.refDetail,
      });
    }
    if (entry.formula && /!(?:[A-Z]+\d+|[A-Z]+:[A-Z]+|\d+:\d+)/.test(entry.formula)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: text.externalSheetDetail,
      });
    }
    if (entry.value.kind !== 'text') continue;
    const cellText = entry.value.value.trim();
    if (cellText.length === 0 && entry.value.value.length > 0) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: text.whitespaceDetail,
      });
    }
    if (/^https?:\/\//i.test(cellText)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: text.urlDetail,
      });
    }
    if (cellText.length > 120 && !/[.!?。！？]\s/.test(cellText)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: text.longTextDetail,
      });
    }
    if (/^[A-Z0-9 _-]{24,}$/.test(cellText)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: text.allCapsDetail,
      });
    }
  }

  if (formulas === 0 && textCells > 0) {
    items.push({
      severity: 'info',
      label: text.workbookStructureLabel,
      detail: text.workbookStructureDetail,
    });
  }

  return items;
}
