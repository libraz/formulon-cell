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
export type RibbonReportLang = 'en' | 'ja';
type ReviewReportText = Strings['reviewReports'];

interface ReviewRange {
  sheet: number;
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

const COMMON_TYPOS: Readonly<Record<string, string>> = {
  accomodate: 'accommodate',
  adress: 'address',
  definately: 'definitely',
  occured: 'occurred',
  recieve: 'receive',
  seperate: 'separate',
  teh: 'the',
  wierd: 'weird',
};

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

const colLabel = (col: number): string => {
  let n = col;
  let out = '';
  do {
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return out;
};

const parseAddrKey = (key: string): { sheet: number; row: number; col: number } | null => {
  const parts = key.split(':').map((part) => Number.parseInt(part, 10));
  if (parts.length !== 3) return null;
  const [sheet, row, col] = parts as [number, number, number];
  if (!Number.isInteger(sheet) || !Number.isInteger(row) || !Number.isInteger(col)) return null;
  return { sheet, row, col };
};

export function reviewCellsFromState(
  state: State,
  sheet = state.data.sheetIndex,
  range?: ReviewRange,
): ReviewCell[] {
  const cells: ReviewCell[] = [];
  for (const [key, cell] of state.data.cells) {
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    if (
      range &&
      (addr.sheet !== range.sheet ||
        addr.row < range.r0 ||
        addr.row > range.r1 ||
        addr.col < range.c0 ||
        addr.col > range.c1)
    )
      continue;
    const label = `${colLabel(addr.col)}${addr.row + 1}`;
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
    cells.push({ label, value, formula: cell.formula, source: 'cell' });
  }
  for (const [key, fmt] of state.format.formats) {
    if (typeof fmt.comment !== 'string' || fmt.comment.trim().length === 0) continue;
    const addr = parseAddrKey(key);
    if (!addr || addr.sheet !== sheet) continue;
    if (
      range &&
      (addr.sheet !== range.sheet ||
        addr.row < range.r0 ||
        addr.row > range.r1 ||
        addr.col < range.c0 ||
        addr.col > range.c1)
    )
      continue;
    const label = `${colLabel(addr.col)}${addr.row + 1} comment`;
    cells.push({ label, value: { kind: 'text', value: fmt.comment }, source: 'comment' });
  }
  return cells.sort((a, b) => a.label.localeCompare(b.label, 'en', { numeric: true }));
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

export function analyzeSpellingCells(
  cells: readonly ReviewCell[],
  lang: RibbonReportLang = 'en',
): RibbonReportItem[] {
  const messages = reviewText(lang);
  const items: RibbonReportItem[] = [];
  for (const entry of cells) {
    if (entry.value.kind !== 'text') continue;
    const cellText = entry.value.value;
    const repeated = cellText.match(/\b([A-Za-z]+)\s+\1\b/i);
    if (repeated) {
      items.push({
        severity: 'warning',
        label: entry.label,
        detail: interpolate(messages.repeatedWord, { word: repeated[0] }),
      });
    }
    if (/\s{2,}/.test(cellText)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: messages.repeatedSpaces,
      });
    }
    if (/[A-Za-z]\s+[,.!?;:]/.test(cellText)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: messages.spaceBeforePunctuation,
      });
    }
    if (/[.!?]\s+[a-z]/.test(cellText)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: messages.lowercaseSentence,
      });
    }
    for (const [wrong, suggestion] of Object.entries(COMMON_TYPOS)) {
      if (new RegExp(`\\b${wrong}\\b`, 'i').test(cellText)) {
        items.push({
          severity: 'warning',
          label: entry.label,
          detail: interpolate(messages.typo, { wrong, suggestion }),
        });
      }
    }
  }
  return items;
}
