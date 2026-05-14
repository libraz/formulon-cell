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
}

export type ScriptCommand = 'uppercase' | 'lowercase' | 'trim' | 'clear';

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

export function analyzeAccessibilityCells(cells: readonly ReviewCell[]): RibbonReportItem[] {
  const items: RibbonReportItem[] = [];
  const nonBlank = cells.filter((entry) => entry.value.kind !== 'blank' || !!entry.formula);
  const formulas = nonBlank.filter((entry) => !!entry.formula).length;
  const textCells = nonBlank.filter((entry) => entry.value.kind === 'text').length;

  if (nonBlank.length === 0) {
    items.push({
      severity: 'info',
      label: 'Empty sheet',
      detail: 'The current sheet has no populated cells to review.',
    });
  }

  for (const entry of nonBlank) {
    if (entry.value.kind === 'error') {
      items.push({
        severity: 'warning',
        label: entry.label,
        detail: `Cell evaluates to ${entry.value.text}. Resolve formula errors before sharing.`,
      });
    }
    if (entry.formula?.includes('#REF!')) {
      items.push({
        severity: 'warning',
        label: entry.label,
        detail: 'Formula contains #REF!, which is usually a broken reference.',
      });
    }
    if (entry.formula && /!(?:[A-Z]+\d+|[A-Z]+:[A-Z]+|\d+:\d+)/.test(entry.formula)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'Formula references another sheet. Confirm the dependency is intentional.',
      });
    }
    if (entry.value.kind !== 'text') continue;
    const text = entry.value.value.trim();
    if (text.length === 0 && entry.value.value.length > 0) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'Cell contains only whitespace.',
      });
    }
    if (/^https?:\/\//i.test(text)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'URL text should have a descriptive label when used as a link.',
      });
    }
    if (text.length > 120 && !/[.!?。！？]\s/.test(text)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'Long text may be hard to scan. Consider wrapping or shortening it.',
      });
    }
    if (/^[A-Z0-9 _-]{24,}$/.test(text)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'All-caps text can be hard to read with assistive technology.',
      });
    }
  }

  if (formulas === 0 && textCells > 0) {
    items.push({
      severity: 'info',
      label: 'Workbook structure',
      detail:
        'No formulas were found on this sheet. Confirm calculated values were not pasted as text.',
    });
  }

  return items;
}

export function analyzeSpellingCells(cells: readonly ReviewCell[]): RibbonReportItem[] {
  const items: RibbonReportItem[] = [];
  for (const entry of cells) {
    if (entry.value.kind !== 'text') continue;
    const text = entry.value.value;
    const repeated = text.match(/\b([A-Za-z]+)\s+\1\b/i);
    if (repeated) {
      items.push({
        severity: 'warning',
        label: entry.label,
        detail: `Repeated word: "${repeated[0]}".`,
      });
    }
    if (/\s{2,}/.test(text)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'Contains repeated spaces.',
      });
    }
    if (/[A-Za-z]\s+[,.!?;:]/.test(text)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'Contains a space before punctuation.',
      });
    }
    if (/[.!?]\s+[a-z]/.test(text)) {
      items.push({
        severity: 'info',
        label: entry.label,
        detail: 'A sentence appears to start with a lowercase letter.',
      });
    }
    for (const [wrong, suggestion] of Object.entries(COMMON_TYPOS)) {
      if (new RegExp(`\\b${wrong}\\b`, 'i').test(text)) {
        items.push({
          severity: 'warning',
          label: entry.label,
          detail: `Possible typo: "${wrong}". Suggested spelling: "${suggestion}".`,
        });
      }
    }
  }
  return items;
}
