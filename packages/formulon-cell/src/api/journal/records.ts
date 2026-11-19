import type { Sheet as EngineSheet, ThreadedComment } from '@libraz/formulon';
import {
  applyDefinedNamesSnapshot,
  captureDefinedNamesSnapshot,
} from '../../commands/named-ranges.js';
import { captureSheetSparklineGroups } from '../../engine/sparkline-sync.js';
import type { JournalCodec, JournalFamily, JournalScope, Restore } from './index.js';

const sheetOf = (scope: JournalScope): EngineSheet => scope.engine.sheet(scope.sheet as number);

const sheetKey = (id: string) => (scope: JournalScope) => `${id}:${scope.sheet}`;

/** A sheet-scoped codec over one engine sheet. */
const sheetCodec = (
  id: string,
  capture: (sheet: EngineSheet, scope: JournalScope) => Restore,
): JournalCodec => ({
  id,
  scopeKey: sheetKey(id),
  capture: (scope) => capture(sheetOf(scope), scope),
});

export const CF: JournalCodec = sheetCodec('cf', (sheet) => {
  const rules = sheet.listCfRules();
  return () => {
    sheet.clearCfRules();
    for (const rule of rules) sheet.addCfRule(rule);
  };
});

export const VALIDATION: JournalCodec = sheetCodec('validation', (sheet) => {
  const rules = sheet.listValidations();
  return () => {
    sheet.clearValidations();
    for (const rule of rules) sheet.addValidation(rule);
  };
});

export const SHEET_PROT: JournalCodec = sheetCodec('sheet-prot', (sheet) => {
  const protection = sheet.getSheetProtection();
  const ranges = sheet.listProtectedRanges();
  return () => {
    sheet.setSheetProtection(protection);
    while (sheet.listProtectedRanges().length > 0) sheet.removeProtectedRange(0);
    for (const range of ranges) sheet.addProtectedRange(range);
  };
});

const WB_PROT: JournalCodec = {
  id: 'wb-prot',
  scopeKey: () => 'wb-prot',
  capture: ({ engine }) => {
    const protection = engine.getWorkbookProtection();
    return () => engine.setWorkbookProtection(protection);
  },
};

export const PRINT: JournalCodec = sheetCodec('print', (sheet) => {
  const headerFooter = sheet.ooxml.getHeaderFooterXml();
  const margins = sheet.ooxml.getPageMarginsXml();
  const setup = sheet.ooxml.getPageSetupXml();
  const options = sheet.ooxml.getPrintOptionsXml();
  const area = sheet.getPrintArea();
  const titles = sheet.getPrintTitles();
  const rowBreaks = sheet.listRowBreaks();
  const columnBreaks = sheet.listColumnBreaks();
  return () => {
    sheet.ooxml.setHeaderFooterXml(headerFooter);
    sheet.ooxml.setPageMarginsXml(margins);
    sheet.ooxml.setPageSetupXml(setup);
    sheet.ooxml.setPrintOptionsXml(options);
    sheet.setPrintArea(area);
    sheet.setPrintTitles(titles);
    sheet.clearBreaks();
    for (const brk of rowBreaks) sheet.setRowBreak(brk.row, brk);
    for (const brk of columnBreaks) sheet.setColumnBreak(brk.col, brk);
  };
});

export const SHEET_CHROME: JournalCodec = sheetCodec('sheet-chrome', (sheet) => {
  const color = sheet.getTabColor();
  const visibility = sheet.getVisibility();
  const sheetPr = sheet.ooxml.getSheetPrXml();
  return () => {
    sheet.setTabColor(color);
    sheet.setVisibility(visibility);
    sheet.ooxml.setSheetPrXml(sheetPr);
  };
});

export const NAMES: JournalCodec = {
  id: 'names',
  scopeKey: () => 'names',
  capture: ({ handle }) => {
    const snapshot = captureDefinedNamesSnapshot(handle);
    return () => applyDefinedNamesSnapshot(handle, snapshot);
  },
};

export const COMMENTS: JournalCodec = sheetCodec('comments', (sheet) => {
  const notes = sheet.listComments().map((note) => ({
    note,
    visible: sheet.getCommentVisible(note.cell.row, note.cell.col),
  }));
  const kept = new Set(notes.map(({ note }) => `${note.cell.row}:${note.cell.col}`));
  return () => {
    for (const { cell } of sheet.listComments()) {
      if (!kept.has(`${cell.row}:${cell.col}`)) sheet.setComment(cell.row, cell.col, null);
    }
    for (const { note, visible } of notes) {
      sheet.setComment(note.cell.row, note.cell.col, {
        text: note.text,
        author: note.author,
      });
      if (visible !== null) sheet.setCommentVisible(note.cell.row, note.cell.col, visible);
    }
  };
});

export const THREADS: JournalCodec = sheetCodec('threads', (sheet) => {
  const threads: readonly ThreadedComment[] = sheet.listThreadedComments();
  return () => {
    // Replies follow their opening comment, so removing back to front never orphans one.
    const current = sheet.listThreadedComments();
    for (let i = current.length - 1; i >= 0; i -= 1) {
      const comment = current[i] as ThreadedComment;
      sheet.removeThreadedComment(comment.id);
    }
    for (const comment of threads) sheet.addThreadedComment(comment);
  };
});

const PEOPLE: JournalCodec = {
  id: 'people',
  scopeKey: () => 'people',
  capture: ({ engine }) => {
    const people = engine.listPeople();
    return () => {
      const kept = new Set(people.map((person) => person.id));
      for (const person of engine.listPeople()) {
        if (!kept.has(person.id)) engine.removePerson(person.id);
      }
      const present = new Set(engine.listPeople().map((person) => person.id));
      for (const person of people) if (!present.has(person.id)) engine.addPerson(person);
    };
  },
};

export const SPARKLINES: JournalCodec = {
  id: 'sparklines',
  scopeKey: sheetKey('sparklines'),
  capture: ({ handle, sheet }) => captureSheetSparklineGroups(handle, sheet as number),
};

/** The rows an AutoFilter apply hides; the sheet filter XML does not carry them. */
const hiddenRows = (sheet: EngineSheet): Set<number> =>
  new Set(
    sheet
      .listRowOverrides()
      .filter((row) => row.hidden)
      .map((row) => row.row),
  );

export const AUTOFILTER: JournalCodec = sheetCodec('autofilter', (sheet) => {
  const xml = sheet.ooxml.getAutoFilterXml();
  const tables = sheet.listTables().map((table) => ({ table, filter: table.getAutoFilter() }));
  const hidden = hiddenRows(sheet);
  return () => {
    sheet.ooxml.setAutoFilterXml(xml);
    for (const { table, filter } of tables) table.setAutoFilter(filter);
    const now = hiddenRows(sheet);
    for (const row of now) if (!hidden.has(row)) sheet.setRowHidden(row, row, false);
    for (const row of hidden) if (!now.has(row)) sheet.setRowHidden(row, row, true);
  };
});

export const RECORDS_JOURNAL: JournalFamily = {
  members: {
    sheet: { setTabColor: SHEET_CHROME, setVisibility: SHEET_CHROME },
    'sheet.ooxml': { setSheetPrXml: SHEET_CHROME },
  },
  kinds: {
    'cf/sheet': CF,
    'validation/sheet': VALIDATION,
    'protection/sheet': SHEET_PROT,
    'protection/workbook': WB_PROT,
    'print/sheet': PRINT,
    'names/workbook': NAMES,
    'comments/sheet': COMMENTS,
    'threadedComments/sheet': THREADS,
    'threadedComments/workbook': PEOPLE,
    'sparklineGroups/sheet': SPARKLINES,
    'autoFilter/sheet': AUTOFILTER,
  },
};
