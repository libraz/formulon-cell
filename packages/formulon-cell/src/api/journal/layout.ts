import type { Sheet } from '@libraz/formulon';
import type { JournalChoice, JournalCodec, JournalFamily, JournalScope } from './index.js';

/** Merges plus row and column overrides of one sheet. */
const rowsColsCodec: JournalCodec = {
  id: 'rows-cols',
  scopeKey: (scope) => `rows-cols:${scope.sheet}`,
  capture: (scope: JournalScope) => {
    const { engine } = scope;
    const index = scope.sheet as number;
    const before = engine.sheet(index);
    const merges = before.listMerges();
    const rows = before.listRowOverrides();
    const cols = before.listColumnOverrides();
    return () => {
      const sheet: Sheet = engine.sheet(index);
      // Reset everything the sheet carries now, then lay the snapshot back down.
      for (const r of sheet.listRowOverrides()) {
        sheet.resetRowHeight(r.row, r.row);
        sheet.setRowHidden(r.row, r.row, false);
        sheet.setRowOutline(r.row, r.row, 0);
        sheet.setRowCollapsed(r.row, r.row, false);
      }
      for (const c of sheet.listColumnOverrides()) {
        sheet.resetColumnWidth(c.first, c.last);
        sheet.setColumnHidden(c.first, c.last, false);
        sheet.setColumnOutline(c.first, c.last, 0);
        sheet.setColumnCollapsed(c.first, c.last, false);
      }
      for (const r of rows) {
        if (r.heightPt !== null) {
          if (r.customHeight) sheet.setRowHeight(r.row, r.row, r.heightPt);
          else sheet.setRowAutoHeight(r.row, r.row, r.heightPt);
        }
        if (r.hidden) sheet.setRowHidden(r.row, r.row, true);
        if (r.outlineLevel > 0) sheet.setRowOutline(r.row, r.row, r.outlineLevel);
        if (r.collapsed) sheet.setRowCollapsed(r.row, r.row, true);
      }
      for (const c of cols) {
        if (c.widthChars !== null) sheet.setColumnWidth(c.first, c.last, c.widthChars);
        if (c.hidden) sheet.setColumnHidden(c.first, c.last, true);
        if (c.outlineLevel > 0) sheet.setColumnOutline(c.first, c.last, c.outlineLevel);
        if (c.collapsed) sheet.setColumnCollapsed(c.first, c.last, true);
      }
      sheet.clearMerges();
      for (const m of merges) sheet.addMerge(m);
    };
  },
};

const ROWS_COLS_MEMBERS = [
  'addMerge',
  'removeMerge',
  'clearMerges',
  'setRowHeight',
  'setRowAutoHeight',
  'resetRowHeight',
  'setRowHidden',
  'setRowOutline',
  'setRowCollapsed',
  'setColumnWidth',
  'resetColumnWidth',
  'setColumnHidden',
  'setColumnOutline',
  'setColumnCollapsed',
];

export const LAYOUT_JOURNAL: JournalFamily = {
  members: {
    sheet: Object.fromEntries(
      ROWS_COLS_MEMBERS.map((name): [string, JournalChoice] => [name, rowsColsCodec]),
    ),
  },
};
