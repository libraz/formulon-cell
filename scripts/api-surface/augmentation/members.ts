// Typechecked against the built declarations by check.mjs: one member from every module that
// augments Workbook or Sheet, reached through the public types only.
import type { Sheet, Workbook } from '@libraz/formulon-cell';

export const workbookMembers = (wb: Workbook) => [
  wb.setCalcMode, // workbook-calc
  wb.listPeople, // workbook-people
  wb.setDefinedName, // sheet-structure (names)
  wb.listPivotCaches, // pivot-cache
];

export const sheetMembers = (sheet: Sheet) => [
  sheet.insertRows, // sheet-structure
  sheet.listCharts, // chart
  sheet.listDrawings, // drawing
  sheet.listPivots, // pivot
  sheet.importCsv, // sheet-cells
  sheet.findAll, // sheet-find
  sheet.protect, // sheet-format
  sheet.listTables, // table
];
